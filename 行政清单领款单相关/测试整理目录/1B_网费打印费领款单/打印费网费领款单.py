# -*- coding: utf-8 -*-
import os
import re
from datetime import datetime
from PyPDF2 import PdfReader
from docxtpl import DocxTemplate

# ============================
# 工具函数
# ============================
def extract_invoice_number(text):
    match = re.search(r"Invoice Number\s+(\d+)", text)
    return match.group(1) if match else None

def extract_total_amount(text):
    match = re.search(r"Total Amount\s+([0-9]+\.[0-9]{2})", text)
    return match.group(1) if match else None

def extract_clicks_charge_period(text):
    match = re.search(r"Clicks Charge Period\s+([0-9]{2} \w{3} [0-9]{4})\s*-\s*([0-9]{2} \w{3} [0-9]{4})", text)
    if match:
        end_date_str = match.group(2)
        end_date = datetime.strptime(end_date_str, "%d %b %Y")
        return end_date.month, end_date.year, end_date_str
    return None, None, None

def extract_invoice_date(text):
    match = re.search(r"INVOICE DATE\s*:\s*(\d{2})/(\d{2})/(\d{4})", text)
    if match:
        month, year = int(match.group(2)), int(match.group(3))
        return month, year
    return None, None

def extract_invoice_no(text):
    match = re.search(r"INVOICE NO\.\s*:\s*(\d+)", text)
    return match.group(1) if match else None

def convert_to_chinese_currency(num):
    digits = "零壹貳叁肆伍陸柒捌玖"
    units = ["", "拾", "佰", "仟"]
    big_units = ["", "萬", "億", "兆"]
    decimal_units = ["角", "分"]
    if num < 0:
        return "负" + convert_to_chinese_currency(-num)
    num_str = f"{num:.2f}"
    integer_part, decimal_part = num_str.split('.')
    integer_part = integer_part.lstrip('0') or '0'
    result = ""
    integer_part = integer_part[::-1]
    for i in range(0, len(integer_part), 4):
        group = integer_part[i:i+4]
        group_str = ""
        zero_flag = False
        for j in range(len(group)):
            n = int(group[j])
            if n == 0:
                if not zero_flag and group_str:
                    group_str = digits[0] + group_str
                zero_flag = True
            else:
                group_str = digits[n] + units[j] + group_str
                zero_flag = False
        group_str = group_str.rstrip(digits[0])
        if group_str:
            result = group_str + big_units[i // 4] + result
    result = result or digits[0]
    result += "元"
    if decimal_part == "00":
        result += "正"
    else:
        jiao = int(decimal_part[0])
        fen = int(decimal_part[1])
        if jiao != 0:
            result += digits[jiao] + decimal_units[0]
        if fen != 0:
            result += digits[fen] + decimal_units[1]
    return result

def generate_common_fields():
    today = datetime.today()
    day, month, year = today.day, today.month, today.year
    if day <= 15:
        m_date = f"15/{month}/{year}"
        m_month = month
        m期 = "2"
    else:
        if month == 12:
            m_month = 1
            year += 1
        else:
            m_month = month + 1
        m_date = f"1/{m_month}/{year}"
        m期 = "1"
    return {
        "领款日期": m_date,
        "m1": str(m_month // 10),
        "m2": str(m_month % 10),
        "期": m期
    }

# ============================
# 主逻辑
# ============================
print("请选择生成的文件类型：")
print("1. 打印费领款单")
print("2. 网费领款单")
choice = input("请输入 1 或 2：").strip()

if choice == "1":
    pdf_path = os.path.join("打印费文件", os.listdir("打印费文件")[0])
    template_path = "HP Inc Hong Kong Limited.docx"
    reader = PdfReader(pdf_path)
    text = "".join(page.extract_text() for page in reader.pages)

    invoice_no = extract_invoice_number(text)
    amount = extract_total_amount(text)
    month, year, end_date_str = extract_clicks_charge_period(text)

    project_name = f"影印費{month}/{year}(InvoiceNo.{invoice_no})"
    amount_str = f"${amount}"
    amount_upper = convert_to_chinese_currency(float(amount))
    common = generate_common_fields()

    context = {
        **common,
        "项目名字编号": project_name,
        "项目金额": amount_str,
        "港币圆数大写": amount_upper
    }
    doc = DocxTemplate(template_path)
    doc.render(context)
    outname = f"output/{year}年{month}月打印费领款单.docx"
    doc.save(outname)
    print("✅ 打印费领款单已生成：", outname)

elif choice == "2":
    pdf_path = os.path.join("网费文件", os.listdir("网费文件")[0])
    template_path = "Information Technology Resource Centre.docx"
    reader = PdfReader(pdf_path)
    text = "".join(page.extract_text() for page in reader.pages)

    invoice_no = extract_invoice_no(text)
    month, year = extract_invoice_date(text)
    project_name = f"山景中心上網費({month}/{year})(NO.{invoice_no})"
    common = generate_common_fields()

    context = {
        **common,
        "项目名字编号": project_name
    }
    doc = DocxTemplate(template_path)
    doc.render(context)
    outname = f"output/{year}年{month}月网费领款单.docx"
    doc.save(outname)
    print("✅ 网费领款单已生成：", outname)
else:
    print("❌ 输入错误，请输入 1 或 2。")