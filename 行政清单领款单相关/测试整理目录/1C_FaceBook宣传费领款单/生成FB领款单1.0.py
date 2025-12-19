import fitz  # PyMuPDF
import os
import datetime
from docxtpl import DocxTemplate

# 📁 路径配置
pdf_path = os.path.join("FaceBook宣传费文件", "invoices.pdf")
template_path = "Knight Creative Limited模板.docx"
output_path = os.path.join("output", "FaceBook宣传费领款单.docx")

# ============================================================
# 🧠 解析PDF，提取“山景服務處”页中目标金额与编号
# ============================================================
def extract_invoice_info(pdf_path):
    doc = fitz.open(pdf_path)
    target_page = None
    invoice_number = None
    hkd_amount = None

    for i, page in enumerate(doc):
        text = page.get_text()
        if "山景服務處" in text:
            target_page = page
            break

    if not target_page:
        raise ValueError("❌ 未在PDF中找到包含 '山景服務處' 的页面")

    # 提取所有行文本
    lines = []
    blocks = target_page.get_text("dict")["blocks"]
    for b in blocks:
        for line in b.get("lines", []):
            line_text = " ".join([span["text"].strip() for span in line["spans"]])
            lines.append(line_text.strip())

    # 定位第二个 “Balance Due” + 下一行为金额
    balance_indices = [i for i, l in enumerate(lines) if l == "Balance Due"]
    for idx in balance_indices:
        if idx + 1 < len(lines):
            next_line = lines[idx + 1].strip()
            if next_line.startswith("HKD"):
                hkd_amount = next_line
                break

    if not hkd_amount:
        raise ValueError("❌ 未在 Balance Due 下方找到金额")

    # 提取发票编号（# INV-xxxxxx）
    for line in lines:
        if line.startswith("# INV-"):
            invoice_number = line
            break

    if not invoice_number:
        raise ValueError("❌ 未找到发票编号（# INV-xxxxx）")

    return hkd_amount.replace("HKD", "").replace(",", "").strip(), invoice_number.replace("# ", "").strip()


# ============================================================
# 💰 金额转中文大写
# ============================================================
def convert_to_chinese_currency(num):
    digits = "零壹貳叁肆伍陸柒捌玖"
    units = ["", "拾", "佰", "仟"]
    big_units = ["", "萬", "億", "兆"]
    decimal_units = ["角", "分"]

    num_str = f"{float(num):.2f}"
    integer_part, decimal_part = num_str.split(".")
    integer_part = integer_part.lstrip("0") or "0"

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


# ============================================================
# 🗓️ 日期字段生成
# ============================================================
def generate_date_fields():
    today = datetime.datetime.today()
    day = today.day
    month = today.month
    year = today.year

    if day <= 15:
        领款日期 = f"15/{month}/{year}"
        month_used = month
        期 = "2"
    else:
        if month == 12:
            month_used = 1
            year += 1
        else:
            month_used = month + 1
        领款日期 = f"1/{month_used}/{year}"
        期 = "1"

    m1 = str(month_used // 10)
    m2 = str(month_used % 10)
    return 领款日期, m1, m2, 期


# ============================================================
# 🚀 主函数
# ============================================================
def main():
    os.makedirs("output", exist_ok=True)

    try:
        amount_str, project_id = extract_invoice_info(pdf_path)
        amount_float = float(amount_str)
        cn_amount = convert_to_chinese_currency(amount_float)
        领款日期, m1, m2, 期 = generate_date_fields()

        # 构造模板上下文
        context = {
            "项目金额": f"${amount_float:,.2f}",
            "项目编号": project_id,
            "港币圆数大写": cn_amount,
            "领款日期": 领款日期,
            "m1": m1,
            "m2": m2,
            "期": 期,
        }

        # 渲染模板
        doc = DocxTemplate(template_path)
        doc.render(context)
        doc.save(output_path)

        print(f"✅ 已成功生成：{output_path}")

    except Exception as e:
        print(f"❌ 出现错误：{e}")


if __name__ == "__main__":
    main()
