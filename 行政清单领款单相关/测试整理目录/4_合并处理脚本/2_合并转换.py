import os
from datetime import datetime
from docx2pdf import convert
from PyPDF2 import PdfMerger, PdfReader

# -----------------------------
# 判断PDF是否为“空白页”
# -----------------------------
def is_blank_page(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            text = page.extract_text()
            if text and text.strip():
                return False
        return True
    except Exception:
        return False

# -----------------------------
# 自动生成合并PDF文件名
# -----------------------------
def get_merged_filename(base_name):
    now = datetime.now()
    year = now.year % 100  # 取后两位数字，如2025 -> 25
    month = now.month
    day = now.day

    if day <= 15:
        period = "第2期"
    else:
        period = "第1期"
        month += 1
        if month == 13:
            month = 1
            year += 1

    return f"{year}年{month}月{period}{base_name}.pdf"

# -----------------------------
# 设置路径
# -----------------------------
project_root = os.path.dirname(os.path.abspath(__file__))
output_dir = os.path.join(project_root, "output")
pdf_base_dir = os.path.join(project_root, "Word转PDF")
pdf_claim_dir = os.path.join(pdf_base_dir, "领款单")
pdf_list_dir = os.path.join(pdf_base_dir, "行政清单")

# 确保输出文件夹存在
os.makedirs(pdf_claim_dir, exist_ok=True)
os.makedirs(pdf_list_dir, exist_ok=True)

# -----------------------------
# Word -> PDF 转换
# -----------------------------
print("📄 正在转换 Word 到 PDF ...")
for filename in os.listdir(output_dir):
    if filename.endswith(".docx"):
        source_path = os.path.join(output_dir, filename)
        if "-领款单" in filename:
            output_path = os.path.join(pdf_claim_dir, filename.replace(".docx", ".pdf"))
            convert(source_path, output_path)
            print(f"✅ 已转换: {filename} -> 领款单 PDF")
        elif "@行政清单" in filename:
            output_path = os.path.join(pdf_list_dir, filename.replace(".docx", ".pdf"))
            convert(source_path, output_path)
            print(f"✅ 已转换: {filename} -> 行政清单 PDF")

# -----------------------------
# 合并 PDF（领款单）
# -----------------------------
print("📎 正在合并 领款单 PDF ...")
merger_claim = PdfMerger()
for filename in sorted(os.listdir(pdf_claim_dir)):
    if filename.endswith(".pdf"):
        path = os.path.join(pdf_claim_dir, filename)
        if not is_blank_page(path):
            merger_claim.append(path)
merged_claim_name = get_merged_filename("领款单")
merged_claim_path = os.path.join(pdf_base_dir, merged_claim_name)
merger_claim.write(merged_claim_path)
merger_claim.close()
print(f"🎉 合并完成：{merged_claim_name}")

# -----------------------------
# 合并 PDF（行政清单）
# -----------------------------
print("📎 正在合并 行政清单 PDF ...")
merger_list = PdfMerger()
for filename in sorted(os.listdir(pdf_list_dir)):
    if filename.endswith(".pdf"):
        path = os.path.join(pdf_list_dir, filename)
        if not is_blank_page(path):
            merger_list.append(path)
merged_list_name = get_merged_filename("行政清单")
merged_list_path = os.path.join(pdf_base_dir, merged_list_name)
merger_list.write(merged_list_path)
merger_list.close()
print(f"🎉 合并完成：{merged_list_name}")
