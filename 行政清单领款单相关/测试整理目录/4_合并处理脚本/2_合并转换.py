# -*- coding: utf-8 -*-
import os
import sys
from datetime import datetime
from docx2pdf import convert
from PyPDF2 import PdfMerger, PdfReader

# ============================================================
# ⚙️ 路径配置
# ============================================================
# 当前脚本所在目录 (4_合并处理脚本)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# 项目根目录 (测试整理目录)
ROOT_DIR = os.path.dirname(BASE_DIR)

# 1. 🔍 文件检索路径 (输入)
SEARCH_DIR_1A = os.path.join(ROOT_DIR, "1A_课程行政清单_领款单", "output")
SEARCH_DIR_1B = os.path.join(ROOT_DIR, "1B_杂费领款单", "output")

# 2. 📂 PDF 输出路径 (分类存放)
PDF_OUT_ROOT = os.path.join(ROOT_DIR, "5_Word转PDF") # 总目录
PDF_OUT_ADMIN = os.path.join(PDF_OUT_ROOT, "行政清单")
PDF_OUT_RECEIPT = os.path.join(PDF_OUT_ROOT, "领款单")

# 确保输出目录存在
os.makedirs(PDF_OUT_ADMIN, exist_ok=True)
os.makedirs(PDF_OUT_RECEIPT, exist_ok=True)

# ============================================================
# 🛠️ 辅助函数
# ============================================================

def is_blank_page(pdf_path):
    """判断 PDF 是否为空白页 (无文字)"""
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            text = page.extract_text()
            if text and text.strip():
                return False # 有文字，不是空白
        return True # 没文字，视为空白
    except Exception:
        return False # 读取出错，保守处理保留文件

def get_merged_filename(base_name):
    """
    根据运行日期生成文件名
    规则：
    <= 15号: 本月 + 第2期
    > 15号: 下月 + 第1期
    """
    now = datetime.now()
    year = now.year % 100  # 取后两位 (2025 -> 25)
    month = now.month
    day = now.day

    if day <= 15:
        period = "第2期"
        # 月份保持本月
    else:
        period = "第1期"
        # 月份+1，处理跨年
        month += 1
        if month == 13:
            month = 1
            year += 1

    return f"{year}年{month}月{period}{base_name}.pdf"

# ============================================================
# 🔄 核心处理逻辑
# ============================================================

def convert_and_sort_files():
    """扫描 Word 文件并转换为 PDF 到对应文件夹"""
    print("🚀 开始 Word 转 PDF...\n")
    
    # 定义要扫描的文件夹列表
    search_dirs = [SEARCH_DIR_1A, SEARCH_DIR_1B]
    
    found_files = False

    for source_dir in search_dirs:
        if not os.path.exists(source_dir):
            print(f"⚠️ 跳过不存在的文件夹: {source_dir}")
            continue

        print(f"📂 正在扫描: {source_dir}")
        for filename in os.listdir(source_dir):
            if filename.startswith("~") or not filename.endswith(".docx"):
                continue # 跳过临时文件和非Word文件
            
            source_file = os.path.join(source_dir, filename)
            target_folder = None

            # 🏷️ 根据后缀分类
            if filename.endswith("行政清单.docx"):
                target_folder = PDF_OUT_ADMIN
            elif filename.endswith("领款单.docx"):
                target_folder = PDF_OUT_RECEIPT
            
            if target_folder:
                found_files = True
                pdf_filename = filename.replace(".docx", ".pdf")
                target_path = os.path.join(target_folder, pdf_filename)
                
                # 执行转换
                try:
                    # print(f"   正在转换: {filename} ...")
                    convert(source_file, target_path)
                    print(f"   ✅ 已转换 -> {os.path.basename(target_folder)}/{pdf_filename}")
                except Exception as e:
                    print(f"   ❌ 转换失败 {filename}: {e}")

    if not found_files:
        print("\n⚠️ 未找到任何需要转换的 .docx 文件。")

def merge_pdfs_in_folder(source_folder, type_name):
    """合并指定文件夹下的 PDF"""
    if not os.path.exists(source_folder):
        return

    # 获取并排序文件
    files = sorted([f for f in os.listdir(source_folder) if f.endswith(".pdf")])
    if not files:
        print(f"\nℹ️ {type_name} 文件夹为空，无需合并。")
        return

    print(f"\n📎 正在合并 {type_name} ({len(files)} 个文件) ...")
    
    merger = PdfMerger()
    count = 0
    
    for filename in files:
        path = os.path.join(source_folder, filename)
        # 过滤空白页
        if not is_blank_page(path):
            merger.append(path)
            count += 1
        else:
            print(f"   ⚠️ 跳过空白页: {filename}")

    if count > 0:
        # 生成输出路径 (保存在根目录 PDF_OUT_ROOT)
        output_filename = get_merged_filename(type_name)
        output_path = os.path.join(PDF_OUT_ROOT, output_filename)
        
        merger.write(output_path)
        merger.close()
        print(f"🎉 合并完成！文件位置: {output_path}")
    else:
        print(f"⚠️ 没有有效内容可合并。")

# ============================================================
# 🚀 主程序入口
# ============================================================
def main():
    # 1. 转换阶段
    convert_and_sort_files()
    
    # 2. 合并阶段
    print("\n" + "="*30)
    merge_pdfs_in_folder(PDF_OUT_RECEIPT, "领款单")   # 合并领款单
    merge_pdfs_in_folder(PDF_OUT_ADMIN, "行政清单")   # 合并行政清单
    
    print("\n✅ 所有任务已完成。")

if __name__ == "__main__":
    main()