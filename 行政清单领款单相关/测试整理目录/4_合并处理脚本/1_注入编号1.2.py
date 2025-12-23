# -*- coding: utf-8 -*-
import os
import csv
from pathlib import Path
from docx import Document

# ============================================================
# ⚙️ 路径配置
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(BASE_DIR)

# CSV 数据源路径
CSV_PATH = os.path.join(ROOT_DIR, "3_會計及財務記賬系統 - 下載文件", "屯門婦聯 - 會計及財務記賬系統 - 支出賬.csv")

# 目标文件夹路径
DIR_1A = os.path.join(ROOT_DIR, "1A_课程行政清单_领款单", "output")
DIR_1B = os.path.join(ROOT_DIR, "1B_杂费领款单", "output")

# ============================================================
# 🛠️ 核心功能函数
# ============================================================

def load_csv_data(csv_path):
    """读取 CSV 文件 (自动去除表头空格)"""
    data = []
    if not os.path.exists(csv_path):
        print(f"❌ 未找到 CSV 文件: {csv_path}")
        return data

    try:
        # 使用 utf-8-sig 防止中文乱码
        with open(csv_path, mode='r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            
            # 🧹 关键步骤：清洗表头，去除所有空格
            # 这样 ' 種類 ' 就会变成 '種類'
            if reader.fieldnames:
                original_headers = reader.fieldnames
                reader.fieldnames = [name.strip() for name in reader.fieldnames]
                print(f"📋 CSV 表头已读取: {reader.fieldnames}")
            
            data = list(reader)
            print(f"✅ 成功读取 CSV 数据，共 {len(data)} 条记录。")

            # 🔍 自检：确认关键列是否存在
            required_cols = ["收款人", "編號", "種類"]
            missing = [col for col in required_cols if col not in reader.fieldnames]
            if missing:
                print(f"❌ 严重警告：CSV 中缺少以下关键列，脚本将无法正常工作: {missing}")
                print(f"   (请检查 CSV 文件是否包含这些列名)")
            
    except Exception as e:
        print(f"❌ 读取 CSV 失败: {e}")
    return data

def inject_code_into_docx(file_path, code):
    """将编号注入到 Word 文档"""
    try:
        doc = Document(file_path)
        is_injected = False
        formatted_code = f"     {code}" # 加空格排版

        for table in doc.tables:
            for row in table.rows:
                for i, cell in enumerate(row.cells):
                    if "電腦編號" in cell.text:
                        if i + 1 < len(row.cells):
                            target_cell = row.cells[i + 1]
                            target_cell.text = formatted_code
                            is_injected = True
                            print(f"   -> 写入成功：{code}")
                            break
                if is_injected: break
            if is_injected: break

        if is_injected:
            doc.save(file_path)
        else:
            print(f"   ⚠️ 未找到 '電腦編號：' 锚点")

    except Exception as e:
        print(f"   ❌ Word 处理出错: {e}")

# ============================================================
# 🚀 主程序
# ============================================================

def main():
    print("🚀 开始执行：注入编号 1.4 (修正列名为 '種類')")
    
    # 1. 加载数据
    csv_rows = load_csv_data(CSV_PATH)
    if not csv_rows:
        return

    # 2. 处理 1A 文件夹 (按人名匹配)
    # 匹配规则：CSV '收款人' == 文件名(去除后缀)
    print(f"\n📂 扫描 1A: {DIR_1A}")
    path_1a = Path(DIR_1A)
    if path_1a.exists():
        for file in path_1a.glob("*-领款单.docx"):
            print(f"📄 处理: {file.name}")
            target_name = file.stem.replace("-领款单", "").strip()
            
            found_code = None
            for row in csv_rows:
                # 🔍 查找收款人
                if row.get('收款人', '').strip() == target_name:
                    found_code = row.get('編號', '').strip()
                    break
            
            if found_code:
                inject_code_into_docx(str(file), found_code)
            else:
                print(f"   ⚠️ 未找到收款人: {target_name}")
    else:
        print(f"⚠️ 文件夹不存在: {DIR_1A}")

    # 3. 处理 1B 文件夹 (按种类匹配)
    # 匹配规则：文件名包含关键词 -> 对应 CSV '種類' 列的值
    print(f"\n📂 扫描 1B: {DIR_1B}")
    path_1b = Path(DIR_1B)
    
    rules_1b = {
        "打印费领款单": "印刷",
        "FaceBook宣传费领款单": "廣告及推廣",
        "网费领款单": "電話及互聯網費"
    }

    if path_1b.exists():
        for file in path_1b.glob("*.docx"):
            matched_category = None
            for keyword, category in rules_1b.items():
                if keyword in file.name:
                    matched_category = category
                    break
            
            if matched_category:
                print(f"📄 杂费文件: {file.name} (寻找种类: {matched_category})")
                
                found_code = None
                for row in csv_rows:
                    # 🟢 修改处：使用 '種類' 列进行匹配
                    csv_type = row.get('種類', '').strip()
                    
                    if csv_type == matched_category:
                        found_code = row.get('編號', '').strip()
                        break # 取第一条匹配的
                
                if found_code:
                    inject_code_into_docx(str(file), found_code)
                else:
                    print(f"   ⚠️ CSV中未找到种类: {matched_category}")
    else:
        print(f"⚠️ 文件夹不存在: {DIR_1B}")

    print("\n✅ 所有任务完成。")

if __name__ == "__main__":
    main()