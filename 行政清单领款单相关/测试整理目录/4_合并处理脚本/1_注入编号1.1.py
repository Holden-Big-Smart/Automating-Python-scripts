import os
import sys
import pandas as pd
from docx import Document
from pathlib import Path

# ==========================================
# 🔧 路径配置修复
# ==========================================

# 1. 获取当前脚本所在的目录 (即: .../测试整理目录/4_合并处理脚本)
current_dir = os.path.dirname(os.path.abspath(__file__))

# 2. 获取项目根目录 (即: .../测试整理目录)
# 假设当前脚本在根目录的下一级子目录中，所以向上退一级
project_root = os.path.dirname(current_dir)

# 3. 构建 CSV 文件的绝对路径
# 注意：根据你的报错信息，文件名是 "屯門婦聯 - 會計及財務記賬系統 - 支出賬.csv"
# 如果实际文件名是 "支出賬.csv"，请修改下面的 csv_filename 变量
csv_folder = "3_會計及財務記賬系統 - 下載文件"
csv_filename = "屯門婦聯 - 會計及財務記賬系統 - 支出賬.csv"
download_csv_path = os.path.join(project_root, csv_folder, csv_filename)

# 4. 构建领款单 output 目录的绝对路径
# 领款单在 "1A_课程行政清单_领款单/output"
output_dir_path = os.path.join(project_root, "1A_课程行政清单_领款单", "output")

print(f"📂 项目根目录: {project_root}")
print(f"📄 读取CSV路径: {download_csv_path}")
print(f"📂 领款单目录: {output_dir_path}\n")

# ==========================================
# 🚀 业务逻辑
# ==========================================

# 检查 CSV 文件是否存在
if not os.path.exists(download_csv_path):
    print(f"❌ 错误：找不到 CSV 文件。\n路径：{download_csv_path}")
    print("👉 请确认文件是否位于 '3_會計及財務記賬系統 - 下載文件' 目录中，且文件名正确。")
    input("按回车键退出...")
    sys.exit(1)

# 读取编号信息
try:
    df = pd.read_csv(download_csv_path, encoding="utf-8-sig")
except Exception as e:
    print(f"❌ 无法读取 CSV 文件，请检查编码或文件格式。")
    raise e

if '收款人' not in df.columns or '編號' not in df.columns:
    print(f"当前列名：{df.columns.tolist()}")
    raise ValueError("❌ CSV 文件缺少 '收款人' 或 '編號' 字段，请检查文件内容。")

# 建立收款人 → 編號 映射 (去除首尾空格以提高匹配率)
name_to_code = dict(zip(
    df['收款人'].astype(str).str.strip(), 
    df['編號'].astype(str).str.strip()
))

# 遍历 output 文件夹中所有领款单
output_path = Path(output_dir_path)

if not output_path.exists():
    print(f"❌ 错误：找不到输出目录 {output_dir_path}")
    print("👉 请先运行 '1A_课程行政清单_领款单' 下的生成脚本以生成文件。")
    input("按回车键退出...")
    sys.exit(1)

unmatched = []
processed_count = 0

print("⏳ 正在处理领款单注入编号...\n")

for file in output_path.glob("*-领款单.docx"):
    # 假设文件名格式为 "姓名-领款单.docx"，提取姓名
    name = file.stem.replace("-领款单", "").strip()
    
    if name in name_to_code:
        code = name_to_code[name]
        try:
            doc = Document(file)
            filled = False

            # 遍历表格寻找目标单元格
            for table in doc.tables:
                for row in table.rows:
                    for i, cell in enumerate(row.cells[:-1]):  # 避免越界
                        if "電腦編號：" in cell.text:
                            # 填入编号 (加空格是为了简单的排版对齐)
                            row.cells[i + 1].text = "     " + code
                            filled = True
                            break
                    if filled: break
                if filled: break

            if filled:
                doc.save(file)
                print(f"✅ 已写入: {name} -> {code}")
                processed_count += 1
            else:
                print(f"⚠️  {name}: 未在文档中找到“電腦編號：”表格位置，跳过。")
                
        except Exception as e:
            print(f"❌ 处理文件失败 {file.name}: {e}")
    else:
        unmatched.append(name)

# 输出结果
if unmatched:
    print("\n⚠️  以下教师未在 CSV 中找到对应编号 (请检查名字是否完全一致)：")
    for name in unmatched:
        print(f" - {name}")
else:
    print("\n✅ 所有教师的电脑编号均已成功匹配并写入！")

print(f"\n📊 共处理文件：{processed_count} 个")
input("\n脚本执行完毕，按任意键关闭...")