import os
import pandas as pd
from docx import Document
from pathlib import Path

# 设置路径
output_dir = "output"
download_csv = "屯門婦聯 - 會計及財務記賬系統 - 下載文件.csv"

# 读取编号信息
try:
    df = pd.read_csv(download_csv, encoding="utf-8-sig")
except Exception as e:
    print(f"❌ 无法读取 CSV 文件：{download_csv}")
    raise e

if '收款人' not in df.columns or '編號' not in df.columns:
    raise ValueError("❌ CSV 文件缺少 '收款人' 或 '編號' 字段")

# 建立收款人 → 編號 映射
name_to_code = dict(zip(df['收款人'].astype(str).str.strip(), df['編號'].astype(str).str.strip()))

# 遍历 output 文件夹中所有领款单
output_path = Path(output_dir)
unmatched = []

for file in output_path.glob("*-领款单.docx"):
    name = file.stem.replace("-领款单", "").strip()
    if name in name_to_code:
        code = name_to_code[name]
        doc = Document(file)
        filled = False

        for table in doc.tables:
            for row in table.rows:
                for i, cell in enumerate(row.cells[:-1]):  # 避免越界
                    if "電腦編號：" in cell.text:
                        row.cells[i + 1].text = "     " + code
                        filled = True
                        break
                if filled:
                    break
            if filled:
                break

        if filled:
            doc.save(file)
            print(f"✅ 已写入电脑编号 → {name}：{code}")
        else:
            print(f"⚠️ 找不到“电脑编号：”表格栏位 → {name}")
    else:
        unmatched.append(name)

# 输出未匹配者
if unmatched:
    print("\\n⚠️ 以下教师未在 CSV 中找到对应编号：")
    for name in unmatched:
        print(" -", name)
else:
    print("\\n✅ 所有教师电脑编号写入成功")

input("\\n脚本执行完毕，按任意键关闭...")
