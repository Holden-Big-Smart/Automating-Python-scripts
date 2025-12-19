import pandas as pd
from datetime import datetime
import os

# ========== 设置 ==========
input_file = "屯門婦聯 - 會員及課程管理系統 - 課程.csv"  # 原始数据文件
output_folder = "output"
output_file = os.path.join(output_folder, "Course-List.xlsx")
target_column = "上課日期"

# ========== 获取当前日期 ==========
today = datetime.today().date()

# ========== 读取数据 ==========
df = pd.read_csv(input_file)

# ========== 检查必要列是否存在 ==========
if target_column not in df.columns:
    raise ValueError(f"缺少「{target_column}」列，请检查文件内容。")

# ========== 存储筛选后的行 ==========
filtered_rows = []

# ========== 遍历每一行（从第二行开始） ==========
for index, row in df.iterrows():
    if pd.isna(row[target_column]):
        continue
    try:
        # 提取"|"后的日期字符串
        parts = str(row[target_column]).split("|")
        if len(parts) != 2:
            continue
        end_date_str = parts[1].split(" ")[0]  # 提取如 "2025-11-27"

        # 转换为日期对象
        end_date = datetime.strptime(end_date_str, "%Y-%m-%d").date()

        # 计算日期差
        delta_days = (end_date - today).days

        # 筛选差值小于等于25天的行
        if delta_days <= 25:
            row_data = row.to_dict()
            row_data["缴费日期差"] = delta_days
            filtered_rows.append(row_data)
    except Exception as e:
        print(f"第{index+2}行处理出错：{e}")

# ========== 构建新DataFrame并排序 ==========
if filtered_rows:
    output_df = pd.DataFrame(filtered_rows)
    # 将"缴费日期差"移至第一列
    cols = ["缴费日期差"] + [col for col in output_df.columns if col != "缴费日期差"]
    output_df = output_df[cols]

    # 排序
    output_df = output_df.sort_values(by="缴费日期差", ascending=True)

    # ========== 输出 ==========
    os.makedirs(output_folder, exist_ok=True)
    output_df.to_excel(output_file, index=False)
    print(f"筛选后的课程已输出到：{output_file}")
else:
    print("没有课程符合日期差小于等于25天的条件。")
