import pandas as pd
from datetime import datetime

# 文件路径
course_file = "屯門婦聯 - 會員及課程管理系統 - 課程.csv"
receipt_file = "屯門婦聯 - 會員及課程管理系統 - 課程收據.csv"

# 读取CSV时强制为字符串，避免格式问题
df_course = pd.read_csv(course_file, dtype=str)
df_receipt = pd.read_csv(receipt_file, dtype=str)

# 清洗空白字符和缺失值
df_course = df_course.fillna("").applymap(str.strip)
df_receipt = df_receipt.fillna("").applymap(str.strip)

# 筛选付款方式为“現金”的数据
df_cash = df_receipt[df_receipt["付款方式"] == "現金"].copy()

# 建立課程名稱 → 類別映射字典（完全匹配）
name_to_category = dict(zip(df_course["名稱"], df_course["類別"]))

# 为现金数据添加“類別”列（匹配不上为""）
df_cash["類別"] = df_cash["課程名稱"].map(name_to_category).fillna("")

# 尝试将總額转为浮点数（用于后续统计）
if "總額" in df_cash.columns:
    df_cash["總額_clean"] = (
        df_cash["總額"]
        .str.replace(",", "")  # 去除千位符
        .str.replace("$", "")
        .str.replace("HK", "")
        .str.extract("(\d+\.?\d*)")[0]
        .astype(float)
    )
else:
    df_cash["總額_clean"] = 0.0

# 构造六个表格
grouped_data = {
    "總課程": df_cash,
    "興趣班組 (IC)": df_cash[df_cash["類別"] == "興趣班組 (IC)"],
    "專業課程(PC)": df_cash[df_cash["類別"] == "專業課程(PC)"],
    "活動(E)": df_cash[df_cash["類別"] == "活動(E)"],
    "其他(O)": df_cash[df_cash["類別"] == "其他(O)"],
    "未匹配": df_cash[df_cash["類別"] == ""],
}

# 写入 Excel 文件
now_str = datetime.now().strftime("%Y%m%d_%H%M")
output_path = f"{now_str}課程分類.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    for sheet_name, df in grouped_data.items():
        df_out = df.copy()

        # 插入“類別”列到首列
        if "類別" in df_out.columns:
            cols = ["類別"] + [col for col in df_out.columns if col != "類別"]
            df_out = df_out[cols]

        # 写入数据
        df_out.to_excel(writer, sheet_name=sheet_name, index=False)

        # 写入“總額合計”
        if not df_out.empty:
            total = df_out["總額_clean"].sum()
            summary_row = pd.DataFrame([[None]*(df_out.shape[1]-2) + ["總額合計", total]], columns=df_out.columns)
            summary_row.to_excel(writer, sheet_name=sheet_name, startrow=len(df_out)+1, index=False, header=False)

print(f"✅ 成功输出分类结果文件：{output_path}")
