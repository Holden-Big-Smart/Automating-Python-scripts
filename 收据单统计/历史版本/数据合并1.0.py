import pandas as pd
import os
from datetime import datetime

# ===== Step 1. 询问日期范围 =====
date_input = input("请输入数据时间（格式：2025/10/30-2025/11/6）：").strip()
print("数据时间为：", date_input)

# ===== Step 2. 读取「課程收據」并处理 =====
course_file = "屯門婦聯 - 會員及課程管理系統 - 課程收據.csv"
df_course = pd.read_csv(course_file)

# (1) 删除指定列
drop_cols_course = ["服務中心","報名中心","會員編號","會員姓名(英文)","會員類別","課程收費","其他收費","扣減","日期","撤銷"]
df_course = df_course.drop(columns=[c for c in drop_cols_course if c in df_course.columns], errors="ignore")

# (2) 打印标题行
print("\n[課程收據] 当前标题行：")
print(list(df_course.columns))

# (3) 按「經辨人員」字母顺序降序排列
if "經辨人員" in df_course.columns:
    df_course = df_course.sort_values(by="經辨人員", ascending=False)
print("\n[課程收據] 排序后經辨人員：")
print(df_course["經辨人員"].tolist())

# (4) 读取「總額」列最后一个数据
df_course["總額"] = df_course["總額"].replace('[\$,]', '', regex=True).astype(float)
課程總額 = df_course["總額"].iloc[-1] if not df_course["總額"].empty else 0
print(f"\n課程總額：${課程總額:.2f}")

# (5) 按經辨人員、付款方式分组统计
group_course = df_course.groupby("經辨人員")
course_vars = {}

for staff, group in group_course:
    total = group["總額"].sum()
    course_vars[f"{staff}_課程總額"] = total

    for pay_method, pay_group in group.groupby("付款方式"):
        subtotal = pay_group["總額"].sum()
        course_vars[f"{staff}_課程{pay_method}"] = subtotal

print("\n[課程收據] 統計結果：")
for k, v in course_vars.items():
    print(f"{k}：${v:.0f}")

# ===== Step 3. 读取「會費收據」并处理 =====
fee_file = "屯門婦聯 - 會員及課程管理系統 - 會費收據.csv"
df_fee = pd.read_csv(fee_file)

# (1) 删除指定列
drop_cols_fee = ["中心","會員編號","會員姓名(英文)","日期","撤銷"]
df_fee = df_fee.drop(columns=[c for c in drop_cols_fee if c in df_fee.columns], errors="ignore")

# (2) 打印标题行
print("\n[會費收據] 当前标题行：")
print(list(df_fee.columns))

# (3) 按「經辨人員」字母顺序降序排列
if "經辨人員" in df_fee.columns:
    df_fee = df_fee.sort_values(by="經辨人員", ascending=False)
print("\n[會費收據] 排序后經辨人員：")
print(df_fee["經辨人員"].tolist())

# (4) 读取「會費」列最后一个数据
df_fee["會費"] = df_fee["會費"].replace('[\$,]', '', regex=True).astype(float)
會費總額 = df_fee["會費"].iloc[-1] if not df_fee["會費"].empty else 0
print(f"\n會費總額：${會費總額:.2f}")

# (5) 按經辨人員、付款方式分组统计
group_fee = df_fee.groupby("經辨人員")
fee_vars = {}

for staff, group in group_fee:
    total = group["會費"].sum()
    fee_vars[f"{staff}_會費總額"] = total

    for pay_method, pay_group in group.groupby("付款方式"):
        subtotal = pay_group["會費"].sum()
        fee_vars[f"{staff}_會費{pay_method}"] = subtotal

print("\n[會費收據] 統計結果：")
for k, v in fee_vars.items():
    print(f"{k}：${v:.0f}")

# ===== Step 4. 合并統計結果 =====
merged_vars = {}
all_staff = set([k.split("_")[0] for k in course_vars.keys()] + [k.split("_")[0] for k in fee_vars.keys()])

for staff in all_staff:
    # 合并總額
    total_course = course_vars.get(f"{staff}_課程總額", 0)
    total_fee = fee_vars.get(f"{staff}_會費總額", 0)
    merged_vars[f"{staff}_合計總額"] = total_course + total_fee

    # 各付款方式合计
    methods = set(
        [k.replace(f"{staff}_課程", "") for k in course_vars if k.startswith(staff)] +
        [k.replace(f"{staff}_會費", "") for k in fee_vars if k.startswith(staff)]
    )
    for method in methods:
        if method == "總額": continue
        course_val = course_vars.get(f"{staff}_課程{method}", 0)
        fee_val = fee_vars.get(f"{staff}_會費{method}", 0)
        merged_vars[f"{staff}_{method}合計"] = course_val + fee_val

print("\n[合併結果]：")
for k, v in merged_vars.items():
    print(f"{k}：${v:.0f}")

# ===== Step 5. 导出结果 =====
output_filename = f"{date_input.replace('/', '.').replace('-', '.')}.xlsx"
with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
    df_course.to_excel(writer, sheet_name="課程收據", index=False)
    df_fee.to_excel(writer, sheet_name="會費收據", index=False)

    # 将合并数据写入最后一个sheet
    merged_df = pd.DataFrame(list(merged_vars.items()), columns=["項目", "金額"])
    merged_df.to_excel(writer, sheet_name="統計結果", index=False)

print(f"\n✅ 已输出文件：{output_filename}")
