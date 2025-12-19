import pandas as pd
from datetime import datetime

# ========= 日期解析：输入格式 日.月-日.月.年（也兼容 / 作为分隔） =========
def parse_date_input(user_input: str):
    """
    解析用户输入的日期范围，支持：
      - '30.10-9.11.2025'
      - '30/10-9/11/2025'
      - '07.11-03.01.2026'
      - '07/11-03/01/2026'
    规则：
      - 最后的“年”属于结束日期；
      - 若 end_month < start_month 则视为跨年：start_year = end_year - 1，否则同年；
      - 返回：
          start_date (datetime)
          end_date (datetime)
          excel_var  -> 'd/m/yyyy-d/m/yyyy'
          display    -> 'd.m-d.m.yyyy'
          filename   -> 'd.m-d.m.yyyy.xlsx'
    """
    s = user_input.strip().replace('/', '.')
    if '-' not in s or s.count('.') < 3:
        raise ValueError("输入格式应为：日.月-日.月.年  例如：30.10-9.11.2025 或 07.11-03.01.2026")

    left, right = s.split('-', 1)
    right_parts = right.split('.')
    if len(right_parts) != 3:
        raise ValueError("右半段必须包含 年：示例 9.11.2025 或 03.01.2026")

    def _to_int(x):
        x = x.strip()
        if not x.isdigit():
            raise ValueError(f"日期中包含非数字：{x}")
        return int(x)

    try:
        sd_str, sm_str = left.split('.')
        ed_str, em_str, y_str = right_parts
        sd, sm = _to_int(sd_str), _to_int(sm_str)
        ed, em = _to_int(ed_str), _to_int(em_str)
        end_year = _to_int(y_str)
    except Exception as e:
        raise ValueError("日期中存在无法解析的数字，请检查格式（示例：30.10-9.11.2025）") from e

    if not (1 <= sd <= 31 and 1 <= ed <= 31 and 1 <= sm <= 12 and 1 <= em <= 12):
        raise ValueError("日期/月份超出有效范围，请检查（日:1-31，月:1-12）。")

    start_year = end_year - 1 if em < sm else end_year
    start_date = datetime(start_year, sm, sd)
    end_date = datetime(end_year, em, ed)
    excel_var = f"{sd}/{sm}/{start_year}-{ed}/{em}/{end_year}"
    display = f"{sd}.{sm}-{ed}.{em}.{end_year}"
    filename = f"{display}.xlsx"
    return start_date, end_date, excel_var, display, filename


# ========= 主流程 =========
def main():
    user_input = input("请输入数据时间（格式：日.月-日.月.年，例如 30.10-9.11.2025 或 07.11-03.01.2026）：").strip()
    start_date, end_date, excel_var, display_str, filename_str = parse_date_input(user_input)

    print(f"数据日期（Excel变量）：{excel_var}")
    print(f"Excel显示日期：{display_str}")
    print(f"输出文件名：{filename_str}")

    # 2) 读取「課程收據」
    course_file = "屯門婦聯 - 會員及課程管理系統 - 課程收據.csv"
    df_course = pd.read_csv(course_file)
    drop_cols_course = ["服務中心","報名中心","會員編號","會員姓名(英文)","會員類別","課程收費","其他收費","扣減","日期","撤銷"]
    df_course = df_course.drop(columns=[c for c in drop_cols_course if c in df_course.columns], errors="ignore")

    print("\n[課程收據] 当前标题行：")
    print(list(df_course.columns))

    # --- 新排序逻辑：按「經辨人員」降序 + 「編號」升序 ---
    if "經辨人員" in df_course.columns:
        if "編號" in df_course.columns:
            df_course = df_course.sort_values(by=["經辨人員", "編號"], ascending=[False, True])
        else:
            df_course = df_course.sort_values(by="經辨人員", ascending=False)

    print("\n[課程收據] 排序后經辨人員（前10条）：")
    if "經辨人員" in df_course.columns:
        print(df_course["經辨人員"].head(10).tolist())
    else:
        print("（未找到「經辨人員」列）")

    if "總額" not in df_course.columns:
        raise KeyError("課程收據缺少「總額」列")
    df_course["總額"] = df_course["總額"].replace(r'[\$,]', '', regex=True).astype(float)
    課程總額 = df_course["總額"].iloc[-1] if not df_course["總額"].empty else 0.0
    print(f"\n課程總額：${課程總額:.2f}")

    if "經辨人員" not in df_course.columns or "付款方式" not in df_course.columns:
        raise KeyError("課程收據缺少「經辨人員」或「付款方式」列")

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

    # 3) 读取「會費收據」
    fee_file = "屯門婦聯 - 會員及課程管理系統 - 會費收據.csv"
    df_fee = pd.read_csv(fee_file)
    drop_cols_fee = ["中心","會員編號","會員姓名(英文)","日期","撤銷"]
    df_fee = df_fee.drop(columns=[c for c in drop_cols_fee if c in df_fee.columns], errors="ignore")

    print("\n[會費收據] 当前标题行：")
    print(list(df_fee.columns))

    # --- 新排序逻辑：按「經辨人員」降序 + 「編號」升序 ---
    if "經辨人員" in df_fee.columns:
        if "編號" in df_fee.columns:
            df_fee = df_fee.sort_values(by=["經辨人員", "編號"], ascending=[False, True])
        else:
            df_fee = df_fee.sort_values(by="經辨人員", ascending=False)

    print("\n[會費收據] 排序后經辨人員（前10条）：")
    if "經辨人員" in df_fee.columns:
        print(df_fee["經辨人員"].head(10).tolist())
    else:
        print("（未找到「經辨人員」列）")

    if "會費" not in df_fee.columns:
        raise KeyError("會費收據缺少「會費」列")
    df_fee["會費"] = df_fee["會費"].replace(r'[\$,]', '', regex=True).astype(float)
    會費總額 = df_fee["會費"].iloc[-1] if not df_fee["會費"].empty else 0.0
    print(f"\n會費總額：${會費總額:.2f}")

    if "經辨人員" not in df_fee.columns or "付款方式" not in df_fee.columns:
        raise KeyError("會費收據缺少「經辨人員」或「付款方式」列")

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

    # 4) 合并结果
    merged_vars = {}
    all_staff = set([k.split("_")[0] for k in course_vars.keys()] + [k.split("_")[0] for k in fee_vars.keys()])
    for staff in all_staff:
        total_course = course_vars.get(f"{staff}_課程總額", 0.0)
        total_fee = fee_vars.get(f"{staff}_會費總額", 0.0)
        merged_vars[f"{staff}_合計總額"] = total_course + total_fee
        methods = set(
            [k.replace(f"{staff}_課程", "") for k in course_vars if k.startswith(staff)] +
            [k.replace(f"{staff}_會費", "") for k in fee_vars if k.startswith(staff)]
        )
        for method in methods:
            if method == "總額":
                continue
            course_val = course_vars.get(f"{staff}_課程{method}", 0.0)
            fee_val = fee_vars.get(f"{staff}_會費{method}", 0.0)
            merged_vars[f"{staff}_{method}合計"] = course_val + fee_val

    print("\n[合併結果]：")
    for k, v in merged_vars.items():
        print(f"{k}：${v:.0f}")

    # 5) 导出 Excel
    with pd.ExcelWriter(filename_str, engine="openpyxl") as writer:
        df_course.to_excel(writer, sheet_name="課程收據", index=False)
        df_fee.to_excel(writer, sheet_name="會費收據", index=False)
        merged_df = pd.DataFrame(list(merged_vars.items()), columns=["項目", "金額"])
        merged_df.to_excel(writer, sheet_name="統計結果", index=False)

        wb = writer.book
        ws = wb.create_sheet("數據信息")
        ws["A1"] = "數據日期"
        ws["B1"] = excel_var
        ws["A2"] = "顯示日期"
        ws["B2"] = display_str
        ws["A3"] = "開始日期(ISO)"
        ws["B3"] = start_date.strftime("%Y-%m-%d")
        ws["A4"] = "結束日期(ISO)"
        ws["B4"] = end_date.strftime("%Y-%m-%d")

    print(f"\n✅ 已输出文件：{filename_str}")


if __name__ == "__main__":
    main()
