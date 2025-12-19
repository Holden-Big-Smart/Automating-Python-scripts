import os
import pandas as pd
import re
from datetime import datetime
from difflib import SequenceMatcher

# === 参数设置 ===
BASE_DIR = os.getcwd()
WAIT_PATH = os.path.join(BASE_DIR, "等待(下期课程)")
PROGRESS_PATH = os.path.join(BASE_DIR, "进行(本期课程)")
OUTPUT_PATH = os.path.join(BASE_DIR, "output")

# 自动定位文件
wait_file = [f for f in os.listdir(WAIT_PATH) if f.endswith(".csv")][0]
progress_file = [f for f in os.listdir(PROGRESS_PATH) if f.endswith(".csv")][0]

wait_df = pd.read_csv(os.path.join(WAIT_PATH, wait_file), dtype=str).fillna("")
progress_df = pd.read_csv(os.path.join(PROGRESS_PATH, progress_file), dtype=str).fillna("")

# 剔除课程名称中的课程编号
def clean_name(name):
    return re.sub(r"\(SIC\d{6}\)", "", name).strip()

# 模糊相似度评分（用于课程名称）
def get_similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()

# 附加字段加分机制
def extra_score(row1, row2):
    score = 0
    if row1.get("逢星期") == row2.get("逢星期"):
        score += 12
    if row1.get("時間") == row2.get("時間"):
        score += 10
    if row1.get("收費") == row2.get("收費"):
        score += 8
    if row1.get("堂數") == row2.get("堂數"):
        score += 6
    if row1.get("上限") == row2.get("上限"):
        score += 4
    return score

# 主函数：筛选“本期中下期未出现”的课程
def find_unmatched_courses():
    unmatched_rows = []
    for idx, row in progress_df.iterrows():
        teacher = row.get("導師", "").strip()
        name_a = clean_name(row.get("名稱", ""))

        # 同一导师下的数据
        wait_subset = wait_df[wait_df["導師"] == teacher]
        matched = False

        for _, row_b in wait_subset.iterrows():
            name_b = clean_name(row_b.get("名稱", ""))
            name_similarity = get_similarity(name_a, name_b)
            name_score = name_similarity * 60

            if name_score < 60:
                continue  # 主体差异大，跳过

            extra = extra_score(row, row_b)
            total_score = name_score + extra

            if total_score >= 85:
                matched = True
                break

        if not matched:
            unmatched_rows.append(row)

    return pd.DataFrame(unmatched_rows)

# 执行并导出
if __name__ == "__main__":
    print("🔍 正在查找下期未出现的本期课程...")
    result_df = find_unmatched_courses()

    if result_df.empty:
        print("✅ 无需更新，所有本期课程均已在下期中列出。")
    else:
        now = datetime.now().strftime("%Y-%m-%d_%H%M")
        out_file = os.path.join(OUTPUT_PATH, f"{now}_待更新课程.csv")
        result_df.to_csv(out_file, index=False, encoding="utf-8-sig")
        print(f"📦 共 {len(result_df)} 条待更新课程，已导出至：{out_file}")
