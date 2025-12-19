import os
import pandas as pd
import re
from datetime import datetime
from difflib import SequenceMatcher

# =========================
# 路径配置
# =========================
BASE_DIR = os.getcwd()
WAIT_PATH = os.path.join(BASE_DIR, "等待(下期课程)")
PROGRESS_PATH = os.path.join(BASE_DIR, "进行(本期课程)")
OUTPUT_PATH = os.path.join(BASE_DIR, "output")

# 自动查找 CSV 文件
wait_file = [f for f in os.listdir(WAIT_PATH) if f.endswith(".csv")][0]
progress_file = [f for f in os.listdir(PROGRESS_PATH) if f.endswith(".csv")][0]

wait_df = pd.read_csv(os.path.join(WAIT_PATH, wait_file), dtype=str).fillna("")
progress_df = pd.read_csv(os.path.join(PROGRESS_PATH, progress_file), dtype=str).fillna("")

# =========================
# 工具函数
# =========================

def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()

def extra_scores(rowA, rowB):
    return {
        "逢星期得分": 12 if rowA.get("逢星期") == rowB.get("逢星期") else 0,
        "時間得分": 10 if rowA.get("時間") == rowB.get("時間") else 0,
        "收費得分": 8 if rowA.get("收費") == rowB.get("收費") else 0,
        "堂數得分": 6 if rowA.get("堂數") == rowB.get("堂數") else 0,
        "上限得分": 4 if rowA.get("上限") == rowB.get("上限") else 0,
    }

# =========================
# 核心逻辑：评分并选出最佳匹配
# =========================

def find_unmatched_courses():
    unmatched = []
    log_records = []

    for _, rowA in progress_df.iterrows():
        teacher = rowA.get("導師", "")
        nameA = rowA.get("名稱", "")

        wait_subset = wait_df[wait_df["導師"] == teacher]

        best_score = -1
        best_match_status = "無匹配"
        for _, rowB in wait_subset.iterrows():
            nameB = rowB.get("名稱", "")
            sim = similarity(nameA, nameB)
            name_score = sim * 60

            if name_score < 50:
                extras = {key: 0 for key in ["逢星期得分", "時間得分", "收費得分", "堂數得分", "上限得分"]}
                total = name_score
                match_result = "否（相似度過低）"
            else:
                extras = extra_scores(rowA, rowB)
                total = name_score + sum(extras.values())
                match_result = "是" if total >= 85 else "否"

            # 保存日志记录
            log_records.append({
                "A_课程名称": nameA,
                "B_课程名称": nameB,
                "导师": teacher,
                "名称相似度": round(sim, 4),
                "主体名称得分": round(name_score, 2),
                **extras,
                "总分": round(total, 2),
                "是否匹配": match_result
            })

            if total > best_score:
                best_score = total
                best_match_status = match_result

        if len(wait_subset) == 0 or best_match_status != "是":
            unmatched.append(rowA)

    return pd.DataFrame(unmatched), pd.DataFrame(log_records)

# =========================
# 主函数入口
# =========================

if __name__ == "__main__":
    print("🔍 正在进行课程匹配比对...")

    unmatched_df, log_df = find_unmatched_courses()
    now = datetime.now().strftime("%Y-%m-%d_%H%M")

    if not unmatched_df.empty:
        output_file = os.path.join(OUTPUT_PATH, f"{now}_待更新课程.csv")
        unmatched_df.to_csv(output_file, index=False, encoding="utf-8-sig")
        print(f"📌 待更新课程共 {len(unmatched_df)} 条，已保存至：{output_file}")
    else:
        print("✅ 所有课程都已在下期课程中找到最佳匹配。")

    log_file = os.path.join(OUTPUT_PATH, f"{now}_匹配日志.csv")
    log_df.to_csv(log_file, index=False, encoding="utf-8-sig")
    print(f"📄 匹配日志已保存至：{log_file}")
