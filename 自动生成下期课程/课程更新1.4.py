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
# 匹配函数：附带完整日志与成功映射记录
# =========================
def match_courses():
    unmatched = []
    log_records = []
    mapping_records = []

    for _, rowA in progress_df.iterrows():
        teacher = rowA.get("導師", "")
        nameA = rowA.get("名稱", "")
        wait_subset = wait_df[wait_df["導師"] == teacher]

        best_score = -1
        best_match_rowB = None
        best_match_name_score = 0
        best_match_extra = {}

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
                best_match_rowB = rowB
                best_match_name_score = name_score
                best_match_extra = extras

        if best_score >= 85 and best_match_rowB is not None:
            mapping_records.append({
                "A_课程名称": nameA,
                "B_课程名称": best_match_rowB["名稱"],
                "导师": teacher,
                "总分": round(best_score, 2),
                "主体名称得分": round(best_match_name_score, 2),
                "附加加分": sum(best_match_extra.values()),
                "匹配备注": "匹配成功（最佳得分）"
            })
        else:
            unmatched.append(rowA)

    return pd.DataFrame(unmatched), pd.DataFrame(log_records), pd.DataFrame(mapping_records)

# =========================
# 主执行入口
# =========================
if __name__ == "__main__":
    print("🔍 正在执行课程匹配与日志生成...")

    unmatched_df, log_df, mapping_df = match_courses()
    now = datetime.now().strftime("%Y-%m-%d_%H%M")

    # 输出未匹配课程
    if not unmatched_df.empty:
        unmatched_path = os.path.join(OUTPUT_PATH, f"{now}_待更新课程.csv")
        unmatched_df.to_csv(unmatched_path, index=False, encoding="utf-8-sig")
        print(f"📌 待更新课程：{len(unmatched_df)} 条 → {unmatched_path}")
    else:
        print("✅ 所有课程都已成功匹配。")

    # 输出匹配日志
    log_path = os.path.join(OUTPUT_PATH, f"{now}_匹配日志.csv")
    log_df.to_csv(log_path, index=False, encoding="utf-8-sig")
    print(f"📄 匹配日志已保存至：{log_path}")

    # 输出匹配成功映射日志
    mapping_path = os.path.join(OUTPUT_PATH, f"{now}_映射日志.csv")
    mapping_df.to_csv(mapping_path, index=False, encoding="utf-8-sig")
    print(f"🔗 映射日志已保存至：{mapping_path}")
