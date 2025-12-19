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

# 去除课程编号 (SICxxxxxx)
def clean_name(name):
    return re.sub(r"\(SIC\d{6}\)", "", name).strip()

# difflib 中文模糊匹配
def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()

# 附加字段加分（最高 40 分）
def extra_scores(rowA, rowB):
    scores = {
        "逢星期得分": 12 if rowA.get("逢星期") == rowB.get("逢星期") else 0,
        "时间得分": 10 if rowA.get("時間") == rowB.get("時間") else 0,
        "收费得分": 8 if rowA.get("收費") == rowB.get("收費") else 0,
        "堂数得分": 6 if rowA.get("堂數") == rowB.get("堂數") else 0,
        "上限得分": 4 if rowA.get("上限") == rowB.get("上限") else 0,
    }
    return scores

# =========================
# 核心匹配函数（含日志）
# =========================

def find_unmatched_courses():
    unmatched = []
    log_records = []

    for idxA, rowA in progress_df.iterrows():
        teacher = rowA.get("導師", "")
        nameA_raw = rowA.get("名稱", "")
        nameA = clean_name(nameA_raw)

        # 从下期课程取出相同导师的记录
        wait_subset = wait_df[wait_df["導師"] == teacher]

        matched = False

        for idxB, rowB in wait_subset.iterrows():
            nameB_raw = rowB.get("名稱", "")
            nameB = clean_name(nameB_raw)

            # 主体名称相似度
            sim = similarity(nameA, nameB)
            name_score = sim * 60

            # 不足 60 分，不进入加分阶段
            if name_score < 45:
                log_records.append({
                    "A_课程名称": nameA_raw,
                    "B_课程名称": nameB_raw,
                    "导师": teacher,
                    "名称相似度": round(sim, 4),
                    "主体名称得分": round(name_score, 2),
                    "逢星期得分": 0,
                    "时间得分": 0,
                    "收费得分": 0,
                    "堂数得分": 0,
                    "上限得分": 0,
                    "总分": round(name_score, 2),
                    "是否匹配": "否（主体相似度过低）"
                })
                continue

            # 附加字段加分
            extras = extra_scores(rowA, rowB)
            total_score = name_score + sum(extras.values())

            # 记录日志
            log_records.append({
                "A_课程名称": nameA_raw,
                "B_课程名称": nameB_raw,
                "导师": teacher,
                "名称相似度": round(sim, 4),
                "主体名称得分": round(name_score, 2),
                **extras,
                "总分": round(total_score, 2),
                "是否匹配": "是" if total_score >= 85 else "否"
            })

            # 匹配成功则不需要继续比对
            if total_score >= 80:
                matched = True
                break

        if not matched:
            unmatched.append(rowA)

    return pd.DataFrame(unmatched), pd.DataFrame(log_records)

# =========================
# 主函数
# =========================

if __name__ == "__main__":
    print("🔍 正在进行比对与日志记录...")

    unmatched_df, log_df = find_unmatched_courses()

    now = datetime.now().strftime("%Y-%m-%d_%H%M")

    # 输出待更新课程
    if not unmatched_df.empty:
        out_course = os.path.join(OUTPUT_PATH, f"{now}_待更新课程.csv")
        unmatched_df.to_csv(out_course, index=False, encoding="utf-8-sig")
        print(f"📌 待更新课程共 {len(unmatched_df)} 条，已输出：{out_course}")
    else:
        print("✅ 所有课程都已在下期课程中找到匹配，无需更新。")

    # 输出日志文件
    out_log = os.path.join(OUTPUT_PATH, f"{now}_匹配日志.csv")
    log_df.to_csv(out_log, index=False, encoding="utf-8-sig")
    print(f"📄 匹配日志已生成：{out_log}")
