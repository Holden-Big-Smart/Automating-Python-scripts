
from docxtpl import DocxTemplate
import json
from datetime import datetime

# 载入 JSON 数据
with open("attendance.json", "r", encoding="utf-8") as f:
    data = json.load(f)["data"]

course = data["course"]
members = data["membership"]
dates = data["class_date"]

# 格式化日期列表 + 提取中文星期几
date_list = [
    datetime.strptime(d["class_date"], "%Y-%m-%d").strftime("%m月%d日")
    for d in dates
]
weekday_map = {
    "Monday": "一",
    "Tuesday": "二",
    "Wednesday": "三",
    "Thursday": "四",
    "Friday": "五",
    "Saturday": "六",
    "Sunday": "日"
}
weekday_zh = weekday_map.get(dates[0]["weekday_fullname"], "？")

# 构造学生数据，加入「序號」字段
students = []
for idx, m in enumerate(members, start=1):
    students.append({
        "序號": idx,
        "姓名": m["name_zh"],
        "會員編號": m["code"],
        "收據編號": m["receipt_code"]
    })

# 渲染上下文数据
context = {
    "課程名稱": course["title"],
    "課程編號": course["code"],
    "開始日期": course["start_date"],
    "結束日期": course["end_date"],
    "星期幾": weekday_zh,
    "上課時間": f"{course['start_time']} - {course['end_time']}",
    "學生人數": len(students),
    "堂數": course["class_count"],
    "導師": course["tutor_name_zh"],
    "導師編號": course["tutor_code"],
    "日期列表": date_list,
    "學員列表": students
}

# 加载模板并渲染
tpl = DocxTemplate("测试模板竖.docx")  # 请使用无 loop.index 的模板
tpl.render(context)

# 保存生成文件
tpl.save("output_課堂點名紙.docx")
print("✅ 渲染成功，檔案已輸出：output_課堂點名紙.docx")
