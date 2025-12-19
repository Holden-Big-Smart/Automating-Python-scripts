import json
import pandas as pd

# 读取 JSON 文件
with open("get_attendance.json", "r", encoding="utf-8") as f:
    data = json.load(f)

# 提取 membership 数据
members = data.get("data", {}).get("membership", [])

# 初始化统一联系人列表
contact_list = []

# 遍历会员信息，提取住宅电话和手机
for member in members:
    name = member.get("name_zh", "").strip()
    residence = member.get("contact_residence", "").strip()
    mobile = member.get("contact_mobile", "").strip()

    if residence:
        contact_list.append([name, residence])
    if mobile:
        contact_list.append([name, mobile])

# 写入 contact.xlsx，一个工作表中写入所有电话
df = pd.DataFrame(contact_list, columns=["姓名", "电话"])
df.to_excel("contact.xlsx", index=False)

print("✅ 已生成 contact.xlsx，姓名与电话统一写入一张表")
