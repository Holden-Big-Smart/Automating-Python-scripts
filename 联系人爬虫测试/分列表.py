import json
import pandas as pd

# 读取 JSON 文件
with open("get_attendance.json", "r", encoding="utf-8") as f:
    data = json.load(f)

# 提取 membership 数据
members = data.get("data", {}).get("membership", [])

# 初始化列表
residence_data = []
mobile_data = []

# 遍历会员信息
for member in members:
    name_zh = member.get("name_zh", "").strip()
    contact_residence = member.get("contact_residence", "").strip()
    contact_mobile = member.get("contact_mobile", "").strip()

    if contact_residence:
        residence_data.append([name_zh, contact_residence])
    if contact_mobile:
        mobile_data.append([name_zh, contact_mobile])

# 写入 Excel 文件
with pd.ExcelWriter("contact.xlsx", engine="openpyxl") as writer:
    pd.DataFrame(mobile_data, columns=["姓名", "手机"]).to_excel(writer, sheet_name="手机", index=False)
    pd.DataFrame(residence_data, columns=["姓名", "住宅电话"]).to_excel(writer, sheet_name="住宅电话", index=False)

print("✅ 已成功写入 contact.xlsx")
