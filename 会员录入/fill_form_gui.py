import tkinter as tk
from selenium import webdriver
from selenium.webdriver.common.by import By
import time

# ---------------------
# 自动填表函数
# ---------------------
def autofill(name, birth):
    # 启动浏览器（注意：替换为你的 chromedriver 路径）
    driver = webdriver.Chrome()

    # 打开目标网页（请替换为你的表单URL）
    driver.get("https://tmdwa-member.sysone.hk:9808/v1.0/membership/main.php?p=1&c=0")

    time.sleep(2)  # 等待网页加载

    try:
        # 填入中文姓名
        name_input = driver.find_element(By.ID, "form_name_zh")
        name_input.clear()
        name_input.send_keys(name)

        # 填入出生日期
        dob_input = driver.find_element(By.ID, "form_birth_of_date")
        dob_input.clear()
        dob_input.send_keys(birth)

        print("✅ 已自动填入表单。")

    except Exception as e:
        print("❌ 出现错误：", e)

# ---------------------
# 构建GUI界面 
# ---------------------
def launch_gui():
    root = tk.Tk()
    root.title("會員資料填寫助手")
    root.geometry("300x200")

    # 姓名
    tk.Label(root, text="會員姓名").pack()
    entry_name = tk.Entry(root)
    entry_name.pack()

    # 出生日期
    tk.Label(root, text="出生日期（格式：YYYY-MM-DD）").pack()
    entry_birth = tk.Entry(root)
    entry_birth.pack()

    # 提交按钮
    def on_submit():
        name = entry_name.get()
        birth = entry_birth.get()
        if name and birth:
            autofill(name, birth)
        else:
            print("⚠️ 請填寫完整資料")

    tk.Button(root, text="確定自動填寫", command=on_submit).pack(pady=10)

    root.mainloop()

# ---------------------
# 程序入口
# ---------------------
if __name__ == "__main__":
    launch_gui()
