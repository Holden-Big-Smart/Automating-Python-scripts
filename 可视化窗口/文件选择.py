import tkinter as tk
from tkinter import filedialog  # 文件选择对话框模块

def on_click():
    label.config(text="你点击了按钮！")

def select_file():
    file_path = filedialog.askopenfilename(
        title="选择一个文件",
        filetypes=[("所有文件", "*.*"), ("文本文件", "*.txt")]
    )
    if file_path:
        file_label.config(text=f"已选择文件：\n{file_path}")

# 创建主窗口
window = tk.Tk()
window.title("我的脚本工具")
window.geometry("600x400+300+100")

# 添加欢迎标签
label = tk.Label(window, text="欢迎使用我的脚本工具", font=("微软雅黑", 14))
label.pack(pady=10)

# 添加“点击我”按钮
button = tk.Button(window, text="点击我", command=on_click, font=("微软雅黑", 12))
button.pack(pady=5)

# 添加“选择文件”按钮
select_button = tk.Button(window, text="选择文件", command=select_file, font=("微软雅黑", 12))
select_button.pack(pady=5)

# 显示选择的文件路径
file_label = tk.Label(window, text="尚未选择文件", font=("微软雅黑", 10), wraplength=500, justify="left")
file_label.pack(pady=10)

# 启动窗口
window.mainloop()
