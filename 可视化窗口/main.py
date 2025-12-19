import tkinter as tk

def on_click():
    label.config(text="你点击了按钮！")

# 创建主窗口
window = tk.Tk()
window.title("我的脚本工具")
window.geometry("600x400+300+100")  # 设置窗口大小：宽 x 高

# 添加标签
label = tk.Label(window, text="欢迎使用我的脚本工具")
label.pack(pady=20)

# 添加按钮
button = tk.Button(window, text="点击我", command=on_click)
button.pack()

# 启动窗口事件循环（必须）
window.mainloop()
