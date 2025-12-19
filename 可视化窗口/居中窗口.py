import tkinter as tk

window = tk.Tk()
window.title("居中窗口")

# 获取屏幕尺寸
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()

# 设置窗口大小
width = 600
height = 400

# 计算居中位置
x = (screen_width - width) // 2
y = (screen_height - height) // 2

# 设置大小和位置
window.geometry(f"{width}x{height}+{x}+{y}")

window.mainloop()
