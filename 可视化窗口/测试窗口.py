import tkinter as tk

def check_size():
    window.update_idletasks()  # 更新窗口尺寸信息
    print("实际窗口宽度:", window.winfo_width())
    print("实际窗口高度:", window.winfo_height())
    print("屏幕分辨率:", window.winfo_screenwidth(), "x", window.winfo_screenheight())

window = tk.Tk()
window.title("窗口尺寸检测")
window.geometry("600x400+100+200")

# 延迟0.5秒打印实际窗口大小
window.after(500, check_size)

window.mainloop()
