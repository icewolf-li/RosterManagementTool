import tkinter as tk
from tkinter import messagebox

# 创建主窗口
root = tk.Tk()
root.title("Tkinter测试")
root.geometry("300x200")

# 添加标签
label = tk.Label(root, text="欢迎使用学委开发的花名册管理工具", 
                font=("SimHei", 10), wraplength=280)
label.pack(pady=20)

# 获取当前目录
import os
current_dir = os.getcwd()

dir_label = tk.Label(root, text=f"当前目录: {current_dir}", 
                   font=("SimHei", 9), wraplength=280, fg="gray")

dir_label.pack(pady=10)

# 添加按钮
button = tk.Button(root, text="测试按钮", command=lambda: messagebox.showinfo("成功", "Tkinter工作正常！"))
button.pack(pady=20)

# 启动主循环
root.mainloop()