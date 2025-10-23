import pandas as pd
import os
import tkinter as tk
from tkinter import messagebox, scrolledtext

# 版本
version = "1.0.3"

# 定义主文件夹路径
dir_path = "计应(中职升本)2501班"
name_list = ['彭恒基', '覃艳玲', '张博涵', '吴振国', '唐雅怡', '袁玉兰', '刘艾玉蕾', '雍雨轩', '周博岩', '孙齐旺', '韩善枨', '韦晓', '覃绪尧', '方兴林', '马倩汝', '曾莹', '徐柳静', '马枝文', '陈慧华', '黎嘉嘉', '李可祺', '孙培相', '黄亮唐', '刘佳雯', '杨凤洁', '黄攀峰', '劳明权', '劳高校', '姚江业', '杨树奎', '幸国梁', '韦彬明', '刘芳妤', '林炎丽', '曾祥娟', '李国成', '宋文婷', '刘东鑫', '陈金泉', '吕佳恒']
name_to_id = {'彭恒基': 2531020130101, '覃艳玲': 2531020130102, '张博涵': 2531020130103, '吴振国': 2531020130104, '唐雅怡': 2531020130105, '袁玉兰': 2531020130106, '刘艾玉蕾': 2531020130107, '雍雨轩': 2531020130108, '周博岩': 2531020130109, '孙齐旺': 2531020130110, '韩善枨': 2531020130111, '韦晓': 2531020130112, '覃绪尧': 2531020130113, '方兴林': 2531020130114, '马倩汝': 2531020130115, '曾莹': 2531020130116, '徐柳静': 2531020130117, '马枝文': 2531020130118, '陈慧华': 2531020130119, '黎嘉嘉': 2531020130120, '李可祺': 2531020130121, '孙培相': 2531020130122, '黄亮唐': 2531020130123, '刘佳雯': 2531020130124, '杨凤洁': 2531020130125, '黄攀峰': 2531020130126, '劳明权': 2531020130127, '劳高校': 2531020130128, '姚江业': 2531020130129, '杨树奎': 2531020130130, '幸国梁': 2531020130131, '韦彬明': 2531020130132, '刘芳妤': 2531020130133, '林炎丽': 2531020130134, '曾祥娟': 2531020130135, '李国成': 2531020130136, '宋文婷': 2531020130137, '刘东鑫': 2531020130138, '陈金泉': 2531020130139, '吕佳恒': 2531020130140}

# 创建班级文件夹及子文件夹
def check_excel_file():
    # 创建主文件夹（如果不存在）
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)
    # 检查并创建所有子文件夹
    for n in name_list:
        son_dir_path = os.path.join(dir_path, n)
        if not os.path.exists(son_dir_path):
            print(f"创建文件夹: {son_dir_path}")
            os.makedirs(son_dir_path)

# 修改姓名文件夹里面文件名称
def rename_file():
    if not os.path.exists(dir_path):
        print(f"主文件夹不存在：{dir_path}")
        return
    # 遍历主文件夹下的所有子文件夹
    for subfolder_name in os.listdir(dir_path):
        subfolder_path = os.path.join(dir_path, subfolder_name)
        
        # 确保是文件夹
        if not os.path.isdir(subfolder_path):
            continue
        
        # 获取对应的学号
        if subfolder_name not in name_to_id:
            print(f"警告：未找到 '{subfolder_name}' 对应的学号，跳过该文件夹")
            continue
        
        student_id = name_to_id[subfolder_name]
        
        # 获取子文件夹中的所有文件
        files = [f for f in os.listdir(subfolder_path) if os.path.isfile(os.path.join(subfolder_path, f))]
        
        # 按文件名排序（可选，确保顺序一致）
        files.sort()
        
        # 重命名文件
        for i, filename in enumerate(files, start=1):
            old_path = os.path.join(subfolder_path, filename)
            file_ext = os.path.splitext(filename)[1]  # 保留扩展名
            new_filename = f"{student_id}（{i}）{file_ext}"
            new_path = os.path.join(subfolder_path, new_filename)
            
            try:
                os.rename(old_path, new_path)
                print(f"重命名：{filename} -> {new_filename}")
            except Exception as e:
                print(f"错误：无法重命名 {filename}，原因：{e}")

# 删除空文件夹
def delete_empty_folders(callback=None):
    if not os.path.exists(dir_path):
        message = f"主文件夹不存在：{dir_path}"
        if callback:
            callback(message)
        return
    
    # 直接遍历主文件夹下的子文件夹
    deleted_count = 0
    for subfolder_name in os.listdir(dir_path):
        subfolder_path = os.path.join(dir_path, subfolder_name)
        
        # 确保是文件夹
        if not os.path.isdir(subfolder_path):
            continue
        
        # 检查文件夹是否为空
        if not os.listdir(subfolder_path):
            try:
                os.rmdir(subfolder_path)
                deleted_count += 1
                if callback:
                    callback(f"删除空文件夹：{subfolder_path}")
            except Exception as e:
                if callback:
                    callback(f"错误：无法删除 {subfolder_path}，原因：{e}")
    
    if callback:
        if deleted_count > 0:
            callback(f"\n删除完成：共删除 {deleted_count} 个空文件夹")
        else:
            callback("未找到空文件夹")


# 创建可视化界面
class RosterManagerApp:
    def __init__(self, root):
        # 设置窗口大小和标题
        root.title("花名册管理工具 v" + version)
        # 尝试设置图标，添加错误处理
        try:
            # 尝试直接使用相对路径
            root.iconbitmap("icon.ico")
        except Exception:
            try:
                # 尝试使用绝对路径
                import os
                icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
                if os.path.exists(icon_path):
                    root.iconbitmap(icon_path)
            except Exception:
                # 如果仍然失败，忽略图标设置错误，程序继续运行
                pass
        root.geometry("300x400")
        root.resizable(False, False)
        
        # 获取当前目录
        current_dir = os.getcwd()
        
        # 创建欢迎标签
        welcome_label = tk.Label(root, text="欢迎使用学委开发的花名册管理工具", 
                                font=("SimHei", 10), wraplength=280)
        welcome_label.pack(pady=10)
        
        # 创建当前目录标签
        dir_label = tk.Label(root, text=f"当前目录: {current_dir}", 
                           font=("SimHei", 9), wraplength=280, fg="gray")
        dir_label.pack(pady=5)
        
        # 创建输出文本框
        self.output_text = scrolledtext.ScrolledText(root, width=35, height=10, 
                                                   font=("SimHei", 9))
        self.output_text.pack(pady=10)
        self.output_text.config(state=tk.DISABLED)
        
        # 创建按钮框架
        button_frame = tk.Frame(root)
        button_frame.pack(pady=20)
        
        # 创建文件夹按钮
        create_button = tk.Button(button_frame, text="创建文件夹", 
                                width=20, height=2, font=("SimHei", 10),
                                command=self.create_folders)
        create_button.pack(pady=5)
        
        # 重命名文件按钮
        rename_button = tk.Button(button_frame, text="文件重命名", 
                                width=20, height=2, font=("SimHei", 10),
                                command=self.rename_files)
        rename_button.pack(pady=5)

        # 删除空文件夹按钮
        delect_button = tk.Button(button_frame, text="删除空文件夹",
                                  width=20, height=2, font=("SimHei", 10),
                                  command=self.delete_empty_folders)
        delect_button.pack(pady=5)
        
        # 退出按钮
        exit_button = tk.Button(root, text="退出", width=15, 
                              font=("SimHei", 9),
                              command=root.quit)
        exit_button.pack(pady=10)
    
    def log_message(self, message):
        """在文本框中显示消息"""
        self.output_text.config(state=tk.NORMAL)
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)  # 滚动到最后
        self.output_text.config(state=tk.DISABLED)
        
    def create_folders(self):
        """创建文件夹功能"""
        try:
            # 清空输出
            self.output_text.config(state=tk.NORMAL)
            self.output_text.delete(1.0, tk.END)
            self.output_text.config(state=tk.DISABLED)
            
            # 调用原有的创建文件夹函数
            check_excel_file(callback=self.log_message)
            messagebox.showinfo("成功", "文件夹创建成功！")
        except Exception as e:
            self.log_message(f"错误: {str(e)}")
            messagebox.showerror("错误", f"创建文件夹失败: {str(e)}")
    
    def rename_files(self):
        """重命名文件功能"""
        try:
            # 清空输出
            self.output_text.config(state=tk.NORMAL)
            self.output_text.delete(1.0, tk.END)
            self.output_text.config(state=tk.DISABLED)
            
            # 调用原有的重命名文件函数
            rename_file(callback=self.log_message)
            messagebox.showinfo("成功", "文件重命名成功！")
        except Exception as e:
            self.log_message(f"错误: {str(e)}")
            messagebox.showerror("错误", f"文件重命名失败: {str(e)}")
    
    def delete_empty_folders(self):
        """删除空文件夹功能"""
        try:
            # 清空输出
            self.output_text.config(state=tk.NORMAL)
            self.output_text.delete(1.0, tk.END)
            self.output_text.config(state=tk.DISABLED)
            
            # 调用删除空文件夹函数
            delete_empty_folders(callback=self.log_message)
            messagebox.showinfo("成功", "空文件夹清理完成！")
        except Exception as e:
            self.log_message(f"错误: {str(e)}")
            messagebox.showerror("错误", f"删除空文件夹失败: {str(e)}")

# 修改现有函数以支持回调
def check_excel_file(callback=None):
    # 创建主文件夹（如果不存在）
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)
        if callback:
            callback(f"创建主文件夹: {dir_path}")
    else:
        if callback:
            callback(f"主文件夹已存在: {dir_path}")
    
    # 检查并创建所有子文件夹
    created_count = 0
    for n in name_list:
        son_dir_path = os.path.join(dir_path, n)
        if not os.path.exists(son_dir_path):
            os.makedirs(son_dir_path)
            created_count += 1
            if callback:
                callback(f"创建文件夹: {son_dir_path}")
    
    if callback and created_count == 0:
        callback("所有学生文件夹已存在，无需创建")

def rename_file(callback=None):
    if not os.path.exists(dir_path):
        message = f"主文件夹不存在：{dir_path}"
        if callback:
            callback(message)
        return
    
    # 遍历主文件夹下的所有子文件夹
    renamed_count = 0
    skipped_count = 0
    
    for subfolder_name in os.listdir(dir_path):
        subfolder_path = os.path.join(dir_path, subfolder_name)
        
        # 确保是文件夹
        if not os.path.isdir(subfolder_path):
            continue
        
        # 获取对应的学号
        if subfolder_name not in name_to_id:
            if callback:
                callback(f"警告：未找到 '{subfolder_name}' 对应的学号，跳过该文件夹")
            skipped_count += 1
            continue
        
        student_id = name_to_id[subfolder_name]
        
        # 获取子文件夹中的所有文件
        files = [f for f in os.listdir(subfolder_path) if os.path.isfile(os.path.join(subfolder_path, f))]
        
        # 按文件名排序
        files.sort()
        
        # 重命名文件
        for i, filename in enumerate(files, start=1):
            old_path = os.path.join(subfolder_path, filename)
            file_ext = os.path.splitext(filename)[1]  # 保留扩展名
            new_filename = f"{student_id}（{i}）{file_ext}"
            new_path = os.path.join(subfolder_path, new_filename)
            
            try:
                os.rename(old_path, new_path)
                renamed_count += 1
                if callback:
                    callback(f"重命名：{filename} -> {new_filename}")
            except Exception as e:
                if callback:
                    callback(f"错误：无法重命名 {filename}，原因：{e}")
    
    if callback:
        callback(f"\n重命名完成：成功重命名 {renamed_count} 个文件，跳过 {skipped_count} 个文件夹")


if __name__ == "__main__":
    # 创建图形界面
    root = tk.Tk()
    app = RosterManagerApp(root)
    root.mainloop()