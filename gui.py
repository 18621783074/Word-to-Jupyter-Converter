# 文件名: gui.py
# 描述: 项目的图形用户界面(GUI)入口点。

import os
import tkinter as tk
from tkinter import filedialog, messagebox
from src.converter import DocxConverter
import threading

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Word to Jupyter 转换器")
        self.root.geometry("500x250") # 设置窗口大小

        # --- 变量 ---
        self.file_path = tk.StringVar()
        self.kernel_name = tk.StringVar()
        self.status_text = tk.StringVar()
        self.status_text.set("请选择一个 .docx 文件")

        # --- 界面布局 ---
        # 文件选择部分
        file_frame = tk.Frame(self.root, pady=10)
        file_frame.pack()
        tk.Label(file_frame, text="Word 文档:").pack(side=tk.LEFT, padx=5)
        tk.Entry(file_frame, textvariable=self.file_path, width=40, state='readonly').pack(side=tk.LEFT)
        tk.Button(file_frame, text="选择文件", command=self.select_file).pack(side=tk.LEFT, padx=5)

        # 内核名称部分
        kernel_frame = tk.Frame(self.root, pady=5)
        kernel_frame.pack()
        tk.Label(kernel_frame, text="内核名称 (可选):").pack(side=tk.LEFT, padx=5)
        tk.Entry(kernel_frame, textvariable=self.kernel_name, width=30).pack(side=tk.LEFT)

        # 转换按钮
        convert_button = tk.Button(self.root, text="开始转换", command=self.start_conversion, font=("", 12, "bold"), pady=10)
        convert_button.pack(pady=20)

        # 状态栏
        status_label = tk.Label(self.root, textvariable=self.status_text, fg="grey")
        status_label.pack(side=tk.BOTTOM, fill=tk.X, ipady=5)

    def select_file(self):
        """弹出文件选择对话框"""
        path = filedialog.askopenfilename(
            title="请选择一个 Word 文档",
            filetypes=[("Word 文档", "*.docx")]
        )
        if path:
            self.file_path.set(path)
            self.status_text.set(f"已选择: {os.path.basename(path)}")

    def start_conversion(self):
        """启动转换过程"""
        docx_path = self.file_path.get()
        if not docx_path:
            messagebox.showwarning("警告", "请先选择一个 Word 文档！")
            return

        kernel = self.kernel_name.get() or None # 如果输入框为空，则为None

        # 为了防止界面在转换过程中“假死”，我们在一个新线程中运行转换逻辑
        thread = threading.Thread(target=self.run_conversion_thread, args=(docx_path, kernel))
        thread.start()

    def run_conversion_thread(self, docx_path, kernel):
        """在后台线程中运行的转换任务"""
        try:
            self.status_text.set("正在转换中，请稍候...")
            # 注意：这里的print输出不会显示在GUI上，我们未来可以改进
            converter = DocxConverter(
                docx_path=docx_path,
                kernel_name=kernel,
                # 将GUI中的方法作为回调函数传递进去
                log_callback=lambda msg: self.status_text.set(msg),
                error_callback=lambda err: messagebox.showerror("错误", err)
            )
            converter.run_conversion() # 修正：调用正确的方法名
            self.status_text.set("✅ 转换成功！.ipynb 文件已保存在原目录下。")
            messagebox.showinfo("成功", "转换成功！")
        except Exception as e:
            self.status_text.set(f"❌ 转换失败: {e}")
            messagebox.showerror("错误", f"发生错误:\n{e}")


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()