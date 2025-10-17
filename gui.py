# 文件名: gui.py
# 描述: 项目的图形用户界面(GUI)入口点，包含预览与微调功能。

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from src.converter import DocxConverter
import threading

class PreviewWindow(tk.Toplevel):
    """一个新窗口，用于预览、编辑和确认转换的内容块。"""
    def __init__(self, parent, app_instance, converter, blocks):
        super().__init__(parent)
        self.app = app_instance
        self.converter = converter
        self.blocks_data = blocks
        self.block_widgets = []

        self.title("预览与微调")
        self.geometry("800x600")

        # --- 创建可滚动的画布 ---
        main_frame = tk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=1)

        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # --- 填充内容块 ---
        self.populate_blocks()

        # --- 底部确认按钮 ---
        bottom_frame = tk.Frame(self)
        bottom_frame.pack(fill=tk.X, pady=10)
        
        confirm_button = tk.Button(bottom_frame, text="确认并开始转换", command=self.confirm_conversion, font=("", 12, "bold"))
        confirm_button.pack(pady=5)

    def populate_blocks(self):
        """动态创建每个内容块的UI控件。"""
        for i, block in enumerate(self.blocks_data):
            row_frame = tk.Frame(self.scrollable_frame, borderwidth=1, relief="solid", pady=5)
            row_frame.pack(fill=tk.X, padx=5, pady=2)

            # --- 左侧控制区 ---
            control_frame = tk.Frame(row_frame)
            control_frame.pack(side=tk.LEFT, padx=5)

            var = tk.BooleanVar(value=True)
            cb = tk.Checkbutton(control_frame, variable=var)
            cb.pack()

            type_label = tk.Label(control_frame, text=f"[{block['type'].upper()}]", fg="blue" if block['type'] == 'code' else "black")
            type_label.pack()
            
            toggle_button = tk.Button(control_frame, text="切换类型", 
                                      command=lambda b=block, lbl=type_label: self.toggle_type(b, lbl))
            toggle_button.pack()

            # --- 右侧内容预览区 ---
            content_text = scrolledtext.ScrolledText(row_frame, height=4, wrap=tk.WORD, font=("Courier New", 9))
            content_text.insert(tk.END, block['content'])
            # 解除文本框的禁用状态，使其可编辑
            # content_text.config(state='disabled')
            content_text.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            
            # 保存文本控件本身，以便后续读取其内容
            self.block_widgets.append({'var': var, 'block': block, 'label': type_label, 'text_widget': content_text})

    def toggle_type(self, block, label):
        """切换一个块的类型 (code/markdown)。"""
        if block['type'] == 'code':
            block['type'] = 'markdown'
            label.config(text="[MARKDOWN]", fg="black")
        else:
            block['type'] = 'code'
            label.config(text="[CODE]", fg="blue")

    def confirm_conversion(self):
        """收集用户的选择并启动最终的转换线程。"""
        final_blocks = []
        for item in self.block_widgets:
            if item['var'].get(): # 检查复选框是否被选中
                # 从文本控件中获取最新的内容
                updated_content = item['text_widget'].get("1.0", tk.END).strip()
                # 更新block字典中的内容
                item['block']['content'] = updated_content
                # 将更新后的block添加到列表中
                final_blocks.append(item['block'])
        
        if not final_blocks:
            messagebox.showwarning("无内容", "您没有选择任何内容块进行转换。")
            return

        # 在主App中启动转换线程
        self.app.run_final_conversion_thread(self.converter, final_blocks)
        self.destroy() # 关闭预览窗口

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Word to Jupyter 转换器")
        self.root.geometry("500x250")

        self.file_path = tk.StringVar()
        self.kernel_name = tk.StringVar()
        self.status_text = tk.StringVar(value="请选择一个 .docx 文件")

        # --- 界面布局 ---
        file_frame = tk.Frame(self.root, pady=10)
        file_frame.pack()
        tk.Label(file_frame, text="Word 文档:").pack(side=tk.LEFT, padx=5)
        tk.Entry(file_frame, textvariable=self.file_path, width=40, state='readonly').pack(side=tk.LEFT)
        tk.Button(file_frame, text="选择文件", command=self.select_file).pack(side=tk.LEFT, padx=5)

        kernel_frame = tk.Frame(self.root, pady=5)
        kernel_frame.pack()
        tk.Label(kernel_frame, text="内核名称 (可选):").pack(side=tk.LEFT, padx=5)
        tk.Entry(kernel_frame, textvariable=self.kernel_name, width=30).pack(side=tk.LEFT)

        preview_button = tk.Button(self.root, text="解析并预览", command=self.start_parsing, font=("", 12, "bold"), pady=10)
        preview_button.pack(pady=20)

        status_label = tk.Label(self.root, textvariable=self.status_text, fg="grey")
        status_label.pack(side=tk.BOTTOM, fill=tk.X, ipady=5)

    def select_file(self):
        path = filedialog.askopenfilename(title="请选择一个 Word 文档", filetypes=[("Word 文档", "*.docx")])
        if path:
            self.file_path.set(path)
            self.status_text.set(f"已选择: {os.path.basename(path)}")

    def start_parsing(self):
        docx_path = self.file_path.get()
        if not docx_path:
            messagebox.showwarning("警告", "请先选择一个 Word 文档！")
            return
        
        kernel = self.kernel_name.get() or None
        
        # 在新线程中运行解析，防止UI冻结
        thread = threading.Thread(target=self.run_parsing_thread, args=(docx_path, kernel))
        thread.start()

    def run_parsing_thread(self, docx_path, kernel):
        """后台线程：仅解析文档，然后打开预览窗口。"""
        try:
            self.status_text.set("正在解析文档，请稍候...")
            converter = DocxConverter(
                docx_path=docx_path,
                kernel_name=kernel,
                log_callback=lambda msg: self.status_text.set(msg),
                error_callback=lambda err: messagebox.showerror("解析错误", err)
            )
            blocks = converter.parse_document()

            if blocks is not None:
                self.status_text.set("解析完成，请在预览窗口中微调。")
                # 在主线程中创建Toplevel窗口
                self.root.after(0, lambda: PreviewWindow(self.root, self, converter, blocks))
            else:
                self.status_text.set("解析失败或文档为空。")

        except Exception as e:
            self.status_text.set(f"❌ 解析失败: {e}")
            messagebox.showerror("错误", f"发生未知错误:\n{e}")

    def run_final_conversion_thread(self, converter, final_blocks):
        """后台线程：根据最终确认的blocks创建notebook。"""
        try:
            self.status_text.set("正在生成 Notebook，请稍候...")
            converter.create_notebook_from_blocks(final_blocks)
            self.status_text.set("✅ 转换成功！.ipynb 文件已保存在原目录下。")
            messagebox.showinfo("成功", "转换成功！")
        except Exception as e:
            self.status_text.set(f"❌ 转换失败: {e}")
            messagebox.showerror("错误", f"发生错误:\n{e}")

if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()