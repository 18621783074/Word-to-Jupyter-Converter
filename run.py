# 文件名: run.py
# 描述: 项目的命令行入口点。

import sys
import tkinter as tk
from tkinter import ttk
from src.converter import DocxConverter
from gui import App

def main_cli():
    """
    处理命令行参数并启动转换过程。
    """
    if len(sys.argv) < 2:
        print("--- Word to Jupyter 转换器 (命令行模式) ---")
        print("\n使用方法: run.exe \"<word文档路径>\" [内核名称]")
        print("\n参数说明:")
        print("  <word文档路径> : (必需) 你的 .docx 文件路径。")
        print("  [内核名称]     : (可选) 用于执行 Notebook 的 Jupyter 内核名称。")
        print("                 如果不提供，脚本将只生成 .ipynb 文件而不执行。")
        print("\n注意: 不带参数运行 run.exe 将启动图形用户界面。")
        sys.exit(1)

    docx_file_path = sys.argv[1]
    kernel = sys.argv[2] if len(sys.argv) > 2 else None
    
    converter = DocxConverter(docx_path=docx_file_path, kernel_name=kernel)
    converter.convert()

def main_gui():
    """
    启动图形用户界面。
    """
    root = tk.Tk()
    style = ttk.Style(root)
    try:
        style.theme_use("arc")
    except tk.TclError:
        print("Note: 'arc' theme not available, using default theme.")
    app = App(root)
    root.mainloop()

if __name__ == '__main__':
    # 如果有命令行参数，则进入CLI模式
    if len(sys.argv) > 1:
        main_cli()
    # 否则，启动GUI
    else:
        main_gui()
