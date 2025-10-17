# 文件名: run.py
# 描述: 项目的命令行入口点。

import sys
from src.converter import DocxConverter

def main():
    """
    处理命令行参数并启动转换过程。
    """
    if len(sys.argv) < 2:
        print("--- Word to Jupyter 转换器 ---")
        print("\n使用方法: python run.py \"<word文档路径>\" [内核名称]")
        print("\n参数说明:")
        print("  <word文档路径> : (必需) 你的 .docx 文件路径。")
        print("  [内核名称]     : (可选) 用于执行 Notebook 的 Jupyter 内核名称。")
        print("                 如果不提供，脚本将只生成 .ipynb 文件而不执行。")
        print("\n示例:")
        print("  # 只转换，不执行")
        print("  python run.py \"Numpy入门.docx\"")
        print("\n  # 转换后，使用名为 'python3' 的内核执行")
        print("  python run.py \"Numpy入门.docx\" python3")
        sys.exit(1)

    docx_file_path = sys.argv[1]
    # 如果提供了第二个参数，则作为内核名；否则为 None
    kernel = sys.argv[2] if len(sys.argv) > 2 else None

    # 创建转换器实例并执行转换
    converter = DocxConverter(docx_path=docx_file_path, kernel_name=kernel)
    converter.convert()

if __name__ == '__main__':
    main()
