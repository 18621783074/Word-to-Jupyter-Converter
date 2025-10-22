# Word to Jupyter Converter

这是一个智能工具，旨在将 Microsoft Word (.docx) 文档无缝转换为功能齐全的 Jupyter Notebook (.ipynb) 文件。它特别适合于将包含代码片段、解释性文本和教程的 Word 文档快速迁移到可执行的 Notebook 环境中。

本工具提供了一个简洁的图形用户界面（GUI）和一个命令行接口（CLI），以满足不同用户的需求。

<p align="center">
  <img src="./screenshots/屏幕截图 2025-10-22 140605.png" alt="应用主界面" width="49%">
  <img src="./screenshots/屏幕截图 2025-10-22 140700.png" alt="预览与调整窗口" width="49%">
</p>


## ✨ 主要功能

- **智能内容识别**: 自动分析 Word 文档中的段落，区分代码块和 Markdown 文本。
- **智能路径修正**: 自动检测代码中（如 `pd.read_csv("data.csv")`）的相对文件路径，并将其修正为能被 Notebook 正确执行的绝对路径，解决了因路径错误导致代码无法运行的常见问题。
- **图形用户界面 (GUI)**: 提供一个直观易用的界面，只需几次点击即可完成转换。
- **预览与调整**: 在最终转换前，提供一个强大的预览窗口，允许您：
    - **编辑**每个代码或文本块的内容。
    - **重新分类**内容块（例如，将误判为文本的代码更正为代码块）。
    - **勾选/取消勾选**某些内容块，只转换您需要的部分。
- **不确定性高亮**: 智能高亮显示那些解析器不太确定的代码块，提醒用户进行人工复查。
- **Jupyter 内核执行**: （可选）在生成 Notebook 后，可以指定一个 Jupyter 内核来自动执行全部代码单元。
- **命令行支持**: 为高级用户和自动化流程提供了完整的命令行接口。
- **独立打包**: 可以被打包成一个独立的 `.exe` 文件，方便在没有 Python 环境的 Windows 电脑上运行。

## 🚀 安装与运行

### 从源代码运行

如果您想从源代码运行此应用，请按照以下步骤操作：

1.  **克隆仓库**
    ```bash
    git clone https://github.com/18621783074/Word-to-Jupyter-Converter.git
    ```

2.  **创建并激活虚拟环境**te
    ```bash
    # Windows
    python -m venv .venv
    .venv\Scripts\activate
    ```

3.  **安装依赖**
    ```bash
    pip install -r requirements.txt
    ```

4.  **运行应用**
    - **启动 GUI**:
      ```bash
      python run.py
      ```
    - **使用命令行**:
      ```bash
      python run.py --file "path/to/your/document.docx" --kernel python3
      ```
      使用 `python run.py --help` 查看所有可用选项。

## 📝 使用方法

### GUI 模式

1.  运行 `python run.py` 启动应用。
2.  点击“选择文件”按钮，选择一个 `.docx` 文档。
3.  (可选) 在“Jupyter 内核”输入框中，填写您希望用来执行 Notebook 的内核名称（例如 `python3`）。如果留空，则只生成文件不执行。
4.  点击“解析并预览”按钮。
5.  在弹出的预览窗口中，检查、修改、调整您的内容块。
6.  点击“确认并转换”完成操作。生成的 `.ipynb` 文件会保存在与原 Word 文档相同的目录下。

### 命令行模式

```
用法: run.py [-h] --file FILE [--kernel KERNEL]

选项:
  -h, --help            显示此帮助信息并退出
  --file FILE, -f FILE  要转换的 Word 文档的路径 (必需)
  --kernel KERNEL, -k KERNEL
                        用于执行 Notebook 的 Jupyter 内核名称 (可选)
```

## 📄 开源许可证

本项目采用 [MIT License](LICENSE) 开源。
