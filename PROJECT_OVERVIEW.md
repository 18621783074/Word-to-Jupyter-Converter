# Word-to-Jupyter 项目结构与逻辑总览

这是一个对当前项目结构的详细解释，旨在帮助你理解每个文件和代码块的作用。

---

## 一、项目文件结构

我们的项目目前包含以下文件和目录：

```
.
├── src/                      # 存放所有核心源代码的目录
│   ├── __init__.py           # 告诉Python这是一个“包”，使其可以被导入
│   └── converter.py          # 包含了所有转换逻辑的核心类
├── data/                     # (测试用) 存放测试数据
│   └── sample_data.csv       # (测试用) 一个示例CSV文件
├── run.py                    # 项目的启动入口，负责与用户交互
├── requirements.txt          # 列出了项目运行所需的所有第三方库
└── PROJECT_OVERVIEW.md       # (本文档) 项目的说明文档
```

### 文件职责详解

- **`run.py`**
  - **角色**: **项目的前台/启动器**
  - **职责**:
    1.  这是用户直接运行的文件 (`python run.py ...`)。
    2.  它负责解析用户在命令行输入的参数（比如Word文档的路径）。
    3.  它不关心**如何**进行转换，只关心**启动**转换过程。它通过创建 `DocxConverter` 类的实例并调用其方法来完成这个任务。

- **`src/converter.py`**
  - **角色**: **项目的后台/引擎**
  - **职责**:
    1.  包含了所有实际的转换工作。
    2.  定义了 `DocxConverter` 类，这个类是整个转换功能的核心。
    3.  它完全独立于用户界面，无论是命令行还是未来的图形界面，都可以调用它来完成工作。

- **`src/` 目录 和 `__init__.py`**
  - **角色**: **代码的组织者**
  - **职责**: `src` 目录将所有功能代码与启动脚本、文档等分离开，使项目更整洁。`__init__.py` 文件（即使是空的）是Python的规定，它告诉Python `src` 目录是一个代码“包”，允许我们从 `run.py` 中使用 `from src.converter import ...` 来导入代码。

- **`requirements.txt`**
  - **角色**: **项目的依赖清单**
  - **职责**: 告诉其他用户或打包工具（如PyInstaller）这个项目需要哪些第三方库才能运行。

---

## 二、核心逻辑：`DocxConverter` 类详解

所有转换魔法都发生在 `src/converter.py` 文件中的 `DocxConverter` 类里面。让我们深入看看它的内部构造。

### 1. `__init__(self, docx_path, kernel_name=None)` - 构造函数

- **作用**: 当我们创建一个转换器实例时（`converter = DocxConverter(...)`），这个函数会被自动调用。
- **逻辑**:
  1.  接收传入的 `docx_path` (Word文档路径) 和可选的 `kernel_name` (Jupyter内核名)。
  2.  将这些信息保存为实例的**属性**（`self.docx_path`, `self.kernel_name`），这样类中的其他方法就可以随时访问它们了。
  3.  根据输入的路径，预先计算好输出文件的路径 (`self.output_path`) 等信息。
  4.  初始化一个空的列表 `self.parsed_blocks`，用于存放从Word中解析出来的代码块和文本块。

### 2. `convert(self)` - 公共主方法

- **作用**: 这是这个类的**唯一**公共接口，是整个转换流程的“总开关”。
- **逻辑**:
  1.  **打印开始信息**。
  2.  **加载Word文档**: 使用 `docx.Document()` 打开文件。
  3.  **调用内部方法**: 按顺序调用其他内部（以`_`开头）方法来执行具体的步骤：
     - `self._parse_paragraphs()`: 解析文档。
     - `self._process_code_paths()`: 修正代码中的路径。
     - `self._create_notebook()`: 创建`.ipynb`文件。
     - `self._run_notebook()`: (如果需要) 执行Notebook。
  4.  **打印结束信息**。

### 3. 内部方法 (以 `_` 开头)

这些方法被设计为类的“内部零件”，不应该被外部直接调用。`convert()` 方法负责按正确的顺序调度它们。

- **`_parse_paragraphs(self, paragraphs)`**:
  - **职责**: 遍历Word文档的所有段落，将它们解析成代码块和Markdown块，并存入 `self.parsed_blocks` 列表中。
  - **调用**: 内部会调用 `_sanitize_text()` 和 `_is_likely_python_code()`。

- **`_is_likely_python_code(self, text)`**:
  - **职责**: 核心识别引擎。通过一系列规则（关键字、符号、黑名单等）判断一段文本是不是Python代码。

- **`_process_code_paths(self)`**:
  - **职责**: 遍历 `self.parsed_blocks` 中所有的代码块，并对每一个代码块调用 `_process_single_code_block()` 进行处理。

- **`_process_single_code_block(self, code_content)`**:
  - **职责**: 使用正则表达式查找单个代码块中的文件路径，并调用 `_find_file_in_project()` 来获取真实路径，最后完成替换。

- **`_create_notebook(self)`**:
  - **职责**: 根据 `self.parsed_blocks` 的内容，使用 `nbformat` 库生成一个Jupyter Notebook对象，并将其写入到 `self.output_path` 指定的文件中。

- **`_run_notebook(self)`**:
  - **职责**: 使用 `subprocess` 模块调用系统的 `jupyter nbconvert` 命令来执行刚刚创建的Notebook文件。

- **其他辅助方法**:
  - `_sanitize_text()`: 文本消毒。
  - `_find_file_in_project()`: 在项目中查找文件。

---

## 三、整体工作流程 (逻辑流)

1.  用户在终端运行 `python run.py "我的文档.docx"`。
2.  `run.py` 的 `main()` 函数被触发。
3.  `main()` 函数创建一个 `DocxConverter` 的实例，并将 `"我的文档.docx"` 作为参数传进去。
4.  `DocxConverter` 的 `__init__()` 方法被调用，初始化所有必要的路径和变量。
5.  `main()` 函数调用 `converter.convert()` 方法。
6.  `convert()` 方法开始按顺序执行内部的 `_parse...`, `_process...`, `_create...` 等方法，一步步完成从解析到生成文件的所有操作。
7.  流程结束。

希望这份文档能帮助你彻底理清我们项目的结构和逻辑！