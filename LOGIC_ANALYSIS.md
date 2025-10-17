# `_is_likely_python_code` 函数深度逻辑分析

本文档旨在详细拆解 `_is_likely_python_code` 函数的内部工作原理。该函数是整个转换工具的智能核心，负责判断任意一段文本更可能是 Python 代码还是普通的 Markdown 文本。

## 核心思想：多层过滤的“安检”模型

您可以将这个函数想象成一个机场的安检流程。每一段文本都是一位“旅客”，需要通过多道安检门。安检的目标是识别出“代码旅客”，并将他们引导至正确的登机口。

-   **快速通道**: 如果一位旅客持有“代码”的头等舱机票（例如包含 `import` 关键字），他会立刻通过安检。
-   **拦截通道**: 如果一位旅客被发现携带了违禁品（例如是一个黑名单单词），他会被立刻拦下，判定为“非代码”。
-   **特殊审查**: 如果一位旅客情况复杂（例如中英文混合），他会被带到“小黑屋”进行更详细的审查。
-   **常规安检**: 大部分旅客需要通过标准的金属探测门和X光机（正则表达式结构扫描）。
-   **默认出口**: 如果一位旅客通过了所有安检门，但身上没有任何“代码”特征，他将从默认的“普通旅客”出口离开。

---

## 详细逻辑流程拆解

以下是函数内部（位于 [`src/converter.py:101-134`](src/converter.py:101)）的逐行工作逻辑：

### 第1道安检门：基础清洁与预检

```python
# src/converter.py:102-103
text_stripped = text.strip()
if not text_stripped: return False
```
-   **目的**: 清理文本，排除无效输入。
-   **逻辑**:
    1.  `text.strip()`: 首先，去掉文本两端的任何空白字符（空格、换行、制表符等）。
    2.  `if not text_stripped`: 如果清理后，文本变成了空字符串，那么它不可能是代码，直接判定为 `False`（非代码）。

### 第2道安检门：黑名单快速拦截

```python
# src/converter.py:104-105
BLACKLIST_SINGLE_WORDS = {'Jupyter', 'Python', ...}
if text_stripped in BLACKLIST_SINGLE_WORDS: return False
```
-   **目的**: 快速排除那些极易被误判为代码的、作为专有名词的单个单词。
-   **逻辑**: 如果整段文本**完全等于**黑名单中的某一个词（例如，某一段的全部内容就是 "Python"），则立即判定为 `False`。这能有效防止文档中的标题或关键词被误判为代码变量。

### 第3道安检门：中英文混合“特别审查室”

这是整个函数中最复杂、最精细的部分，专门处理包含中文字符的文本。

```python
# src/converter.py:106
if re.search(r'[\u4e00-\u9fa5]', text_stripped):
    # ... 进入特别审查流程 ...
```
-   **逻辑**: 首先，检查文本中是否**含有任何**中文字符。如果**不含**，则直接跳过整个审查室，进入第4道安检门。如果**含有**，则启动以下复杂流程：

    1.  **创建“无字符串”版本**:
        ```python
        # src/converter.py:107
        text_no_strings = re.sub(r'(\'.*?\')|(\".*?\")', '', text_stripped)
        ```
        -   **目的**: 判断中文是出现在代码的“裸露”部分，还是被安全地包裹在字符串内部。
        -   **逻辑**: 创建一个临时副本，用正则表达式挖掉所有被单引号或双引号包裹的字符串内容。

    2.  **检查“裸露”中文**:
        ```python
        # src/converter.py:108
        if re.search(r'[\u4e00-\u9fa5]', text_no_strings):
            # ... 发现裸露中文，进入最终审查 ...
        else:
            # 所有中文都在字符串里，这是合法的
            return True 
        ```
        -   **逻辑**: 检查“无字符串”版本是否**仍然**含有中文。
            -   如果**不含**，说明所有中文都安全地在字符串里（例如 `print("你好")`），这是合法的 Python 语法，直接判定为 `True`（代码）。
            -   如果**含有**，说明中文出现在了代码的裸露部分，情况可疑，需要进入最终审查。

    3.  **最终审查：代码与中文注释混合体**
        ```python
        # src/converter.py:109-121
        if '#' in text_no_strings:
            # ... 逻辑 ...
            return True # or False
        else:
            # 含有裸露中文，但又不是注释，判定为非代码
            return False
        ```
        -   **逻辑**: 这是对含有“裸露”中文的文本的最后机会。
            -   `if '#' in text_no_strings`: 检查是否存在注释符号 `#`。
                -   如果**存在**，程序会认为这可能是一个“代码 + 中文注释”的组合（例如 `x = 1 # 这是注释`）。它会进一步分析 `#` 前面的部分是否像无害的代码。如果是，就判定为 `True`。如果 `#` 前面的部分看起来像一句长中文，就判定为 `False`。
                -   如果**不存在** `#`，那么程序就认为这些裸露的中文是非法的（在 Python 中，裸露的非 ASCII 字符通常会导致语法错误），因此直接判定为 `False`。

### 第4道安检门：Shell 命令排除

```python
# src/converter.py:123-124
shell_commands = ['cd', 'ls', 'pip', ...]
if text_stripped.split() and text_stripped.split()[0] in shell_commands: return False
```
-   **目的**: 排除那些在教程中常见的、容易与 Python 代码混淆的命令行操作。
-   **逻辑**: 检查文本的第一个单词是否是 `shell_commands` 列表中的一员。如果是，判定为 `False`。

### 第5道安检门：关键字与常见模式“头等舱通道”

```python
# src/converter.py:126-130
python_patterns = [r'\bimport\s+', r'\bfrom\s+', ...]
if any(re.search(pattern, text_stripped) for pattern in python_patterns): return True
```
-   **目的**: 这是最强有力的代码识别依据。如果文本中出现了明确的 Python 关键字或常用库的特征用法，那么它**极有可能**是代码。
-   **逻辑**: 使用一组正则表达式，检查文本中是否**包含**任何一个模式（例如 `import `、`def `、`np.`、`pd.` 等）。只要有一个模式匹配成功，就立刻判定为 `True`，安检结束。

### 第6道安检门：代码结构“全身扫描”

```python
# src/converter.py:131-133
if re.match(r'^[a-zA-Z_]...(=|\+=|...).*$', text_stripped): return True
if re.match(r'^\s*[\w\d\.\_]+...[\+\-\*\/]...$', text_stripped): return True
if re.match(r'^[a-zA-Z_][\w\d\._\[\]\'\"]*$', text_stripped): return True
```
-   **目的**: 作为最后的防线，捕捉那些没有明显关键字，但具有典型代码“长相”的文本。
-   **逻辑**: 通过三个正则表达式，匹配几种常见的代码结构：
    1.  **赋值/调用/索引语句**: 匹配以变量名开头，后跟 `=`、`+=`、`[`、`(` 或 `.` 的行。
    2.  **数学运算**: 匹配形如 `a * b` 或 `var1 / var2` 的简单数学表达式。
    3.  **单个变量/属性访问**: 匹配看起来像一个单独的变量名或属性访问的行（例如 `my_variable` 或 `my_df['column']`）。

### 最终出口：默认判定为“非代码”

```python
# src/converter.py:134
return False
```
-   **逻辑**: 如果一段文本（一位“旅客”）经历了以上所有安检流程，但都没有被任何一条规则判定为 `True`，那么它最终会从这个默认出口离开，被判定为 `False`（非代码/Markdown）。