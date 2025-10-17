# 文件名: converter.py
# 描述: 包含了从Word到Jupyter Notebook转换的核心逻辑。

import os
import re
import sys
import subprocess
from pathlib import Path

import docx
import nbformat
from nbformat.v4 import new_notebook, new_code_cell, new_markdown_cell

class DocxConverter:
    def __init__(self, docx_path, kernel_name=None, log_callback=None, error_callback=None):
        self.docx_path = docx_path
        self.kernel_name = kernel_name
        self.docx_dir = os.path.dirname(os.path.abspath(docx_path))
        self.output_filename = os.path.splitext(os.path.basename(docx_path))[0] + '.ipynb'
        self.output_path = os.path.join(self.docx_dir, self.output_filename)
        self.parsed_blocks = []
        
        # 设置回调函数
        self.log = log_callback if log_callback else print
        self.error = error_callback if error_callback else lambda msg: print(f"ERROR: {msg}")

    def run_conversion(self):
        """执行完整的转换流程"""
        self.log(f"--- 任务开始: 处理 '{os.path.basename(self.docx_path)}' ---")
        
        try:
            doc = docx.Document(self.docx_path)
        except Exception as e:
            self.error(f"无法打开Word文档 '{self.docx_path}'.\n详情: {e}")
            return

        # 1. 智能解析文档
        self._parse_paragraphs(doc.paragraphs)
        if not self.parsed_blocks:
            self.log("未在文档中找到任何有效内容。")
            return

        # 2. 对代码块进行路径修正
        self._process_code_paths()

        # 3. 创建 Notebook
        if not self._create_notebook():
            return

        # 4. 如果指定了内核，则执行 Notebook
        if self.kernel_name:
            self._run_notebook()
        
        self.log(f"\n--- 任务完成 ---")
        self.log(f"最终文件 '{self.output_filename}' 已保存在Word文档相同目录下。")

    def _sanitize_text(self, text):
        # 关键修复1：将所有非中断空格 (U+00A0) 替换为标准的空格
        text = text.replace('\u00A0', ' ')
        # 关键修复2：将所有软回车 (垂直制表符 \v) 替换为标准的换行符 \n
        text = text.replace('\v', '\n')
        # 最终修复：将所有回车符 \r (carriage return) 也替换为标准的换行符 \n
        text = text.replace('\r', '\n')
        # 然后，再移除其他残留的、我们确实不希望出现的不可见控制字符
        return re.sub(r'[\x00-\x08\x0c\x0e-\x1f]', '', text)

    def _find_file_in_project(self, filename):
        start_path = Path(self.docx_dir).resolve()
        found_files = list(start_path.rglob(filename))
        if found_files:
            return found_files[0].resolve().as_posix()
        return None

    def _process_code_paths(self):
        self.log("\n--- 开始智能路径修正 ---")
        for block in self.parsed_blocks:
            if block['type'] == 'code':
                block['content'] = self._process_single_code_block(block['content'])
        self.log("--- 路径修正结束 ---")

    def _process_single_code_block(self, code_content):
        pattern = re.compile(
            r"""
            ( \b(?:pd|pandas)\.read_\w+\b | \b(?:np|numpy)\.(?:loadtxt|load|genfromtxt)\b | \bopen\b )
            \s*\( \s* (?:[\w\d_]+\s*=\s*)? (['"]) (.*?) \2 """,
            re.VERBOSE | re.IGNORECASE
        )
        modified_code = code_content
        for match in reversed(list(pattern.finditer(code_content))):
            original_full_match = match.group(0)
            original_path = match.group(3)
            filename = os.path.basename(original_path)
            if not filename or filename == ".": continue
            self.log(f"INFO: 在代码中发现文件引用: '{filename}' (原始路径: '{original_path}')")
            real_path = self._find_file_in_project(filename)
            if real_path:
                new_full_match = original_full_match.replace(original_path, real_path, 1)
                start, end = match.span()
                modified_code = modified_code[:start] + new_full_match + modified_code[end:]
                self.log(f"SUCCESS: 路径已自动修正 -> '{real_path}'")
            else:
                self.log(f"WARNING: 未能在项目目录中找到 '{filename}'。代码中的路径将保持原样。")
        return modified_code

    def _is_likely_python_code(self, text):
        text_stripped = text.strip()
        if not text_stripped: return False
        BLACKLIST_SINGLE_WORDS = {'Jupyter', 'Python', 'Numpy', 'Pandas', 'Matplotlib', 'Scipy', 'Tensorflow', 'Pytorch', 'Linux', 'Windows', 'MacOS','or','and','not'}
        if text_stripped in BLACKLIST_SINGLE_WORDS: return False
        if re.search(r'[\u4e00-\u9fa5]', text_stripped):
            text_no_strings = re.sub(r'(\'.*?\')|(\".*?\")', '', text_stripped)
            if re.search(r'[\u4e00-\u9fa5]', text_no_strings):
                if '#' in text_no_strings:
                    parts = text_no_strings.split('#', 1)
                    code_part = parts[0].strip()
                    comment_part = parts[1] if len(parts) > 1 else ""
                    # 检查注释部分是否有中文
                    if re.search(r'[\u4e00-\u9fa5]', comment_part):
                        # 如果#号前是看起来无害的代码，就判定为代码
                        if not code_part: return True
                        if '=' not in code_part and '(' not in code_part and len(code_part.split()) > 2:
                            return False
                        return True
                else:
                    return False
            return True
        shell_commands = ['cd', 'ls', 'pip', 'conda', 'ipython', 'jupyter', 'mkdir']
        if text_stripped.split() and text_stripped.split()[0] in shell_commands: return False
        # 使用正则表达式来更健壮地匹配关键字，\s+可以匹配任何空白字符（包括word中的特殊空格）
        python_patterns = [
            r'\bimport\s+', r'\bfrom\s+', r'\bdef\s+', r'\bclass\s+', r'\bif\s+', r'\bfor\s+',
            r'\bwhile\s+', r'\breturn\b', r'print\s*\(', r'np\.', r'pd\.', r'plt\.'
        ]
        if any(re.search(pattern, text_stripped) for pattern in python_patterns): return True
        if re.match(r'^[a-zA-Z_][\w\.]*\s*(=|\+=|-=|\*=|/=|\[|\(|\.).*$', text_stripped): return True
        if re.match(r'^\s*[\w\d\.\_]+\s*[\+\-\*\/\%]{1,2}\s*[\w\d\.\_]+\s*$', text_stripped) and ' ' not in text_stripped: return True
        if re.match(r'^[a-zA-Z_][\w\d\._\[\]\'\"]*$', text_stripped): return True
        return False

    def _parse_paragraphs(self, paragraphs):
        self.log("\n--- 开始智能解析 Word 文档 ---")
        block_buffer = []
        current_block_type = None
        for para in paragraphs:
            # 预处理：将一个段落内可能存在的“代码+中文”混合体分裂
            # 这个正则表达式寻找一个右括号、右方括号或右花括号，后面紧跟着一个中文字符，
            # 并在它们之间插入一个换行符。
            sanitized_text = self._sanitize_text(para.text)
            processed_text = re.sub(r'([\)\]\}])([\u4e00-\u9fa5])', r'\1\n\2', sanitized_text)
            
            # 现在，我们按真正的换行符来处理每一行
            lines = processed_text.split('\n')

            for line in lines:
                text = line.strip()
                if not text or text.lower() == 'copy': continue
                
                para_type = 'code' if self._is_likely_python_code(text) else 'markdown'
                
                if current_block_type is None: current_block_type = para_type
                
                if para_type == current_block_type:
                    block_buffer.append(text)
                else:
                    if block_buffer:
                        content = '\n'.join(block_buffer)
                        self.parsed_blocks.append({'type': current_block_type, 'content': content})
                    current_block_type = para_type
                    block_buffer = [text]
        if block_buffer:
            content = '\n'.join(block_buffer)
            self.parsed_blocks.append({'type': current_block_type, 'content': content})
        self.log(f"解析完成，共生成 {len(self.parsed_blocks)} 个内容块。")

    def _create_notebook(self):
        nb = new_notebook()
        for block in self.parsed_blocks:
            if block['type'] == 'code':
                nb.cells.append(new_code_cell(block['content']))
            elif block['type'] == 'markdown':
                nb.cells.append(new_markdown_cell(block['content']))
        try:
            with open(self.output_path, 'w', encoding='utf-8') as f:
                nbformat.write(nb, f)
            self.log(f"✅ 成功创建 Notebook: '{self.output_path}'")
            return True
        except Exception as e:
            self.error(f"写入 Notebook 文件失败: {e}")
            return False

    def _run_notebook(self):
        self.log("\n--- 开始执行 Notebook ---")
        self.log(f"指定内核: {self.kernel_name}")
        command = [
            sys.executable, '-m', 'jupyter', 'nbconvert',
            '--to', 'notebook', '--execute', '--inplace',
            f'--ExecutePreprocessor.kernel_name={self.kernel_name}',
            '--allow-errors', self.output_path
        ]
        self.log(f"[执行命令]: {' '.join(command)}")
        try:
            result = subprocess.run(command, capture_output=True, text=True, check=False, encoding='utf-8', errors='ignore')
            if "Kernel not found" in result.stderr:
                self.error(f"内核 '{self.kernel_name}' 未找到！请检查内核名称是否正确。\n--- Jupyter 返回的错误信息 ---\n{result.stderr}")
                return
            if result.returncode == 0:
                self.log("✅ Notebook 执行成功。")
            else:
                self.log("⚠️ Notebook 执行时遇到错误（已记录在输出文件中）。")
            self.log("-" * 25)
        except FileNotFoundError:
            self.error("'jupyter' 命令未找到。请确保 Jupyter 已安装并在系统路径中。")
        except Exception as e:
            self.error(f"执行时发生未知错误: {e}")