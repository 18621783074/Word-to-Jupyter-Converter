# -*- coding: utf-8 -*-
"""
Word to Jupyter Converter - Core Logic
Description: Contains the main DocxConverter class that handles parsing the
             Word document, identifying code/markdown blocks, and generating
             the Jupyter Notebook.
Author: Stella
"""
import os
import re
import sys
import subprocess
from pathlib import Path

import docx
import nbformat
from nbformat.v4 import new_notebook, new_code_cell, new_markdown_cell


class DocxConverter:
    """
    Handles the core logic of converting a .docx file to a .ipynb file.
    """
    def __init__(self, docx_path, kernel_name=None, log_callback=None, error_callback=None):
        self.docx_path = docx_path
        self.kernel_name = kernel_name
        self.docx_dir = os.path.dirname(os.path.abspath(docx_path))
        self.output_filename = os.path.splitext(os.path.basename(docx_path))[0] + '.ipynb'
        self.output_path = os.path.join(self.docx_dir, self.output_filename)
        self.parsed_blocks = []

        # Setup callbacks for logging and errors to communicate with the GUI
        self.log = log_callback if log_callback else print
        self.error = error_callback if error_callback else lambda msg: print(f"ERROR: {msg}")

    def parse_document(self):
        """
        Step 1 & 2: Parse the Word document, correct file paths in code,
        and return a list of content blocks for the GUI to preview.
        """
        self.log("--- Starting Document Analysis ---")
        try:
            doc = docx.Document(self.docx_path)
        except Exception as e:
            self.error(f"Could not open Word document '{self.docx_path}'.\nDetails: {e}")
            return None

        # 1. Intelligently parse paragraphs into blocks
        self._parse_paragraphs(doc.paragraphs)

        if not self.parsed_blocks:
            self.log("No valid content found in the document.")
            return []

        # 2. Process code blocks to correct hardcoded file paths
        self._process_code_paths()

        self.log("Document analysis complete.")
        return self.parsed_blocks

    def create_notebook_from_blocks(self, blocks):
        """
        Step 3 & 4: Create and optionally execute a Notebook from a given
        list of content blocks (potentially modified by the user in the GUI).
        """
        if not blocks:
            self.log("No content blocks provided; cannot create Notebook.")
            return False, "NO_BLOCKS"

        self.parsed_blocks = blocks
        self.log(f"\n--- Creating Notebook from {len(self.parsed_blocks)} selected blocks ---")

        # 3. Create the .ipynb file
        if not self._create_notebook():
            return False, "CREATE_FAILED"

        # 4. Execute the notebook if a kernel name is provided
        execution_status = "NOT_RUN"
        if self.kernel_name:
            execution_status = self._run_notebook()

        self.log(f"\n--- Task Complete ---")
        self.log(f"Final file '{self.output_filename}' has been saved in the same directory as the Word document.")
        return True, execution_status

    def run_conversion(self):
        """
        Executes the full conversion process. Provided for compatibility
        with non-GUI interfaces like a command-line script.
        """
        self.log(f"--- Task Started: Processing '{os.path.basename(self.docx_path)}' ---")

        parsed_blocks = self.parse_document()

        if parsed_blocks is not None:
            self.create_notebook_from_blocks(parsed_blocks)

    def _sanitize_text(self, text):
        """
        Cleans text extracted from Word by replacing special whitespace
        characters and removing invisible control characters.
        """
        # Replace non-breaking spaces (U+00A0) with standard spaces
        text = text.replace('\u00A0', ' ')
        # Replace vertical tabs (soft newlines) with standard newlines
        text = text.replace('\v', '\n')
        # Replace carriage returns with standard newlines
        text = text.replace('\r', '\n')
        # Remove other problematic control characters
        return re.sub(r'[\x00-\x08\x0c\x0e-\x1f]', '', text)

    def _find_file_in_project(self, filename):
        """ Recursively searches for a file within the project directory. """
        start_path = Path(self.docx_dir).resolve()
        found_files = list(start_path.rglob(filename))
        if found_files:
            # Return the path in a consistent, cross-platform (POSIX) format
            return found_files[0].resolve().as_posix()
        return None

    def _process_code_paths(self):
        """ Iterates through code blocks and applies path correction. """
        self.log("\n--- Starting Smart Path Correction ---")
        for block in self.parsed_blocks:
            if block['type'] == 'code':
                block['content'] = self._process_single_code_block(block['content'])
        self.log("--- Path Correction Finished ---")

    def _process_single_code_block(self, code_content):
        """
        Uses regex to find and replace file paths in a single code block.
        """
        # This regex finds common data loading functions and captures the file path.
        pattern = re.compile(
            r"""
            # Group 1: The function call (e.g., pd.read_csv, np.load, open)
            ( \b(?:pd|pandas)\.read_\w+\b | \b(?:np|numpy)\.(?:loadtxt|load|genfromtxt)\b | \bopen\b )
            \s*\( \s*
            (?:[\w\d_]+\s*=\s*)? # Optional keyword argument (e.g., filepath=)
            (['"])               # Group 2: The opening quote (' or ")
            (.*?)                # Group 3: The actual path (non-greedy)
            \2                   # Matches the same closing quote
            """,
            re.VERBOSE | re.IGNORECASE
        )
        modified_code = code_content
        # Iterate backwards to avoid messing up indices of subsequent matches
        for match in reversed(list(pattern.finditer(code_content))):
            original_full_match = match.group(0)
            original_path = match.group(3)
            filename = os.path.basename(original_path)

            if not filename or filename == ".":
                continue

            self.log(f"Info: Found file reference: '{filename}' (Original path: '{original_path}')")
            real_path = self._find_file_in_project(filename)

            if real_path:
                # Ensure path uses forward slashes for consistency in Python strings
                safe_real_path = real_path.replace('\\', '/')
                new_full_match = original_full_match.replace(original_path, safe_real_path, 1)
                start, end = match.span()
                modified_code = modified_code[:start] + new_full_match + modified_code[end:]
                self.log(f"Success: Auto-corrected path -> '{safe_real_path}'")
            else:
                self.log(f"Warning: Could not find '{filename}' in the project directory. Path will be left as is.")
        return modified_code

    def _is_likely_python_code(self, text):
        """
        A heuristic-based classifier to determine if a line of text is code.

        Returns:
            A tuple: (is_code, is_uncertain)
            - (True, False): Confidently code.
            - (True, True):  Likely code, but uncertain.
            - (False, False): Confidently markdown.
        """
        text_stripped = text.strip()
        if not text_stripped:
            return (False, False)

        # High-confidence keywords and patterns (definitely code)
        high_confidence_patterns = [
            r'\bimport\s+', r'\bfrom\s+', r'\bdef\s+', r'\bclass\s+',
            r'print\s*\(', r'np\.', r'pd\.', r'plt\.'
        ]
        if any(re.search(pattern, text_stripped) for pattern in high_confidence_patterns):
            return (True, False)

        # High-confidence markdown patterns (definitely not code)
        if text_stripped.startswith(('#', '##', '###')):
             return (False, False)
        BLACKLIST_SINGLE_WORDS = {
            'Jupyter', 'Python', 'Numpy', 'Pandas', 'Matplotlib', 'Scipy',
            'Tensorflow', 'Pytorch', 'Linux', 'Windows', 'MacOS'
        }
        if text_stripped in BLACKLIST_SINGLE_WORDS:
            return (False, False)
        
        # Handle lines with Chinese characters
        if re.search(r'[\u4e00-\u9fa5]', text_stripped):
            text_no_strings = re.sub(r'(\'.*?\')|(\".*?\")', '', text_stripped)
            if re.search(r'[\u4e00-\u9fa5]', text_no_strings):
                if '#' in text_no_strings:
                    parts = text_no_strings.split('#', 1)
                    code_part = parts[0].strip()
                    if code_part:
                        return (True, False) # Code with Chinese comment
                return (False, False) # Chinese is in the code part, not a comment
            return (True, True) # Chinese only in string literals is uncertain

        # Exclude common shell commands
        shell_commands = ['cd', 'ls', 'pip', 'conda', 'ipython', 'jupyter', 'mkdir']
        if text_stripped.split() and text_stripped.split()[0] in shell_commands:
            return (False, False)
            
        # Medium-confidence patterns (likely code, but could be ambiguous)
        medium_confidence_patterns = [
            r'\bif\s+', r'\bfor\s+', r'\bwhile\s+', r'\breturn\b',
            r'^[a-zA-Z_][\w\.]*\s*(=|\+=|-=|\*=|/=|\[|\(|\.).*$', # Assignments / calls
            r'^\s*[\w\d\._]+\s*[\+\-\*\/%]+\s*[\w\d\._]+\s*$', # Simple arithmetic
            r'^[a-zA-Z_][\w\d\._\[\]\'\"]*$' # A single variable name
        ]
        if any(re.search(pattern, text_stripped) for pattern in medium_confidence_patterns):
            # A single word could just be a name in a sentence.
            if ' ' not in text_stripped and '.' not in text_stripped:
                return (True, True)
            return (True, True)

        return (False, False)

    def _parse_paragraphs(self, paragraphs):
        """
        Parses paragraphs from the docx file, intelligently grouping them
        into code and markdown blocks.
        """
        self.log("\n--- Parsing Word Document Content ---")
        self.parsed_blocks = []  # Clear previous results
        block_buffer = []
        current_block_type = None

        for para in paragraphs:
            sanitized_text = self._sanitize_text(para.text)
            processed_text = re.sub(r'([\)\]\}])([\u4e00-\u9fa5])', r'\1\n\2', sanitized_text)
            lines = processed_text.split('\n')

            for line in lines:
                text = line.strip()
                if not text or text.lower() == 'copy':
                    continue

                is_code, is_uncertain = self._is_likely_python_code(text)
                para_type = 'code' if is_code else 'markdown'

                if current_block_type is None:
                    current_block_type = para_type

                if para_type == current_block_type:
                    block_buffer.append({'text': text, 'uncertain': is_uncertain})
                else:
                    if block_buffer:
                        block_is_uncertain = any(item['uncertain'] for item in block_buffer)
                        content = '\n'.join(item['text'] for item in block_buffer)
                        self.parsed_blocks.append({
                            'type': current_block_type,
                            'content': content,
                            'uncertain': block_is_uncertain if current_block_type == 'code' else False
                        })
                    current_block_type = para_type
                    block_buffer = [{'text': text, 'uncertain': is_uncertain}]
        
        if block_buffer:
            block_is_uncertain = any(item['uncertain'] for item in block_buffer)
            content = '\n'.join(item['text'] for item in block_buffer)
            self.parsed_blocks.append({
                'type': current_block_type,
                'content': content,
                'uncertain': block_is_uncertain if current_block_type == 'code' else False
            })

        self.log(f"Parsing complete. Generated {len(self.parsed_blocks)} content blocks.")

    def _create_notebook(self):
        """ Creates a .ipynb file from the parsed blocks. """
        nb = new_notebook()
        for block in self.parsed_blocks:
            if block['type'] == 'code':
                nb.cells.append(new_code_cell(block['content']))
            elif block['type'] == 'markdown':
                nb.cells.append(new_markdown_cell(block['content']))
        try:
            with open(self.output_path, 'w', encoding='utf-8') as f:
                nbformat.write(nb, f)
            self.log(f"✅ Notebook created successfully: '{self.output_path}'")
            return True
        except Exception as e:
            self.error(f"Failed to write Notebook file: {e}")
            return False

    def _run_notebook(self):
        """ Executes the created Notebook using nbconvert. """
        self.log("\n--- Executing Notebook ---")
        self.log(f"Using kernel: {self.kernel_name}")
        command = [
            'jupyter', 'nbconvert',
            '--to', 'notebook', '--execute', '--inplace',
            f'--ExecutePreprocessor.kernel_name={self.kernel_name}',
            '--allow-errors', self.output_path
        ]
        self.log(f"Executing command: {' '.join(command)}")
        try:
            # Hide console window on Windows when running from a bundled exe
            startupinfo = None
            if sys.platform == 'win32':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

            result = subprocess.run(
                command, capture_output=True, text=True, check=False,
                encoding='utf-8', errors='ignore', startupinfo=startupinfo
            )

            if "Kernel not found" in result.stderr:
                self.error(f"Kernel '{self.kernel_name}' not found!\n{result.stderr}")
                return f"KERNEL_NOT_FOUND:{self.kernel_name}"

            if result.returncode == 0:
                self.log("✅ Notebook executed successfully.")
                return "SUCCESS"
            else:
                self.log("⚠️ Errors occurred during Notebook execution (see output file for details).")
                return "EXECUTION_ERROR"

        except FileNotFoundError:
            self.error("'jupyter' command not found. Please ensure Jupyter is installed and in your system's PATH.")
            return "JUPYTER_NOT_FOUND"
        except Exception as e:
            self.error(f"An unknown error occurred during execution: {e}")
            return f"UNKNOWN_ERROR:{e}"