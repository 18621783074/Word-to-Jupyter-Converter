# -*- coding: utf-8 -*-
"""
Word to Jupyter Converter - GUI
Description: The main graphical user interface for the application.
Author: Stella
"""
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import base64
import io

from PIL import Image, ImageTk, ImageDraw, ImageFont

from src.converter import DocxConverter
from src.assets import BACKGROUND_IMAGE_B64


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class PreviewWindow(tk.Toplevel):
    """
    A Toplevel window that allows users to preview, edit, re-classify,
    and select the parsed content blocks before final conversion.
    """
    def __init__(self, parent, app_instance, converter, blocks):
        super().__init__(parent)
        self.app = app_instance
        self.converter = converter
        self.blocks_data = blocks
        self.block_widgets = []

        self.title("预览与调整")
        self.geometry("800x600")

        # --- Style for highlighting uncertain blocks ---
        self.style = ttk.Style(self)
        self.style.configure('Uncertain.TFrame', background='#fffacd', bordercolor='orange') # Lemon Chiffon with orange border

        main_frame = tk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=1)

        self.canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self._bind_mouse_scroll(self.canvas)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.populate_blocks()

        bottom_frame = tk.Frame(self)
        bottom_frame.pack(fill=tk.X, pady=10)
        confirm_button = ttk.Button(bottom_frame, text="确认并转换", command=self.confirm_conversion, style='Accent.TButton')
        confirm_button.pack(pady=5)

    def _bind_mouse_scroll(self, widget):
        """ Binds mouse wheel events for cross-platform scrolling. """
        widget.bind("<MouseWheel>", self._on_mousewheel)  # Windows
        widget.bind("<Button-4>", self._on_mousewheel)    # Linux scroll up
        widget.bind("<Button-5>", self._on_mousewheel)    # Linux scroll down

    def _on_mousewheel(self, event):
        """ Handles cross-platform mouse wheel scrolling. """
        if hasattr(event, 'delta') and event.delta != 0:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")
    def populate_blocks(self):
        for i, block in enumerate(self.blocks_data):
            is_uncertain = block.get('uncertain', False) and block['type'] == 'code'
            
            # Determine the style for the frame
            frame_style = 'Uncertain.TFrame' if is_uncertain else 'TFrame'
            
            row_frame = ttk.Frame(self.scrollable_frame, borderwidth=1, relief="solid", padding=5, style=frame_style)
            row_frame.pack(fill=tk.X, padx=5, pady=2)
            
            control_frame = ttk.Frame(row_frame, style=frame_style) # Ensure child frame also has the style
            control_frame.pack(side=tk.LEFT, padx=5, fill='y')

            var = tk.BooleanVar(value=True)
            cb = ttk.Checkbutton(control_frame, variable=var)
            cb.pack(pady=2)

            type_label = ttk.Label(control_frame, text=f"[{block['type'].upper()}]", foreground="blue" if block['type'] == 'code' else "black")
            type_label.pack(pady=2)

            # Pass the row_frame to the toggle command
            toggle_button = ttk.Button(control_frame, text="切换类型", command=lambda b=block, lbl=type_label, rf=row_frame: self.toggle_type(b, lbl, rf))
            toggle_button.pack(pady=2)

            if is_uncertain:
                review_label = ttk.Label(control_frame, text="建议复查", foreground="#E69B00", font=("", 8, "bold"))
                review_label.pack(pady=(5,0))

            content_text = scrolledtext.ScrolledText(row_frame, height=4, wrap=tk.WORD, font=("Courier New", 9))
            content_text.insert(tk.END, block['content'])
            content_text.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

            self.block_widgets.append({
                'var': var,
                'block': block,
                'label': type_label,
                'text_widget': content_text,
                'frame': row_frame,
                'control_frame': control_frame
            })

    def toggle_type(self, block, label, frame):
        """ Toggles the block type and removes the 'uncertain' highlight. """
        if block['type'] == 'code':
            block['type'] = 'markdown'
            label.config(text="[MARKDOWN]", foreground="black")
        else:
            block['type'] = 'code'
            label.config(text="[CODE]", foreground="blue")

        # User has made a decision, so it's no longer uncertain.
        if block.get('uncertain', False):
            block['uncertain'] = False
            # Find and remove the "Review!" label if it exists
            for widget in frame.winfo_children():
                if isinstance(widget, ttk.Frame): # Find control_frame
                    for child in widget.winfo_children():
                         if isinstance(child, ttk.Label) and "建议复查" in child.cget("text"):
                            child.destroy()
            # Reset the style to default
            frame.config(style='TFrame')
            # The child frame also needs its style reset
            for widget in frame.winfo_children():
                if isinstance(widget, ttk.Frame):
                    widget.config(style='TFrame')

    def confirm_conversion(self):
        """ Gathers the final block data and initiates the final conversion. """
        final_blocks = []
        for item in self.block_widgets:
            if item['var'].get():
                # Get updated content from the text widget
                updated_content = item['text_widget'].get("1.0", tk.END).strip()
                item['block']['content'] = updated_content
                item['block']['uncertain'] = False # Mark as confirmed
                final_blocks.append(item['block'])

        if not final_blocks:
            messagebox.showwarning("没有内容", "您没有选择任何要转换的内容块。")
            return

        # Call the main app's method to perform the final conversion
        self.app.run_final_conversion(self.converter, final_blocks)
        self.destroy()


class App:
    """ The main application class for the Word to Jupyter converter GUI. """
    def __init__(self, root):
        self.root = root
        self.root.title("Word-to-Jupyter 转换器 v3.0")

        self.WIN_WIDTH, self.WIN_HEIGHT = 600, 360
        self.root.geometry(f"{self.WIN_WIDTH}x{self.WIN_HEIGHT}")
        self.root.resizable(False, False)

        # Tkinter variables
        self.file_path = tk.StringVar()
        self.kernel_name = tk.StringVar(value="python3")  # Default kernel name
        self.status_text = tk.StringVar(value="请选择一个 .docx 文件")

        self._setup_background()
        self._create_widgets()

    def _create_ui_background_if_needed(self):
        """
        Checks if the UI background image exists. If not, it generates one
        by rendering text onto the source background image. This function
        is adapted from the original prepare_assets.py script.
        """
        output_path = resource_path(os.path.join('assets', 'background_ui.png'))

        if os.path.exists(output_path):
            return  # Background already exists, do nothing.

        try:
            # --- Decode the embedded Base64 image ---
            image_data = base64.b64decode(BACKGROUND_IMAGE_B64)
            source_image = Image.open(io.BytesIO(image_data))
            
            # --- Configuration for the 600x360 window size ---
            DOC_LABEL_POS = (112, 103)
            KERNEL_LABEL_POS = (112, 146)
            AUTHOR_LABEL_POS = (self.WIN_WIDTH - 13, self.WIN_HEIGHT - 13)

            try:
                # Use a common font, fallback to default if not found
                FONT = ImageFont.truetype("msyh.ttc", 13, encoding="utf-8")
                SMALL_FONT = ImageFont.truetype("msyh.ttc", 9, encoding="utf-8")
            except IOError:
                FONT = ImageFont.load_default()
                SMALL_FONT = ImageFont.load_default()

            TEXT_COLOR = "#333333"
            SHADOW_COLOR = "#FFFFFF"
            AUTHOR_COLOR = "#444444"

            with source_image as img:
                img = img.resize((self.WIN_WIDTH, self.WIN_HEIGHT), Image.Resampling.LANCZOS)
                draw = ImageDraw.Draw(img)

                def draw_text_with_outline(position, text, font, fill, shadow, anchor):
                    """ Helper function to draw text with a subtle outline for better readability. """
                    x, y = position
                    draw.text((x - 1, y - 1), text, font=font, fill=shadow, anchor=anchor)
                    draw.text((x + 1, y - 1), text, font=font, fill=shadow, anchor=anchor)
                    draw.text((x - 1, y + 1), text, font=font, fill=shadow, anchor=anchor)
                    draw.text((x + 1, y + 1), text, font=font, fill=shadow, anchor=anchor)
                    draw.text((x, y), text, font=font, fill=fill, anchor=anchor)

                draw_text_with_outline(DOC_LABEL_POS, "Word 文档:", FONT, TEXT_COLOR, SHADOW_COLOR, anchor='lm')
                draw_text_with_outline(KERNEL_LABEL_POS, "Jupyter 内核:", FONT, TEXT_COLOR, SHADOW_COLOR, anchor='lm')
                draw_text_with_outline(AUTHOR_LABEL_POS, "Created by Stella", SMALL_FONT, AUTHOR_COLOR, SHADOW_COLOR, anchor="rb")

                img.save(output_path, 'PNG')
                print(f"UI background generated automatically: '{output_path}'")

        except Exception as e:
            print(f"Error generating UI background: {e}")

    def _setup_background(self):
        """
        Sets up the main window's background image.
        It first ensures the UI-specific background is generated, then loads it.
        """
        # 1. Generate the background with text labels if it doesn't exist
        self._create_ui_background_if_needed()

        # 2. Load the final UI background image
        bg_path = resource_path(os.path.join('assets', 'background_ui.png'))
        if os.path.exists(bg_path):
            try:
                self.bg_image_pil = Image.open(bg_path).resize((self.WIN_WIDTH, self.WIN_HEIGHT), Image.Resampling.LANCZOS)
                self.bg_image_tk = ImageTk.PhotoImage(self.bg_image_pil)
                bg_label = tk.Label(self.root, image=self.bg_image_tk)
                bg_label.place(x=0, y=0, relwidth=1, relheight=1)
            except Exception as e:
                print(f"Error loading background image: {e}")
                self.bg_image_pil = None
        else:
            # Fallback to a default theme if the background is missing
            self.bg_image_pil = None

    def _create_panel(self, x, y, width, height, radius=20, alpha=100):
        """
        Creates a semi-transparent rounded rectangle panel on top of the
        background image to improve UI element visibility.
        """
        if not self.bg_image_pil:
            # Fallback for when the background image fails to load
            panel = tk.Frame(self.root, bg='#DDDDDD', width=width, height=height)
            panel.place(x=x, y=y, anchor='center')
            return

        # Crop the section of the background where the panel will be placed
        bg_crop = self.bg_image_pil.crop((x - width / 2, y - height / 2, x + width / 2, y + height / 2))

        # Create a new image for the overlay
        overlay = Image.new('RGBA', bg_crop.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(overlay)

        # Draw the rounded rectangle on the overlay
        rect_color = (255, 255, 255, alpha)
        draw.rounded_rectangle((0, 0, width, height), radius, fill=rect_color)

        # Composite the overlay onto the cropped background
        panel_img_pil = Image.alpha_composite(bg_crop.convert('RGBA'), overlay)

        # Convert to Tkinter PhotoImage and display
        self.panel_img_tk = ImageTk.PhotoImage(panel_img_pil)
        panel_label = tk.Label(self.root, image=self.panel_img_tk, borderwidth=0)
        panel_label.place(x=x, y=y, anchor='center')

    def _create_widgets(self):
        """ Creates and places all the widgets in the main window. """
        # --- Panel and Widget Layout ---
        panel_x, panel_y = self.WIN_WIDTH / 2, self.WIN_HEIGHT / 2
        panel_w, panel_h = 470, 240
        self._create_panel(panel_x, panel_y, panel_w, panel_h, radius=22, alpha=90)

        # Y-coordinates for widgets
        doc_y = 103
        kernel_y = 146

        # --- Widget Creation ---
        # Document selection
        doc_entry = ttk.Entry(self.root, textvariable=self.file_path, width=30, state='readonly')
        doc_entry.place(x=410, y=doc_y, anchor='e')
        doc_button = ttk.Button(self.root, text="选择文件", command=self.select_file, width=12)
        doc_button.place(x=420, y=doc_y, anchor='w')

        # Kernel name entry
        kernel_entry = ttk.Entry(self.root, textvariable=self.kernel_name, width=30)
        kernel_entry.place(x=410, y=kernel_y, anchor='e')

        # Kernel hint label
        hint_y = kernel_y + 22
        hint_label = tk.Label(self.root, text="(可选, 留空则跳过执行)",
                              font=("", 8), bg="#DFE5EA", fg="gray", borderwidth=0)
        hint_label.place(x=410, y=hint_y, anchor='e')

        # Main action button
        button_y = kernel_y + 75
        style = ttk.Style()
        style.configure('Accent.TButton', font=("", 11, "bold"))
        preview_button = ttk.Button(self.root, text="解析并预览", command=self.run_parsing, style='Accent.TButton')
        preview_button.place(x=panel_x, y=button_y, anchor='center', width=170, height=42)

        # Status bar
        status_y = button_y + 55
        status_label = tk.Label(self.root, textvariable=self.status_text, foreground="#333333", background="#DFE5EA", borderwidth=0)
        status_label.place(x=panel_x, y=status_y, anchor='center')

    def select_file(self):
        """ Opens a file dialog to select a .docx file. """
        path = filedialog.askopenfilename(
            title="选择一个 Word 文档",
            filetypes=[("Word 文档", "*.docx")]
        )
        if path:
            self.file_path.set(path)
            self.status_text.set(f"已选择: {os.path.basename(path)}")

    def run_parsing(self, event=None):
        """
        Initiates the document parsing process. This runs synchronously
        and will freeze the GUI temporarily.
        """
        docx_path = self.file_path.get()
        if not docx_path:
            messagebox.showwarning("警告", "请先选择一个 Word 文档！")
            return
        kernel = self.kernel_name.get() or None

        try:
            self.status_text.set("正在解析文档，请稍候...")
            self.root.update_idletasks()  # Force UI update

            converter = DocxConverter(
                docx_path=docx_path,
                kernel_name=kernel,
                log_callback=lambda msg: self.status_text.set(msg),
                error_callback=lambda err: messagebox.showerror("解析错误", err)
            )
            blocks = converter.parse_document()

            if blocks is not None:
                self.status_text.set("解析完成。请在预览窗口中调整结果。")
                PreviewWindow(self.root, self, converter, blocks)
            else:
                # This case is hit if parsing fails critically or the doc is empty.
                self.status_text.set("解析失败或文档为空。")

        except Exception as e:
            self.status_text.set(f"❌ 解析失败: {e}")
            messagebox.showerror("错误", f"发生意外错误:\n{e}")

    def run_final_conversion(self, converter, final_blocks):
        """
        Initiates the final notebook creation and execution. This runs
        synchronously and will freeze the GUI temporarily.
        """
        try:
            self.status_text.set("正在生成和执行 Notebook，请稍候...")
            self.root.update_idletasks()  # Force UI update

            created, status = converter.create_notebook_from_blocks(final_blocks)

            if not created:
                self.status_text.set(f"❌ 转换失败: {status}")
                messagebox.showerror("创建失败", f"创建 .ipynb 文件时出错:\n{status}")
                return

            # Handle different execution statuses with user-friendly messages
            if status == "SUCCESS":
                self.status_text.set("✅ 转换和执行成功！")
                messagebox.showinfo("成功", "Notebook 创建并执行成功！")
            elif status == "JUPYTER_NOT_FOUND":
                self.status_text.set("✅ 已转换 (未找到 Jupyter)")
                messagebox.showwarning("转换成功", "Notebook 已创建，但未在您的系统中找到 Jupyter，因此跳过了执行步骤。")
            elif status.startswith("KERNEL_NOT_FOUND"):
                kernel_name = status.split(":", 1)[1]
                self.status_text.set(f"✅ 已转换 (未找到内核 '{kernel_name}')")
                messagebox.showwarning("执行失败", f"Notebook 已创建，但未找到 Jupyter 内核 '{kernel_name}'。")
            elif status == "EXECUTION_ERROR":
                self.status_text.set("✅ 已转换 (但执行时有错误)")
                messagebox.showwarning("执行出错", "Notebook 已创建，但在执行过程中发生错误。请检查 .ipynb 文件以获取详细信息。")
            elif status == "NOT_RUN":
                self.status_text.set("✅ 转换成功！")
                messagebox.showinfo("成功", "Notebook 文件创建成功！ (因未指定内核，已跳过执行)。")
            else:
                self.status_text.set(f"❌ 未知的执行错误: {status}")
                messagebox.showerror("未知错误", f"执行期间发生未知错误:\n{status}")

        except Exception as e:
            self.status_text.set(f"❌ 严重错误: {e}")
            messagebox.showerror("严重错误", f"在最终转换过程中发生严重错误:\n{e}")


if __name__ == '__main__':
    root = tk.Tk()
    style = ttk.Style(root)
    # Use a modern theme if available, otherwise fallback gracefully.
    try:
        style.theme_use("arc")
    except tk.TclError:
        print("Note: 'arc' theme not available, using default theme.")
    app = App(root)
    root.mainloop()