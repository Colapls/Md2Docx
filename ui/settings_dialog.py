"""
设置对话框：字体设置和段落设置
"""
import tkinter as tk
from tkinter import ttk, messagebox
from converter.docx_generator import FontSettings, ParagraphSettings
from docx.enum.text import WD_LINE_SPACING


class FontSettingsTab(ttk.Frame):
    """字体设置标签页"""

    def __init__(self, parent, current_font_settings: FontSettings = None):
        super().__init__(parent)

        self.current = current_font_settings or FontSettings()

        # 获取系统字体列表
        self.system_fonts = self._get_system_fonts()

        self._create_widgets()
        self._load_settings()

    def _get_system_fonts(self) -> list:
        """获取系统字体列表"""
        try:
            import tkinter.font as tkfont
            return list(tkfont.families())
        except Exception:
            return ["宋体", "微软雅黑", "黑体", "楷体", "Times New Roman", "Arial"]

    def _create_widgets(self):
        """创建控件"""
        # 中文字体
        ttk.Label(self, text="正文中文字体:").grid(row=0, column=0, sticky='w', padx=10, pady=10)
        self.chinese_font_var = tk.StringVar()
        self.chinese_font_combo = ttk.Combobox(
            self,
            textvariable=self.chinese_font_var,
            values=self.system_fonts,
            state='readonly',
            width=25
        )
        self.chinese_font_combo.grid(row=0, column=1, sticky='ew', padx=10, pady=10)

        # 英文字体
        ttk.Label(self, text="正文字号:").grid(row=1, column=0, sticky='w', padx=10, pady=10)
        self.font_size_var = tk.StringVar(value="12")
        font_sizes = [str(i) for i in range(10, 17)]
        self.font_size_combo = ttk.Combobox(
            self,
            textvariable=self.font_size_var,
            values=font_sizes,
            state='readonly',
            width=25
        )
        self.font_size_combo.grid(row=1, column=1, sticky='ew', padx=10, pady=10)

        # 英文字体
        ttk.Label(self, text="正文章英文字体:").grid(row=2, column=0, sticky='w', padx=10, pady=10)
        self.english_font_var = tk.StringVar()
        self.english_font_combo = ttk.Combobox(
            self,
            textvariable=self.english_font_var,
            values=self.system_fonts,
            state='readonly',
            width=25
        )
        self.english_font_combo.grid(row=2, column=1, sticky='ew', padx=10, pady=10)

        # 列权重
        self.columnconfigure(1, weight=1)

    def _load_settings(self):
        """加载当前设置"""
        if self.current:
            self.chinese_font_var.set(self.current.chinese_font)
            self.english_font_var.set(self.current.english_font)
            self.font_size_var.set(str(self.current.font_size))

    def get_font_settings(self) -> FontSettings:
        """获取字体设置"""
        settings = FontSettings()
        settings.chinese_font = self.chinese_font_var.get()
        settings.english_font = self.english_font_var.get()
        try:
            settings.font_size = int(self.font_size_var.get())
        except ValueError:
            settings.font_size = 12
        return settings


class ParagraphSettingsTab(ttk.Frame):
    """段落设置标签页"""

    def __init__(self, parent, current_settings: ParagraphSettings = None):
        super().__init__(parent)

        self.current = current_settings or ParagraphSettings()

        self._create_widgets()
        self._load_settings()

    def _create_widgets(self):
        """创建控件"""
        # 行间距
        ttk.Label(self, text="正文章行间距:").grid(row=0, column=0, sticky='w', padx=10, pady=10)
        self.line_spacing_var = tk.StringVar(value="1.5")
        line_spacing_values = ["1.0", "1.5", "2.0"]
        self.line_spacing_combo = ttk.Combobox(
            self,
            textvariable=self.line_spacing_var,
            values=line_spacing_values,
            state='readonly',
            width=25
        )
        self.line_spacing_combo.grid(row=0, column=1, sticky='ew', padx=10, pady=10)

        # 首行缩进
        ttk.Label(self, text="首行缩进:").grid(row=1, column=0, sticky='w', padx=10, pady=10)
        self.first_line_indent_var = tk.BooleanVar(value=True)
        self.first_line_indent_check = ttk.Checkbutton(
            self,
            text="启用首行缩进",
            variable=self.first_line_indent_var
        )
        self.first_line_indent_check.grid(row=1, column=1, sticky='w', padx=10, pady=5)

        # 缩进值
        ttk.Label(self, text="缩进字符数:").grid(row=2, column=0, sticky='w', padx=10, pady=10)
        self.indent_size_var = tk.StringVar(value="2")
        indent_frame = ttk.Frame(self)
        indent_frame.grid(row=2, column=1, sticky='w', padx=10, pady=5)
        self.indent_size_spin = ttk.Spinbox(
            indent_frame,
            from_=0,
            to=10,
            textvariable=self.indent_size_var,
            width=10
        )
        self.indent_size_spin.pack(side='left')

        # 标题设置
        ttk.Separator(self, orient='horizontal').grid(row=3, column=0, columnspan=2,
                                                       sticky='ew', pady=15)

        ttk.Label(self, text="标题行间距:").grid(row=4, column=0, sticky='w', padx=10, pady=10)
        self.heading_line_spacing_var = tk.StringVar(value="1.5")
        heading_line_spacing_values = ["1.0", "1.5", "2.0"]
        self.heading_line_spacing_combo = ttk.Combobox(
            self,
            textvariable=self.heading_line_spacing_var,
            values=heading_line_spacing_values,
            state='readonly',
            width=25
        )
        self.heading_line_spacing_combo.grid(row=4, column=1, sticky='ew', padx=10, pady=10)

        # 段前距
        ttk.Label(self, text="标题段前距(磅):").grid(row=5, column=0, sticky='w', padx=10, pady=10)
        self.space_before_var = tk.StringVar(value="12")
        self.space_before_spin = ttk.Spinbox(
            self,
            from_=0,
            to=50,
            textvariable=self.space_before_var,
            width=25
        )
        self.space_before_spin.grid(row=5, column=1, sticky='ew', padx=10, pady=10)

        # 段后距
        ttk.Label(self, text="标题段后距(磅):").grid(row=6, column=0, sticky='w', padx=10, pady=10)
        self.space_after_var = tk.StringVar(value="6")
        self.space_after_spin = ttk.Spinbox(
            self,
            from_=0,
            to=50,
            textvariable=self.space_after_var,
            width=25
        )
        self.space_after_spin.grid(row=6, column=1, sticky='ew', padx=10, pady=10)

        # 列权重
        self.columnconfigure(1, weight=1)

    def _load_settings(self):
        """加载当前设置"""
        if self.current:
            self.line_spacing_var.set(str(self.current.line_spacing))
            self.first_line_indent_var.set(self.current.first_line_indent_enabled)
            self.indent_size_var.set(str(self.current.first_line_indent))
            self.heading_line_spacing_var.set(str(self.current.heading_line_spacing))
            self.space_before_var.set(str(int(self.current.heading_space_before.pt)))
            self.space_after_var.set(str(int(self.current.heading_space_after.pt)))

    def get_paragraph_settings(self) -> ParagraphSettings:
        """获取段落设置"""
        from docx.shared import Pt

        settings = ParagraphSettings()
        settings.line_spacing = float(self.line_spacing_var.get())
        settings.line_spacing_type = WD_LINE_SPACING.MULTIPLE
        settings.first_line_indent_enabled = self.first_line_indent_var.get()
        settings.first_line_indent = int(self.indent_size_var.get())
        settings.heading_line_spacing = float(self.heading_line_spacing_var.get())

        try:
            settings.heading_space_before = Pt(float(self.space_before_var.get()))
        except ValueError:
            settings.heading_space_before = Pt(12)

        try:
            settings.heading_space_after = Pt(float(self.space_after_var.get()))
        except ValueError:
            settings.heading_space_after = Pt(6)

        return settings


class SettingsDialog(tk.Toplevel):
    """设置对话框"""

    def __init__(self, parent, font_settings: FontSettings = None,
                 paragraph_settings: ParagraphSettings = None):
        super().__init__(parent)

        self.title("设置")
        self.geometry("450x400")
        self.resizable(False, False)

        # 居中显示
        self.transient(parent)
        self.grab_set()

        self.font_settings = font_settings
        self.paragraph_settings = paragraph_settings
        self.result_font_settings = None
        self.result_paragraph_settings = None

        self._create_widgets()

        # 等待对话框关闭
        self.wait_window()

    def _create_widgets(self):
        """创建控件"""
        # 创建笔记本（Tab 控制）
        notebook = ttk.Notebook(self)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # 字体设置标签页
        self.font_tab = FontSettingsTab(notebook, self.font_settings)
        notebook.add(self.font_tab, text="字体设置")

        # 段落设置标签页
        self.paragraph_tab = ParagraphSettingsTab(notebook, self.paragraph_settings)
        notebook.add(self.paragraph_tab, text="段落设置")

        # 按钮框架
        button_frame = ttk.Frame(self)
        button_frame.pack(fill='x', padx=10, pady=10)

        # 确定按钮
        ok_button = ttk.Button(button_frame, text="确定", command=self._on_ok)
        ok_button.pack(side='right', padx=5)

        # 取消按钮
        cancel_button = ttk.Button(button_frame, text="取消", command=self._on_cancel)
        cancel_button.pack(side='right')

    def _on_ok(self):
        """确定按钮"""
        self.result_font_settings = self.font_tab.get_font_settings()
        self.result_paragraph_settings = self.paragraph_tab.get_paragraph_settings()
        self.destroy()

    def _on_cancel(self):
        """取消按钮"""
        self.destroy()

    def get_settings(self):
        """获取设置结果"""
        return self.result_font_settings, self.result_paragraph_settings
