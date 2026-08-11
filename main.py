"""
Markdown 转 DOCX 可视化工具
主程序入口
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
from typing import Dict, Any

from converter.markdown_converter import MarkdownConverter
from converter.docx_generator import FontSettings, ParagraphSettings


# 中文字号映射（Word 标准字号表）
CHINESE_FONT_SIZES = [
    ("初号", 42), ("小初", 36),
    ("一号", 26), ("小一", 24),
    ("二号", 22), ("小二", 18),
    ("三号", 16), ("小三", 15),
    ("四号", 14), ("小四", 12),
    ("五号", 10.5), ("小五", 9),
    ("六号", 7.5), ("小六", 6.5),
    ("七号", 5.5), ("八号", 5),
]
CHINESE_SIZE_LIST = [f"{name} ({size})" for name, size in CHINESE_FONT_SIZES]


def get_font_size_from_chinese(chinese_str: str) -> int:
    """从中文字号获取数字值"""
    for name, size in CHINESE_FONT_SIZES:
        if name in chinese_str:
            return size
    return 12


class FontSettingPanel(ttk.Frame):
    """字体设置面板"""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)

        self.chinese_font_var = tk.StringVar(value="宋体")
        self.english_font_var = tk.StringVar(value="Times New Roman")
        self.font_size_var = tk.StringVar(value="五号 (10.5)")

        self._create_widgets()

    def _create_widgets(self):
        """创建控件"""
        row1 = ttk.Frame(self)
        row1.pack(fill='x', pady=2)
        ttk.Label(row1, text="中文字体:", width=8).pack(side='left')
        chinese_combo = ttk.Combobox(row1, textvariable=self.chinese_font_var,
                                      width=15, state='readonly')
        chinese_combo['values'] = self._get_font_list()
        chinese_combo.pack(side='left', fill='x', expand=True)

        row2 = ttk.Frame(self)
        row2.pack(fill='x', pady=2)
        ttk.Label(row2, text="英文字体:", width=8).pack(side='left')
        english_combo = ttk.Combobox(row2, textvariable=self.english_font_var,
                                      width=15, state='readonly')
        english_combo['values'] = self._get_english_font_list()
        english_combo.pack(side='left', fill='x', expand=True)

        row3 = ttk.Frame(self)
        row3.pack(fill='x', pady=2)
        ttk.Label(row3, text="字号:", width=8).pack(side='left')
        size_combo = ttk.Combobox(row3, textvariable=self.font_size_var,
                                   width=15, state='readonly')
        size_combo['values'] = CHINESE_SIZE_LIST
        size_combo.pack(side='left', fill='x', expand=True)

    def _get_font_list(self):
        """获取中文字体列表"""
        try:
            import tkinter.font as tkfont
            fonts = list(tkfont.families())
            common = ["宋体", "微软雅黑", "黑体", "楷体", "仿宋", "方正姚体", "方正舒体"]
            for c in common:
                if c in fonts:
                    fonts.remove(c)
                    fonts.insert(0, c)
            return fonts
        except:
            return ["宋体", "微软雅黑", "黑体", "楷体", "仿宋"]

    def _get_english_font_list(self):
        """获取英文字体列表"""
        return ["Times New Roman", "Arial", "Calibri", "Consolas", "Courier New",
                "Georgia", "Impact", "Verdana"]

    def get_font_settings(self) -> FontSettings:
        """获取字体设置"""
        settings = FontSettings()
        settings.chinese_font = self.chinese_font_var.get()
        settings.english_font = self.english_font_var.get()
        settings.font_size = get_font_size_from_chinese(self.font_size_var.get())
        return settings

    def set_font_settings(self, settings: FontSettings):
        """设置字体"""
        self.chinese_font_var.set(settings.chinese_font)
        self.english_font_var.set(settings.english_font)
        for name, size in CHINESE_FONT_SIZES:
            if abs(size - settings.font_size) < 0.5:
                self.font_size_var.set(f"{name} ({size})")
                return
        self.font_size_var.set("五号 (10.5)")


class HeadingFontPanel(ttk.Frame):
    """标题字体设置面板"""

    def __init__(self, parent, heading_count: int = 6, **kwargs):
        super().__init__(parent, **kwargs)

        self.heading_count = heading_count
        self.heading_panels = {}

        self._create_widgets()

    def _create_widgets(self):
        """创建控件"""
        top_frame = ttk.Frame(self)
        top_frame.pack(fill='x', pady=(0, 5))

        ttk.Label(top_frame, text="标题级别:", font=('', 9, 'bold')).pack(side='left')

        self.level_var = tk.StringVar(value="1")
        level_combo = ttk.Combobox(top_frame, textvariable=self.level_var,
                                    width=8, state='readonly')
        level_combo['values'] = [str(i) for i in range(1, self.heading_count + 1)]
        level_combo.pack(side='left', padx=10)
        level_combo.bind('<<ComboboxSelected>>', self._on_level_changed)

        # 为每个级别创建面板
        for i in range(1, self.heading_count + 1):
            panel = FontSettingPanel(self)
            panel.pack(fill='x', pady=2)
            self.heading_panels[i] = panel

            # 设置默认值
            default_settings = FontSettings()
            default_settings.font_size = 16 - (i - 1) * 2 if i <= 6 else 10
            default_settings.font_size = max(8, default_settings.font_size)
            default_settings.chinese_font = "黑体" if i == 1 else "宋体"
            panel.set_font_settings(default_settings)

        self._show_level(1)

    def _on_level_changed(self, event=None):
        """切换标题级别"""
        level = int(self.level_var.get())
        self._show_level(level)

    def _show_level(self, level: int):
        """显示指定级别的设置"""
        for lvl, panel in self.heading_panels.items():
            if lvl == level:
                panel.pack(fill='x', pady=2)
            else:
                panel.pack_forget()

    def get_heading_font(self, level: int) -> FontSettings:
        """获取指定级别的字体设置"""
        if level < 1:
            level = 1
        if level > self.heading_count:
            level = self.heading_count
        return self.heading_panels[level].get_font_settings()


class CollapsibleFrame(ttk.Frame):
    """可折叠框架"""

    def __init__(self, parent, title: str, **kwargs):
        super().__init__(parent, **kwargs)

        self.title = title
        self.is_expanded = tk.BooleanVar(value=False)

        self._create_widgets()

    def _create_widgets(self):
        """创建控件"""
        # 标题栏
        self.header = ttk.Frame(self)
        self.header.pack(fill='x')

        self.toggle_btn = ttk.Button(self.header, text="▶ " + self.title,
                                      command=self.toggle, width=15)
        self.toggle_btn.pack(side='left')

        # 内容区
        self.content_frame = ttk.Frame(self, padding="5")

    def toggle(self):
        """切换展开/折叠"""
        if self.is_expanded.get():
            self.content_frame.pack_forget()
            self.toggle_btn.config(text="▶ " + self.title)
            self.is_expanded.set(False)
        else:
            self.content_frame.pack(fill='x', pady=5)
            self.toggle_btn.config(text="▼ " + self.title)
            self.is_expanded.set(True)

    def expand(self):
        """展开"""
        if not self.is_expanded.get():
            self.toggle()

    def get_content_frame(self) -> ttk.Frame:
        """获取内容框架"""
        return self.content_frame


class Application(tk.Tk):
    """主应用程序窗口"""

    def __init__(self):
        super().__init__()

        self.title("Markdown 转 DOCX")
        self.geometry("620x820")
        self.resizable(True, True)

        self.selected_file_path = None

        self._create_widgets()
        self._set_default_font()

    def _set_default_font(self):
        """设置默认字体"""
        try:
            style = ttk.Style()
            style.configure('.', font=('微软雅黑', 9))

            # Big green "Convert" button.
            style.configure(
                'Green.TButton',
                font=('微软雅黑', 12, 'bold'),
                foreground='black',
                background='#2E8B57',      # sea green
                bordercolor='#1F6B43',
                padding=(20, 8),
            )
            style.map(
                'Green.TButton',
                foreground=[('disabled', '#666666'),
                            ('pressed',  '#000000'),
                            ('active',   '#1a1a1a')],
                background=[('disabled', '#7a7a7a'),
                            ('pressed',  '#1F6B43'),
                            ('active',   '#3CA66C')],
            )

            # Bigger progress bar text + chunk color.
            style.configure('Big.Horizontal.TProgressbar', thickness=18)
        except Exception:
            pass

    def _create_widgets(self):
        """创建界面控件"""
        outer = ttk.Frame(self, padding="10")
        outer.pack(fill='both', expand=True)

        # 中间可滚动区域 —— 容纳所有"配置"面板。
        # 用 Canvas+Scrollbar 实现，展开折叠面板时仍可通过滚动查看内容。
        body_frame = ttk.Frame(outer)
        body_frame.pack(fill='both', expand=True)

        self._body_canvas = tk.Canvas(
            body_frame, highlightthickness=0, borderwidth=0)
        self._body_canvas.pack(side='left', fill='both', expand=True)

        body_scroll = ttk.Scrollbar(
            body_frame, orient='vertical',
            command=self._body_canvas.yview)
        body_scroll.pack(side='right', fill='y')

        self._body_canvas.configure(yscrollcommand=body_scroll.set)

        # 鼠标滚轮支持（Windows / macOS）。
        self._body_canvas.bind_all(
            '<MouseWheel>',
            lambda e: self._body_canvas.yview_scroll(
                int(-1 * (e.delta / 120)), 'units'),
        )
        self._body_canvas.bind_all(
            '<Button-4>',
            lambda e: self._body_canvas.yview_scroll(-1, 'units'),
        )
        self._body_canvas.bind_all(
            '<Button-5>',
            lambda e: self._body_canvas.yview_scroll(1, 'units'),
        )

        # 配置面板的容器 Frame。
        main_frame = ttk.Frame(self._body_canvas)
        self._body_canvas_window = self._body_canvas.create_window(
            (0, 0), window=main_frame, anchor='nw',
        )
        main_frame.bind(
            '<Configure>',
            lambda e: self._body_canvas.configure(
                scrollregion=self._body_canvas.bbox('all'),
            ),
        )
        self._body_canvas.bind(
            '<Configure>',
            lambda e: self._body_canvas.itemconfigure(
                self._body_canvas_window, width=e.width,
            ),
        )

        # ========== 文件选择区域 ==========
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="8")
        file_frame.pack(fill='x', pady=(0, 10))

        file_inner = ttk.Frame(file_frame)
        file_inner.pack(fill='x')

        self.file_path_var = tk.StringVar(value="未选择文件")
        file_label = ttk.Label(file_inner, textvariable=self.file_path_var,
                                foreground='gray', font=('', 9))
        file_label.pack(side='left', fill='x', expand=True)

        select_button = ttk.Button(file_inner, text="选择文件", command=self._select_file)
        select_button.pack(side='right', padx=(10, 0))

        # ========== 字体设置区域 ==========
        font_frame = ttk.LabelFrame(main_frame, text="字体设置", padding="8")
        font_frame.pack(fill='x', pady=(0, 10))

        # 标题字体
        heading_label = ttk.Label(font_frame, text="标题字体 (H1-H6)", font=('', 9, 'bold'))
        heading_label.pack(anchor='w', pady=(0, 5))

        self.heading_font_panel = HeadingFontPanel(font_frame, heading_count=6)
        self.heading_font_panel.pack(fill='x')

        ttk.Separator(font_frame, orient='horizontal').pack(fill='x', pady=10)

        # 正文字体
        body_label = ttk.Label(font_frame, text="正文字体", font=('', 9, 'bold'))
        body_label.pack(anchor='w', pady=(0, 5))
        self.body_font_panel = FontSettingPanel(font_frame)
        self.body_font_panel.pack(fill='x')

        # ========== 段落设置区域 ==========
        para_frame = ttk.LabelFrame(main_frame, text="段落设置", padding="8")
        para_frame.pack(fill='x', pady=(0, 10))

        # 标题段落
        heading_para_label = ttk.Label(para_frame, text="标题段落", font=('', 9, 'bold'))
        heading_para_label.pack(anchor='w', pady=(0, 5))

        heading_para_inner = ttk.Frame(para_frame)
        heading_para_inner.pack(fill='x')
        ttk.Label(heading_para_inner, text="行间距:").pack(side='left')
        self.heading_line_spacing_var = tk.StringVar(value="1.25")
        heading_spacing_combo = ttk.Combobox(heading_para_inner, textvariable=self.heading_line_spacing_var,
                                             width=12, state='readonly')
        heading_spacing_combo['values'] = ["1.0", "1.25", "1.5", "1.75", "2.0", "2.5", "3.0"]
        heading_spacing_combo.pack(side='left', padx=(5, 15))

        ttk.Label(heading_para_inner, text="段前距:").pack(side='left')
        self.heading_space_before_var = tk.StringVar(value="12")
        ttk.Spinbox(heading_para_inner, from_=0, to=50,
                    textvariable=self.heading_space_before_var, width=5).pack(side='left', padx=(5, 10))
        ttk.Label(heading_para_inner, text="磅").pack(side='left')

        heading_para_inner2 = ttk.Frame(para_frame)
        heading_para_inner2.pack(fill='x', pady=3)
        ttk.Label(heading_para_inner2, text="段后距:").pack(side='left')
        self.heading_space_after_var = tk.StringVar(value="6")
        ttk.Spinbox(heading_para_inner2, from_=0, to=50,
                    textvariable=self.heading_space_after_var, width=5).pack(side='left', padx=(5, 10))
        ttk.Label(heading_para_inner2, text="磅").pack(side='left')

        ttk.Separator(para_frame, orient='horizontal').pack(fill='x', pady=8)

        # 正文段落
        body_para_label = ttk.Label(para_frame, text="正文段落", font=('', 9, 'bold'))
        body_para_label.pack(anchor='w', pady=(0, 5))

        body_para_inner = ttk.Frame(para_frame)
        body_para_inner.pack(fill='x')
        ttk.Label(body_para_inner, text="行间距:").pack(side='left')
        self.body_line_spacing_var = tk.StringVar(value="1.25")
        body_spacing_combo = ttk.Combobox(body_para_inner, textvariable=self.body_line_spacing_var,
                                          width=12, state='readonly')
        body_spacing_combo['values'] = ["1.0", "1.25", "1.5", "1.75", "2.0", "2.5", "3.0"]
        body_spacing_combo.pack(side='left', padx=(5, 15))

        self.body_first_line_indent_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(body_para_inner, text="首行缩进",
                        variable=self.body_first_line_indent_var).pack(side='left')

        self.body_indent_size_var = tk.StringVar(value="2")
        ttk.Spinbox(body_para_inner, from_=0, to=10,
                    textvariable=self.body_indent_size_var, width=3).pack(side='left', padx=5)
        ttk.Label(body_para_inner, text="字符").pack(side='left')

        # 转换选项 —— 分两行摆放，避免横向溢出
        options_frame = ttk.Frame(para_frame)
        options_frame.pack(fill='x', pady=(10, 0))

        # 第 1 行：常用勾选（去空格 + 标题重新编号 + 处理表格）
        row1 = ttk.Frame(options_frame)
        row1.pack(fill='x', pady=(0, 4))

        self.remove_spaces_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row1, text="自动去除多余空格",
                        variable=self.remove_spaces_var).pack(side='left', padx=(0, 12))

        self.renumber_headings_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row1, text="标题重新编号",
                        variable=self.renumber_headings_var).pack(side='left', padx=(0, 12))

        self.process_tables_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            row1, text="处理表格",
            variable=self.process_tables_var,
            command=self._on_table_toggle,
        ).pack(side='left', padx=(0, 12))

        # 第 2 行：公式转换模式 + 表格对齐（仅处理表格开启时显示）
        row2 = ttk.Frame(options_frame)
        row2.pack(fill='x')

        # 探测 pandoc 是否可用
        from converter.pandoc_helper import is_pandoc_available
        self._pandoc_available = is_pandoc_available()

        self.formula_mode_frame = ttk.Frame(row2)
        self.formula_mode_frame.pack(side='left')

        ttk.Label(self.formula_mode_frame, text="公式转换:").pack(side='left', padx=(0, 5))

        self.formula_mode_var = tk.StringVar(
            value='pandoc' if self._pandoc_available else 'builtin'
        )
        self.formula_mode_combo = ttk.Combobox(
            self.formula_mode_frame,
            textvariable=self.formula_mode_var,
            values=['pandoc', 'builtin'],
            width=10,
            state='readonly',
        )
        self.formula_mode_combo.pack(side='left')

        if self._pandoc_available:
            status_text = " (pandoc 已检测到)"
            status_color = 'green'
        else:
            status_text = " (未安装 pandoc，使用内置)"
            status_color = 'gray'
        self.formula_status_label = ttk.Label(
            self.formula_mode_frame, text=status_text,
            foreground=status_color, font=('', 8),
        )
        self.formula_status_label.pack(side='left', padx=(5, 0))

        # 旧变量保留为兼容性读取（默认勾选状态=启用公式转换）
        self.convert_formulas_var = tk.BooleanVar(value=True)

        # ========== 表格设置（独立一栏，默认隐藏） ==========
        self.table_frame = ttk.LabelFrame(main_frame, text="表格设置", padding="8")
        # 不立即 pack，由 _on_table_toggle 控制

        # Map display label -> internal OOXML alignment value.
        self._TABLE_ALIGN_LABELS = {
            '靠左': 'left',
            '居中': 'center',
            '靠右': 'right',
        }

        # ---- 表格字体（不折叠，直接显示）----
        font_label = ttk.Label(self.table_frame, text="表格字体", font=('', 9, 'bold'))
        font_label.pack(anchor='w', pady=(0, 5))

        self.table_font_panel = FontSettingPanel(self.table_frame)
        self.table_font_panel.pack(fill='x', pady=(0, 10))

        # 默认设置表格字体为小五号
        table_default = FontSettings()
        table_default.font_size = 9
        self.table_font_panel.set_font_settings(table_default)

        # ---- 表格段落（不折叠，直接显示）----
        para_label = ttk.Label(self.table_frame, text="表格段落", font=('', 9, 'bold'))
        para_label.pack(anchor='w', pady=(0, 5))

        # 单元格对齐
        align_row = ttk.Frame(self.table_frame)
        align_row.pack(fill='x', pady=(0, 5))

        ttk.Label(align_row, text="单元格对齐:").pack(side='left', padx=(0, 5))

        self.table_alignment_var = tk.StringVar(value='居中')
        self.table_alignment_combo = ttk.Combobox(
            align_row,
            textvariable=self.table_alignment_var,
            values=['靠左', '居中', '靠右'],
            width=6,
            state='readonly',
        )
        self.table_alignment_combo.pack(side='left')

        # 行间距
        spacing_row = ttk.Frame(self.table_frame)
        spacing_row.pack(fill='x')

        ttk.Label(spacing_row, text="行间距:").pack(side='left')

        self.table_line_spacing_var = tk.StringVar(value="1.0")
        table_spacing_combo = ttk.Combobox(
            spacing_row,
            textvariable=self.table_line_spacing_var,
            width=12, state='readonly',
        )
        table_spacing_combo['values'] = ["1.0", "1.25", "1.5", "1.75", "2.0"]
        table_spacing_combo.pack(side='left', padx=(5, 0))

        # ========== 转换按钮 + 进度条（固定在底部，永远可见） ==========
        button_frame = ttk.Frame(outer)
        button_frame.pack(fill='x', pady=(6, 0))

        self.convert_button = ttk.Button(
            button_frame, text="转    换",
            command=self._convert, style='Green.TButton',
        )
        self.convert_button.pack(pady=5, ipadx=10, ipady=4)

        # ========== 进度条 ==========
        progress_frame = ttk.Frame(outer)
        progress_frame.pack(fill='x', pady=(0, 6))

        self.progress_label_var = tk.StringVar(value='')
        self.progress_label = ttk.Label(
            progress_frame, textvariable=self.progress_label_var,
            foreground='#1F6B43', font=('微软雅黑', 9),
        )
        self.progress_label.pack(anchor='w')

        self.progress_bar = ttk.Progressbar(
            progress_frame, orient='horizontal', mode='indeterminate',
            style='Big.Horizontal.TProgressbar',
        )
        # Hidden until a conversion starts.
        self.progress_bar.pack(fill='x', pady=(2, 0))
        self.progress_bar.pack_forget()
        self.progress_label.pack_forget()

        # ========== 状态显示（永远钉在滚动区底部） ==========
        status_frame = ttk.LabelFrame(main_frame, text="状态", padding="8")
        # side='bottom' ensures the status panel is anchored to the
        # bottom of the scrollable area, regardless of what was packed
        # above it. We pack this LAST so its bottom-anchor sticks.
        status_frame.pack(side='bottom', fill='x', pady=(10, 0))

        self.status_text = tk.Text(status_frame, height=4, state='disabled',
                                    wrap='word', relief='flat', font=('', 9))
        self.status_text.pack(fill='x')

        scrollbar = ttk.Scrollbar(self.status_text, command=self.status_text.yview)
        scrollbar.pack(side='right', fill='y')
        self.status_text.config(yscrollcommand=scrollbar.set)

        # ========== 底部信息（外层固定显示） ==========
        info_label = ttk.Label(outer, text="支持 .md 文件转换为 .docx 格式",
                               foreground='gray', font=('', 8))
        info_label.pack(side='bottom', pady=(2, 0))

    # ---- progress bar helpers ---------------------------------------
    def _progress_begin(self, text: str = '正在准备转换...') -> None:
        """Show the progress bar + label and start its indeterminate
        animation. Called once a conversion starts."""
        try:
            self.progress_label_var.set(text)
            self.progress_label.pack(
                anchor='w', before=self.progress_bar,  # no-op if already packed
            )
            # Re-pack in case they were hidden.
            self.progress_label.pack(anchor='w')
            self.progress_bar.pack(fill='x', pady=(2, 0))
            self.progress_bar.start(12)
        except Exception:
            pass

    def _progress_step(self, text: str) -> None:
        """Update the label, leave the bar running."""
        try:
            self.progress_label_var.set(text)
            self.update_idletasks()
        except Exception:
            pass

    def _progress_end(self, text: str = '') -> None:
        """Stop the animation and hide the bar again."""
        try:
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
            self.progress_label.pack_forget()
            self.progress_label_var.set(text)
        except Exception:
            pass

    def _on_table_toggle(self):
        """Show the dedicated '表格设置' LabelFrame only when the user
        has ticked the '处理表格' checkbox. The LabelFrame itself is a
        single pack target, so we don't have to chase sub-widgets
        individually."""
        if self.process_tables_var.get():
            self.table_frame.pack(fill='x', pady=(0, 10))
        else:
            self.table_frame.pack_forget()

    def _select_file(self):
        """选择文件"""
        filepath = filedialog.askopenfilename(
            title="选择 Markdown 文件",
            filetypes=[("Markdown 文件", "*.md"), ("所有文件", "*.*")]
        )

        if filepath:
            self.selected_file_path = filepath
            filename = os.path.basename(filepath)
            self.file_path_var.set(filename)
            self._update_status(f"已选择文件: {filename}")

    def _update_status(self, message: str):
        """更新状态文本"""
        self.status_text.config(state='normal')
        self.status_text.insert('end', f"{message}\n")
        self.status_text.see('end')
        self.status_text.config(state='disabled')

    def _get_settings(self):
        """获取当前所有设置"""
        from docx.shared import Pt

        # 6级标题字体
        heading_fonts = [self.heading_font_panel.get_heading_font(i) for i in range(1, 7)]

        # 正文字体
        body_font = self.body_font_panel.get_font_settings()

        # 表格字体（从折叠面板获取）
        table_font = self.table_font_panel.get_font_settings()

        # 图名表名字体（比正文字号小一号，跟正文字体一致）
        caption_font = FontSettings()
        caption_font.chinese_font = body_font.chinese_font
        caption_font.english_font = body_font.english_font
        caption_font.font_size = max(8, body_font.font_size - 1.5)  # 小一号

        # 段落设置
        heading_para = ParagraphSettings()
        heading_para.line_spacing = float(self.heading_line_spacing_var.get())
        heading_para.heading_line_spacing = float(self.heading_line_spacing_var.get())
        heading_para.heading_space_before = Pt(float(self.heading_space_before_var.get()))
        heading_para.heading_space_after = Pt(float(self.heading_space_after_var.get()))

        body_para = ParagraphSettings()
        body_para.line_spacing = float(self.body_line_spacing_var.get())
        body_para.first_line_indent_enabled = self.body_first_line_indent_var.get()
        body_para.first_line_indent = int(self.body_indent_size_var.get())

        table_para = ParagraphSettings()
        table_para.line_spacing = float(self.table_line_spacing_var.get())

        # 图名表名：行间距与正文一致，默认居中
        caption_para = ParagraphSettings()
        caption_para.line_spacing = float(self.body_line_spacing_var.get())
        caption_para.first_line_indent_enabled = False

        caption_alignment = "center"  # 默认居中

        # 转换选项
        remove_spaces = self.remove_spaces_var.get()
        convert_formulas = self.convert_formulas_var.get()

        return {
            'heading_fonts': heading_fonts,
            'body_font': body_font,
            'table_font': table_font,
            'caption_font': caption_font,
            'heading_para': heading_para,
            'body_para': body_para,
            'table_para': table_para,
            'caption_para': caption_para,
            'caption_alignment': caption_alignment,
            'remove_spaces': remove_spaces,
            'convert_formulas': convert_formulas,
        }

    def _convert(self):
        """执行转换"""
        if not self.selected_file_path:
            messagebox.showwarning("警告", "请先选择 Markdown 文件")
            return

        if not os.path.exists(self.selected_file_path):
            messagebox.showerror("错误", "文件不存在")
            return

        # Lock the button + show progress bar so the user can't double-click
        # and so they get visible feedback during the (potentially multi-second)
        # conversion.
        self.convert_button.config(state='disabled', text='转换中…')
        self._progress_begin('正在读取文件...')

        try:
            self._update_status("正在读取文件...")

            with open(self.selected_file_path, 'r', encoding='utf-8') as f:
                markdown_text = f.read()

            settings = self._get_settings()

            # ---------- 选择公式转换路径 ----------
            # 1) pandoc 路径：让 pandoc 转 docx（含 OMML 原生公式），
            #    然后用 style_patcher 打字体/缩进/标题/A4 等补丁。
            #    公式保真度最高（texmath → OMML），但需要系统装 pandoc。
            # 2) 内置路径：用 python-docx + latex2mathml 自己渲染。
            use_pandoc = (
                self.formula_mode_var.get() == 'pandoc'
                and self._pandoc_available
                and settings.get('convert_formulas', True)
            )

            default_name = os.path.splitext(
                os.path.basename(self.selected_file_path)
            )[0] + '.docx'

            output_path = filedialog.asksaveasfilename(
                title="保存 DOCX 文件",
                defaultextension=".docx",
                initialfile=default_name,
                filetypes=[("DOCX 文件", "*.docx"), ("所有文件", "*.*")]
            )

            if not output_path:
                self._update_status("用户取消保存")
                return

            if use_pandoc:
                self._progress_step('使用 pandoc 高保真公式转换 (1/3)...')
                self._update_status("使用 pandoc 高保真公式转换...")
                self._convert_via_pandoc(
                    self.selected_file_path, output_path, settings
                )
            else:
                self._progress_step('使用内置公式转换 (1/3)...')
                self._update_status("使用内置公式转换...")
                self._convert_via_builtin(markdown_text, output_path, settings)

            self._progress_step('完成 (3/3)')
            self._update_status(f"转换完成！\n文件已保存至: {output_path}")
            messagebox.showinfo("成功", f"文件已保存至:\n{output_path}")

        except Exception as e:
            error_msg = f"转换失败: {str(e)}"
            self._update_status(error_msg)
            messagebox.showerror("错误", error_msg)
            import traceback
            traceback.print_exc()
        finally:
            # Restore button + hide progress bar regardless of outcome.
            self._progress_end()
            self.convert_button.config(state='normal', text='转    换')

    def _convert_via_pandoc(self, md_path: str, output_path: str,
                            settings: Dict[str, Any]) -> None:
        """Pandoc pipeline: pandoc -> docx -> style patch."""
        import tempfile
        from converter.pandoc_helper import md_to_docx_via_pandoc
        from converter.style_patcher import patch_docx_styles

        # Pass per-conversion toggles through to the patcher.
        settings = dict(settings)
        settings['renumber_headings'] = bool(self.renumber_headings_var.get())
        settings['process_tables'] = bool(self.process_tables_var.get())
        # Translate Chinese label back to OOXML alignment token.
        align_label = self.table_alignment_var.get() or '居中'
        settings['table_cell_alignment'] = self._TABLE_ALIGN_LABELS.get(
            align_label, 'center')

        with tempfile.TemporaryDirectory() as tmp:
            raw = os.path.join(tmp, 'pandoc_raw.docx')
            self._progress_step('调用 pandoc 翻译 LaTeX 公式 (2/3)...')
            md_to_docx_via_pandoc(md_path, raw)
            self._progress_step('应用样式补丁 (3/3)...')
            patch_docx_styles(raw, output_path, settings)

    def _convert_via_builtin(self, markdown_text: str, output_path: str,
                             settings: Dict[str, Any]) -> None:
        """Original pipeline: python-docx + latex2mathml."""
        self._progress_step('解析 Markdown (2/3)...')
        self._update_status("正在解析 Markdown...")

        converter = MarkdownConverter()
        elements = converter.convert(markdown_text)

        self._update_status(f"解析完成，共 {len(elements)} 个元素")
        self._progress_step('生成 DOCX (3/3)...')
        self._update_status("正在生成 DOCX...")

        from converter.docx_generator import generate_docx_with_settings
        generate_docx_with_settings(elements, settings, output_path)


def main():
    """主函数"""
    app = Application()
    app.mainloop()


if __name__ == '__main__':
    main()
