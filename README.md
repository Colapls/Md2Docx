# Markdown 转 DOCX 工具

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)]()
[![License](https://img.shields.io/badge/License-MIT-green.svg)]()
[![Built with](https://img.shields.io/badge/Built%20with-PyInstaller-orange.svg)](https://www.pyinstaller.org/)

一个面向中文科研/工程文档的 Markdown → DOCX 转换桌面工具。
AI生成的 Markdown文件，**一键**得到排版规范、能直接发给同事的 Word 文档。

> 之所以不直接用 Word 写：Markdown 让你专注内容；用这个工具，**Word 只负责最后的排版**。

---

## ✨ 功能亮点

### �� 文本格式
- **字体 / 字号**：中文字体、英文字体、正文字号全部独立可配
- **段落**：行间距、首行缩进、标题段前段后距
- **标题自动编号**：一键去掉原有编号，按 1 / 1.1 / 1.1.1 重新编号
- **去横线**：自动移除 `---` 转成的水平横线
- **空格清理**：自动合并字母/数字旁的空格、`中文 英文` 之间的空格
- **中文引号**：自动将 `"..."` 转成 `"..."` 并按出现顺序左右配对

### �� 数学公式
- **`$...$` 行内公式 + `$$...$$` 块公式**
- 通过 **pandoc** 把 LaTeX 翻译成 **Office Math Markup Language (OMML)** —— Word/WPS 双击公式即可进入公式编辑器继续修改
- **公式居中显示**
- 没装 pandoc 时自动降级到内置 `latex2mathml`（输出 MathML，Word 也能识别）

### �� 表格
- ✅ 自动给所有单元格加边框
- ✅ 单元格文字对齐：靠左 / 居中 / 靠右
- ✅ 表格整体居中显示
- ✅ 自动在表格上方插入 `表1 表名`、`表2 表名` 形式的中文表名
- ✅ 表名字体与正文一致，字号比正文小一号，无斜体
- ✅ 表格后自动加一行空行，拉开与正文的距离
- ✅ 表格内容（字体/字号/行间距/无缩进）**完全独立于正文设置**

### ��️ 界面
- 文件选择 → 字体/段落设置 → 转换，**三步完成**
- 帮助按钮 + 状态栏实时反馈
- 滚动容器 + 固定底部的转换按钮，**永远可见**
- 转换按钮 ... 显眼，没错
- 进度条 + 步骤提示（中/英文）

---

## �� 截图

TODO: 截图占位

```
┌────────────────────────────────────────────┐
│  Markdown 转 DOCX                          │
│                                            │
│  文件选择：[.../example.md          ] [浏览] │
│                                            │
│  ▼ 字体设置                                │
│    ▼ 标题1: 黑体 16pt                       │
│      ...                                   │
│    正文字体: 宋体 10.5pt                     │
│                                            │
│  ▼ 段落设置                                │
│    行间距: 1.5   首行缩进: 2 字符             │
│                                            │
│  ▼ 表格设置 (勾选"处理表格"后展开)            │
│    单元格对齐: 居中                          │
│    表格字体: 楷体 9pt                        │
│    行间距: 1.0                              │
│                                            │
│  ☑ 自动去除多余空格 ☑ 标题重新编号            │
│  ☑ 处理表格                                  │
│  公式转换: [pandoc    ▼]                     │
│                                            │
│  [           转    换              ]          │
│  ▓▓▓▓▓▓▓▓▓▓▓▓▓▓░░░░░░░░  调用 pandoc 翻译公式  │
│                                            │
│  ▼ 状态                                    │
│  ...                                       │
└────────────────────────────────────────────┘
```

---

## �� 快速开始

### 方法一：下载预编译版本（推荐普通用户）

前往 [Releases](https://github.com/yourname/Md2Docx/releases) 页面下载最新 `Markdown2Docx.exe`，
双击运行即可。**无需安装 Python**。

> ⚠️ 首次启动可能需要 1~2 秒。

### 方法二：从源码运行（推荐开发者）

需要 Python 3.8+。

```bash
# 1. 克隆仓库
git clone https://github.com/yourname/Md2Docx.git
cd Md2Docx

# 2. 安装依赖
pip install -r requirements.txt

# 3. 运行
python main.py
```

### 方法三：自行打包 exe

```bash
# Windows
build.bat

# 或手动
pip install pyinstaller
python -m PyInstaller build.spec
```

打包完成后，`dist\Markdown2Docx.exe` 即可分发。

---

## �� 依赖说明

| 包 | 用途 | 必需 |
|---|---|---|
| `python-docx` | 读写 .docx 文件 | ✅ |
| `latex2mathml` | 内置公式转换（fallback） | ✅ |
| `markdown` | 备用 Markdown 解析 | ✅ |
| `pyinstaller` | 打包成 exe | 仅打包时 |
| **pandoc** | **公式转换（OMML）** | ⭐ 推荐 |

> �� **强烈建议安装 pandoc**。装上之后公式就是 Word 原生公式对象，可双击编辑。
> 没装也能用——内置 `latex2mathml` 会用 MathML 兜底。

安装 pandoc：
- Windows: `choco install pandoc` 或 [pandoc.org](https://pandoc.org/installing.html)
- macOS: `brew install pandoc`
- Linux: `sudo apt install pandoc`

---

## ��️ 项目结构

```
Md2Docx/
├── main.py                      # GUI 入口（tkinter）
├── converter/
│   ├── pandoc_helper.py         # pandoc 探测 / 命令行调用
│   ├── style_patcher.py         # docx 样式补丁（核心）
│   ├── docx_generator.py        # 字体 / 段落设置数据类
│   └── markdown_converter.py    # 备用 Markdown 解析
├── ui/
│   └── settings_dialog.py       # 设置对话框
├── utils/
│   └── text_cleanup.py          # 空格 / 引号 / 编号清理
├── test_files/                  # 集成测试用例
├── build.spec                   # PyInstaller 配置
├── build.bat                    # 一键打包脚本
├── requirements.txt
├── runtime_hook_tk.py           # PyInstaller runtime hook
└── README.md
```

**核心思路：**

```
.md ── pandoc ──> raw.docx (含 OMML 公式)
                       │
                       ▼
              style_patcher.py 给 docx 打样式补丁
              （改 styles.xml + 改 document.xml）
                       │
                       ▼
                 final.docx
```

为什么不直接用 `python-docx` 写公式？因为 `python-docx` 写不出正确的 OMML。
也不直接把 `<m:oMath>` 节点硬编码，因为 LaTeX 语法翻译太复杂。
所以走 pandoc + 样式补丁这条路线。

---

## ⚙️ 配置说明

### 字体设置

| 项 | 说明 |
|---|---|
| 中文字体 | 宋体 / 黑体 / 楷体 / 微软雅黑 / ... |
| 英文字体 | Times New Roman / Arial / Calibri / ... |
| 字号 | Word 标准中文字号（初号 ~ 八号） |

### 段落设置

| 项 | 说明 |
|---|---|
| 行间距 | 1.0 / 1.25 / 1.5 / 1.75 / 2.0 倍 |
| 首行缩进 | 是否启用 + 缩进字符数（中文常用 2） |
| 标题行间距 | 标题段落的行间距 |
| 标题段前/段后距 | 单位：磅 |

### 表格设置（勾选"处理表格"后显示）

| 项 | 说明 |
|---|---|
| 单元格对齐 | 靠左 / 居中 / 靠右 |
| 表格字体 | 与正文相同字段（中英 + 字号） |
| 表格行间距 | 与正文独立 |

### 转换选项

| 项 | 说明 |
|---|---|
| ☑ 自动去除多余空格 | 字母/数字旁的空格、`中文 英文` 之间的空格 |
| ☑ 标题重新编号 | 去掉原有编号，按 1 / 1.1 / 1.1.1 重编 |
| ☑ 处理表格 | 启用表格边框 / 对齐 / 表名 |
| 公式转换 | pandoc（推荐）/ 内置 MathML |

---

## �� 高级用法

### 自定义 pandoc 路径

程序优先从 `PATH` 中查找 pandoc。如果装在非标准位置，可以：

```python
# main.py
from converter.pandoc_helper import md_to_docx_via_pandoc
md_to_docx_via_pandoc(input_path, output_path,
                      pandoc_exe=r'D:\tools\pandoc\pandoc.exe')
```

### 修改样式补丁

样式补丁的核心逻辑在 [converter/style_patcher.py](converter/style_patcher.py)：

- `patch_docx_styles()` —— 入口
- `ensure_table_content_style()` —— 表格段落样式（防止 BodyText 首行缩进泄漏）
- `_patch_tables()` —— 表格边框 / 对齐 / 表名
- `_auto_number_headings()` —— 标题自动编号
- `_replace_in_wt()` —— 跨段落文本操作（空格 / 引号）

### 二次开发

`main.py` 的 `Application._collect_settings()` 返回一个 dict，
传给 `converter.style_patcher.patch_docx_styles()` 即可。
你可以写自己的脚本调用同一套接口做批量转换：

```python
from converter.style_patcher import patch_docx_styles
from converter.docx_generator import FontSettings, ParagraphSettings

settings = {
    'body_font': FontSettings(font_size=10.5, chinese_font='宋体'),
    'heading_fonts': [...],
    'body_para': ParagraphSettings(line_spacing=1.5),
    'process_tables': True,
    'table_cell_alignment': 'center',
    'renumber_headings': True,
}

patch_docx_styles('input.docx', 'output.docx', settings)
```

---

## �� 测试

`test_files/` 目录有几个真实科研文档（Word 复杂公式 + 长中文 + 表格），适合回归测试。

```bash
# 跑一次完整转换
python main.py
# 选择 test_files/6.5.3_数字孪生驱动的快速改造流程.md
# 输出到任意位置，对比 .docx 的渲染效果
```

---

## �� 已知问题

- **pandoc 翻译失败的公式**：极少数 LaTeX 宏包命令 texmath 不认识，会原样输出。能转的公式：分式、积分、矩阵、求和、希腊字母、单位、上下标、根号 —— 95% 日常科研/工程场景够用。
- **emoji 输出**：Markdown 里的 emoji 会以图片形式落到 docx 中，文件会变大。
- **公式块末的空格**：pandoc 转 OMML 时偶尔会在 `$$ ... $$` 之后留下空行，工具未做处理（如有需要可手动删）。

---

## ��️ 路线图

- [ ] 表格从 Caption 字段读取"表名"（目前固定为"表名"）
- [ ] 图片自动居中 + 尺寸调整
- [ ] 代码块带背景色 + 等宽字体
- [ ] 引用块样式
- [ ] 脚注 / 尾注
- [ ] 批量转换（目录级）
- [ ] 命令行模式（无 GUI 转换）

---

## �� 贡献

欢迎 PR / Issue。提交前请：

1. 在 `test_files/` 中放一个能复现的 Markdown 样例
2. 用 `python main.py` 跑一遍确认无回归
3. 写清楚改动原因

---

## �� 许可证

本项目使用 [MIT License](LICENSE)。

---

## �� 致谢

- [pandoc](https://pandoc.org/) —— 公式转换
- [texmath](https://github.com/jgm/tex