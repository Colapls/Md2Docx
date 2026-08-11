"""
Markdown 解析器：将 Markdown 转换为中间表示
"""
import re
import markdown
from markdown.extensions import extra, codehilite, tables, fenced_code
from typing import List, Dict, Any, Tuple, Optional
from dataclasses import dataclass
from enum import Enum


class ElementType(Enum):
    """元素类型枚举"""
    PARAGRAPH = "paragraph"
    HEADING = "heading"
    BLOCKQUOTE = "blockquote"
    CODE_BLOCK = "code_block"
    LIST = "list"
    LIST_ITEM = "list_item"
    HORIZONTAL_RULE = "horizontal_rule"
    FORMULA_BLOCK = "formula_block"
    FORMULA_INLINE = "formula_inline"
    TABLE = "table"


@dataclass
class Heading:
    """标题元素"""
    level: int  # 1-6
    text: str


@dataclass
class Paragraph:
    """段落元素"""
    text: str
    is_italic: bool = False
    is_bold: bool = False


@dataclass
class FormulaInline:
    """行内公式"""
    formula: str


@dataclass
class FormulaBlock:
    """公式块"""
    formula: str


@dataclass
class CodeBlock:
    """代码块"""
    code: str
    language: str


@dataclass
class ListItem:
    """列表项"""
    text: str
    level: int  # 缩进级别


@dataclass
class TableCell:
    """表格单元格"""
    text: str
    is_header: bool


@dataclass
class TableRow:
    """表格行"""
    cells: List[TableCell]


@dataclass
class Table:
    """表格"""
    headers: List[str]
    rows: List[List[str]]


@dataclass
class BlockQuote:
    """引用块"""
    text: str


class MarkdownConverter:
    """Markdown 转换为中间表示"""

    def __init__(self):
        self.md = markdown.Markdown(
            extensions=[
                'markdown.extensions.extra',
                'markdown.extensions.codehilite',
                'markdown.extensions.tables',
                'markdown.extensions.fenced_code',
                'markdown.extensions.nl2br',
            ]
        )
        # 预处理：提取公式块
        self.formula_blocks: List[Tuple[str, str]] = []  # (placeholder, formula)

    def preprocess_formulas(self, text: str) -> Tuple[str, Dict[int, bool]]:
        """
        预处理文本，提取公式并替换为占位符

        Args:
            text: 原始 markdown 文本

        Returns:
            (替换占位符后的文本, 块公式索引集合)
        """
        formula_count = [0]
        block_indices = set()

        def replace_block_formula(match):
            formula = match.group(1)
            placeholder = f'@@FORMULABLOCK{formula_count[0]}@@'
            self.formula_blocks.append((placeholder, formula))
            block_indices.add(formula_count[0])
            formula_count[0] += 1
            # 块公式在 markdown 中独占一行：替换为占位符并保留换行符以确保占位符作为独立段落存在
            return f'@@FORMULABLOCK{formula_count[0] - 1}@@'

        def replace_inline_formula(match):
            formula = match.group(1)
            placeholder = f'@@FORMULAINLINE{formula_count[0]}@@'
            self.formula_blocks.append((placeholder, formula))
            formula_count[0] += 1
            return placeholder

        # 处理 $$...$$ 块公式
        text = re.sub(r'\$\$(.+?)\$\$', replace_block_formula, text, flags=re.DOTALL)
        # 处理 $...$ 行内公式
        text = re.sub(r'\$(.+?)\$', replace_inline_formula, text)

        return text, block_indices

    def convert(self, markdown_text: str) -> List[Dict[str, Any]]:
        """
        将 Markdown 文本转换为中间表示

        Args:
            markdown_text: Markdown 格式文本

        Returns:
            元素列表
        """
        # 重置状态
        self.formula_blocks = []
        self.formula_block_indices = set()

        # 预处理：提取公式
        processed_text, block_indices = self.preprocess_formulas(markdown_text)
        self.formula_block_indices = block_indices

        # 解析 HTML
        html = self.md.convert(processed_text)

        # 解析元素
        elements = self._parse_html_elements(html)

        return elements

    def _parse_html_elements(self, html: str) -> List[Dict[str, Any]]:
        """
        解析 HTML 元素

        Args:
            html: HTML 文本

        Returns:
            元素列表
        """
        elements = []

        # 分割 HTML 标签
        current_text = ""
        in_pre = False
        in_code = False

        lines = html.split('\n')

        for line in lines:
            stripped = line.strip()

            # 处理代码块
            if '<pre>' in stripped or '<code>' in stripped:
                in_pre = True
            if '</pre>' in stripped or '</code>' in stripped:
                in_pre = False
                in_code = False
                continue

            # 跳过空行
            if not stripped:
                continue

            # 处理标题
            heading_match = re.match(r'<h([1-6])>(.+?)</h[1-6]>', stripped, re.DOTALL)
            if heading_match:
                level = int(heading_match.group(1))
                text = self._strip_tags(heading_match.group(2))
                elements.append({
                    'type': ElementType.HEADING,
                    'level': level,
                    'text': text
                })
                continue

            # 处理引用块
            if stripped.startswith('<blockquote'):
                text = self._strip_tags(stripped)
                elements.append({
                    'type': ElementType.BLOCKQUOTE,
                    'text': text
                })
                continue

            # 处理水平线
            if '<hr' in stripped:
                elements.append({
                    'type': ElementType.HORIZONTAL_RULE,
                })
                continue

            # 处理列表
            if '<ul>' in stripped or '<ol>' in stripped:
                continue
            if '</ul>' in stripped or '</ol>' in stripped:
                continue

            list_match = re.match(r'<li>(.+?)</li>', stripped, re.DOTALL)
            if list_match:
                text = self._strip_tags(list_match.group(1))
                elements.append({
                    'type': ElementType.LIST_ITEM,
                    'text': text
                })
                continue

            # 处理段落
            if '<p>' in stripped:
                # 检查段落是否只包含块公式占位符
                block_only = self._extract_block_only(stripped)
                if block_only is not None:
                    elements.append({
                        'type': ElementType.FORMULA_BLOCK,
                        'formula': block_only
                    })
                    continue

                text = self._strip_tags(stripped)
                elements.append({
                    'type': ElementType.PARAGRAPH,
                    'text': text
                })
                continue

            # 检查独立的块公式占位符行（不在 <p> 标签内的公式）
            standalone = self._extract_standalone_block(stripped)
            if standalone is not None:
                elements.append({
                    'type': ElementType.FORMULA_BLOCK,
                    'formula': standalone
                })
                continue

        return elements

    def _extract_block_only(self, html: str) -> Optional[str]:
        """如果段落内只包含一个块公式占位符，返回公式内容；否则返回 None"""
        inner = re.search(r'<p>(.*?)</p>', html, re.DOTALL)
        if not inner:
            return None
        body = inner.group(1).strip()
        # 仅有一个块公式占位符
        m = re.fullmatch(r'@@FORMULABLOCK(\d+)@@', body)
        if m:
            idx = int(m.group(1))
            if idx in self.formula_block_indices:
                return self.formula_blocks[idx][1]
        return None

    def _extract_standalone_block(self, stripped: str) -> Optional[str]:
        """检查独立行上的块公式占位符"""
        m = re.fullmatch(r'@@FORMULABLOCK(\d+)@@', stripped)
        if m:
            idx = int(m.group(1))
            if idx in self.formula_block_indices:
                return self.formula_blocks[idx][1]
        return None

    def _strip_tags(self, html: str) -> str:
        """
        移除 HTML 标签

        Args:
            html: HTML 文本

        Returns:
            纯文本
        """
        # 恢复公式占位符中的 $ 符号
        text = re.sub(r'@@FORMULABLOCK(\d+)@@',
                      lambda m: f'$${self.formula_blocks[int(m.group(1))][1]}$$',
                      html)
        text = re.sub(r'@@FORMULAINLINE(\d+)@@',
                     lambda m: f'${self.formula_blocks[int(m.group(1))][1]}$',
                     text)

        # 移除其他 HTML 标签
        text = re.sub(r'<[^>]+>', '', text)
        return text.strip()

    def get_formula_blocks(self) -> List[Tuple[str, str]]:
        """获取公式块列表"""
        return self.formula_blocks
