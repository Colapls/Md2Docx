"""
Style patcher for docx files produced by pandoc.

After pandoc turns Markdown into docx, the document structure is in
place (real OMML equations, headings, tables, etc.) but the typography
matches pandoc's defaults, not the user's settings.

This module rewrites the OOXML inside the docx zip so that:

  * docDefaults uses the user's body font/size/line-spacing.
  * Heading1..Heading6 use the per-level heading fonts/sizes.
  * BodyText paragraphs get the user's first-line indent.
  * sectPr (page setup) uses A4 with reasonable margins.
  * Anything pandoc left as theme-font references becomes an explicit
    font name so Word / WPS doesn't fall back to Calibri.
"""
from __future__ import annotations

import io
import re
import shutil
import zipfile
from typing import Dict, Any, Optional


# ---------------------------------------------------------------------------
# Low-level XML helpers (string-based, kept simple on purpose).
# ---------------------------------------------------------------------------

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def _half_pt(pt: float) -> int:
    """Word stores font sizes in half-points (sz val)."""
    return int(round(float(pt) * 2))


def _line_twips(multiple: float) -> int:
    """Convert a multiple line-spacing (e.g. 1.25) into line value with lineRule=auto.

    For lineRule="auto", line = 240 * multiple. 240 = single line in twips.
    """
    return int(round(240.0 * float(multiple)))


def _first_line_indent_chars(chars: int, font_pt: float) -> int:
    """For w:ind firstLine, Word expects twentieths-of-a-point.
    Use firstLineChars when we want the indent expressed in characters
    (which is what Word's "首行缩进 N 字符" means)."""
    # firstLineChars is in hundredths of a character (1 char = 100).
    return int(round(float(chars) * 100))


# ---------------------------------------------------------------------------
# docDefaults block (replaces theme-font rFonts with explicit fonts).
# ---------------------------------------------------------------------------

def _font_get(font, key, default):
    """Read a font setting from either a dict or a FontSettings-like object."""
    if isinstance(font, dict):
        return font.get(key, default)
    return getattr(font, key, default)


def build_doc_defaults(body_font, body_line_spacing: float) -> str:
    cn = _font_get(body_font, 'chinese_font', '宋体')
    en = _font_get(body_font, 'english_font', 'Times New Roman')
    sz = _font_get(body_font, 'font_size', 10.5)
    line = _line_twips(body_line_spacing)

    return (
        '<w:docDefaults>'
        '<w:rPrDefault><w:rPr>'
        f'<w:rFonts w:ascii="{en}" w:hAnsi="{en}" w:eastAsia="{cn}" w:cs="{en}"/>'
        f'<w:sz w:val="{_half_pt(sz)}"/>'
        f'<w:szCs w:val="{_half_pt(sz)}"/>'
        '<w:lang w:bidi="ar-SA" w:eastAsia="zh-CN" w:val="en-US"/>'
        '</w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr>'
        f'<w:spacing w:after="0" w:line="{line}" w:lineRule="auto"/>'
        '</w:pPr></w:pPrDefault>'
        '</w:docDefaults>'
    )


# ---------------------------------------------------------------------------
# Per-heading style block. We emit explicit font/size/color/bold info so
# Word stops falling back to theme fonts.
# ---------------------------------------------------------------------------

def _heading_block(style_id: str, name: str, font,
                   line_spacing: float, space_before_pt: float,
                   space_after_pt: float, *, bold: bool) -> str:
    cn = _font_get(font, 'chinese_font', '黑体')
    en = _font_get(font, 'english_font', 'Times New Roman')
    sz = _font_get(font, 'font_size', 14)
    line = _line_twips(line_spacing)
    sb = int(round(float(space_before_pt) * 20))   # pt -> twips
    sa = int(round(float(space_after_pt) * 20))

    bold_xml = '<w:b/><w:bCs/>' if bold else ''

    return (
        f'<w:style w:styleId="{style_id}" w:type="paragraph">'
        f'<w:name w:val="{name}"/>'
        '<w:basedOn w:val="Normal"/>'
        '<w:next w:val="BodyText"/>'
        '<w:qFormat/>'
        '<w:pPr>'
        f'<w:spacing w:before="{sb}" w:after="{sa}" w:line="{line}" w:lineRule="auto"/>'
        '<w:outlineLvl w:val="' + str(int(re.search(r'\d+', style_id).group())) + '"/>'
        '</w:pPr>'
        '<w:rPr>'
        f'<w:rFonts w:ascii="{en}" w:hAnsi="{en}" w:eastAsia="{cn}" w:cs="{en}"/>'
        f'{bold_xml}'
        f'<w:sz w:val="{_half_pt(sz)}"/>'
        f'<w:szCs w:val="{_half_pt(sz)}"/>'
        '</w:rPr>'
        '</w:style>'
    )


def build_heading_styles(heading_fonts, heading_line_spacing: float,
                         space_before_pt: float, space_after_pt: float) -> str:
    blocks = []
    levels = [('Heading1', 'heading 1', True),
              ('Heading2', 'heading 2', True),
              ('Heading3', 'heading 3', True),
              ('Heading4', 'heading 4', True),
              ('Heading5', 'heading 5', True),
              ('Heading6', 'heading 6', False)]
    for i, (sid, name, bold) in enumerate(levels):
        if i < len(heading_fonts):
            font = heading_fonts[i]
        else:
            font = heading_fonts[-1]
        blocks.append(_heading_block(
            sid, name, font, heading_line_spacing,
            space_before_pt, space_after_pt, bold=bold))
    return ''.join(blocks)


# ---------------------------------------------------------------------------
# Body text style block (BodyText is what pandoc uses for paragraphs).
# We also touch FirstParagraph, which is based on BodyText, so any
# indent we set on BodyText cascades to FirstParagraph automatically.
# ---------------------------------------------------------------------------

def build_body_text_style(body_font,
                          line_spacing: float,
                          indent_chars: int,
                          indent_enabled: bool) -> str:
    cn = _font_get(body_font, 'chinese_font', '宋体')
    en = _font_get(body_font, 'english_font', 'Times New Roman')
    sz = _font_get(body_font, 'font_size', 10.5)
    line = _line_twips(line_spacing)

    if indent_enabled and indent_chars > 0:
        indent_xml = (
            f'<w:ind w:firstLineChars="{_first_line_indent_chars(indent_chars, sz)}" '
            'w:firstLine="0"/>'
        )
        # Actually firstLineChars needs a fallback firstLine twips for
        # renderers that ignore firstLineChars. Compute it in twips:
        first_line_twips = int(round(float(indent_chars) * float(sz) * 20))
        indent_xml = (
            f'<w:ind w:firstLineChars="{_first_line_indent_chars(indent_chars, sz)}" '
            f'w:firstLine="{first_line_twips}"/>'
        )
    else:
        indent_xml = '<w:ind w:firstLine="0"/>'

    return (
        '<w:style w:styleId="BodyText" w:type="paragraph">'
        '<w:name w:val="Body Text"/>'
        '<w:basedOn w:val="Normal"/>'
        '<w:link w:val="BodyTextChar"/>'
        '<w:qFormat/>'
        '<w:pPr>'
        f'<w:spacing w:after="0" w:line="{line}" w:lineRule="auto"/>'
        f'{indent_xml}'
        '<w:jc w:val="both"/>'
        '</w:pPr>'
        '<w:rPr>'
        f'<w:rFonts w:ascii="{en}" w:hAnsi="{en}" w:eastAsia="{cn}" w:cs="{en}"/>'
        f'<w:sz w:val="{_half_pt(sz)}"/>'
        f'<w:szCs w:val="{_half_pt(sz)}"/>'
        '</w:rPr>'
        '</w:style>'
    )


# ---------------------------------------------------------------------------
# Section properties (A4 page, normal margins).
# ---------------------------------------------------------------------------

def build_sect_pr() -> str:
    # A4: 11906 x 16838 twips. Margins: top/bottom 1440 (1 inch),
    # left/right 1440 (1 inch). Header/footer 720 (0.5 inch).
    return (
        '<w:sectPr>'
        '<w:footnotePr><w:numFmt w:val="decimal"/><w:numRestart w:val="eachSect"/></w:footnotePr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="708"/>'
        '<w:docGrid w:linePitch="360"/>'
        '</w:sectPr>'
    )


# ---------------------------------------------------------------------------
# Post-processing on document.xml content (run inside w:t nodes).
# ---------------------------------------------------------------------------

# Patterns that match an existing leading number/prefix on a heading.
# Same set the original python-docx path used.
_HEADING_NUMBER_PATTERNS = [
    # 中文括号编号：（一）、（1）
    r'^[\（\(][一二三四五六七八九十百千零\d]+[\）\)][\s、.]*',
    # 拉丁字母括号编号：（a）、（A）
    r'^[\（\(][a-zA-Zα-ωΑ-Ω]+[\）\)][\s、.]*',
    # 中文数字+标点：一、二、三、
    r'^[一二三四五六七八九十百千零]+[、.][\s]*',
    # 多级数字编号：1.、1.1、1.1.1、1-
    r'^\d+(\.\d+)*[-.、\s]+',
    # 单级数字编号：1、2、3、
    r'^\d+[.、)）\s]+',
    # 拉丁字母编号：a.、A.、a)
    r'^[a-zA-Z][.、)）\s]+',
    # 第几章/第几节：第一章、第一节
    r'^第[一二三四五六七八九十百千\d]+[节章节]?[\s]*',
]


def _strip_existing_numbering(text: str) -> str:
    for pat in _HEADING_NUMBER_PATTERNS:
        text = re.sub(pat, '', text)
    return text.strip()


# Match a complete <w:t ...>TEXT</w:t> element so we can rewrite its
# inner text while preserving the surrounding w:r.
_W_T_RE = re.compile(r'(<w:t(?:\s[^>]*)?>)([^<]*)(</w:t>)')


def _replace_in_wt(doc_xml: str, transform) -> str:
    """Run `transform(text) -> new_text` on every w:t text fragment."""
    def _sub(match):
        open_tag, text, close_tag = match.group(1), match.group(2), match.group(3)
        new_text = transform(text)
        if new_text == text:
            return match.group(0)
        return f'{open_tag}{new_text}{close_tag}'
    return _W_T_RE.sub(_sub, doc_xml)


def _convert_smart_quotes(text: str) -> str:
    """Pairwise convert ASCII quotes into Chinese curly quotes.

    Pandoc's smart-quote pass on Chinese text is buggy: it often turns
    every " into a right curly quote. We rebuild the pairing here.
    """
    out = []
    even = True
    for ch in text:
        if ch == '"':
            out.append('“' if even else '”')
            even = not even
        elif ch == "'":
            out.append('‘' if even else '’')
            even = not even
        else:
            out.append(ch)
    return ''.join(out)


def fix_smart_quotes_in_document(doc_xml: str) -> str:
    """Convert ASCII " and ' inside text fragments into Chinese curly
    quotes, alternating left/right. Also normalizes pandoc's broken
    pairing by collapsing all curly quotes back to ASCII first."""
    # First: collapse any curly quotes back to ASCII so our pairwise
    # replacement gets a clean slate.
    def _collapse(text: str) -> str:
        return (text.replace('“', '"')
                    .replace('”', '"')
                    .replace('‘', "'")
                    .replace('’', "'"))
    doc_xml = _replace_in_wt(doc_xml, _collapse)
    # Then: pairwise ASCII -> curly.
    return _replace_in_wt(doc_xml, _convert_smart_quotes)


# Heading numbering state (re-used across all heading paragraphs).
class _HeadingNumberer:
    def __init__(self):
        self.counters = [0] * 9

    def next_number_for(self, level: int) -> str:
        if level < 1:
            level = 1
        if level > 9:
            level = 9
        self.counters[level - 1] += 1
        for i in range(level, 9):
            self.counters[i] = 0
        return '.'.join(str(self.counters[i]) for i in range(level)) + ' '


def _auto_number_headings(doc_xml: str) -> str:
    """For every paragraph whose pStyle is Heading1..Heading6, strip the
    leading existing number and prepend our own 1 / 1.1 / 1.1.1 style
    number. State advances across paragraphs."""
    numberer = _HeadingNumberer()
    # Match an entire <w:p>...</w:p> block that contains a heading pStyle.
    p_re = re.compile(r'<w:p\b[^>]*>.*?</w:p>', re.DOTALL)
    style_re = re.compile(r'<w:pStyle\s+w:val="(Heading[1-6])"\s*/>')

    def _rewrite_paragraph(p_xml: str) -> str:
        m = style_re.search(p_xml)
        if not m:
            return p_xml
        level = int(m.group(1)[-1])
        number = numberer.next_number_for(level)

        # Collect all w:t text fragments inside this paragraph and
        # rewrite them as:  number + stripped-text + (any remaining text).
        wt_re = re.compile(r'(<w:t(?:\s[^>]*)?>)([^<]*)(</w:t>)')
        fragments = wt_re.findall(p_xml)
        if not fragments:
            return p_xml

        # Combine all text, strip the leading number, split back into
        # fragments using a fixed rule: first fragment gets
        # "number + stripped", others keep their original text.
        all_text = ''.join(text for _, text, _ in fragments)
        stripped = _strip_existing_numbering(all_text)

        # If stripped is empty (entire heading was just a number),
        # bail out — don't produce an empty heading.
        if not stripped:
            return p_xml

        # Rebuild: prefix first fragment with "number"; clear other
        # fragments because the user text is now in the first one.
        new_fragments = []
        first_done = False
        for open_tag, _, close_tag in fragments:
            if not first_done:
                new_fragments.append(f'{open_tag}{number}{stripped}{close_tag}')
                first_done = True
            else:
                new_fragments.append(f'{open_tag}{close_tag}')

        # Splice: replace each <w:t> in order.
        out = []
        idx = 0
        last_end = 0
        for m2 in wt_re.finditer(p_xml):
            out.append(p_xml[last_end:m2.start()])
            out.append(new_fragments[idx])
            last_end = m2.end()
            idx += 1
        out.append(p_xml[last_end:])
        return ''.join(out)

    return p_re.sub(lambda m: _rewrite_paragraph(m.group(0)), doc_xml)


def _remove_horizontal_rules(doc_xml: str) -> str:
    """Drop paragraphs whose only content is a horizontal-rule VML pict
    (pandoc emits <w:pict><v:rect ... o:hr="t" ... /></w:pict> for ---)."""
    # Match paragraphs that contain a v:rect with the o:hr attribute
    # (pandoc's horizontal rule marker). The whole <w:p>...</w:p> is
    # removed including any leading/trailing whitespace runs.
    hr_re = re.compile(
        r'<w:p\b[^>]*>(?:(?!</w:p>).)*?<v:rect\b[^>]*o:hr="t"[^>]*/>(?:(?!</w:p>).)*?</w:p>',
        re.DOTALL,
    )
    return hr_re.sub('', doc_xml)


# Whitespace cleanup rules applied to every <w:t> text fragment.
# Mirrors utils.text_cleanup.cleanup_spaces so behaviour matches the
# original python-docx path.
#
# Important: the rule for ASCII + space + CJK is intentionally
# ONE-WAY only — we collapse "中文 数字" → "中文数字" but we leave
# "数字 中文" alone, because heading auto-numbering inserts
# "1 标题" into that exact shape.
_SPACE_RULES = [
    # 字母/数字之间的多余空格： "a b" -> "ab",  "1 2" -> "12"
    (re.compile(r'([a-zA-Z0-9])\s+([a-zA-Z0-9])'), r'\1\2'),
    # 中文 → 英文/数字："中 a" -> "中a"
    (re.compile(r'([一-龥])\s+([a-zA-Z0-9])'), r'\1\2'),
    # 多个连续空格压缩为 1 个
    (re.compile(r' {2,}'), ' '),
]


def _collapse_spaces(text: str) -> str:
    """Strip the kinds of stray whitespace that show up around
    alphanumeric / CJK character boundaries. Same rules as
    utils.text_cleanup.cleanup_spaces."""
    for pat, repl in _SPACE_RULES:
        text = pat.sub(repl, text)
    return text.lstrip(' ').rstrip(' ')


def fix_spaces_in_document(doc_xml: str) -> str:
    """Apply _collapse_spaces to every <w:t> text fragment."""
    return _replace_in_wt(doc_xml, _collapse_spaces)


# ---------------------------------------------------------------------------
# Table processing: borders, cell alignment, and a generated caption
# paragraph ("表N 表名") inserted right above each table.
# ---------------------------------------------------------------------------

# OOXML alignment values: 'left' / 'center' / 'right' / 'both'.
_TABLE_CELL_JC = {
    'left':   'left',
    'center': 'center',
    'right':  'right',
}

# A single empty paragraph injected right after every table so the
# following body text doesn't sit flush against the table.
# Uses Normal pStyle (not BodyText) so it won't pick up the 2-char
# first-line indent.
_BLANK_PARAGRAPH = '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>'


def _build_tbl_borders() -> str:
    """Build a <w:tblBorders> block with all six sides set to single
    hairlines. Injected into each <w:tblPr>."""
    sides = ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')
    parts = []
    for side in sides:
        parts.append(
            f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        )
    return '<w:tblBorders>' + ''.join(parts) + '</w:tblBorders>'


def _build_tbl_caption_paragraph(label: str, body_font) -> str:
    """A standalone centered paragraph used as a table caption
    ("表1 表名"). Uses the same fonts as the body text, with a size
    one notch smaller, no italic."""
    cn = _font_get(body_font, 'chinese_font', '宋体')
    en = _font_get(body_font, 'english_font', 'Times New Roman')
    # Caption size = body size minus one step (>= 8pt minimum).
    body_sz = _font_get(body_font, 'font_size', 10.5)
    cap_sz = max(8, body_sz - 1.5)
    return (
        '<w:p>'
        '<w:pPr><w:pStyle w:val="Caption"/>'
        '<w:spacing w:before="120" w:after="60"/>'
        '<w:jc w:val="center"/>'
        '</w:pPr>'
        f'<w:r><w:rPr>'
        f'<w:rFonts w:ascii="{en}" w:hAnsi="{en}" w:eastAsia="{cn}" w:cs="{en}"/>'
        f'<w:sz w:val="{_half_pt(cap_sz)}"/>'
        f'<w:szCs w:val="{_half_pt(cap_sz)}"/>'
        '</w:rPr>'
        f'<w:t xml:space="preserve">{label}</w:t></w:r>'
        '</w:p>'
    )


def _patch_tables(doc_xml: str, body_font, table_font=None,
                  line_spacing: float = 1.0,
                  cell_alignment: str = 'center',
                  add_caption: bool = True) -> str:
    """For each <w:tbl>...</w:tbl>:
      * inject <w:tblBorders> (full hairline grid) into its <w:tblPr>;
      * set <w:jc w:val="center"/> on the table itself so it floats to
        the middle of the page;
      * rebind every <w:pStyle w:val="Compact"/> inside cells to our
        own TableContent style (which is basedOn Normal, has NO first
        line indent and uses the table paragraph line spacing) so body
        paragraph settings (indent / fonts) never leak into cells;
      * force a zero-indent + the configured line spacing + cell
        alignment on each inner <w:p>/<w:pPr>;
      * rewrite run rFonts / size to the configured table font so the
        cell text never falls back to BodyText fonts;
      * insert a centered "表N 表名" caption paragraph immediately
        before the table (auto-incrementing N).
    Returns the rewritten doc_xml.
    """
    jc = _TABLE_CELL_JC.get(cell_alignment, 'center')
    borders = _build_tbl_borders()
    # If the caller didn't pass a separate table font, fall back to the
    # body font for the run rPr.
    tf = table_font or body_font
    cn = _font_get(tf, 'chinese_font', '宋体')
    en = _font_get(tf, 'english_font', 'Times New Roman')
    sz = _font_get(tf, 'font_size', 10.5)
    sz_half = _half_pt(sz)
    # line spacing -> twips; 1.0 = 240, 1.5 = 360 etc.
    line_twips = int(round(240.0 * float(line_spacing)))

    run_rpr = (
        f'<w:rPr>'
        f'<w:rFonts w:ascii="{en}" w:hAnsi="{en}" '
        f'w:eastAsia="{cn}" w:cs="{en}"/>'
        f'<w:sz w:val="{sz_half}"/>'
        f'<w:szCs w:val="{sz_half}"/>'
        '</w:rPr>'
    )

    table_re = re.compile(r'<w:tbl>.*?</w:tbl>', re.DOTALL)
    counter = [0]

    def _rewrite_table(tbl_xml: str) -> str:
        counter[0] += 1
        caption = (f'表{counter[0]} 表名' if add_caption else '')

        # 1) inject <w:tblBorders> at the end of <w:tblPr> and ensure
        #    the table is horizontally centered on the page.
        tblpr_re = re.compile(r'<w:tblPr>(.*?)</w:tblPr>', re.DOTALL)
        def _inject_tblpr(match):
            inner = match.group(1)
            # Drop any existing tblJc to avoid duplicates.
            inner = re.sub(r'<w:jc[^/]*/>', '', inner)
            return f'<w:tblPr>{inner}{borders}<w:jc w:val="center"/></w:tblPr>'
        tbl_xml = tblpr_re.sub(_inject_tblpr, tbl_xml, count=1)

        # 2) inject <w:vAlign w:val="center"/> into every <w:tcPr/>
        #    and rewrite every inner <w:p>/<w:pPr> to:
        #      - use the TableContent pStyle (basedOn Normal, no indent)
        #      - drop any existing firstLine indent
        #      - use the configured line spacing
        #      - use the requested cell alignment
        #    Plus rewrite every <w:r>/<w:rPr> to the table font/size.
        tc_re = re.compile(r'<w:tc>(.*?)</w:tc>', re.DOTALL)
        def _fix_cell(match):
            cell = match.group(1)
            # tcPr: ensure non-empty, then add vAlign
            tcpr_re = re.compile(r'<w:tcPr\s*/>')
            if tcpr_re.search(cell):
                cell = tcpr_re.sub(
                    '<w:tcPr><w:vAlign w:val="center"/></w:tcPr>',
                    cell, count=1,
                )
            else:
                cell = re.sub(
                    r'<w:tcPr>',
                    '<w:tcPr><w:vAlign w:val="center"/>',
                    cell, count=1,
                )

            # Rewrite every <w:p> inside this cell.
            p_re = re.compile(r'<w:p>(.*?)</w:p>', re.DOTALL)
            def _fix_paragraph(pm):
                inner = pm.group(1)
                # 2a) pPr: replace pStyle + zero out ind + line spacing + jc
                ppr_re = re.compile(r'<w:pPr>(.*?)</w:pPr>', re.DOTALL)
                new_ppr = (
                    '<w:pPr>'
                    '<w:pStyle w:val="TableContent"/>'
                    f'<w:spacing w:after="0" w:line="{line_twips}" w:lineRule="auto"/>'
                    '<w:ind w:firstLine="0" w:firstLineChars="0"/>'
                    f'<w:jc w:val="{jc}"/>'
                    '</w:pPr>'
                )
                if ppr_re.search(inner):
                    inner = ppr_re.sub(new_ppr, inner, count=1)
                else:
                    inner = new_ppr + inner

                # 2b) rewrite every <w:r>...</w:r> to use the table
                #     font + size (replace or insert rPr).
                run_re = re.compile(r'<w:r(?:\s[^>]*)?>(.*?)</w:r>', re.DOTALL)
                def _fix_run(rm):
                    # rm.group(0) is the entire <w:r>...</w:r> span
                    # (including the opening tag), so we can safely
                    # search for both the opening tag and any rPr.
                    span = rm.group(0)
                    rrpr_re = re.compile(r'<w:rPr>(.*?)</w:rPr>', re.DOTALL)
                    if rrpr_re.search(span):
                        new_span = rrpr_re.sub(run_rpr, span, count=1)
                    else:
                        # Insert rPr right after the <w:r ...> tag.
                        new_span = re.sub(
                            r'(<w:r(?:\s[^>]*)?>)',
                            r'\1' + run_rpr,
                            span, count=1,
                        )
                    return new_span
                inner = run_re.sub(_fix_run, inner)
                return f'<w:p>{inner}</w:p>'

            cell = p_re.sub(_fix_paragraph, cell)
            return f'<w:tc>{cell}</w:tc>'
        tbl_xml = tc_re.sub(_fix_cell, tbl_xml)

        # 3) caption: prepend a caption paragraph.
        # 4) Always append a single empty paragraph after the table so
        #    there's a blank line between the table and the following
        #    body text (tables otherwise sit flush against the next
        #    paragraph).
        if add_caption:
            return (_build_tbl_caption_paragraph(caption, body_font)
                    + tbl_xml
                    + _BLANK_PARAGRAPH)
        return tbl_xml + _BLANK_PARAGRAPH

    return table_re.sub(lambda m: _rewrite_table(m.group(0)), doc_xml)


# Build the Caption style XML using the user's body font so captions
# visually match the body text (same font, one size smaller, no italic).
def _build_caption_style_xml(body_font) -> str:
    cn = _font_get(body_font, 'chinese_font', '宋体')
    en = _font_get(body_font, 'english_font', 'Times New Roman')
    body_sz = _font_get(body_font, 'font_size', 10.5)
    cap_sz = max(8, body_sz - 1.5)
    return (
        '<w:style w:styleId="Caption" w:type="paragraph">'
        '<w:name w:val="caption"/>'
        '<w:basedOn w:val="Normal"/>'
        '<w:next w:val="Normal"/>'
        '<w:qFormat/>'
        '<w:pPr>'
        '<w:spacing w:before="120" w:after="120"/>'
        '<w:jc w:val="center"/>'
        '</w:pPr>'
        '<w:rPr>'
        f'<w:rFonts w:ascii="{en}" w:hAnsi="{en}" '
        f'w:eastAsia="{cn}" w:cs="{en}"/>'
        f'<w:sz w:val="{_half_pt(cap_sz)}"/>'
        f'<w:szCs w:val="{_half_pt(cap_sz)}"/>'
        '</w:rPr>'
        '</w:style>'
    )


def ensure_caption_style(styles_xml: str, body_font=None) -> str:
    """Inject (or REPLACE) the Caption style so it inherits the user's
    body font, a size one step smaller, and no italic.

    We always overwrite any existing Caption style so that pandoc's
    default italic caption doesn't leak through."""
    style_xml = _build_caption_style_xml(body_font) if body_font is not None \
        else (
            '<w:style w:styleId="Caption" w:type="paragraph">'
            '<w:name w:val="caption"/>'
            '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/>'
            '<w:qFormat/>'
            '<w:pPr><w:spacing w:before="120" w:after="120"/>'
            '<w:jc w:val="center"/></w:pPr>'
            '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
            '</w:style>'
        )
    # If a Caption style already exists, replace it in place.
    cap_pat = re.compile(
        r'<w:style\b[^>]*w:styleId="Caption"[^>]*>.*?</w:style>',
        re.DOTALL,
    )
    if cap_pat.search(styles_xml):
        return cap_pat.sub(style_xml, styles_xml, count=1)
    return styles_xml.replace('</w:styles>', style_xml + '</w:styles>')


# ---------------------------------------------------------------------------
# TableContent paragraph style: cells must NOT inherit BodyText's first
# line indent or its line spacing. This style is basedOn Normal (so it
# gets docDefaults fonts but no indent / spacing) and explicitly zeroes
# out firstLine indent.
# ---------------------------------------------------------------------------

def _build_table_content_style_xml(line_spacing: float = 1.0,
                                   cell_alignment: str = 'center') -> str:
    line_twips = int(round(240.0 * float(line_spacing)))
    return (
        '<w:style w:styleId="TableContent" w:type="paragraph" w:customStyle="1">'
        '<w:name w:val="Table Content"/>'
        '<w:basedOn w:val="Normal"/>'
        '<w:qFormat/>'
        '<w:pPr>'
        f'<w:spacing w:after="0" w:line="{line_twips}" w:lineRule="auto"/>'
        '<w:ind w:firstLine="0" w:firstLineChars="0"/>'
        f'<w:jc w:val="{_TABLE_CELL_JC.get(cell_alignment, "center")}"/>'
        '</w:pPr>'
        '</w:style>'
    )


def ensure_table_content_style(styles_xml: str, line_spacing: float = 1.0,
                               cell_alignment: str = 'center') -> str:
    """Inject (or REPLACE) the TableContent style. Always overwrite so
    the user's current line_spacing / cell_alignment takes effect."""
    new_xml = _build_table_content_style_xml(line_spacing, cell_alignment)
    pat = re.compile(
        r'<w:style\b[^>]*w:styleId="TableContent"[^>]*>.*?</w:style>',
        re.DOTALL,
    )
    if pat.search(styles_xml):
        return pat.sub(new_xml, styles_xml, count=1)
    return styles_xml.replace('</w:styles>', new_xml + '</w:styles>')


def detach_compact_from_body_text(styles_xml: str) -> str:
    """Pandoc's Compact style is basedOn BodyText, which inherits the
    2-character first-line indent. If the user has table processing off
    but the doc still contains Compact references, we still want cells
    to behave as table cells. Rewrite Compact to be basedOn Normal and
    zero out its indent."""
    pat = re.compile(
        r'(<w:style\b[^>]*w:styleId="Compact"[^>]*>.*?<w:basedOn w:val=")([^"]+)(")',
        re.DOTALL,
    )
    styles_xml = pat.sub(r'\1Normal\3', styles_xml)

    # Force firstLine indent to zero (compact-specific block may carry
    # one or none; be safe and clear).
    compact_pat = re.compile(
        r'(<w:style\b[^>]*w:styleId="Compact"[^>]*>.*?)(</w:style>)',
        re.DOTALL,
    )
    def _rewrite(m):
        body = m.group(1)
        # Strip any existing <w:ind .../> in the Compact pPr.
        body = re.sub(r'<w:ind\b[^/]*/>', '', body)
        # Insert a zero-indent inside the pPr block (or create one).
        ppr_pat = re.compile(r'<w:pPr>(.*?)</w:pPr>', re.DOTALL)
        if ppr_pat.search(body):
            body = ppr_pat.sub(
                lambda mm: f'<w:pPr><w:ind w:firstLine="0" w:firstLineChars="0"/>{mm.group(1)}</w:pPr>',
                body, count=1,
            )
        else:
            body = body + '<w:pPr><w:ind w:firstLine="0" w:firstLineChars="0"/></w:pPr>'
        return body + m.group(2)
    return compact_pat.sub(_rewrite, styles_xml)


# ---------------------------------------------------------------------------
# The actual patch entry point.
# ---------------------------------------------------------------------------

def patch_docx_styles(src_path: str, dst_path: str, settings: Dict[str, Any]) -> str:
    """
    Patch a docx file in-place (or to dst_path if different).

    `settings` follows the same shape as the GUI builds:
      {
        'heading_fonts': [FontSettings() * 6],
        'body_font':     FontSettings(),
        'heading_para':  ParagraphSettings(),
        'body_para':     ParagraphSettings(),
      }
    """
    # Read everything from the source zip into memory.
    with zipfile.ZipFile(src_path, 'r') as zin:
        parts = {n: zin.read(n) for n in zin.namelist()}

    # ---- styles.xml ----
    if 'word/styles.xml' not in parts:
        raise RuntimeError('word/styles.xml not found in source docx')
    styles_xml = parts['word/styles.xml'].decode('utf-8')

    body_font = settings['body_font']
    body_para = settings['body_para']
    heading_fonts = settings['heading_fonts']
    heading_para = settings['heading_para']

    # Optional knobs (default off so callers that don't know about them
    # keep the previous behaviour).
    renumber_headings = bool(settings.get('renumber_headings', True))
    remove_hr = bool(settings.get('remove_horizontal_rules', True))
    fix_quotes = bool(settings.get('fix_smart_quotes', True))
    fix_spaces = bool(settings.get('fix_spaces', True))
    process_tables = bool(settings.get('process_tables', False))
    table_cell_alignment = settings.get('table_cell_alignment', 'center')
    table_font = settings.get('table_font')
    table_para = settings.get('table_para')
    table_line_spacing = float(getattr(table_para, 'line_spacing', 1.0)) if table_para else 1.0

    new_defaults = build_doc_defaults(body_font, body_para.line_spacing)
    new_headings = build_heading_styles(
        heading_fonts,
        heading_para.heading_line_spacing,
        # pt -> raw twips-ish: heading_space_* are already docx.shared.Pt
        heading_para.heading_space_before.pt,
        heading_para.heading_space_after.pt,
    )
    new_body_text = build_body_text_style(
        body_font,
        body_para.line_spacing,
        body_para.first_line_indent,
        body_para.first_line_indent_enabled,
    )

    # Replace the existing docDefaults block.
    styles_xml = re.sub(
        r'<w:docDefaults>.*?</w:docDefaults>',
        new_defaults,
        styles_xml,
        count=1,
        flags=re.DOTALL,
    )

    # Replace each Heading1..Heading6 style block.
    for i in range(1, 7):
        sid = f'Heading{i}'
        # pandoc's <w:style ... w:styleId="HeadingN" ... > ... </w:style>
        pattern = re.compile(
            r'<w:style\b[^>]*w:styleId="' + sid + r'"[^>]*>.*?</w:style>',
            re.DOTALL,
        )
        if pattern.search(styles_xml):
            styles_xml = pattern.sub(
                _heading_block(sid,
                               f'heading {i}',
                               heading_fonts[i - 1] if i - 1 < len(heading_fonts) else heading_fonts[-1],
                               heading_para.heading_line_spacing,
                               heading_para.heading_space_before.pt,
                               heading_para.heading_space_after.pt,
                               bold=True),
                styles_xml,
                count=1,
            )

    # Replace BodyText.
    bt_pat = re.compile(
        r'<w:style\b[^>]*w:styleId="BodyText"[^>]*>.*?</w:style>',
        re.DOTALL,
    )
    if bt_pat.search(styles_xml):
        styles_xml = bt_pat.sub(new_body_text, styles_xml, count=1)

    # Add a Caption style if the table-processing flag is on, so the
    # generated "表N 表名" paragraphs have a style to bind to.
    if process_tables:
        styles_xml = ensure_caption_style(styles_xml, body_font=body_font)
        # Insert our own TableContent style so each cell paragraph can
        # bind to a paragraph style that's explicitly indent-free.
        styles_xml = ensure_table_content_style(
            styles_xml,
            line_spacing=table_line_spacing,
            cell_alignment=table_cell_alignment,
        )

    # Detach pandoc's Compact style from BodyText so that — whether or
    # not the user enabled table processing — cell paragraphs never
    # inherit BodyText's first-line indent.
    styles_xml = detach_compact_from_body_text(styles_xml)

    parts['word/styles.xml'] = styles_xml.encode('utf-8')

    # ---- document.xml: replace sectPr ----
    if 'word/document.xml' not in parts:
        raise RuntimeError('word/document.xml not found in source docx')
    doc_xml = parts['word/document.xml'].decode('utf-8')

    # 1) Horizontal rules first, so the heading counter doesn't see
    #    stray paragraphs around them.
    if remove_hr:
        doc_xml = _remove_horizontal_rules(doc_xml)

    # 2) Heading auto-numbering: strip existing numbers and prepend
    #    "1 / 1.1 / 1.1.1" before each Heading1..Heading6 paragraph.
    if renumber_headings:
        doc_xml = _auto_number_headings(doc_xml)

    # 3) Whitespace cleanup: collapse spaces between alphanumerics and
    #    around CJK characters. Done after numbering so the "1 " /
    #    "1.1 " prefix isn't disturbed, but before quote conversion
    #    so quote pairing stays consistent.
    if fix_spaces:
        doc_xml = fix_spaces_in_document(doc_xml)

    # 4) Smart quotes: convert ASCII " ' into Chinese curly quotes,
    #    pairing left/right.
    if fix_quotes:
        doc_xml = fix_smart_quotes_in_document(doc_xml)

    # 4) Rewrite FirstParagraph references to BodyText so that every
    #    body paragraph gets our configured indent (FirstParagraph in
    #    pandoc's default style is meant to skip the first indent,
    #    which is the opposite of Chinese typography conventions).
    doc_xml = doc_xml.replace(
        '<w:pStyle w:val="FirstParagraph"/>',
        '<w:pStyle w:val="BodyText"/>',
    )

    # 5) Replace sectPr with our A4 + standard margins block.
    doc_xml = re.sub(
        r'<w:sectPr>.*?</w:sectPr>',
        build_sect_pr(),
        doc_xml,
        count=1,
        flags=re.DOTALL,
    )

    # 6) Optional: table processing (borders + cell alignment + caption).
    #    Done last because it inserts new paragraphs ahead of <w:tbl>.
    if process_tables:
        doc_xml = _patch_tables(
            doc_xml,
            body_font=body_font,
            table_font=table_font,
            line_spacing=table_line_spacing,
            cell_alignment=table_cell_alignment,
            add_caption=True,
        )

    parts['word/document.xml'] = doc_xml.encode('utf-8')

    # ---- repackage ----
    out_path = dst_path if dst_path != src_path else src_path
    if dst_path != src_path:
        # Write to a temp buffer first, then copy to be safe.
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for n, data in parts.items():
                zout.writestr(n, data)
        with open(out_path, 'wb') as f:
            f.write(buf.getvalue())
    else:
        # In-place rewrite.
        tmp = src_path + '.tmp'
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for n, data in parts.items():
                zout.writestr(n, data)
        shutil.move(tmp, src_path)

    return out_path
