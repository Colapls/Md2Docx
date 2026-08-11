"""
文本清理工具：处理引号转换和多余空格删除
"""
import re


def convert_quotes(text: str) -> str:
    """
    将英文双引号转换为中文双引号""

    Args:
        text: 输入文本

    Returns:
        转换后的文本
    """
    # 中文双引号
    left_quote = '“'  # "
    right_quote = '”'  # "
    single_left = '‘'  # '
    single_right = '’'  # '

    # 处理双引号 - 配对替换
    count = [0]
    def replace_double_quote(match):
        if count[0] % 2 == 0:
            count[0] += 1
            return left_quote
        else:
            count[0] += 1
            return right_quote

    text = re.sub(r'"', replace_double_quote, text)

    # 处理单引号
    count2 = [0]
    def replace_single_quote(match):
        if count2[0] % 2 == 0:
            count2[0] += 1
            return single_left
        else:
            count2[0] += 1
            return single_right

    text = re.sub(r"'", replace_single_quote, text)

    return text


def cleanup_spaces(text: str) -> str:
    """
    删除多余空格，特别是字母或数字旁边的空格

    Args:
        text: 输入文本

    Returns:
        清理后的文本
    """
    # 删除字母数字旁边多余空格 (a b -> ab)
    text = re.sub(r'([a-zA-Z0-9])\s+([a-zA-Z0-9])', r'\1\2', text)

    # 删除中文与英文/数字之间的多余空格
    text = re.sub(r'([一-龥])\s+([a-zA-Z0-9])', r'\1\2', text)
    text = re.sub(r'([a-zA-Z0-9])\s+([一-龥])', r'\1\2', text)

    # 删除多个空格为一个
    text = re.sub(r' {2,}', ' ', text)

    # 删除行首行尾空格
    text = text.lstrip(' ').rstrip(' ')

    return text


def cleanup_text(text: str) -> str:
    """
    综合文本清理：先转换引号，再删除多余空格

    Args:
        text: 输入文本

    Returns:
        清理后的文本
    """
    text = convert_quotes(text)
    text = cleanup_spaces(text)
    return text


def cleanup_paragraph(paragraph_text: str) -> str:
    """
    清理段落文本，保留必要的换行结构

    Args:
        paragraph_text: 段落文本

    Returns:
        清理后的段落
    """
    # 清理每行
    lines = paragraph_text.split('\n')
    cleaned_lines = [cleanup_text(line) for line in lines]
    return '\n'.join(cleaned_lines)
