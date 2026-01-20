#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/12 15:18
# @Author  : afish
# @File    : utils.py
"""
获取 段落/字体 属性
"""
from typing import Optional, Tuple

from docx.enum.text import WD_LINE_SPACING
from docx.oxml.shared import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from loguru import logger


def paragraph_get_alignment(paragraph: Paragraph):
    """
    获取段落对齐方式
    Params:
        paragraph: 段落对象，通常是 python-docx 的 Paragraph 对象

    Return:
        str: 对齐方式的描述字符串
    """
    alignment = paragraph.paragraph_format.alignment
    return alignment if alignment else '未设置'


def _get_effective_line_height(paragraph: Paragraph) -> Optional[float]:
    """
    计算段落的有效行高（单位：pt）
    """
    fmt = paragraph.paragraph_format

    # 1. 获取字号（pt）
    font_size_pt = 12.0  # 默认小四
    if paragraph.runs:
        for run in paragraph.runs:
            if run.font.size is not None:
                font_size_pt = run.font.size.pt
                break
    else:
        style = paragraph.style
        if hasattr(style, 'font') and style.font.size is not None:
            font_size_pt = style.font.size.pt

    # 2. 获取行距规则和值
    line_spacing = fmt.line_spacing
    rule = fmt.line_spacing_rule

    if rule == WD_LINE_SPACING.MULTIPLE:
        # 多倍行距：行高 = 倍数 × 字号
        multiplier = float(line_spacing) if line_spacing is not None else 1.0
        return multiplier * font_size_pt
    elif rule in (WD_LINE_SPACING.EXACTLY, WD_LINE_SPACING.AT_LEAST):
        # 固定行高：直接使用该值（此时“行”概念弱化，但仍可作为单位）
        if line_spacing is not None and hasattr(line_spacing, 'pt'):
            return line_spacing.pt
        else:
            return font_size_pt  # fallback
    else:
        # 默认单倍行距
        return font_size_pt


def paragraph_get_space_before(paragraph: Paragraph):
    """
    获取段落前间距(行)
    Params:
        paragraph: 段落对象，通常是 python-docx 的 Paragraph 对象

    Return:
        float: 间距的
    """
    space_before = paragraph.paragraph_format.space_before
    if space_before is None:
        return None

    line_height = _get_effective_line_height(paragraph)
    if line_height is None or line_height <= 0:
        return None

    return round(space_before.pt / line_height, 1)


def paragraph_get_space_after(paragraph: Paragraph):
    """
    获取段落后间距(行)
    Params:
       paragraph: 段落对象，通常是 python-docx 的 Paragraph 对象

    Return:
        float: 间距的
    """
    space_after = paragraph.paragraph_format.space_after
    if space_after is None:
        return None

    line_height = _get_effective_line_height(paragraph)
    if line_height is None or line_height <= 0:
        return None

    return round(space_after.pt / line_height, 1)


def paragraph_get_line_spacing(paragraph: Paragraph):
    """
    获取段落行间距
    Params:
       paragraph: 段落对象，通常是 python-docx 的 Paragraph 对象

    Return:
        float: 间距的大小
    """
    fmt = paragraph.paragraph_format
    rule = fmt.line_spacing_rule
    spacing = fmt.line_spacing

    # 只处理 MULTIPLE（多倍行距）的情况
    if rule == WD_LINE_SPACING.MULTIPLE:
        if spacing is None:
            return 1.0  # Word 默认单倍行距
        return float(spacing)
    else:
        return float(spacing)


def paragraph_get_first_line_indent(paragraph: Paragraph, font_size_pt=12.0):
    """
    获取段落首行缩进
    Params:
       paragraph: 段落对象，通常是 python-docx 的 Paragraph 对象

    Return:
        int: 首行缩进的字符大小（近似值），如果无法计算返回0
    """
    try:
        para_format = paragraph.paragraph_format
        first_line_indent = para_format.first_line_indent
        # 如果首行缩进为None，返回0
        if first_line_indent is None:
            return 0

        # 获取缩进值（以pt为单位）
        if hasattr(first_line_indent, 'pt'):
            # 是python-docx的Length对象
            indent_pt = first_line_indent.pt
            if indent_pt is None:
                return 0
        else:
            # 尝试转换为数值
            try:
                indent_pt = float(first_line_indent)
            except (ValueError, TypeError):
                return 0

        # 获取段落字体大小（以pt为单位）
        # 先尝试获取段落第一个run的字体大小
        if paragraph.runs and len(paragraph.runs) > 0:
            for run in paragraph.runs:
                if run.font and run.font.size:
                    if hasattr(run.font.size, 'pt'):
                        font_size_pt = run.font.size.pt
                    else:
                        try:
                            font_size_pt = float(run.font.size)
                        except (ValueError, TypeError):
                            pass

                    if font_size_pt and font_size_pt > 0:
                        break

        # 如果第一个run没有字体大小，尝试从样式获取
        if font_size_pt == 12.0 and hasattr(paragraph, 'style') and paragraph.style:
            try:
                if hasattr(paragraph.style.font, 'size'):
                    if hasattr(paragraph.style.font.size, 'pt'):
                        font_size_pt = paragraph.style.font.size.pt
                    else:
                        try:
                            font_size_pt = float(paragraph.style.font.size)
                        except (ValueError, TypeError):
                            pass
            except AttributeError:
                pass

        # 计算字符大小
        # 假设中文等宽，字符宽度等于字体大小
        # 但实际上，中文标点和全角字符占1个字符宽度，英文字符占0.5个字符宽度
        # 这里我们简单计算：字符数 = 缩进值(pt) / 字体大小(pt)

        if font_size_pt and font_size_pt > 0:
            # 计算字符数
            char_count = indent_pt / font_size_pt

            # 考虑到实际排版，可能需要调整
            # 通常中文段落缩进是2个字符，但Word中缩进可能会有点偏差
            # 我们将结果四舍五入到最接近的整数

            # 如果缩进是负数（悬挂缩进），也正确处理
            if char_count >= 0:
                char_count_int = round(char_count)
            else:
                char_count_int = -round(abs(char_count))

            return char_count_int
        else:
            return 0

    except Exception as e:
        logger.error(f"获取首行缩进字符大小时出错: {e}")
        return 0


def paragraph_get_builtin_style_name(paragraph: Paragraph):
    """
    获取段落样式名称(全字母小写)
    Params:
       paragraph: 段落对象，通常是 python-docx 的 Paragraph 对象

    Return:
        str: 样式名称
    """
    style = paragraph.style
    if style is None:
        return ""
    return style.name.lower()


def run_get_font_name(run: Run) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    获取run的字体
    Params:
        run: python-docx 的 Run 对象

    Return:
        tuple: (eastAsia, ascii, hAnsi)
        - ascii: 英文字体（如 "Times New Roman"）
        - hAnsi: 高 ANSI 字符字体（通常与 ascii 相同）
        - 未设置时返回 None
    """
    rPr = run._element.rPr
    if rPr is None:
        return None, None, None

    rFonts = rPr.rFonts
    if rFonts is None:
        return None, None, None

    # 使用 qn() 获取带命名空间的属性名
    east_asia = rFonts.get(qn('w:eastAsia'))
    ascii_font = rFonts.get(qn('w:ascii'))
    h_ansi = rFonts.get(qn('w:hAnsi'))
    return east_asia, ascii_font, h_ansi


def run_get_font_size(run: Run):
    """
    获取run的字体大小
    Params:
        run: python-docx 的 Run 对象

    Return:
        float: 字体大小，单位为pt
    """
    font_size = run.font.size
    if font_size is None:
        return None
    return font_size.pt


def run_get_font_color(run: Run) -> Optional[Tuple[int, int, int]]:
    """
    获取run的字体颜色
    Params:
        run: python-docx 的 Run 对象

    Return:
        tuple or None: (r, g, b) 元组，每个分量为 0-255 的整数。
                       若未设置颜色或使用主题色，返回 (0, 0, 0)。
    """
    color = run.font.color
    if color is None or color.rgb is None:
        return 0, 0, 0

    rgb_hex = color.rgb  # 如 'FF0000'
    if not isinstance(rgb_hex, str) or len(rgb_hex) != 6:
        return 0, 0, 0

    try:
        r = int(rgb_hex[0:2], 16)
        g = int(rgb_hex[2:4], 16)
        b = int(rgb_hex[4:6], 16)
        return r, g, b
    except ValueError:
        return 0, 0, 0


def run_get_font_bold(run: Run) -> bool:
    """
    获取run的字体是否加粗
    Params:
        run: python-docx 的 Run 对象

    Return:
        bool: 是否加粗。若未显式设置，返回 False。
    """
    return bool(run.font.bold)


def run_get_font_italic(run: Run) -> bool:
    """
    获取run的字体是否斜体
    Params:
        run: python-docx 的 Run 对象

    Return:
        bool: 是否斜体。若未显式设置，返回 False。
    """
    return bool(run.font.italic)


def run_get_font_underline(run: Run) -> bool:
    """
    获取run的字体是否下划线
    Params:
        run: python-docx 的 Run 对象

    Return:
        bool: 是否下划线。若未显式设置，返回 False。
    """
    return bool(run.font.underline)
