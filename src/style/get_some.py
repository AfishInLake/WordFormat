#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/12 15:18
# @Author  : afish
# @File    : utils.py
"""
获取 段落/字体 属性
"""
from typing import Optional, Tuple

from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH
from docx.oxml.shared import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from loguru import logger


def _wd_align_to_str(alignment) -> str:
    """将 WD_ALIGN_PARAGRAPH 枚举转为中文描述"""
    mapping = {
        WD_ALIGN_PARAGRAPH.LEFT: '左对齐',
        WD_ALIGN_PARAGRAPH.CENTER: '居中',
        WD_ALIGN_PARAGRAPH.RIGHT: '右对齐',
        WD_ALIGN_PARAGRAPH.JUSTIFY: '两端对齐',
        WD_ALIGN_PARAGRAPH.DISTRIBUTE: '分散对齐',
    }
    return mapping.get(alignment, f'未知({alignment})')


def paragraph_get_alignment(paragraph: Paragraph) -> str:
    """
    获取段落的有效对齐方式（考虑直接格式 + 样式继承）

    Args:
        paragraph: python-docx 的 Paragraph 对象

    Returns:
        str: 对齐方式描述，如 '左对齐', '居中', '右对齐', '两端对齐'，若均未设置则返回 '未设置'
    """
    # 1. 先看段落是否显式设置了对齐
    direct_alignment = paragraph.paragraph_format.alignment
    if direct_alignment is not None:
        return _wd_align_to_str(direct_alignment)

    # 2. 否则，从段落样式中获取
    style = paragraph.style
    while style is not None:
        if hasattr(style, 'paragraph_format') and style.paragraph_format.alignment is not None:
            return _wd_align_to_str(style.paragraph_format.alignment)
        # 尝试向上查找基础样式（部分版本支持 _base_style）
        base_style = getattr(style, '_base_style', None)
        if base_style is None:
            break
        style = base_style

    # 3. 所有地方都没设置 → Word 默认是左对齐，但为严谨返回“未设置”
    return '未设置'


def _get_effective_line_height(paragraph) -> Optional[float]:
    """
    计算段落的有效行高（单位：pt）
    """
    fmt = paragraph.paragraph_format

    # 1. 获取字号（优先从样式获取，再考虑runs）
    font_size_pt = 12.0  # 默认

    # 首先尝试从段落样式获取
    style = paragraph.style
    if style and hasattr(style, 'font') and style.font.size is not None:
        font_size_pt = style.font.size.pt
    else:
        # 检查runs中的字体大小
        if paragraph.runs:
            for run in paragraph.runs:
                if run.font and run.font.size is not None:
                    font_size_pt = run.font.size.pt
                    break

    # 2. 获取行距规则和值
    line_spacing = fmt.line_spacing
    rule = fmt.line_spacing_rule

    # 如果段落没有设置行距，检查样式中的行距
    if line_spacing is None and style and hasattr(style, 'paragraph_format'):
        style_fmt = style.paragraph_format
        if hasattr(style_fmt, 'line_spacing'):
            line_spacing = style_fmt.line_spacing
            if hasattr(style_fmt, 'line_spacing_rule'):
                rule = style_fmt.line_spacing_rule

    # 3. 根据规则计算行高
    if rule == WD_LINE_SPACING.MULTIPLE:
        # 多倍行距
        if line_spacing is not None:
            try:
                multiplier = float(line_spacing)
                return multiplier * font_size_pt
            except (ValueError, TypeError):
                pass

    elif rule in (WD_LINE_SPACING.EXACTLY, WD_LINE_SPACING.AT_LEAST):
        # 固定行高或最小行高
        if line_spacing is not None and hasattr(line_spacing, 'pt'):
            return line_spacing.pt

    # 默认单倍行距
    return font_size_pt


def _get_style_spacing(style, spacing_type='before'):
    """
    递归查找样式中的段前/段后间距（支持Lines和twips两种格式）
    :param style: docx.styles.style._ParagraphStyle 对象
    :param spacing_type: 'before' 段前 / 'after' 段后
    :return: (lines值, twips值) 元组，无则返回(0.0, 0)
    """
    if not style:
        return (0.0, 0)

    # 1. 获取样式的XML元素
    style_elem = style.element
    style_pPr = style_elem.find(qn('w:pPr'))
    if not style_pPr:
        # 递归查基样式
        return _get_style_spacing(style.base_style, spacing_type)

    # 2. 查找样式中的spacing元素
    spacing = style_pPr.find(qn('w:spacing'))
    if not spacing:
        # 递归查基样式
        return _get_style_spacing(style.base_style, spacing_type)

    # 3. 优先读取Lines属性（beforeLines/afterLines）
    lines_attr = spacing.get(qn(f'w:{spacing_type}Lines'))
    lines_val = int(lines_attr) / 100.0 if lines_attr is not None else 0.0

    # 4. 读取twips属性（before/after）
    twips_attr = spacing.get(qn(f'w:{spacing_type}'))
    twips_val = int(twips_attr) if twips_attr is not None else 0

    if lines_val > 0:
        return (lines_val, twips_val)
    # 无Lines值时，返回twips值，继续递归查基样式（避免漏基样式的设置）
    base_lines, base_twips = _get_style_spacing(style.base_style, spacing_type)
    return (lines_val if lines_val > 0 else base_lines, twips_val if twips_val > 0 else base_twips)


def paragraph_get_space_before(paragraph):
    """
    精准获取段前间距（单位：行）
    完全复刻Word逻辑：段落自身 > 样式继承（Lines→twips） > 内置默认
    """
    font_size_pt = _get_font_size_pt(paragraph)
    actual_line_twips = font_size_pt * 20  # 当前字体1行=字体大小×20twips
    p = paragraph._element
    pPr = p.find(qn('w:pPr'))

    # 第一步：查段落自身的设置
    self_lines = 0.0
    self_twips = 0
    if pPr is not None:
        spacing = pPr.find(qn('w:spacing'))
        if spacing is not None:
            # 1.1 优先查Lines属性
            before_lines_attr = spacing.get(qn('w:beforeLines'))
            if before_lines_attr is not None:
                self_lines = int(before_lines_attr) / 100.0
            # 1.2 查twips属性（兜底）
            before_twips_attr = spacing.get(qn('w:before'))
            self_twips = int(before_twips_attr) if before_twips_attr is not None else 0

    # 第二步：自身无值，查样式继承
    style_lines, style_twips = _get_style_spacing(paragraph.style, 'before')
    final_lines = self_lines if self_lines > 0 else style_lines
    final_twips = self_twips if self_twips > 0 else style_twips

    # 第三步：有Lines值直接返回
    if final_lines > 0:
        return round(final_lines, 1)

    # 第四步：无Lines值，用twips换算（核心修复：不再直接返回0.0，而是用样式的twips换算）
    if final_twips > 0:
        before_lines = final_twips / actual_line_twips
        return round(before_lines, 1)

    # 第五步：终极兜底（Word内置默认值，可根据需求调整）
    # 比如：标题样式默认0.5行，正文默认0行，可根据实际场景加判断
    style_name = paragraph.style.name.lower()
    if any(key in style_name for key in ['标题', 'heading', 'title']):
        return 0.5  # 标题默认0.5行
    return 0.0  # 正文默认0行


# 段后间距函数（同理，仅修改spacing_type为after）
def paragraph_get_space_after(paragraph):
    font_size_pt = _get_font_size_pt(paragraph)
    actual_line_twips = font_size_pt * 20
    p = paragraph._element
    pPr = p.find(qn('w:pPr'))

    # 第一步：查段落自身的设置
    self_lines = 0.0
    self_twips = 0
    if pPr is not None:
        spacing = pPr.find(qn('w:spacing'))
        if spacing is not None:
            after_lines_attr = spacing.get(qn('w:afterLines'))
            if after_lines_attr is not None:
                self_lines = int(after_lines_attr) / 100.0
            after_twips_attr = spacing.get(qn('w:after'))
            self_twips = int(after_twips_attr) if after_twips_attr is not None else 0

    # 第二步：自身无值，查样式继承
    style_lines, style_twips = _get_style_spacing(paragraph.style, 'after')
    final_lines = self_lines if self_lines > 0 else style_lines
    final_twips = self_twips if self_twips > 0 else style_twips

    # 第三步：有Lines值直接返回
    if final_lines > 0:
        return round(final_lines, 1)

    # 第四步：无Lines值，用twips换算（核心修复）
    if final_twips > 0:
        after_lines = final_twips / actual_line_twips
        return round(after_lines, 1)

    # 第五步：终极兜底（Word内置默认值）
    style_name = paragraph.style.name.lower()
    if any(key in style_name for key in ['标题', 'heading', 'title']):
        return 0.5
    return 0.0


def _get_space_from_style(paragraph, spacing_type):
    """
    从样式中获取间距设置（磅值）
    """
    style = paragraph.style
    if not style:
        return 0.0

    # 递归查找样式链中的间距设置
    current_style = style
    while current_style:
        if hasattr(current_style, 'paragraph_format'):
            style_fmt = current_style.paragraph_format

            if spacing_type == 'after':
                spacing = style_fmt.space_after
            else:
                spacing = style_fmt.space_before

            if spacing is not None and hasattr(spacing, 'pt'):
                return spacing.pt

        # 检查基样式
        if hasattr(current_style, 'base_style') and current_style.base_style:
            current_style = current_style.base_style
        else:
            break

    return 0.0


def paragraph_get_line_spacing(paragraph):
    """
    获取段落行间距（倍数）
    Params:
       paragraph: 段落对象，通常是 python-docx 的 Paragraph 对象
    Return:
        float: 行间距的倍数
    """
    # 首先尝试从直接段落格式获取
    fmt = paragraph.paragraph_format
    rule = fmt.line_spacing_rule
    spacing = fmt.line_spacing

    # 如果没有直接设置，检查样式
    if (rule is None or rule == WD_LINE_SPACING.SINGLE) and (spacing is None or spacing == 1.0):
        style = paragraph.style
        if style and hasattr(style, 'paragraph_format'):
            style_fmt = style.paragraph_format
            # 优先使用样式的行距设置
            if hasattr(style_fmt, 'line_spacing_rule'):
                rule = style_fmt.line_spacing_rule
            if hasattr(style_fmt, 'line_spacing'):
                spacing = style_fmt.line_spacing

    # 只处理 MULTIPLE（多倍行距）的情况
    if rule == WD_LINE_SPACING.MULTIPLE:
        if spacing is None:
            return 1.0  # Word 默认单倍行距
        try:
            return float(spacing)
        except (ValueError, TypeError):
            return 1.0
    elif rule == WD_LINE_SPACING.ONE_POINT_FIVE:
        return 1.5  # 1.5倍行距
    elif rule == WD_LINE_SPACING.DOUBLE:
        return 2.0  # 双倍行距
    elif rule == WD_LINE_SPACING.SINGLE:
        return 1.0  # 单倍行距
    else:
        # 对于 EXACTLY, AT_LEAST 或其他规则
        # 需要计算行高相对于字号的倍数
        font_size_pt = _get_font_size_pt(paragraph)

        if rule in (WD_LINE_SPACING.EXACTLY, WD_LINE_SPACING.AT_LEAST):
            if spacing is not None and hasattr(spacing, 'pt'):
                line_height_pt = spacing.pt
                if font_size_pt > 0:
                    return line_height_pt / font_size_pt
        # 默认返回单倍行距
        return 1.0


def _get_font_size_pt(paragraph):
    """获取段落的字体大小（单位：pt），兼容无字体设置的情况"""
    default_font_size_pt = 12.0
    result = []
    if paragraph.runs:
        for run in paragraph.runs:
            if run.font and run.font.size is not None:
                result.append(run.font.size.pt)
    return max(result) if result else default_font_size_pt


def paragraph_get_first_line_indent(paragraph: Paragraph, font_size_pt=12.0):
    """
    获取段落首行缩进
    Params:
       paragraph: 段落对象，通常是 python-docx 的 Paragraph 对象

    Return:
        int: 首行缩进的字符大小（近似值），如果无法计算返回0
    """
    # FIXME: 首行缩进按字符计算有误，需要修改
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
    if font_size is not None:
        return font_size.pt
    # 直接取段落样式的字号（大多数情况足够）
    style = run._parent.style
    if style and style.font.size is not None:
        return style.font.size.pt
    return 12.0


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
