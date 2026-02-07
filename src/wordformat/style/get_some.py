#! /usr/bin/env python
# @Time    : 2026/1/12 15:18
# @Author  : afish
# @File    : utils.py
"""
获取 段落/字体 属性
"""

from docx.enum.text import WD_LINE_SPACING
from docx.oxml.shared import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from loguru import logger


def paragraph_get_alignment(paragraph: Paragraph) -> object:
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
        return direct_alignment

    # 2. 否则，从段落样式中获取
    style = paragraph.style
    while style is not None:
        if (
            hasattr(style, "paragraph_format")
            and style.paragraph_format.alignment is not None
        ):
            return style.paragraph_format.alignment
        # 尝试向上查找基础样式（部分版本支持 _base_style）
        base_style = getattr(style, "_base_style", None)
        if base_style is None:
            break
        style = base_style

    # 3. 所有地方都没设置 → Word 默认是左对齐，但为严谨返回“未设置”
    return None


def _get_effective_line_height(paragraph) -> float | None:  # noqa c901
    """
    计算段落的有效行高（单位：pt）
    """
    fmt = paragraph.paragraph_format

    # 1. 获取字号（优先从样式获取，再考虑runs）
    font_size_pt = 12.0  # 默认

    # 首先尝试从段落样式获取
    style = paragraph.style
    if style and hasattr(style, "font") and style.font.size is not None:
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
    if line_spacing is None and style and hasattr(style, "paragraph_format"):
        style_fmt = style.paragraph_format
        if hasattr(style_fmt, "line_spacing"):
            line_spacing = style_fmt.line_spacing
            if hasattr(style_fmt, "line_spacing_rule"):
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
        if line_spacing is not None and hasattr(line_spacing, "pt"):
            return line_spacing.pt

    # 默认单倍行距
    return font_size_pt


def _get_style_spacing(style, spacing_type="before"):
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
    style_pPr = style_elem.find(qn("w:pPr"))
    if style_pPr is None:
        # 递归查基样式
        return _get_style_spacing(style.base_style, spacing_type)

    # 2. 查找样式中的spacing元素
    spacing = style_pPr.find(qn("w:spacing"))
    if spacing is None:
        # 递归查基样式
        return _get_style_spacing(style.base_style, spacing_type)

    # 3. 优先读取Lines属性（beforeLines/afterLines）
    lines_attr = spacing.get(qn(f"w:{spacing_type}Lines"))
    lines_val = int(lines_attr) / 100.0 if lines_attr is not None else 0.0

    # 4. 读取twips属性（before/after）
    twips_attr = spacing.get(qn(f"w:{spacing_type}"))
    twips_val = int(twips_attr) if twips_attr is not None else 0

    if lines_val > 0:
        return lines_val, twips_val
    # 无Lines值时，返回twips值，继续递归查基样式（避免漏基样式的设置）
    base_lines, base_twips = _get_style_spacing(style.base_style, spacing_type)
    return (
        lines_val if lines_val > 0 else base_lines,
        twips_val if twips_val > 0 else base_twips,
    )


def paragraph_get_space_before(paragraph):
    """
    精准获取段前间距（单位：行）
    完全复刻Word逻辑：段落自身 > 样式继承（Lines→twips） > 内置默认
    核心修复：解决小数值twips被误判为非零行数的问题（如100twips→0.4行）
    """
    font_size_pt = _get_font_size_pt(paragraph)
    actual_line_twips = font_size_pt * 20  # 当前字体1行=字体大小×20twips
    # 核心阈值：小于单行高度50%的twips视为"0行"（匹配Word视觉逻辑）
    MIN_EFFECTIVE_TWIPS_RATIO = 0.5
    effective_twips_threshold = actual_line_twips * MIN_EFFECTIVE_TWIPS_RATIO

    p = paragraph._element
    pPr = p.find(qn("w:pPr"))

    # 第一步：查段落自身的设置
    self_lines = 0.0
    self_twips = 0
    if pPr is not None:
        spacing = pPr.find(qn("w:spacing"))
        if spacing is not None:
            # 1.1 优先查Lines属性（Word中Lines优先级高于twips）
            before_lines_attr = spacing.get(qn("w:beforeLines"))
            if before_lines_attr is not None:
                try:
                    self_lines = (
                        int(before_lines_attr) / 100.0
                    )  # Word中Lines单位是1/100行
                except (ValueError, TypeError):
                    self_lines = 0.0  # 异常值兜底为0
            # 1.2 查twips属性（兜底）+ 类型校验
            before_twips_attr = spacing.get(qn("w:before"))
            try:
                self_twips = (
                    int(before_twips_attr) if before_twips_attr is not None else 0
                )
            except (ValueError, TypeError):
                self_twips = 0  # 非数字twips值视为0

    # 第二步：自身无值，查样式继承
    style_lines, style_twips = _get_style_spacing(paragraph.style, "before")
    final_lines = self_lines if self_lines > 0 else style_lines
    final_twips = self_twips if self_twips > 0 else style_twips

    # 第三步：有Lines值直接返回（Lines是Word显性设置，优先级最高）
    if final_lines > 0:
        return round(final_lines, 1)

    # 第四步：无Lines值，用twips换算（核心修复逻辑）
    if final_twips > 0:
        # 修复1：小数值twips视为0行（如100twips<12pt字体的阈值120twips→0行）
        if final_twips < effective_twips_threshold:
            return 0.0
        # 修复2：有效twips值才换算为行数（避免无效小值误判）
        before_lines = final_twips / actual_line_twips
        return round(before_lines, 1)
    return 0.0  # 正文样式默认0行


def paragraph_get_space_after(paragraph):
    """
    获取段落段后间距（行），修复默认twips值误判问题
    核心修复：
    1. 区分「有效twips值」和「默认twips值」（0/100等小值视为0行）
    2. 严格匹配Word的「0行」逻辑：twips<单行高度的1/2视为0行
    """
    font_size_pt = _get_font_size_pt(paragraph)
    actual_line_twips = font_size_pt * 20  # 1pt=20twips，单行高度
    MIN_EFFECTIVE_TWIPS_RATIO = 0.5  # 小于单行高度50%的twips视为0行
    effective_twips_threshold = actual_line_twips * MIN_EFFECTIVE_TWIPS_RATIO

    p = paragraph._element
    pPr = p.find(qn("w:pPr"))

    # 第一步：查段落自身的设置
    self_lines = 0.0
    self_twips = 0
    if pPr is not None:
        spacing = pPr.find(qn("w:spacing"))
        if spacing is not None:
            # 读取Lines值（优先）
            after_lines_attr = spacing.get(qn("w:afterLines"))
            if after_lines_attr is not None:
                self_lines = float(int(after_lines_attr) / 100.0)
            # 读取twips值（注意："0"转为0，非数字视为0）
            after_twips_attr = spacing.get(qn("w:after"))
            try:
                self_twips = (
                    int(after_twips_attr) if after_twips_attr is not None else 0
                )
            except (ValueError, TypeError):
                self_twips = 0

    # 第二步：自身无值，查样式继承
    style_lines, style_twips = _get_style_spacing(paragraph.style, "after")
    final_lines = self_lines if self_lines > 0 else style_lines
    final_twips = self_twips if self_twips > 0 else style_twips

    # 第三步：有Lines值直接返回（Lines优先级最高）
    if final_lines > 0:
        return round(final_lines, 1)

    # 第四步：修复核心——判断twips是否为有效非零值
    if final_twips > 0:
        # 关键：小于阈值的twips视为0行（解决100twips=0.4行的错误）
        if final_twips < effective_twips_threshold:
            return 0.0
        # 有效twips值，换算为行数
        after_lines = final_twips / actual_line_twips
        return round(after_lines, 1)

    # 第五步：终极兜底（Word内置默认值）
    style_name = paragraph.style.name.lower()
    # 修复：仅当无任何设置时才用兜底值，避免覆盖0行逻辑
    if any(key in style_name for key in ["标题", "heading", "title"]):
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
        if hasattr(current_style, "paragraph_format"):
            style_fmt = current_style.paragraph_format

            if spacing_type == "after":
                spacing = style_fmt.space_after
            else:
                spacing = style_fmt.space_before

            if spacing is not None and hasattr(spacing, "pt"):
                return spacing.pt

        # 检查基样式
        if hasattr(current_style, "base_style") and current_style.base_style:
            current_style = current_style.base_style
        else:
            break

    return 0.0


def paragraph_get_line_spacing(paragraph):  # noqa c901
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
    if (rule is None or rule == WD_LINE_SPACING.SINGLE) and (
        spacing is None or spacing == 1.0
    ):
        style = paragraph.style
        if style and hasattr(style, "paragraph_format"):
            style_fmt = style.paragraph_format
            # 优先使用样式的行距设置
            if hasattr(style_fmt, "line_spacing_rule"):
                rule = style_fmt.line_spacing_rule
            if hasattr(style_fmt, "line_spacing"):
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
            if spacing is not None and hasattr(spacing, "pt"):
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


def paragraph_get_first_line_indent(paragraph: Paragraph):  # noqa c901
    """
    精准获取首行缩进，优先解析XML字符单位（firstLineChars），兼容物理单位
    :param para: 目标段落对象
    :return: 字典结果，包含：
             - type: 缩进单位类型（char/pt），char=字符单位，pt=物理单位
             - value: 对应单位的缩进值（字符数为浮点数，pt值为整数）
             - char_visual: 视觉等效字符数（统一换算，方便对比）
    """
    try:
        p = paragraph._element
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            return 0.0

        # 获取XML中的ind节点（缩进核心节点）
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            return 0.0

        # 步骤1：优先解析字符单位 firstLineChars（核心：值=字符数×100）
        first_line_chars = ind.get(qn("w:firstLineChars"))
        if first_line_chars and first_line_chars.isdigit():
            chars_num = int(first_line_chars) / 100  # 200 → 2.0字符
            # return {
            #     "type": "char",
            #     "value": round(chars_num, 1),
            #     "char_visual": round(chars_num, 1)
            # }
            return chars_num

        # 步骤2：无字符单位，解析物理单位 firstLine（单位：twips，1pt=20twips）
        first_line_twips = ind.get(qn("w:firstLine"))
        if first_line_twips and first_line_twips.isdigit():
            indent_pt = int(first_line_twips) // 20  # twips → pt（取整，避免小数）
            # 换算为视觉等效字符数（字符数=pt值/字体pt值）
            font_pt = _get_font_size_pt(paragraph)
            char_visual = indent_pt / font_pt if font_pt > 0 else 0.0
            # return {
            #     "type": "pt",
            #     "value": indent_pt,
            #     "char_visual": round(char_visual, 1)
            # }
            return round(char_visual, 1)

        # 无任何缩进设置
        return 0.0

    except Exception as e:
        logger.error(f"获取首行缩进失败：{e}")
        return 0.0


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


def run_get_font_name(run: Run) -> str | None:
    """
    获取 Run 对象的东亚字体（eastAsia font）名称。
    Params:
        run: python-docx 的 Run 对象

    Return:
       str: 东亚字体名称（如 "Microsoft YaHei"），如果未设置则返回 None。
    """
    rPr = run._element.rPr
    if rPr is not None:
        rFonts = rPr.rFonts
        if rFonts is not None:
            # 获取 w:eastAsia 属性
            east_asia = rFonts.get(qn("w:eastAsia"))
            return east_asia if east_asia else None
    return None


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


def run_get_font_color(run: Run) -> tuple[int, int, int] | None:
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
