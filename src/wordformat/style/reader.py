#! /usr/bin/env python
# @Time    : 2026/1/12 15:18
# @Author  : afish
# @File    : reader.py
"""获取段落 / 字体的**有效**格式属性。

所有读取统一走 `style.inheritance.StyleResolver`，沿 OOXML 继承链
（直接格式 → 字符/段落样式 basedOn 链 → docDefaults → 主题字体）解析，
不再各自手写样式回退，避免漏掉继承链。新增属性只需在 inheritance.py 加一个提取器。
"""

from docx.enum.text import WD_LINE_SPACING
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from loguru import logger
from lxml.etree import _Element

from wordformat.style.inheritance import (
    StyleResolver,
    x_alignment,
    x_bold,
    x_color_rgb,
    x_first_line_indent,
    x_font_ascii,
    x_font_ea,
    x_italic,
    x_left_indent,
    x_line_spacing,
    x_right_indent,
    x_size_pt,
    x_space_after,
    x_space_before,
    x_underline,
)


def _real_elem(obj) -> bool:
    """obj 是否为持有真实 lxml 元素的段落/run（排除 Mock 等无效输入）。"""
    return isinstance(getattr(obj, "_element", None), _Element)


def _para(paragraph, extractor, default=None):
    if not _real_elem(paragraph):
        return default
    try:
        return StyleResolver.for_paragraph(paragraph).resolve_para(
            paragraph, extractor, default
        )
    except Exception as e:
        logger.debug(f"段落属性解析失败：{e}")
        return default


def _run(run, extractor, default=None):
    if not _real_elem(run):
        return default
    try:
        return StyleResolver.for_run(run).resolve_run(run, extractor, default)
    except Exception as e:
        logger.debug(f"run 属性解析失败：{e}")
        return default


def _run_font(run, extractor):
    if not _real_elem(run):
        return None
    try:
        return StyleResolver.for_run(run).resolve_font(run, extractor)
    except Exception as e:
        logger.debug(f"字体解析失败：{e}")
        return None


# ── 段落级 ────────────────────────────────────────────────────────
def paragraph_get_alignment(paragraph: Paragraph):
    """段落有效对齐方式（沿继承链）；均未设置返回 None。"""
    return _para(paragraph, x_alignment, None)


def paragraph_get_space_before(paragraph) -> float | None:
    """段前间距（单位：行）；无显式值返回 None。"""
    return _para(paragraph, x_space_before, None)


def paragraph_get_space_after(paragraph) -> float | None:
    """段后间距（单位：行）；无显式值返回 None。"""
    return _para(paragraph, x_space_after, None)


def paragraph_get_line_spacing(paragraph) -> float | None:
    """行距倍数；固定值/最小值返回 None。"""
    res = _para(paragraph, x_line_spacing, None)
    if not res:
        return None
    try:
        rule, factor = res
    except (TypeError, ValueError):
        return None
    return factor if rule == WD_LINE_SPACING.MULTIPLE else None


def paragraph_get_line_spacing_rule(paragraph):
    """行距类型（WD_LINE_SPACING）。倍数 1.0/1.5/2.0 归一为单倍/1.5/2 倍。"""
    res = _para(paragraph, x_line_spacing, None)
    if not res:
        return WD_LINE_SPACING.MULTIPLE
    try:
        rule, factor = res
    except (TypeError, ValueError):
        return WD_LINE_SPACING.MULTIPLE
    if rule == WD_LINE_SPACING.MULTIPLE:
        return {
            1.0: WD_LINE_SPACING.SINGLE,
            1.5: WD_LINE_SPACING.ONE_POINT_FIVE,
            2.0: WD_LINE_SPACING.DOUBLE,
        }.get(factor, WD_LINE_SPACING.MULTIPLE)
    return rule


def paragraph_get_first_line_indent(paragraph: Paragraph) -> float | None:
    """首行缩进（字符）；首行=正值，悬挂=负值，无则 None。"""
    return _para(paragraph, x_first_line_indent, None)


def paragraph_get_builtin_style_name(paragraph: Paragraph) -> str:
    """段落样式名（全小写）。"""
    style = paragraph.style
    if style is None:
        return ""
    return style.name.lower()


# ── 字符级 ────────────────────────────────────────────────────────
def run_get_font_name(run: Run) -> str | None:
    """东亚字体（eastAsia），沿继承链并解析主题字体；未设置 None。"""
    return _run_font(run, x_font_ea)


def run_get_font_name_en(run: Run) -> str | None:
    """西文字体（ascii），沿继承链并解析主题字体；未设置 None。"""
    return _run_font(run, x_font_ascii)


def run_get_font_size_pt(run: Run) -> float:
    """字号（pt），沿继承链；均未设置回退 12.0。"""
    return _run(run, x_size_pt, 12.0)


def run_get_font_color(run: Run) -> tuple[int, int, int] | None:
    """字体颜色 (r,g,b)；主题色返回 None（不确定），未设置返回 (0,0,0)。"""
    return _run(run, x_color_rgb, (0, 0, 0))


def run_get_font_bold(run: Run) -> bool:
    """有效加粗（沿继承链：直接→字符样式→段落样式→docDefaults）。"""
    return bool(_run(run, x_bold, False))


def run_get_font_italic(run: Run) -> bool:
    """有效斜体（沿继承链解析）。"""
    return bool(_run(run, x_italic, False))


def run_get_font_underline(run: Run) -> bool:
    """有效下划线（沿继承链解析）。"""
    return bool(_run(run, x_underline, False))


class GetIndent:
    """段落左/右缩进（单位：字符），沿继承链解析。"""

    @staticmethod
    def line_indent(paragraph: Paragraph, indent_type: str = "left") -> float | None:
        if indent_type not in ("left", "right"):
            logger.error("indent_type 必须是 'left' 或 'right'")
            raise ValueError("indent_type 必须是 'left' 或 'right'")
        extractor = x_left_indent if indent_type == "left" else x_right_indent
        val = _para(paragraph, extractor, None)
        if val is None:
            return None
        return max(0.0, val)

    @staticmethod
    def left_indent(paragraph: Paragraph) -> float | None:
        return GetIndent.line_indent(paragraph, "left")

    @staticmethod
    def right_indent(paragraph: Paragraph) -> float | None:
        return GetIndent.line_indent(paragraph, "right")
