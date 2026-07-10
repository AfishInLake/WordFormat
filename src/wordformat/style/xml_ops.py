#! /usr/bin/env python
"""Word OOXML 样式定义级别的 rPr / pPr XML 原子操作。

本模块仅供底层 XML 写入，不依赖 python-docx 对象模型（Paragraph / Run），
可被 set_some.py、stages.py 等模块复用。

值转换遵循 set_some.py 中 _Set* 类的模式：
- 行距：倍数→w:line（×240），pt→twips（×20）
- 段间距：行→w:beforeLines/w:afterLines（×100）
- 缩进：字符→w:leftChars/w:rightChars/w:firstLineChars（×100）
"""

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from wordformat.style.defs import Alignment, LineSpacingRule

# --- w:rPr 字符格式 ---


def rPr_set_font(rPr, cn_name=None, en_name=None):
    """在任意 w:rPr 元素上设置中英文字体（w:rFonts）。"""
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    if cn_name:
        rFonts.set(qn("w:eastAsia"), str(cn_name))
    if en_name:
        rFonts.set(qn("w:ascii"), str(en_name))
        rFonts.set(qn("w:hAnsi"), str(en_name))


def rPr_set_font_size(rPr, pt_val):
    """在任意 w:rPr 元素上设置字号（w:sz / w:szCs）。pt_val 单位为磅。"""
    half_pt = str(int(round(pt_val * 2)))
    sz = rPr.find(qn("w:sz"))
    if sz is None:
        sz = OxmlElement("w:sz")
        rPr.append(sz)
    sz.set(qn("w:val"), half_pt)
    szCs = rPr.find(qn("w:szCs"))
    if szCs is None:
        szCs = OxmlElement("w:szCs")
        rPr.append(szCs)
    szCs.set(qn("w:val"), half_pt)


def rPr_set_font_color(rPr, rgb_tuple):
    """在任意 w:rPr 元素上设置字体颜色，清除旧的主题色 w:color。"""
    hex_color = f"{rgb_tuple[0]:02X}{rgb_tuple[1]:02X}{rgb_tuple[2]:02X}"
    for old_color in rPr.findall(qn("w:color")):
        rPr.remove(old_color)
    new_color = OxmlElement("w:color")
    new_color.set(qn("w:val"), hex_color)
    rPr.append(new_color)


def rPr_set_bold(rPr, bold):
    """在任意 w:rPr 元素上设置/移除加粗（w:b / w:bCs）。"""
    b = rPr.find(qn("w:b"))
    bCs = rPr.find(qn("w:bCs"))
    if bold:
        if b is None:
            rPr.append(OxmlElement("w:b"))
        if bCs is None:
            rPr.append(OxmlElement("w:bCs"))
    else:
        if b is not None:
            rPr.remove(b)
        if bCs is not None:
            rPr.remove(bCs)


def rPr_set_italic(rPr, italic):
    """在任意 w:rPr 元素上设置/移除斜体（w:i）。"""
    i = rPr.find(qn("w:i"))
    if italic:
        if i is None:
            rPr.append(OxmlElement("w:i"))
    else:
        if i is not None:
            rPr.remove(i)


def rPr_set_underline(rPr, underline):
    """在任意 w:rPr 元素上设置/移除下划线（w:u）。"""
    u = rPr.find(qn("w:u"))
    if underline:
        if u is None:
            u = OxmlElement("w:u")
            u.set(qn("w:val"), "single")
            rPr.append(u)
    else:
        if u is not None:
            rPr.remove(u)


# --- w:pPr 段落格式 ---


def pPr_set_alignment(pPr, wd_alignment):
    """在任意 w:pPr 元素上设置对齐方式（w:jc）。

    wd_alignment: WD_ALIGN_PARAGRAPH 枚举值，通过 Alignment.rel_value 获取。
    """
    xml_val = Alignment.XML_VAL_MAP.get(wd_alignment, "left")
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        pPr.append(jc)
    jc.set(qn("w:val"), xml_val)


def line_spacing_val_to_xml(rel_value, rel_unit):
    """将 LineSpacing 的 (value, unit) 转为 Word w:line 内部值。

    遵循 _SetLineSpacing 的模式：倍数→240ths，pt→twips（1pt=20twips）。
    """
    if rel_unit == "倍":
        return rel_value * 240
    if rel_unit in ("pt",):
        return int(Pt(rel_value)) // 635  # EMU → twips
    return rel_value * 240


def line_rule_to_xml(wd_line_spacing_rule):
    """将 WD_LINE_SPACING 枚举值转为 w:lineRule XML 字符串。

    wd_line_spacing_rule: 通过 LineSpacingRule.rel_value 获取。
    """
    return LineSpacingRule.XML_RULE_MAP.get(wd_line_spacing_rule, "auto")


# --- 元素创建辅助 ---


def ensure_rPr(style_element):
    """确保 style 元素中存在 w:rPr，不存在则创建并返回。"""
    rPr = style_element.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        style_element.insert(0, rPr)
    return rPr


def ensure_pPr(style_element):
    """确保 style 元素中存在 w:pPr，不存在则创建并返回。"""
    pPr = style_element.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        style_element.append(pPr)
    return pPr
