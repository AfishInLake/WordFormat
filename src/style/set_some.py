#! /usr/bin/env python
# @Time    : 2026/1/20 18:30
# @Author  : afish
# @File    : set_some.py
"""
设置 段落/字体 属性
"""

from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run


def run_set_font_name(run: Run, font_name: str):
    """
    设置 Run 对象的中文字体名称（修复空值问题，兼容无样式配置的 Run）

    参数:
        run (docx.text.run.Run): 要设置字体的 Run 对象。
        font_name (str): 字体名称，如 "Microsoft YaHei"、"SimSun"、"Arial" 等。
    """
    # 1. 获取 Run 对应的 XML 元素
    r = run._element
    # 2. 检查并创建 rPr 节点（字体样式根节点，不存在则创建）
    if r.rPr is None:
        # 手动创建 rPr 节点（docx 原生XML结构，兼容所有版本）
        rPr = parse_xml(
            r'<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
        )
        r.append(rPr)
    # 3. 检查并创建 rFonts 节点（字体名称配置节点，不存在则创建）
    if r.rPr.rFonts is None:
        # 手动创建 rFonts 节点
        rFonts = parse_xml(
            r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
        )
        r.rPr.append(rFonts)
    # 4. 安全设置中文字体（eastAsia 对应中文/东亚字体）
    r.rPr.rFonts.set(qn("w:eastAsia"), font_name)


def _paragraph_space_by_lines(
    paragraph: Paragraph, before_lines: float = 0.0, after_lines: float = 0.0
) -> None:
    """
    设置段落【段前/段后间距】为指定行数（Word原生行单位，非Pt磅值）
    核心：单次删旧建全新spacing节点，与行距完全解耦，支持单/双侧同时设置
    :param paragraph: 目标段落对象
    :param before_lines: 段前间距行数，浮点型（如0.1/0.7，>0时生效）
    :param after_lines: 段后间距行数，浮点型（如0.1/0.7，>0时生效）
    """
    pPr = paragraph._element.get_or_add_pPr()
    # 步骤1：读取已有属性（若存在spacing节点，保留未被修改的属性）
    existing_before = 0.0
    existing_after = 0.0
    # 查找已有spacing节点，避免直接删除
    for spacing in pPr.xpath(".//w:spacing"):
        # 读取已有段前属性，转换为行数值
        if spacing.get(qn("w:beforeLines")):
            existing_before = int(spacing.get(qn("w:beforeLines"))) / 100
        # 读取已有段后属性，转换为行数值
        if spacing.get(qn("w:afterLines")):
            existing_after = int(spacing.get(qn("w:afterLines"))) / 100
        # 先暂存，后续统一删除

    # 步骤2：删除所有旧spacing节点（保持原有核心逻辑，避免属性冲突）
    for spacing in pPr.xpath(".//w:spacing"):
        pPr.remove(spacing)

    # 步骤3：合并属性值（新设置值>0则用新值，否则用原有值）
    final_before = before_lines if before_lines > 0 else existing_before
    final_after = after_lines if after_lines > 0 else existing_after

    # 步骤4：新建spacing节点，同时设置合并后的双侧属性
    spacing = OxmlElement("w:spacing")
    if final_before > 0:
        spacing.set(qn("w:beforeLines"), str(int(final_before * 100)))
    if final_after > 0:
        spacing.set(qn("w:afterLines"), str(int(final_after * 100)))

    pPr.append(spacing)


def set_paragraph_space_before_by_lines(paragraph: Paragraph, lines: float) -> None:
    """
    设置段落【段前间距】为指定行数（Word 原生 w:beforeLines 行单位，非 Pt 磅值）
    :param paragraph: 目标段落对象
    :param lines: 段前间距行数，浮点型（如 0.0/1.0/0.5），直接映射 Word w:beforeLines
    """
    _paragraph_space_by_lines(paragraph, before_lines=lines)


def set_paragraph_space_after_by_lines(paragraph: Paragraph, lines: float) -> None:
    """
    设置段落【段后间距】为指定行数（Word 原生 w:afterLines 行单位，非 Pt 磅值）
    :param paragraph: 目标段落对象
    :param lines: 段后间距行数，浮点型（如 0.0/1.0/0.5），直接映射 Word w:afterLines
    """
    _paragraph_space_by_lines(paragraph, after_lines=lines)
