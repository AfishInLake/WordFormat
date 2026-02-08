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
from loguru import logger


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


def set_paragraph_space_before(paragraph: Paragraph, before_lines: float):
    """
    设置段落**段前间距**（单位：行），与paragraph_get_space_before逻辑完全对齐
    核心规则：Lines属性存储（×100）→ 自动创建缺失节点 → 兼容Word原生逻辑
    :param paragraph: 目标Paragraph对象
    :param before_lines: 段前间距（行），支持浮点数（如0.5/1/1.5）
    """
    _set_paragraph_spacing(paragraph, spacing_type="before", target_lines=before_lines)


def set_paragraph_space_after(paragraph: Paragraph, after_lines: float):
    """
    设置段落**段后间距**（单位：行），与paragraph_get_space_after逻辑完全对齐
    :param paragraph: 目标Paragraph对象
    :param after_lines: 段后间距（行），支持浮点数（如0.5/1/1.5）
    """
    _set_paragraph_spacing(paragraph, spacing_type="after", target_lines=after_lines)


def _set_paragraph_spacing(
    paragraph: Paragraph, spacing_type: str, target_lines: float
):
    """
    核心通用设置逻辑：段前/段后间距统一处理（行→Lines属性×100）
    完全复刻Word逻辑：
    1. 优先设置w:XXLines属性（值=行×100，Word原生存储规则）
    2. 自动创建pPr/spacing缺失的XML节点，无空值报错
    3. 清除兜底的w:XX twips属性，避免与Lines属性冲突
    4. 支持浮点数行值（如0.5行→50，1.2行→120）
    """
    try:
        # 校验行值：非负数，过大值兜底（避免排版异常）
        target_lines = max(float(target_lines), 0.0)
        if target_lines > 10.0:
            logger.warning(f"间距{target_lines}行过大，自动兜底为10行")
            target_lines = 10.0

        # 核心：Word中Lines属性以「1/100行」为单位存储（如0.5行→50，1行→100）
        lines_100unit = int(round(target_lines * 100))  # 四舍五入为整数，匹配Word存储

        p = paragraph._element
        # 步骤1：创建/获取pPr节点（段落属性根节点，无则创建）
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            pPr = parse_xml(
                r'<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
            )
            p.append(pPr)

        # 步骤2：创建/获取spacing节点（间距节点，无则创建）
        spacing = pPr.find(qn("w:spacing"))
        if spacing is None:
            spacing = parse_xml(
                r'<w:spacing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
            )
            pPr.append(spacing)

        # 步骤3：设置核心属性w:XXLines（优先级最高，与获取函数对应）
        spacing.set(qn(f"w:{spacing_type}Lines"), str(lines_100unit))

        # 步骤4：清除兜底的w:XX twips属性，避免冲突（Word会优先读取Lines，清除后更纯净）
        if spacing.get(qn(f"w:{spacing_type}")) is not None:
            spacing.attrib.pop(qn(f"w:{spacing_type}"))

    except Exception as e:
        logger.error(f"设置段落{spacing_type}间距{target_lines}行失败: {e}")


def set_paragraph_first_line_indent(  # noqa C901
    para: Paragraph, value: float | int, unit: str = "char"
):
    """
    设置首行缩进，支持字符/物理单位，写入Word原生XML字段（firstLineChars/firstLine）
    :param para: 目标段落对象
    :param value: 缩进值（字符单位为浮点数/整数，如2/2.0；物理单位为整数pt，如24）
    :param unit: 缩进单位，char=字符单位（默认），pt=物理单位
    :return: 设置成功返回True，失败返回False
    """
    try:
        if unit not in ["char", "pt"]:
            logger.error(f"不支持的缩进单位：{unit}，仅支持char/pt")
            return False
        if value < 0:
            logger.warning("缩进值不能为负，自动置为0")
            value = 0

        p = para._element
        pPr = p.get_or_add_pPr()

        # 保留 left / right
        existing_left = None
        existing_right = None
        ind = pPr.find(qn("w:ind"))
        if ind is not None:
            existing_left = ind.get(qn("w:left"))
            existing_right = ind.get(qn("w:right"))
            pPr.remove(ind)

        attrs = {}
        if existing_left is not None:
            attrs["w:left"] = existing_left
        if existing_right is not None:
            attrs["w:right"] = existing_right

        if value == 0:
            # ✅ 现实主义方案：同时写 firstLine=0 和 firstLineChars=0
            # 虽然冗余，但能确保 WPS/Word 都清零
            attrs["w:firstLine"] = "0"
            attrs["w:firstLineChars"] = "0"
            attrs["w:hanging"] = "0"  # 顺手清 hanging
        else:
            if unit == "char":
                chars_int = int(round(float(value) * 100))
                attrs["w:firstLineChars"] = str(chars_int)
                # 同时写 firstLine 作为 fallback（可选）
                # twips = chars_int * 20 // 100  # 粗略换算
                # attrs["w:firstLine"] = str(twips)
            else:
                indent_twips = int(value) * 20
                attrs["w:firstLine"] = str(indent_twips)
                # 可选：也写 firstLineChars 作为 fallback
                # chars = int(indent_twips / 20)  # 不精确，慎用

        attr_str = " ".join(f'{k}="{v}"' for k, v in attrs.items())
        ind_xml = f'<w:ind xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" {attr_str}/>'  # noqa E501

        new_ind = parse_xml(ind_xml)
        pPr.append(new_ind)
        return True

    except Exception as e:
        logger.error(f"设置首行缩进失败：{e}")
        return False
