#! /usr/bin/env python
"""
标题自动编号模块

功能：
1. 清除标题段落中的手动编号文本（正则匹配）
2. 应用 Word 自动编号（通过 XML 操作）

流程：
  格式化完成后 → 对 heading 节点 → 清除手动编号 → 应用自动编号
"""

import re
from copy import deepcopy
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Mm, Inches, Pt
from wordformat.style.utils import extract_unit_from_string
from loguru import logger


# EMU 到 twips 的换算系数
# docx.shared 的 Cm/Mm/Inches/Pt 返回 EMU，w:ind 需要 twips
# 1 twip = 635 EMU（即 914400 EMU/inch ÷ 1440 twips/inch）
_EMU_PER_TWIP = 635


def _convert_to_twips(value_str: str) -> int:
    """
    将带单位的缩进值转换为 twips（Word 内部单位）。
    使用 docx.shared 官方 API 进行单位换算，确保精度。

    仅用于物理单位（厘米、毫米、英寸、磅），不处理字符单位。
    例如："0.75cm" → 425, "12磅" → 240
    """
    result = extract_unit_from_string(value_str)
    if not result.is_valid or result.value is None:
        logger.warning(f"无法解析缩进值: {value_str}，使用默认值 0")
        return 0

    unit = result.standard_unit
    val = result.value

    try:
        if unit == "cm":
            emu = int(Cm(val))
        elif unit == "mm":
            emu = int(Mm(val))
        elif unit == "inch":
            emu = int(Inches(val))
        elif unit == "pt":
            emu = int(Pt(val))
        else:
            logger.warning(f"不支持的物理单位: {unit}（{value_str}），使用默认值 0")
            return 0
    except Exception as e:
        logger.warning(f"单位换算失败: {value_str}，错误: {e}，使用默认值 0")
        return 0

    # EMU → twips（使用 round 避免整数除法精度损失）
    return round(emu / _EMU_PER_TWIP)


def _set_indent_value(ind_element, indent_type: str, value_str: str):
    """
    设置 w:ind 元素的缩进值，自动区分字符单位和物理单位。

    字符单位：使用 w:leftChars / w:hangingChars（1字符=100单位）
    物理单位：使用 w:left / w:hanging（twips）

    Args:
        ind_element: w:ind OxmlElement
        indent_type: "left" 或 "hanging"
        value_str: 带单位的值，如 "0字符"、"0.75cm"、"420磅"
    """
    result = extract_unit_from_string(value_str)
    if not result.is_valid or result.value is None:
        logger.warning(f"无法解析缩进值: {value_str}，跳过设置")
        return

    unit = result.standard_unit
    val = result.value

    if unit in ("字符", "char"):
        # 字符单位：使用专用 *Chars 属性（1字符 = 100 单位）
        chars_int = int(round(float(val) * 100))
        attr_name = qn(f"w:{indent_type}Chars")
        ind_element.set(attr_name, str(chars_int))
        # 不设置物理属性，Word 会优先使用 *Chars
    elif unit in ("cm", "mm", "inch", "pt"):
        # 物理单位：使用 twips 属性
        twips = _convert_to_twips(value_str)
        ind_element.set(qn(f"w:{indent_type}"), str(twips))
    else:
        logger.warning(f"不支持的缩进单位: {unit}（{value_str}），跳过设置")


# 中文字号到半磅值（half-points）的映射
# Word 内部字号单位为 half-points（半磅），如 12pt = 24 half-points
_FONT_SIZE_HALF_PT_MAP = {
    "一号": 52, "小一": 48, "二号": 44, "小二": 36,
    "三号": 32, "小三": 30, "四号": 28, "小四": 24,
    "五号": 21, "小五": 18, "六号": 15, "七号": 11,
}


def _build_numbering_rPr(headings_config, level_key: str):
    """
    为编号构建 w:rPr 元素，使编号的字体/字号跟随标题样式。

    Args:
        headings_config: 标题配置对象（config.headings）
        level_key: "level_1" / "level_2" / "level_3"

    Returns:
        w:rPr OxmlElement 或 None
    """
    level_cfg = getattr(headings_config, level_key, None)
    if level_cfg is None:
        return None

    # 获取字体和字号
    chinese_font = getattr(level_cfg, "chinese_font_name", None)
    english_font = getattr(level_cfg, "english_font_name", None)
    font_size = getattr(level_cfg, "font_size", None)
    bold = getattr(level_cfg, "bold", None)

    if not chinese_font and not english_font and not font_size and bold is None:
        return None

    rPr = OxmlElement("w:rPr")

    # 设置字体
    rFonts = OxmlElement("w:rFonts")
    has_font = False
    if chinese_font:
        rFonts.set(qn("w:eastAsia"), chinese_font)
        has_font = True
    if english_font:
        rFonts.set(qn("w:ascii"), english_font)
        rFonts.set(qn("w:hAnsi"), english_font)
        has_font = True
    if has_font:
        rPr.append(rFonts)

    # 设置字号（Word 内部单位为 half-points）
    if font_size:
        half_pt = _FONT_SIZE_HALF_PT_MAP.get(font_size)
        if half_pt is None:
            try:
                half_pt = int(float(font_size) * 2)
            except (ValueError, TypeError):
                half_pt = None
        if half_pt is not None:
            sz = OxmlElement("w:sz")
            sz.set(qn("w:val"), str(half_pt))
            rPr.append(sz)
            szCs = OxmlElement("w:szCs")
            szCs.set(qn("w:val"), str(half_pt))
            rPr.append(szCs)

    # 设置加粗
    if bold is True:
        b = OxmlElement("w:b")
        rPr.append(b)
        bCs = OxmlElement("w:bCs")
        rPr.append(bCs)

    return rPr


def strip_manual_numbering(paragraph, pattern: str) -> bool:
    """
    清除段落开头的手动编号文本。

    仅修改第一个 run 的文本，保留其余 run 的格式不变。
    如果段落文本不匹配正则，不做任何修改。

    Args:
        paragraph: docx 段落对象
        pattern: 正则表达式，用于匹配段落开头的编号

    Returns:
        bool: 是否成功清除了编号
    """
    if not pattern or not paragraph.runs:
        return False

    full_text = paragraph.text
    match = re.match(pattern, full_text)
    if not match:
        return False

    # 计算需要删除的字符数
    stripped_len = match.end()

    # 从第一个 run 开始逐字符删除
    remaining = stripped_len
    for run in paragraph.runs:
        if remaining <= 0:
            break
        run_text = run.text
        if len(run_text) <= remaining:
            # 整个 run 都要删除
            remaining -= len(run_text)
            run.text = ""
        else:
            # 只删除 run 的前 remaining 个字符
            run.text = run_text[remaining:]
            remaining = 0

    # 清除第一个 run 开头可能残留的空白
    if paragraph.runs and paragraph.runs[0].text:
        paragraph.runs[0].text = paragraph.runs[0].text.lstrip()

    logger.debug(f"已清除手动编号: '{match.group()}' → 段落剩余: '{paragraph.text[:30]}...'")
    return True


def apply_auto_numbering(paragraph, num_id: str, ilvl: str = "0"):
    """
    为段落应用 Word 自动编号。

    通过在段落的 <w:pPr> 中添加 <w:numPr> 元素，
    引用文档中已有的 numbering 定义。

    Args:
        paragraph: docx 段落对象
        num_id: 编号定义 ID（对应 numbering.xml 中的 w:numId）
        ilvl: 编号级别（默认 "0"）
    """
    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        paragraph._element.insert(0, pPr)

    # 移除已有的 numPr（避免重复）
    existing_numPr = pPr.find(qn("w:numPr"))
    if existing_numPr is not None:
        pPr.remove(existing_numPr)

    # 创建新的 numPr
    numPr = OxmlElement("w:numPr")
    ilvl_elem = OxmlElement("w:ilvl")
    ilvl_elem.set(qn("w:val"), ilvl)
    numId_elem = OxmlElement("w:numId")
    numId_elem.set(qn("w:val"), num_id)
    numPr.append(ilvl_elem)
    numPr.append(numId_elem)
    pPr.append(numPr)

    logger.debug(f"已应用自动编号: numId={num_id}, ilvl={ilvl}")


def create_numbering_definition(document, config, headings_config=None) -> dict[str, str]:
    """
    在文档中创建自动编号定义（如果不存在）。

    根据配置中的 template 生成 abstractNum 和 num 定义，
    返回 {level_key: num_id} 映射。

    Args:
        document: docx Document 对象
        config: NumberingConfig 配置对象
        headings_config: 标题配置（用于设置编号字体/字号跟随标题样式）

    Returns:
        dict: {"level_1": "100", "level_2": "101", "level_3": "102"}
    """
    if not config.enabled:
        return {}

    # 获取或创建 numbering part
    try:
        numbering_part = document.part.numbering_part
    except (AttributeError, KeyError, NotImplementedError):
        # 文档没有 numbering part，需要手动创建
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from docx.opc.packuri import PackURI
        from docx.parts.numbering import NumberingPart

        numbering_elm = OxmlElement("w:numbering")

        # 手动创建 NumberingPart 实例并建立关系
        numbering_part = NumberingPart(
            PackURI("/word/numbering.xml"),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
            numbering_elm,
            document.part.package,
        )
        document.part.relate_to(numbering_part, RT.NUMBERING)

    numbering_elm = numbering_part._element

    # 查找已有的最大 abstractNumId 和 numId
    max_abstract_num_id = -1
    max_num_id = 0
    for elem in numbering_elm.findall(qn("w:abstractNum")):
        num_id = int(elem.get(qn("w:abstractNumId"), "0"))
        max_abstract_num_id = max(max_abstract_num_id, num_id)
    for elem in numbering_elm.findall(qn("w:num")):
        num_id = int(elem.get(qn("w:numId"), "0"))
        max_num_id = max(max_num_id, num_id)

    result = {}
    level_configs = [
        ("level_1", config.level_1, 0),
        ("level_2", config.level_2, 1),
        ("level_3", config.level_3, 2),
    ]

    for level_key, level_config, ilvl in level_configs:
        if not level_config.enabled or not level_config.template:
            continue

        abstract_num_id = max_abstract_num_id + 1
        max_abstract_num_id += 1
        num_id = max_num_id + 1
        max_num_id += 1

        # 创建 abstractNum
        abstract_num = OxmlElement("w:abstractNum")
        abstract_num.set(qn("w:abstractNumId"), str(abstract_num_id))

        # 创建多级 lvl（需要创建所有上级级别才能正确计数）
        for lvl_i in range(ilvl + 1):
            lvl = OxmlElement("w:lvl")
            lvl.set(qn("w:ilvl"), str(lvl_i))

            start = OxmlElement("w:start")
            start.set(qn("w:val"), "1")
            lvl.append(start)

            numFmt = OxmlElement("w:numFmt")
            # 根据模板判断格式
            template = level_config.template
            if "第" in template and "章" in template:
                numFmt.set(qn("w:val"), "chineseCountingThousand")
            else:
                numFmt.set(qn("w:val"), "decimal")
            lvl.append(numFmt)

            lvlText = OxmlElement("w:lvlText")
            # 根据级别生成 lvlText
            if lvl_i == ilvl:
                lvlText.set(qn("w:val"), level_config.template)
            else:
                # 上级级别使用简单的 %N 格式
                lvlText.set(qn("w:val"), f"%{lvl_i + 1}")
            lvl.append(lvlText)

            # 编号后缀：tab（制表符）、space（空格）、nothing（无）
            suff = OxmlElement("w:suff")
            suff_val = level_config.suffix or "space"
            suff.set(qn("w:val"), suff_val)
            lvl.append(suff)

            lvlJc = OxmlElement("w:lvlJc")
            lvlJc.set(qn("w:val"), "left")
            lvl.append(lvlJc)

            # pPr - 缩进设置
            pPr = OxmlElement("w:pPr")
            ind = OxmlElement("w:ind")

            # 计算编号缩进（left）：编号起始位置距左边距
            if level_config.numbering_indent:
                _set_indent_value(ind, "left", level_config.numbering_indent)
            else:
                ind.set(qn("w:left"), "0")

            # 计算文本缩进（hanging）：文本相对于编号起始位置的偏移量
            if level_config.text_indent:
                _set_indent_value(ind, "hanging", level_config.text_indent)
            else:
                # 默认：每级缩进约 0.3cm（210 twips）
                ind.set(qn("w:hanging"), str((lvl_i + 1) * 210))

            pPr.append(ind)
            lvl.append(pPr)

            # rPr - 编号的字体/字号跟随标题样式
            if headings_config and lvl_i == ilvl:
                rPr = _build_numbering_rPr(headings_config, level_key)
                if rPr is not None:
                    lvl.append(rPr)

            abstract_num.append(lvl)

        numbering_elm.append(abstract_num)

        # 创建 num 引用
        num = OxmlElement("w:num")
        num.set(qn("w:numId"), str(num_id))
        abstract_num_id_ref = OxmlElement("w:abstractNumId")
        abstract_num_id_ref.set(qn("w:val"), str(abstract_num_id))
        num.append(abstract_num_id_ref)
        numbering_elm.append(num)

        result[level_key] = str(num_id)
        logger.debug(f"创建编号定义: {level_key} → numId={num_id}, template={level_config.template}")

    return result


def process_heading_numbering(root_node, document, config, headings_config=None):
    """
    处理所有标题节点的编号：清除手动编号 + 应用自动编号。

    Args:
        root_node: 文档树根节点
        document: docx Document 对象
        config: NumberingConfig 配置对象
        headings_config: 标题配置（用于编号字体/字号跟随标题样式）
    """
    if not config.enabled:
        return

    # 创建编号定义（传入标题配置以同步编号样式）
    num_id_map = create_numbering_definition(document, config, headings_config)
    if not num_id_map:
        return

    # 标题节点类型到配置级别的映射
    heading_map = {
        "level_1": ("level_1", "0"),
        "level_2": ("level_2", "1"),
        "level_3": ("level_3", "2"),
    }

    def traverse(node):
        """递归遍历文档树"""
        # 获取节点的 category
        category = node.value.get("category", "") if isinstance(node.value, dict) else ""

        for level_key, (config_key, ilvl) in heading_map.items():
            if category == f"heading_{level_key}":
                level_config = getattr(config, config_key, None)
                if not level_config or not level_config.enabled:
                    break

                paragraph = getattr(node, "paragraph", None)
                if not paragraph:
                    break

                # 1. 清除手动编号
                if level_config.strip_pattern:
                    strip_manual_numbering(paragraph, level_config.strip_pattern)

                # 2. 应用自动编号
                num_id = num_id_map.get(config_key)
                if num_id:
                    apply_auto_numbering(paragraph, num_id, ilvl)

                break

        for child in node.children:
            traverse(child)

    traverse(root_node)
    logger.info("标题自动编号处理完成")
