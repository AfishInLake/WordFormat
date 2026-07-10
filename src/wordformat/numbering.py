#! /usr/bin/env python
"""
标题自动编号 & 参考文献条目编号模块

功能：
1. 清除标题段落中的手动编号文本（正则匹配）
2. 应用 Word 自动编号（通过 XML 操作）
3. 清除参考文献条目中的手动编号
4. 应用参考文献条目的自动编号

流程：
  格式化完成后 → 对 heading / ReferenceEntry 节点 → 清除手动编号 → 应用自动编号
"""

import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Mm, Pt
from loguru import logger

from wordformat.style.units import extract_unit_from_string

# EMU 到 twips 的换算系数
_EMU_PER_TWIP = 635


def _auto_strip_numbering(paragraph, ilvl: int) -> bool:
    """自动检测并清除段落开头的手动编号文本。

    根据标题级别 'ilvl' 构建匹配列表，按优先级依次尝试。
    无需用户编写正则表达式。

    支持：
    - 中文编号：第一章、第一节、一、二／1. 、1.1.1、1.1.2
    - 英文编号：1. / 1.1 / 1.1.1 / I. / A.
    - 括号编号：（一）、(1)、（1）、1)
    - 带空格或标点后缀：1.1   、 一、
    """
    if not paragraph.runs:
        return False

    text = paragraph.text.lstrip("\u3000 ")  # 全角/半角空格
    if not text:
        return False

    # ---------- 公共原子 ----------
    _CH = "[一二三四五六七八九十百千零壹贰叁肆伍陆柒捌玖拾佰仟]"
    _CHN = (
        "(?:一|二|三|四|五|六|七|八|九|十|百|千|零|壹|贰|叁|肆|伍|陆|柒|捌|玖|拾|佰|仟)"
    )
    _NUM = r"\d+"
    _ROMAN = "[IVXLCDM]+"

    # 中文数字/阿拉伯数字 混合组
    _CHD = f"(?:{_CHN}|[0-9])"
    _CHDM = f"(?:{_CHN}|[0-9])+"

    # 按 ilvl 构建模式列表（优先级从高到低）
    patterns: list[str] = []

    if ilvl == 0:  # 一级标题
        patterns = [
            rf"^第{_CHDM}章\s*",  # 第一章、第1章
            rf"^第{_CHDM}节\s*",  # 第一节
            rf"^第{_CHDM}部分\s*",  # 第一部分
            rf"^（{_CHDM}）\s*",  # （一）、（1）
            rf"^\({_CHDM}\)\s*",  # (一)、(1)
            rf"^{_CHDM}\)\s*",  # 一)、1)
            rf"^{_ROMAN}\.\s+",  # I.  / V.
            rf"^{_CH}\s*[、，,\.]?\s*",  # 一、 / 一 / 一.
        ]
    elif ilvl == 1:  # 二级标题
        patterns = [
            rf"^{_NUM}\.{_NUM}\.{_NUM}\s*",  # 1.1.1 (先匹配更长的)
            rf"^{_NUM}\.{_NUM}\s+",  # 1.1   (带空格)
            rf"^{_NUM}\.{_NUM}\s*",  # 1.1
            rf"^{_CH}\.{_NUM}\s*",  # 一.1
            rf"^{_CH}\s*[、，,\.]?\s*",  # 一、/ 一.
        ]
    else:  # 三级及以上
        patterns = [
            rf"^{_NUM}\.{_NUM}\.{_NUM}\s*",  # 1.1.1
            rf"^{_NUM}\.{_NUM}\.{_NUM}\s+",  # 1.1.1
            rf"^{_NUM}\.{_NUM}\s*",  # 1.1
            rf"^{_NUM}\s*[、，,\.]?\s*",  # 1. / 1、
        ]

    # 去除所有尾部空格和零宽断言
    for pattern in patterns:
        match = re.match(pattern, text)
        if not match:
            continue
        stripped_len = len(match.group())
        _remove_chars(paragraph, stripped_len)
        return True

    return False


def _strip_reference_numbering(paragraph) -> bool:
    """检测并清除参考文献条目开头的手动编号。

    支持常见参考文献编号格式：
      [1]  [1]  1.  1)  (1)  [1]  ［1］  ①
    返回 True 表示已清除编号。
    """
    if not paragraph.runs:
        return False

    text = paragraph.text.lstrip("　 ")
    if not text:
        return False

    patterns = [
        r"^\[\d+\][\s　]*",  # [1]  [1]
        r"^［\d+］[\s　]*",  # ［1］（全角方括号）
        r"^\(\d+\)[\s　]*",  # (1)
        r"^（\d+）[\s　]*",  # （1）（全角括号）
        r"^\d+\.[\s　]+",  # 1.  (必须有空格，避免匹配年份如 2024.)
        r"^\d+\)[\s　]*",  # 1)
        r"^[①②③④⑤⑥⑦⑧⑨⑩][\s　]*",  # ①（带圈数字）
    ]

    for pattern in patterns:
        match = re.match(pattern, text)
        if match:
            _remove_chars(paragraph, len(match.group()))
            return True

    return False


def _remove_chars(paragraph, count: int) -> None:
    """从段落开头删除指定数量的字符，跨 run 操作。"""
    remaining = count
    for run in paragraph.runs:
        if remaining <= 0:
            break
        rt = run.text
        if len(rt) <= remaining:
            remaining -= len(rt)
            run.text = ""
        else:
            run.text = rt[remaining:]
            remaining = 0
    # 清除第一个非空 run 开头的空白
    for run in paragraph.runs:
        if run.text:
            run.text = run.text.lstrip("\u3000 ")
            break


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
def _build_numbering_rPr(headings_config, level_key: str):
    """为编号构建 w:rPr 元素，使编号的字体/字号跟随标题样式。"""
    from wordformat.style.defs import FontColor, FontSize
    from wordformat.style.xml_ops import (
        rPr_set_bold,
        rPr_set_font,
        rPr_set_font_color,
        rPr_set_font_size,
    )

    level_cfg = getattr(headings_config, level_key, None)
    if level_cfg is None:
        return None

    chinese_font = getattr(level_cfg, "chinese_font_name", None)
    english_font = getattr(level_cfg, "english_font_name", None)
    font_size = getattr(level_cfg, "font_size", None)
    font_color = getattr(level_cfg, "font_color", None)
    bold = getattr(level_cfg, "bold", None)

    if not any([chinese_font, english_font, font_size, font_color, bold is not None]):
        return None

    rPr = OxmlElement("w:rPr")

    if chinese_font or english_font:
        rPr_set_font(rPr, cn_name=chinese_font, en_name=english_font)

    if font_size:
        try:
            pt_val = FontSize(font_size).rel_value
            if pt_val == font_size:
                # 非中文标签（如 "15.5" 或 14），直接作为磅值使用
                pt_val = float(font_size)
            rPr_set_font_size(rPr, pt_val)
        except (ValueError, TypeError):
            pass

    if font_color:
        try:
            rPr_set_font_color(rPr, FontColor(font_color).rel_value)
        except Exception:
            pass

    if bold is not None:
        rPr_set_bold(rPr, bold)

    return rPr


def _traverse_numbering(
    node, heading_map, heading_num_map, config, ref_enabled, reference_num_id, counters
):
    """递归遍历节点树，应用标题和参考文献编号。"""
    from wordformat.rules.references import ReferenceEntry

    category = node.value.get("category", "") if isinstance(node.value, dict) else ""
    paragraph = getattr(node, "paragraph", None)

    if paragraph:
        # 处理标题节点
        for level_key, (ilvl_str, config_key) in heading_map.items():
            if category == f"heading_{level_key}":
                level_config = getattr(config, config_key, None)
                if level_config and level_config.enabled:
                    _auto_strip_numbering(paragraph, int(ilvl_str))
                    num_id = heading_num_map.get(config_key)
                    if num_id:
                        apply_auto_numbering(paragraph, num_id, ilvl_str)
                    counters["heading"] += 1
                break

        # 处理参考文献条目节点
        if ref_enabled and isinstance(node, ReferenceEntry):
            _strip_reference_numbering(paragraph)
            if reference_num_id:
                apply_auto_numbering(paragraph, reference_num_id, "0")
            counters["ref"] += 1

    for child in node.children:
        _traverse_numbering(
            child,
            heading_map,
            heading_num_map,
            config,
            ref_enabled,
            reference_num_id,
            counters,
        )


def process_heading_numbering(root_node, document, config, headings_config=None):
    """处理所有标题和参考文献条目的自动编号。"""
    if not config.enabled:
        return

    definitions = create_numbering_definition(document, config, headings_config)

    heading_map = {
        "level_1": ("0", "level_1"),
        "level_2": ("1", "level_2"),
        "level_3": ("2", "level_3"),
    }

    ref_config = getattr(config, "references", None)
    ref_enabled = isinstance(ref_config, dict) and ref_config.get("enabled", False)

    counters = {"heading": 0, "ref": 0}
    _traverse_numbering(
        root_node,
        heading_map,
        definitions.get("headings", {}),
        config,
        ref_enabled,
        definitions.get("references"),
        counters,
    )
    logger.info(
        f"编号处理完成：标题 {counters['heading']} 个，"
        f"参考文献条目 {counters['ref']} 个"
    )


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


def create_numbering_definition(document, config, headings_config=None) -> dict:
    """
    在文档中创建自动编号定义（如果不存在）。

    根据配置中的 template 生成 abstractNum 和 num 定义。

    返回:
      {
        "headings": {"level_1": num_id, ...},
        "references": num_id or None,
      }

    关键设计：
      - 所有标题级别共用同一个 abstractNum，确保多级编号的计数器正确联动
      - 参考文献条目使用独立的 abstractNum，不与标题编号互相干扰
    """
    if not config.enabled:
        return {"headings": {}, "references": None}

    # 获取或创建 numbering part
    try:
        numbering_part = document.part.numbering_part
    except (AttributeError, KeyError, NotImplementedError):
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from docx.opc.packuri import PackURI
        from docx.parts.numbering import NumberingPart

        numbering_elm = OxmlElement("w:numbering")
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

    heading_num_map = {}
    reference_num_id = None

    # ========================
    # 1. 标题编号定义
    # ========================
    level_configs = [
        ("level_1", config.level_1, 0),
        ("level_2", config.level_2, 1),
        ("level_3", config.level_3, 2),
    ]
    enabled_heading_levels = [
        (key, lcfg, ilvl)
        for key, lcfg, ilvl in level_configs
        if lcfg.enabled and lcfg.template
    ]

    if enabled_heading_levels:
        abstract_num_id = max_abstract_num_id + 1
        num_id = max_num_id + 1
        max_abstract_num_id += 1
        max_num_id += 1

        abstract_num = OxmlElement("w:abstractNum")
        abstract_num.set(qn("w:abstractNumId"), str(abstract_num_id))

        for level_key, level_config, ilvl in enabled_heading_levels:
            lvl = _build_numbering_level(level_key, level_config, ilvl, headings_config)
            abstract_num.append(lvl)

        numbering_elm.append(abstract_num)

        num = OxmlElement("w:num")
        num.set(qn("w:numId"), str(num_id))
        abstract_num_id_ref = OxmlElement("w:abstractNumId")
        abstract_num_id_ref.set(qn("w:val"), str(abstract_num_id))
        num.append(abstract_num_id_ref)
        numbering_elm.append(num)

        logger.debug(
            f"创建标题编号定义: abstractNumId={abstract_num_id}, numId={num_id}"
        )
        heading_num_map = {key: str(num_id) for key, _, _ in enabled_heading_levels}

    # ========================
    # 2. 参考文献条目编号定义
    # ========================
    ref_config = getattr(config, "references", None)
    if (
        isinstance(ref_config, dict)
        and ref_config.get("enabled", False)
        and ref_config.get("template")
    ):
        ref_abstract_num_id = max_abstract_num_id + 1
        ref_num_id = max_num_id + 1

        ref_abstract_num = OxmlElement("w:abstractNum")
        ref_abstract_num.set(qn("w:abstractNumId"), str(ref_abstract_num_id))

        # 参考文献只有一级（ilvl=0），numFmt 始终为 decimal
        ref_lvl = _build_numbering_level(None, ref_config, 0, None)
        ref_abstract_num.append(ref_lvl)

        numbering_elm.append(ref_abstract_num)

        ref_num = OxmlElement("w:num")
        ref_num.set(qn("w:numId"), str(ref_num_id))
        ref_abstract_num_id_ref = OxmlElement("w:abstractNumId")
        ref_abstract_num_id_ref.set(qn("w:val"), str(ref_abstract_num_id))
        ref_num.append(ref_abstract_num_id_ref)
        numbering_elm.append(ref_num)

        reference_num_id = str(ref_num_id)
        logger.debug(
            f"创建参考文献编号定义: abstractNumId={ref_abstract_num_id}, numId={ref_num_id}"
        )

    return {"headings": heading_num_map, "references": reference_num_id}


def _build_numbering_level(
    level_key: str | None,
    level_config,
    ilvl: int,
    headings_config,
) -> "OxmlElement":
    """构建单个 w:lvl 元素，供标题和参考文献共用。

    Args:
        level_key: 标题级别键（如 "level_1"），参考文献时为 None
        level_config: 标题编号配置 dict
        ilvl: 编号级别（0-based）
        headings_config: 标题配置（仅标题需要，参考文献传 None）
    """
    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), str(ilvl))

    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    lvl.append(start)

    numFmt = OxmlElement("w:numFmt")
    template = level_config.template or ""
    if "第" in template and "章" in template:
        numFmt.set(qn("w:val"), "chineseCountingThousand")
    else:
        numFmt.set(qn("w:val"), "decimal")
    lvl.append(numFmt)

    lvlText = OxmlElement("w:lvlText")
    lvlText.set(qn("w:val"), template)
    lvl.append(lvlText)

    suff = OxmlElement("w:suff")
    suff_val = level_config.suffix or "space"
    suff.set(qn("w:val"), suff_val)
    lvl.append(suff)

    lvlJc = OxmlElement("w:lvlJc")
    lvlJc.set(qn("w:val"), "left")
    lvl.append(lvlJc)

    pPr = OxmlElement("w:pPr")
    ind = OxmlElement("w:ind")

    if level_config.numbering_indent:
        _set_indent_value(ind, "left", level_config.numbering_indent)
    else:
        ind.set(qn("w:left"), "0")

    if level_config.text_indent:
        _set_indent_value(ind, "hanging", level_config.text_indent)
    else:
        ind.set(qn("w:hanging"), str((ilvl + 1) * 210))

    pPr.append(ind)
    lvl.append(pPr)

    # rPr — 编号字体/字号跟随样式（仅标题有此需求）
    if headings_config and level_key:
        rPr = _build_numbering_rPr(headings_config, level_key)
        if rPr is not None:
            lvl.append(rPr)

    return lvl
