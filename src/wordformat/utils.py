#! /usr/bin/env python
# @Time    : 2025/12/22 22:20
# @Author  : afish
# @File    : utils.py
import os
from typing import Any

import yaml
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from loguru import logger


def load_yaml_with_merge(file_path: str) -> dict[str, Any]:
    """
    加载 YAML 文件，并正确处理 <<: *anchor 合并语法。

    要求 YAML 文件使用标准的 YAML merge key (<<)。
    """
    with open(file_path, encoding="utf-8") as f:
        # 使用 FullLoader 支持 << 合并
        config = yaml.load(f, Loader=yaml.FullLoader)
    return config


def ensure_is_directory(path):
    """
    检查 path 是否为一个已存在的文件夹路径。
    如果不是（是文件、或路径不存在），则抛出 ValueError。
    """
    if not os.path.exists(path):
        raise ValueError(f"路径不存在: '{path}'")
    if not os.path.isdir(path):
        raise ValueError(f"路径不是一个文件夹（它可能是一个文件）: '{path}'")


def get_file_name(file_name: str) -> str:
    basename = os.path.basename(file_name)
    filename_without_ext, _ = os.path.splitext(basename)  # 提取docx文件名称
    return filename_without_ext


def get_paragraph_numbering_text(paragraph: Paragraph) -> str:
    """
    提取段落的自动编号文字（如"第一章"、"1.1"、"1."等）。

    Word 的自动编号存储在 numbering.xml 中，para.text 不包含编号文字。
    本函数从段落的 XML 中读取 numId 和 ilvl，查找对应的 lvlText 模板，
    然后根据当前编号计数器替换占位符，生成实际的编号文字。

    Args:
        paragraph: docx 段落对象

    Returns:
        编号文字字符串，无编号时返回空字符串
    """
    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is None:
        return ""
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        return ""

    numId_elem = numPr.find(qn("w:numId"))
    ilvl_elem = numPr.find(qn("w:ilvl"))
    if numId_elem is None:
        return ""

    num_id = numId_elem.get(qn("w:val"))
    ilvl = int(ilvl_elem.get(qn("w:val"))) if ilvl_elem is not None else 0

    if num_id is None or num_id == "0":
        return ""

    # 获取 numbering part
    try:
        numbering_part = paragraph.part.numbering_part
    except (AttributeError, KeyError, NotImplementedError):
        return ""

    numbering_elm = numbering_part._element

    # 查找 num -> abstractNum -> lvl -> lvlText
    num_elem = None
    for el in numbering_elm.findall(qn("w:num")):
        if el.get(qn("w:numId")) == num_id:
            num_elem = el
            break
    if num_elem is None:
        return ""

    abstract_num_id_ref = num_elem.find(qn("w:abstractNumId"))
    if abstract_num_id_ref is None:
        return ""

    abstract_num_id = abstract_num_id_ref.get(qn("w:val"))

    abstract_num = None
    for el in numbering_elm.findall(qn("w:abstractNum")):
        if el.get(qn("w:abstractNumId")) == abstract_num_id:
            abstract_num = el
            break
    if abstract_num is None:
        return ""

    lvl = None
    for el in abstract_num.findall(qn("w:lvl")):
        if el.get(qn("w:ilvl")) == str(ilvl):
            lvl = el
            break
    if lvl is None:
        return ""

    lvl_text_elem = lvl.find(qn("w:lvlText"))
    if lvl_text_elem is None:
        return ""

    lvl_text = lvl_text_elem.get(qn("w:val"), "")
    num_fmt_elem = lvl.find(qn("w:numFmt"))
    num_fmt = (
        num_fmt_elem.get(qn("w:val"), "decimal")
        if num_fmt_elem is not None
        else "decimal"
    )

    # 计算当前级别的编号值
    # 需要遍历文档中所有使用同一 abstractNum 的段落来计数
    level_counters = _count_numbering_levels(numbering_elm, abstract_num_id, paragraph)

    current_num = level_counters.get(ilvl, 1)

    # 根据格式化类型转换数字
    _format_number(current_num, num_fmt)

    # 替换 lvlText 中的占位符
    # %1 -> 当前级别, %2 -> 下一级别, etc.
    result = lvl_text
    for lvl_idx, lvl_val in sorted(level_counters.items()):
        placeholder = f"%{lvl_idx + 1}"
        if placeholder in result:
            fmt_for_level = _get_level_fmt(abstract_num, lvl_idx)
            result = result.replace(placeholder, _format_number(lvl_val, fmt_for_level))

    return result


def _count_numbering_levels(
    numbering_elm, abstract_num_id: str, target_paragraph: Paragraph
) -> dict[int, int]:
    """
    遍历文档段落，计算目标段落所在编号链的各级计数器值。

    返回 {ilvl: count} 字典，如 {0: 1, 1: 2} 表示一级编号为1，二级编号为2
    """
    counters = {}
    target_element = target_paragraph._element

    # 找到所有引用同一 abstractNum 的 numId
    num_ids = set()
    for num_elem in numbering_elm.findall(qn("w:num")):
        abstract_num_id_ref = num_elem.find(qn("w:abstractNumId"))
        if (
            abstract_num_id_ref is not None
            and abstract_num_id_ref.get(qn("w:val")) == abstract_num_id
        ):
            num_ids.add(num_elem.get(qn("w:numId")))

    if not num_ids:
        return counters

    # 遍历文档 body 中的所有段落
    body = target_paragraph._element.getparent()  # body element
    if body is None:
        return counters

    # 获取 isRestart 标志
    def get_is_restart(lvl_element):
        """检查该级别是否在每个上级编号重启"""
        if lvl_element is None:
            return True
        restart = lvl_element.find(qn("w:isLgl") if False else qn("w:lvlRestart"))
        # lvlRestart 不存在时默认行为：下级在上级重启
        return restart is None

    for para_elm in body.findall(qn("w:p")):
        pPr = para_elm.find(qn("w:pPr"))
        if pPr is None:
            continue
        numPr = pPr.find(qn("w:numPr"))
        if numPr is None:
            continue
        numId_elem = numPr.find(qn("w:numId"))
        if numId_elem is None:
            continue
        para_num_id = numId_elem.get(qn("w:val"))
        if para_num_id not in num_ids:
            continue

        ilvl_elem = numPr.find(qn("w:ilvl"))
        para_ilvl = int(ilvl_elem.get(qn("w:val"))) if ilvl_elem is not None else 0

        # 增加当前级别计数
        counters[para_ilvl] = counters.get(para_ilvl, 0) + 1

        # 重置下级计数
        for key in list(counters.keys()):
            if key > para_ilvl:
                del counters[key]

        # 检查是否到达目标段落
        if para_elm is target_element:
            break

    return counters


def _format_number(num: int, num_fmt: str) -> str:
    """根据 numFmt 将数字格式化为对应字符串"""
    format_map = {
        "decimal": str,
        "upperRoman": lambda n: _to_roman(n).upper(),
        "lowerRoman": lambda n: _to_roman(n).lower(),
        "upperLetter": lambda n: chr(ord("A") + n - 1) if 1 <= n <= 26 else str(n),
        "lowerLetter": lambda n: chr(ord("a") + n - 1) if 1 <= n <= 26 else str(n),
        "chineseCountingThousand": lambda n: _to_chinese_num(n),
        "ideographTraditional": lambda n: _to_chinese_num(n),
        "chineseCounting": lambda n: _to_chinese_num(n),
    }
    formatter = format_map.get(num_fmt, str)
    return formatter(num)


def _get_level_fmt(abstract_num, ilvl: int) -> str:
    """获取指定级别的 numFmt"""
    lvl = None
    for el in abstract_num.findall(qn("w:lvl")):
        if el.get(qn("w:ilvl")) == str(ilvl):
            lvl = el
            break
    if lvl is not None:
        num_fmt_elem = lvl.find(qn("w:numFmt"))
        if num_fmt_elem is not None:
            return num_fmt_elem.get(qn("w:val"), "decimal")
    return "decimal"


def _to_roman(num: int) -> str:
    """将整数转换为罗马数字"""
    if num <= 0:
        return "0"
    val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    syms = ["m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i"]
    roman = ""
    i = 0
    while num > 0:
        for _ in range(num // val[i]):
            roman += syms[i]
            num -= val[i]
        i += 1
    return roman


def _to_chinese_num(num: int) -> str:
    """将整数转换为中文数字（一到一百）"""
    if num <= 0:
        return str(num)
    digits = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
    if num < 10:
        return digits[num]
    if num < 100:
        tens = num // 10
        ones = num % 10
        if tens == 1:
            result = "十"
        else:
            result = digits[tens] + "十"
        if ones > 0:
            result += digits[ones]
        return result
    if num == 100:
        return "一百"
    return str(num)


def _from_roman(roman: str) -> int:
    """罗马数字转整数，如 'I'→1, 'IV'→4, 'X'→10"""
    roman = roman.strip().lower()
    if not roman:
        raise ValueError("Empty roman numeral")
    roman_map = {"i": 1, "v": 5, "x": 10, "l": 50, "c": 100, "d": 500, "m": 1000}
    result = 0
    prev_val = 0
    for char in reversed(roman):
        val = roman_map.get(char)
        if val is None:
            raise ValueError(f"Invalid roman numeral: '{roman}'")
        if val < prev_val:
            result -= val
        else:
            result += val
        prev_val = val
    return result


def _from_chinese_num(chinese: str) -> int:
    """中文数字转整数，如 '一'→1, '十二'→12, '一百'→100, '一百二十三'→123"""
    chinese = chinese.strip()
    if not chinese:
        raise ValueError("Empty chinese numeral")

    digit_map = {
        "零": 0,
        "一": 1,
        "二": 2,
        "三": 3,
        "四": 4,
        "五": 5,
        "六": 6,
        "七": 7,
        "八": 8,
        "九": 9,
        "壹": 1,
        "贰": 2,
        "叁": 3,
        "肆": 4,
        "伍": 5,
        "陆": 6,
        "柒": 7,
        "捌": 8,
        "玖": 9,
    }
    unit_map = {"十": 10, "拾": 10, "百": 100, "佰": 100, "千": 1000, "仟": 1000}

    # 纯数字字符
    if all(c in digit_map for c in chinese):
        result = 0
        for c in chinese:
            result = result * 10 + digit_map[c]
        return result

    result = 0
    section = 0
    for char in chinese:
        if char in digit_map:
            section = digit_map[char]
        elif char in unit_map:
            unit = unit_map[char]
            if section == 0:
                section = 1
            result += section * unit
            section = 0
        else:
            raise ValueError(f"Invalid chinese numeral character: '{char}'")
    result += section
    return result


def parse_caption_text(text: str) -> dict | None:
    """解析题注文本为结构化组件。

    支持格式：[续][标签][章节号][分隔符][题注编号] [题注名称]
    章节号支持阿拉伯数字、中文数字、罗马数字。
    分隔符支持 . - : — –
    支持 "续表"/"续图" 前缀，解析后 is_continued 为 True。

    Returns:
        {"label", "chapter_text", "separator", "number_text", "name",
         "chapter_num", "number_num", "is_continued"} 或 None
    """
    import re

    text = text.strip()
    if not text:
        return None

    # 检测 "续" 前缀（续表/续图）
    is_continued = text.startswith("续")
    if is_continued:
        text = text[1:].strip()

    if not text:
        return None

    # 分隔符：句点 . 、连字符 - 、冒号 : 、长划线 — (U+2014)、短划线 – (U+2013)
    SEP = r"[.\-:—–]"
    # 标签后可选空格（支持 "图 1.2 xxx" 和 "图1.2 xxx"）
    LABEL = r"([图表])\s*"

    # 模式1：图/表 + 阿拉伯数字章节 + 分隔符 + 阿拉伯数字编号 + 可选空格 + 名称
    m = re.match(rf"^{LABEL}(\d+)({SEP})(\d+)[\s　]+(.+)$", text)
    if m:
        label, ch_text, sep, num_text, name = m.groups()
        return {
            "label": label,
            "chapter_text": ch_text,
            "chapter_num": int(ch_text),
            "separator": sep,
            "number_text": num_text,
            "number_num": int(num_text),
            "name": name.strip(),
            "is_continued": is_continued,
        }

    # 模式2：图/表 + 中文数字章节 + 分隔符 + 阿拉伯数字编号 + 可选空格 + 名称
    chinese_num_chars = r"[一二三四五六七八九十百千零壹贰叁肆伍陆柒捌玖拾佰仟]+"
    m = re.match(rf"^{LABEL}({chinese_num_chars})({SEP})(\d+)[\s　]+(.+)$", text)
    if m:
        label, ch_text, sep, num_text, name = m.groups()
        try:
            ch_num = _from_chinese_num(ch_text)
        except ValueError:
            ch_num = None
        return {
            "label": label,
            "chapter_text": ch_text,
            "chapter_num": ch_num,
            "separator": sep,
            "number_text": num_text,
            "number_num": int(num_text),
            "name": name.strip(),
            "is_continued": is_continued,
        }

    # 模式3：图/表 + 罗马数字章节 + 分隔符 + 阿拉伯数字编号 + 可选空格 + 名称
    roman_chars = r"[IVXLCDMivxlcdm]+"
    m = re.match(rf"^{LABEL}({roman_chars})({SEP})(\d+)[\s　]+(.+)$", text)
    if m:
        label, ch_text, sep, num_text, name = m.groups()
        try:
            ch_num = _from_roman(ch_text)
        except ValueError:
            ch_num = None
        return {
            "label": label,
            "chapter_text": ch_text,
            "chapter_num": ch_num,
            "separator": sep,
            "number_text": num_text,
            "number_num": int(num_text),
            "name": name.strip(),
            "is_continued": is_continued,
        }

    # 尝试匹配编号部分也使用中文数字或罗马数字
    m = re.match(
        rf"^{LABEL}(\d+)({SEP})({chinese_num_chars}|{roman_chars})[\s　]+(.+)$", text
    )
    if m:
        label, ch_text, sep, num_text, name = m.groups()
        ch_num = int(ch_text)
        try:
            is_chinese = any(c in _digit_map for c in num_text)
            num_num = (
                _from_chinese_num(num_text) if is_chinese else _from_roman(num_text)
            )
        except ValueError:
            num_num = None
        return {
            "label": label,
            "chapter_text": ch_text,
            "chapter_num": ch_num,
            "separator": sep,
            "number_text": num_text,
            "number_num": num_num,
            "name": name.strip(),
            "is_continued": is_continued,
        }

    return None


# 供 parse_caption_text 内部使用的 digit_map
_digit_map = {
    "零": 0,
    "一": 1,
    "二": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
    "壹": 1,
    "贰": 2,
    "叁": 3,
    "肆": 4,
    "伍": 5,
    "陆": 6,
    "柒": 7,
    "捌": 8,
    "玖": 9,
}


def remove_all_numbering(doc):
    """
    强制解除样式与列表的绑定
    Args:
        doc:
    Returns:

    """
    title_style_names = ["Heading 1", "Heading 2", "Heading 3"]

    for style_name in title_style_names:
        if style_name in doc.styles:
            style = doc.styles[style_name]
            style_element = style._element

            # 删除 <w:pPr> 中的 numPr（样式级别的编号）
            pPr = style_element.find(qn("w:pPr"))
            if pPr is not None:
                numPr = pPr.find(qn("w:numPr"))
                if numPr is not None:
                    pPr.remove(numPr)

                # 可选：也删除 outlineLvl（大纲级别，有时触发编号）
                outlineLvl = pPr.find(qn("w:outlineLvl"))
                if outlineLvl is not None:
                    pPr.remove(outlineLvl)

            logger.debug(f"已解除样式 '{style_name}' 的编号绑定")


def ensure_directory_exists(path):
    """
    检查路径是否存在，如果不存在则创建对应的文件夹。

    参数:
        path (str): 需要检查或创建的文件夹路径

    说明:
        - 如果路径已存在且是文件夹，则不做任何操作
        - 如果路径不存在，则递归创建所有必需的父目录
        - 如果路径存在但是是文件，则抛出 ValueError
    """
    if os.path.exists(path):
        if not os.path.isdir(path):
            raise ValueError(f"路径已存在但不是文件夹：'{path}'")
    else:
        os.makedirs(path, exist_ok=True)
        logger.info(f"已创建文件夹：'{path}'")
