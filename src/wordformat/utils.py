#! /usr/bin/env python
# @Time    : 2025/12/22 22:20
# @Author  : afish
# @File    : utils.py
import hashlib
import os
from collections import Counter
from typing import Any

import yaml
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from loguru import logger
from lxml import etree


def check_duplicate_fingerprints(data):
    """
    检查 JSON 列表中字典的 'fingerprint' 字段是否有重复。

    参数:
        data (list): 包含字典的列表，每个字典应有 'fingerprint' 键

    打印:
        重复的 fingerprint 及其出现次数
    """
    # 提取所有 fingerprint 值
    try:
        fingerprints = [item["fingerprint"] for item in data]
    except KeyError as e:
        raise ValueError(f"数据中缺少 'fingerprint' 字段: {e}") from e

    # 统计每个 fingerprint 的出现次数
    counter = Counter(fingerprints)

    # 找出重复项（出现次数 > 1）
    duplicates = {fp: count for fp, count in counter.items() if count > 1}

    if duplicates:
        logger.warning("发现重复的 fingerprint：")
        for fp, count in duplicates.items():
            logger.warning(f"  - {fp} （出现 {count} 次）")
    else:
        logger.info("未发现重复的 fingerprint。")


def get_paragraph_xml_fingerprint(paragraph: Paragraph):
    xml_str = etree.tostring(paragraph._element, encoding="utf-8", method="xml")
    return hashlib.sha256(xml_str).hexdigest()


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
    except (AttributeError, KeyError):
        return ""

    numbering_elm = numbering_part._element

    # 查找 num -> abstractNum -> lvl -> lvlText
    num_elem = numbering_elm.find(qn(f"w:num[@w:numId='{num_id}']"))
    if num_elem is None:
        return ""

    abstract_num_id_ref = num_elem.find(qn("w:abstractNumId"))
    if abstract_num_id_ref is None:
        return ""

    abstract_num_id = abstract_num_id_ref.get(qn("w:val"))

    abstract_num = numbering_elm.find(
        qn(f"w:abstractNum[@w:abstractNumId='{abstract_num_id}']")
    )
    if abstract_num is None:
        return ""

    lvl = abstract_num.find(qn(f"w:lvl[@w:ilvl='{ilvl}']"))
    if lvl is None:
        return ""

    lvl_text_elem = lvl.find(qn("w:lvlText"))
    if lvl_text_elem is None:
        return ""

    lvl_text = lvl_text_elem.get(qn("w:val"), "")
    num_fmt_elem = lvl.find(qn("w:numFmt"))
    num_fmt = num_fmt_elem.get(qn("w:val"), "decimal") if num_fmt_elem is not None else "decimal"

    # 计算当前级别的编号值
    # 需要遍历文档中所有使用同一 abstractNum 的段落来计数
    level_counters = _count_numbering_levels(numbering_elm, abstract_num_id, paragraph)

    current_num = level_counters.get(ilvl, 1)

    # 根据格式化类型转换数字
    formatted_num = _format_number(current_num, num_fmt)

    # 替换 lvlText 中的占位符
    # %1 -> 当前级别, %2 -> 下一级别, etc.
    result = lvl_text
    for lvl_idx, lvl_val in sorted(level_counters.items()):
        placeholder = f"%{lvl_idx + 1}"
        if placeholder in result:
            fmt_for_level = _get_level_fmt(abstract_num, lvl_idx)
            result = result.replace(placeholder, _format_number(lvl_val, fmt_for_level))

    return result


def _count_numbering_levels(numbering_elm, abstract_num_id: str, target_paragraph: Paragraph) -> dict[int, int]:
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
        if abstract_num_id_ref is not None and abstract_num_id_ref.get(qn("w:val")) == abstract_num_id:
            num_ids.add(num_elem.get(qn("w:numId")))

    if not num_ids:
        return counters

    # 遍历文档 body 中的所有段落
    body = target_paragraph.part._element.getparent()  # body element
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
    lvl = abstract_num.find(qn(f"w:lvl[@w:ilvl='{ilvl}']"))
    if lvl is not None:
        num_fmt_elem = lvl.find(qn("w:numFmt"))
        if num_fmt_elem is not None:
            return num_fmt_elem.get(qn("w:val"), "decimal")
    return "decimal"


def _to_roman(num: int) -> str:
    """将整数转换为罗马数字"""
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
    units = ["", "十", "百", "千"]
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
    return str(num)


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
