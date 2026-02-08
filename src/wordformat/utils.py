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
