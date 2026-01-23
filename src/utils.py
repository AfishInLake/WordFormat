#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2025/12/22 22:20
# @Author  : afish
# @File    : utils.py
import hashlib
from typing import Any, Dict

import yaml
from docx.text.paragraph import Paragraph
from lxml import etree


def get_paragraph_fingerprint(paragraph: Paragraph):
    """使用段落的前20个字生成指纹"""
    return hashlib.sha256(paragraph.text[:20].strip().encode("utf-8")).hexdigest()


def load_yaml_with_merge(file_path: str) -> Dict[str, Any]:
    """
    加载 YAML 文件，并正确处理 <<: *anchor 合并语法。

    要求 YAML 文件使用标准的 YAML merge key (<<)。
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        # 使用 FullLoader 支持 << 合并
        config = yaml.load(f, Loader=yaml.FullLoader)
    return config

