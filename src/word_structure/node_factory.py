#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 19:37
# @Author  : afish
# @File    : node_factory.py
from typing import Dict, Any, Optional

from loguru import logger

from src.rules.body import BodyText
from src.rules.node import FormatNode
from src.word_structure.settings import CATEGORY_TO_CLASS, BODY_CATEGORIES


def create_node(item: Dict[str, Any], level: int) -> Optional[FormatNode]:
    """
    根据 item['category'] 创建对应的 FormatNode 子类实例。
    """
    category = item.get('category')
    if not category:
        raise ValueError("Item missing 'category' field")

    if category in CATEGORY_TO_CLASS:
        cls = CATEGORY_TO_CLASS[category]
        return cls(value=item, expected_rule={}, level=level)
    elif category in BODY_CATEGORIES:
        return BodyText(value=item, expected_rule={}, level=999)
    else:
        logger.warning(f"Unknown category: {category}, skipping.")
        return None
