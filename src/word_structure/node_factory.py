#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 19:37
# @Author  : afish
# @File    : node_factory.py
from typing import Dict, Any, Optional

from loguru import logger

from src.rules.node import FormatNode
from src.word_structure.settings import CATEGORY_TO_CLASS


def create_node(item: Dict[str, Any], level: int, config: Dict[str, Any]) -> Optional[FormatNode]:
    """
    根据 item['category'] 创建对应的 FormatNode 子类实例。
    """
    category = item.get('category')
    if not category:
        raise ValueError("Item missing 'category' field")

    if category in CATEGORY_TO_CLASS:
        cls = CATEGORY_TO_CLASS[category]
        instance = cls(value=item, expected_rule={}, level=level)
        instance.load_config(config)
        return instance
    else:
        logger.warning(f"Unknown category: {category}, skipping.")
        return None
