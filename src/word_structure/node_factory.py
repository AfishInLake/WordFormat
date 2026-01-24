#! /usr/bin/env python
# @Time    : 2026/1/11 19:37
# @Author  : afish
# @File    : node_factory.py
from typing import Any

from loguru import logger

from src.rules.node import FormatNode
from src.word_structure.settings import CATEGORY_TO_CLASS


def create_node(
    item: dict[str, Any], level: int, config: dict[str, Any]
) -> FormatNode | None:
    """
    根据 item['category'] 创建对应的 FormatNode 子类实例。
    """
    category = item.get("category")
    if not category:
        raise ValueError(f"Item {item} missing 'category' field")

    if category in CATEGORY_TO_CLASS:
        cls = CATEGORY_TO_CLASS[category]
        instance = cls(value=item, expected_rule={}, level=level)
        instance.load_config(config)
        return instance
    else:
        logger.warning(f"Unknown category: {category}, skipping.")
        return None
