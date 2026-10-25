#! /usr/bin/env python
# @Time    : 2026/1/11 19:37
# @Author  : afish
# @File    : node_factory.py
from typing import Any, Optional

from loguru import logger

from wordformat.rules.node import FormatNode
from wordformat.structure.registry import Registry
from wordformat.structure.settings import CATEGORY_TO_CLASS


def create_node(
    item: dict[str, Any],
    level: int,
    config: dict[str, Any],
    registry: Optional[Registry] = None,
) -> FormatNode | None:
    """
    根据 item['category'] 创建对应的 FormatNode 子类实例。

    :param registry: 领域注册表实例（get_registry(domain)）；None 时使用
        默认领域（论文）注册表，行为与历史一致。
    """
    category = item.get("category")
    if not category:
        raise ValueError(f"Item {item} missing 'category' field")

    cls = (
        registry.get(category)
        if registry is not None
        else CATEGORY_TO_CLASS.get(category)
    )
    if cls is None:
        logger.warning(f"Unknown category: {category}, skipping.")
        return None

    instance = cls(value=item, expected_rule={}, level=level)
    instance.load_config(config)
    return instance
