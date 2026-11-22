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
    strict: bool = False,
) -> FormatNode | None:
    """
    根据 item['category'] 创建对应的 FormatNode 子类实例。

    :param registry: 领域注册表实例（get_registry(domain)）；None 时使用
        默认领域（论文）注册表，行为与历史一致。
    :param strict: 未知类别是否直接报错。用户显式输入路径（jd/sync）
        应传 True 以在入口 fail-fast；模型分类结果保留容错（默认）。
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
        table = (
            {k for k, _ in registry.items()}
            if registry is not None
            else set(CATEGORY_TO_CLASS)
        )
        message = f"未注册的节点类别 {category!r}；可用类别: {sorted(table)}"
        if strict:
            raise ValueError(message)
        # 模型分类结果容错：跳过并留痕（准确类别由后处理修正）
        logger.warning(f"{message}，跳过该条目")
        return None

    instance = cls(value=item, expected_rule={}, level=level)
    instance.load_config(config)
    return instance
