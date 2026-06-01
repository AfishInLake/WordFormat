#! /usr/bin/env python
# @Time    : 2026/1/11 19:37
# @Author  : afish
# @File    : node_factory.py
from typing import Any

from loguru import logger

from wordformat.rules.node import FormatNode
from wordformat.word_structure.settings import CATEGORY_TO_CLASS


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
        # 图片节点：从 JSON 的 meta 声明式字段初始化 content
        if category == "image":
            meta = item.get("meta", {}) if isinstance(item.get("meta"), dict) else {}
            instance.content["image_source"] = meta.get("image_source") or item.get(
                "image_source", ""
            )
            instance.content["image_sha256"] = meta.get("image_sha256") or item.get(
                "image_sha256", ""
            )
            instance.content["width_emu"] = meta.get("width_emu") or item.get(
                "width_emu", 0
            )
            instance.content["height_emu"] = meta.get("height_emu") or item.get(
                "height_emu", 0
            )
            instance.content["alignment"] = meta.get("alignment") or item.get(
                "alignment", "center"
            )
            instance.content["image_path"] = meta.get("image_path") or item.get(
                "image_path", ""
            )
        # 公式节点：从 JSON meta 初始化 latex
        if category == "formula":
            meta = item.get("meta", {}) if isinstance(item.get("meta"), dict) else {}
            instance.content["latex"] = meta.get("latex") or item.get("latex", "")
            instance.content["display"] = meta.get("display", True)
        # 表格节点：从 JSON meta 初始化 rows/cols/cells
        if category == "table":
            meta = item.get("meta", {}) if isinstance(item.get("meta"), dict) else {}
            instance.content["rows"] = meta.get("rows") or item.get("rows", 1)
            instance.content["cols"] = meta.get("cols") or item.get("cols", 1)
            instance.content["cells"] = meta.get("cells") or item.get("cells", [])
        instance.load_config(config)
        return instance
    else:
        logger.warning(f"Unknown category: {category}, skipping.")
        return None
