#! /usr/bin/env python
# @Time    : 2026/1/11 19:39
# @Author  : afish
# @File    : tree_builder.py
# wordformat/tree_builder.py

from typing import Optional

from wordformat.rules.node import FormatNode
from wordformat.structure.registry import Registry
from wordformat.structure.settings import CATEGORY_TO_CLASS, LEVEL_MAP
from wordformat.tree import Stack

_NON_HEADING = {
    "body_text",
    "math_block",
    "figure_image",
    "table_object",
    "toc_entry",
}

_HEADING_CATEGORIES = {
    k: v for k, v in CATEGORY_TO_CLASS.items() if k not in _NON_HEADING
}


class DocumentTreeBuilder:
    """负责将扁平列表构建成层级树结构"""

    HEADING_CATEGORIES = _HEADING_CATEGORIES
    CONFIG = {}

    def __init__(self, registry: Optional[Registry] = None):
        self.stack = Stack()
        self._config = {}
        # 领域注册表实例；None = 默认领域（论文），与历史行为一致
        self._registry = registry

    def _heading_categories(self) -> dict:
        """标题类别表：领域注册表优先（多领域标签集不同时树构建仍正确），
        默认实例回退类级静态表（导入时快照，兼容既有引用）。"""
        if self._registry is not None:
            return {k: v for k, v in self._registry.items() if k not in _NON_HEADING}
        return self.HEADING_CATEGORIES

    def build_tree(self, items: list[dict]) -> FormatNode:
        """
        从扁平段落列表构建文档语义树。
        返回根节点。
        """
        root = self._create_root_node()
        self.stack.push(root)

        for item in items:
            node = self._create_node_from_item(item)
            if node is None:
                continue

            if self._is_heading_category(item["category"]):
                self._attach_heading_node(node)
            else:
                self._attach_body_node(node)

        return root

    def _create_root_node(self) -> FormatNode:
        return FormatNode(
            value={"category": "top", "paragraph": "[ROOT]"}, expected_rule={}, level=0
        )

    def _create_node_from_item(self, item: dict) -> FormatNode | None:
        from wordformat.structure.node_factory import create_node

        level = self._determine_level(item["category"])
        return create_node(item=item, level=level, config=self._config)

    def _determine_level(self, category: str) -> int:
        """根据 category 映射到逻辑层级"""
        if self._registry is not None:
            return self._registry.level_of(category)
        return LEVEL_MAP.get(category, 999)

    # 标题类类别：目录/附录已注册进 CATEGORY_TO_CLASS，自动落在 HEADING_CATEGORIES 中，
    # 其子树隔离依赖 LEVEL_MAP 给定的层级（heading_mulu/heading_fulu = 1）
    def _is_heading_category(self, category: str) -> bool:
        return category in self._heading_categories()

    def _attach_heading_node(self, node: FormatNode):
        """处理标题类节点：维护栈层级"""
        while not self.stack.is_empty():
            top = self.stack.peek()
            if hasattr(top, "level") and top.level >= node.level:
                self.stack.pop()
            else:
                break

        parent = self.stack.peek()
        parent.add_child_node(node)
        self.stack.push(node)

    def _attach_body_node(self, node: FormatNode):
        """处理正文类节点：挂到最近的标题下"""
        parent = (
            self.stack.peek() if not self.stack.is_empty() else self._create_root_node()
        )
        parent.add_child_node(node)
