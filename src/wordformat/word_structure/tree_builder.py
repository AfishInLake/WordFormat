#! /usr/bin/env python
# @Time    : 2026/1/11 19:39
# @Author  : afish
# @File    : tree_builder.py
# wordformat/tree_builder.py

from wordformat.rules.node import FormatNode
from wordformat.tree import Stack
from wordformat.word_structure.settings import CATEGORY_TO_CLASS, LEVEL_MAP


class DocumentTreeBuilder:
    """负责将扁平列表构建成层级树结构"""

    HEADING_CATEGORIES = CATEGORY_TO_CLASS
    CONFIG = {}

    def __init__(self):
        self.stack = Stack()

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
        from wordformat.word_structure.node_factory import create_node

        level = self._determine_level(item["category"])
        return create_node(item=item, level=level, config=self.CONFIG)

    def _determine_level(self, category: str) -> int:
        """根据 category 映射到逻辑层级"""
        return LEVEL_MAP.get(category, 999)

    def _is_heading_category(self, category: str) -> bool:
        return category in self.HEADING_CATEGORIES

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
