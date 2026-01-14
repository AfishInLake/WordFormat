#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 19:39
# @Author  : afish
# @File    : tree_builder.py
# src/tree_builder.py
from typing import List, Optional

from .settings import CATEGORY_TO_CLASS
from src.rules.node import FormatNode
from src.tree import Stack


class DocumentTreeBuilder:
    """负责将扁平列表构建成层级树结构"""

    HEADING_CATEGORIES = CATEGORY_TO_CLASS

    def __init__(self):
        self.stack = Stack()

    def build_tree(self, items: List[dict]) -> FormatNode:
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

            if self._is_heading_category(item['category']):
                self._attach_heading_node(node)
            else:
                self._attach_body_node(node)

        return root

    def _create_root_node(self) -> FormatNode:
        return FormatNode(
            value={'category': 'top', 'paragraph': '[ROOT]'},
            expected_rule={},
            level=0
        )

    def _create_node_from_item(self, item: dict) -> Optional[FormatNode]:
        from src.word_structure.node_factory import create_node
        level = self._determine_level(item['category'])
        return create_node(item, level)

    def _determine_level(self, category: str) -> int:
        """根据 category 映射到逻辑层级"""
        level_map = {
            'heading_level_1': 1,
            'heading_level_2': 2,
            'heading_level_3': 3,
            'heading_fulu': 4,
            'references_title': 1,
            'acknowledgements_title': 1,
            'abstract_chinese_title': 1,
            'abstract_english_title': 1,
            'keywords_chinese': 3,
            'keywords_english': 3,
        }
        return level_map.get(category, 999)

    def _is_heading_category(self, category: str) -> bool:
        return category in self.HEADING_CATEGORIES

    def _attach_heading_node(self, node: FormatNode):
        """处理标题类节点：维护栈层级"""
        while not self.stack.is_empty():
            top = self.stack.peek()
            if hasattr(top, 'level') and top.level >= node.level:
                self.stack.pop()
            else:
                break

        parent = self.stack.peek()
        parent.add_child_node(node)
        self.stack.push(node)

    def _attach_body_node(self, node: FormatNode):
        """处理正文类节点：挂到最近的标题下"""
        parent = self.stack.peek() if not self.stack.is_empty() else self._create_root_node()
        parent.add_child_node(node)
