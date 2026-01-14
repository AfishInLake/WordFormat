#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/10 14:07
# @Author  : afish
# @File    : node.py
from typing import Any, List, Dict, Sequence

from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run


class TreeNode:
    """树的节点类"""

    def __init__(self, value: Any):
        self.value = value
        self.children: List['TreeNode'] = []

    def add_child(self, child_value: Any) -> 'TreeNode':
        """添加一个子节点，并返回该子节点（便于链式调用）"""
        child_node = TreeNode(child_value)
        self.children.append(child_node)
        return child_node

    def add_child_node(self, child_node: 'TreeNode') -> None:
        """直接添加一个 TreeNode 作为子节点"""
        self.children.append(child_node)

    def __repr__(self) -> str:
        return f"TreeNode({self.value})"


class FormatNode(TreeNode):
    """所有格式检查节点的基类"""

    def __init__(self, value, level: int | float, paragraph: Paragraph = None, expected_rule: Dict[str, Any] = None):
        super().__init__(value=value)  # value 就是 paragraph
        self.level: int | float = level
        self.paragraph: Paragraph = paragraph
        self.expected_rule = expected_rule

    def update_paragraph(self, paragraph: Paragraph | Dict):
        self.paragraph = paragraph

    def check_format(self, doc: Document) -> List[Dict[str, Any]]:
        """虚方法：由子类实现具体的格式检查逻辑"""
        raise NotImplementedError("Subclasses should implement this!")

    def add_comment(self, doc: Document, runs: Run | Sequence[Run], text: str):
        doc.add_comment(runs=runs, text=text, author="论文解析器", initials="afish")
