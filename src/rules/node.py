#! /usr/bin/env python
# @Time    : 2026/1/10 14:07
# @Author  : afish
# @File    : node.py
from collections.abc import Sequence
from typing import Any

from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run


class TreeNode:
    """树的节点类"""

    NODE_TYPE = "node"

    def __init__(self, value: Any):
        self.value = value
        self.__config = {}
        self.children: list[TreeNode] = []
        self.__set_fingerprint()

    @property
    def config(self):
        return self.__config

    def load_config(self, full_config: dict | Any) -> None:
        """
        根据 cls.NODE_TYPE（点分路径）从 full_config 中提取子配置。

        例如：
          NODE_TYPE = 'abstract.keywords.chinese'
          则从 full_config['abstract']['keywords']['chinese'] 提取
        """
        if not isinstance(full_config, dict):
            self.__config = {}
            return

        # 解析路径：支持 'a.b.c' 或直接 'top_level_key'
        path_parts = self.NODE_TYPE.split(".")

        current = full_config
        try:
            for part in path_parts:
                if not isinstance(current, dict):
                    raise KeyError(
                        f"Expected dict at path {'.'.join(path_parts[: path_parts.index(part) + 1])}"
                    )
                current = current[part]
            # 如果最终值不是 dict，也允许（但通常应是 dict）
            self.__config = current if isinstance(current, dict) else {}
        except (KeyError, TypeError):
            # 路径不存在或结构不匹配，返回空配置
            self.__config = {}

    def __set_fingerprint(self):
        if self.value and isinstance(self.value, dict):
            if self.value.get("category") == "top":
                return
            try:
                self.fingerprint = self.value["fingerprint"]
            except KeyError as err:
                raise ValueError(f"{self.value} must have a 'fingerprint' key") from err

    def add_child(self, child_value: Any) -> "TreeNode":
        """添加一个子节点，并返回该子节点（便于链式调用）"""
        child_node = TreeNode(child_value)
        self.children.append(child_node)
        return child_node

    def add_child_node(self, child_node: "TreeNode") -> None:
        """直接添加一个 TreeNode 作为子节点"""
        self.children.append(child_node)

    def __repr__(self) -> str:
        return f"TreeNode({self.value})"


class FormatNode(TreeNode):
    """所有格式检查节点的基类"""

    def __init__(
        self,
        value,
        level: int | float,
        paragraph: Paragraph = None,
        expected_rule: dict[str, Any] = None,
    ):
        super().__init__(value=value)  # value 就是 paragraph
        self.level: int | float = level
        self.paragraph: Paragraph = paragraph
        self.expected_rule = expected_rule

    def update_paragraph(self, paragraph: Paragraph | dict):
        self.paragraph = paragraph

    def check_format(self, doc: Document) -> list[dict[str, Any]]:
        """虚方法：由子类实现具体的格式检查逻辑"""
        raise NotImplementedError("Subclasses should implement this!")

    def add_comment(self, doc: Document, runs: Run | Sequence[Run], text: str):
        if text.strip():
            doc.add_comment(runs=runs, text=text, author="论文解析器", initials="afish")
