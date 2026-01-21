#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 20:19
# @Author  : afish
# @File    : utils.py

from typing import Callable

from src.rules.body import BodyText
from src.rules.node import FormatNode


def find_and_modify_first(root: FormatNode, condition: Callable[[FormatNode], bool]):
    """
    找到第一个满足 condition 的节点，调用 modifier(node) 修改它，并返回该节点。
    :param root: 树的根节点（FormatNode 实例）
    :param condition: 函数，接收 node，返回 bool
    :return: 被修改的节点（FormatNode） if found, else None
    """
    from collections import deque
    queue = deque([root])
    while queue:
        node = queue.popleft()
        if condition(node):
            return node
        queue.extend(node.children)
    return None


def promote_bodytext_in_subtrees_of_type(
        root: 'FormatNode',
        parent_type: type,
        target_type: type
):
    """
    遍历整棵树：
      - 找到所有类型为 parent_type 的节点；
      - 对每个这样的节点，递归遍历其所有子孙；
      - 将其中所有 BodyText 类型的节点，替换为 target_type。

    :param root: 树的根节点
    :param parent_type: 父节点类型（如 AbstractTitleCN, ReferencesNode）
    :param target_type: 目标替换类型（如 AbstractContentCN, ReferenceEntry）
    """

    def upgrade_subtree(node):
        """递归升级 node 的整个子树中的 BodyText 节点"""
        for i, child in enumerate(node.children):
            if isinstance(child, BodyText):
                node.children[i].__class__ = target_type
                # 可选：打印日志
                # text = child.value.get('paragraph', '')[:40]
                # print(f"  ↑ 升级为 {target_type.__name__}: {text}...")
            # 递归处理子树（即使刚被升级，也要继续查 children）
            upgrade_subtree(child)

    def traverse_all(node):
        """全局遍历，寻找 parent_type 节点"""
        if isinstance(node, parent_type):
            upgrade_subtree(node)  # 升级该节点的整个子树
        for child in node.children:
            traverse_all(child)

    traverse_all(root)
