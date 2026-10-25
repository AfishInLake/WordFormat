#! /usr/bin/env python
# @Time    : 2026/1/11 20:19
# @Author  : afish
# @File    : utils.py

from collections.abc import Callable

from wordformat.rules.body import BodyText
from wordformat.rules.node import FormatNode
from wordformat.tree import bfs_walk


def find_and_modify_first(root: FormatNode, condition: Callable[[FormatNode], bool]):
    """BFS 查找第一个满足 condition 的节点并返回。"""
    for node in bfs_walk(root):
        if condition(node):
            return node
    return None


def promote_bodytext_in_subtrees_of_type(
    root: "FormatNode", parent_type: type, target_type: type
):
    """
    遍历整棵树：
      - 找到所有类型为 parent_type 的节点；
      - 对每个这样的节点，递归遍历其所有子孙；
      - 将其中所有 BodyText 类型的节点，替换为 target_type。

    替换实例类的同时同步 value 中的 category 身份字段（从注册表反查注册名），
    保证基于 category 的判断（如 numbering 的 references_content 识别）能覆盖提升品。

    :param root: 树的根节点
    :param parent_type: 父节点类型（如 AbstractTitleCN, ReferencesNode）
    :param target_type: 目标替换类型（如 AbstractContentCN, ReferenceEntry）
    """
    from wordformat.structure.settings import CATEGORY_TO_CLASS

    target_category = next(
        (cat for cat, cls in CATEGORY_TO_CLASS.items() if cls is target_type),
        None,
    )

    def upgrade_subtree(node):
        """递归升级 node 的整个子树中的 BodyText 节点"""
        for i, child in enumerate(node.children):
            if isinstance(child, BodyText):
                node.children[i].__class__ = target_type
                # 结构契约同步：实例类变更后 category 身份字段保持一致
                if isinstance(child.value, dict) and target_category:
                    child.value["category"] = target_category
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
