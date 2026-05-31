#! /usr/bin/env python
"""统一的树构建器 —— 从扁平段落列表构建层级树（纯数据，零外部依赖）。

这是 tree.py 中 _build_simple_tree 和 word_structure/tree_builder.py 中
DocumentTreeBuilder 的统一实现。两者唯一的区别是节点类型：
  - _build_simple_tree 创建 core.TreeNode
  - DocumentTreeBuilder 通过 node_factory 创建 FormatNode 子类

统一后，建树逻辑集中在此，节点创建委托给可选的 factory 参数。
"""

from typing import Callable, Optional

from wordformat.core.category import CategoryRegistry, get_registry
from wordformat.core.tree import Stack, TreeNode

# 节点工厂类型：接收 (item: dict, level: int) → TreeNode
NodeFactory = Callable[[dict, int], TreeNode]


def build_tree(
    items: list[dict],
    *,
    registry: Optional[CategoryRegistry] = None,
    factory: Optional[NodeFactory] = None,
    root_value: Optional[dict] = None,
) -> TreeNode:
    """从扁平段落列表构建层级树。

    Args:
        items: 扁平段落列表，每项为 {"category": "...", ...}
        registry: 类别注册表（默认使用全局注册表）
        factory: 可选的节点工厂（默认创建 core.TreeNode）
        root_value: 根节点的 value（默认 {"category": "top"}）

    Returns:
        树的根节点
    """
    if registry is None:
        registry = get_registry()
    if root_value is None:
        root_value = {"category": "top", "paragraph": "[ROOT]"}
    if factory is None:

        def _default_factory(item, level):
            return TreeNode(item)

        factory = _default_factory

    stack: Stack[TreeNode] = Stack()
    root = TreeNode(root_value)
    stack.push(root)

    for item in items:
        category = item.get("category", "body_text")
        level = registry.get_level(category)
        node = factory(item, level)

        if registry.is_heading(category):
            # 标题节点：维护栈层级
            while not stack.is_empty():
                top = stack.peek()
                top_cat = (
                    top.value.get("category", "") if isinstance(top.value, dict) else ""
                )
                if registry.get_level(top_cat) >= level:
                    stack.pop()
                else:
                    break
            parent = stack.peek_safe() or root
            parent.add_child_node(node)
            stack.push(node)
        else:
            # 正文节点：挂到最近的标题下
            parent = stack.peek() if not stack.is_empty() else root
            parent.add_child_node(node)

    return root
