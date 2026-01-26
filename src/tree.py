#! /usr/bin/env python
# @Time    : 2026/1/8 15:00
# @Author  : afish
# @File    : tree.py
from collections.abc import Callable, Iterator
from typing import Any, Generic, Optional, TypeVar

from src.rules.node import TreeNode


class Tree:
    """多叉树类"""

    def __init__(self, root_value: Any):
        self.root = TreeNode(root_value)

    def is_empty(self) -> bool:
        return self.root is None

    # ===== 遍历方法 =====

    def preorder(self) -> Iterator[Any]:
        """前序遍历（根 → 子树从左到右）"""

        def _preorder(node: TreeNode):
            if node:
                yield node.value
                for child in node.children:
                    yield from _preorder(child)

        yield from _preorder(self.root)

    def postorder(self) -> Iterator[Any]:
        """后序遍历（子树从左到右 → 根）"""

        def _postorder(node: TreeNode):
            if node:
                for child in node.children:
                    yield from _postorder(child)
                yield node.value

        yield from _postorder(self.root)

    def level_order(self) -> Iterator[Any]:
        """层序遍历（广度优先）"""
        if not self.root:
            return
        queue = [self.root]
        while queue:
            node = queue.pop(0)
            yield node.value
            queue.extend(node.children)

    # ===== 查找与工具方法 =====
    def find_by_condition(
        self, condition: Callable[[Any], bool]
    ) -> Optional["TreeNode"]:
        """
        根据条件函数查找节点。
        :param condition: 接收 node.value，返回 bool
        :return: 匹配的第一个节点（DFS顺序），或 None
        """

        def _dfs(node: "TreeNode"):
            if node and condition(node.value):
                return node
            for child in node.children:
                result = _dfs(child)
                if result:
                    return result
            return None

        return _dfs(self.root)

    def height(self) -> int:
        """返回树的高度（根到最深叶子的边数）"""

        def _height(node: TreeNode) -> int:
            if not node.children:
                return 0
            return 1 + max(_height(child) for child in node.children)

        return _height(self.root)

    def size(self) -> int:
        """返回树中节点总数"""
        return sum(1 for _ in self.preorder())

    def __repr__(self) -> str:
        return f"Tree(root={self.root.value})"


T = TypeVar("T")


class Stack(Generic[T]):
    """一个简单的后进先出（LIFO）栈"""

    def __init__(self) -> None:
        self._items: list[T] = []

    def push(self, item: T) -> None:
        """将元素压入栈顶"""
        self._items.append(item)

    def pop(self) -> T:
        """弹出并返回栈顶元素，若栈为空则抛出 IndexError"""
        if self.is_empty():
            raise IndexError("pop from empty stack")
        return self._items.pop()

    def peek(self) -> T:
        """返回栈顶元素但不弹出，若栈为空则抛出 IndexError"""
        if self.is_empty():
            raise IndexError("peek from empty stack")
        return self._items[-1]

    def peek_safe(self) -> T | None:
        """安全地查看栈顶元素，若栈为空返回 None"""
        return self._items[-1] if self._items else None

    def is_empty(self) -> bool:
        """判断栈是否为空"""
        return len(self._items) == 0

    def size(self) -> int:
        """返回栈中元素个数"""
        return len(self._items)

    def clear(self) -> None:
        """清空栈"""
        self._items.clear()

    def __repr__(self) -> str:
        return f"Stack({self._items})"

    def __bool__(self) -> bool:
        """支持 if stack: 判断非空"""
        return not self.is_empty()


def print_tree(node: TreeNode, prefix: str = "", is_last: bool = True) -> None:
    """
    以树形结构打印多叉树（类似 Linux 的 tree 命令）

    示例输出：
    ├── 第一章
    │   ├── 1.1 背景
    │   │   └── 正文...
    │   └── 1.2 意义
    """
    # 打印当前节点
    connector = "└── " if is_last else "├── "
    value = node.value

    # 尝试提取可读内容
    if isinstance(value, dict):
        cat = value.get("category", "unknown")
        para = value.get("paragraph", "")[:50]  # 截断长文本
        display = f"【{cat}】 {para}"
    elif hasattr(value, "paragraph") and isinstance(value.paragraph, dict):
        cat = value.paragraph.get("category", "unknown")
        para = value.paragraph.get("paragraph", "")[:50]
        display = f"【{cat}】 {para}"
    else:
        display = str(value)[:60]

    print(prefix + connector + display)  # noqa t201

    # 递归打印子节点
    if hasattr(node, "children"):
        children = node.children
        for i, child in enumerate(children):
            is_last_child = i == len(children) - 1
            # 下一级前缀
            extension = "    " if is_last else "│   "
            new_prefix = prefix + extension
            print_tree(child, new_prefix, is_last_child)
