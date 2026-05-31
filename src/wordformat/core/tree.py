#! /usr/bin/env python
"""纯数据层：TreeNode、Tree、Stack —— 零外部依赖（仅 stdlib）。

TreeNode 是树结构的基本节点，可独立于 docx / pydantic / yaml 使用。
Tree 是面向 TreeNode 的多叉树封装，提供遍历和查找。
Stack 是泛型 LIFO 栈。
"""

from collections.abc import Callable, Iterator
from typing import Any, Generic, Optional, TypeVar


class TreeNode:
    """树的节点类 —— 纯数据，零外部依赖。

    Attributes:
        value:    原始数据（分类结果 dict）
        children: 子节点列表
        type:     内容类型 — "paragraph" | "table" | "image" | "formula"
        content:  类型特有数据（表格单元格、图片路径等）
    """

    NODE_TYPE = "node"

    def __init__(self, value: Any):
        self.value = value
        self.children: list[TreeNode] = []
        self.type: str = "paragraph"
        self.content: dict = {}

    def add_child(self, child_value: Any) -> "TreeNode":
        """添加一个子节点（从 value 创建），并返回该子节点。"""
        child_node = TreeNode(child_value)
        self.children.append(child_node)
        return child_node

    def add_child_node(self, child_node: "TreeNode") -> None:
        """直接添加一个 TreeNode 作为子节点。"""
        self.children.append(child_node)

    def to_dict(self) -> dict:
        """导出为层级 dict（含 children 递归）。"""
        result: dict = {
            "value": self.value,
            "children": [c.to_dict() for c in self.children],
        }
        if self.type != "paragraph":
            result["type"] = self.type
        if self.content:
            result["content"] = self.content
        return result

    @classmethod
    def from_dict(cls, data: dict) -> "TreeNode":
        """从层级 dict 还原树。"""
        node = cls(data["value"])
        node.type = data.get("type", "paragraph")
        node.content = data.get("content", {})
        for child_data in data.get("children", []):
            node.add_child_node(cls.from_dict(child_data))
        return node

    def __repr__(self) -> str:
        return f"TreeNode({self.value})"


class Tree:
    """多叉树类"""

    def __init__(self, root_value: Any):
        self.root = TreeNode(root_value)

    def is_empty(self) -> bool:
        return len(self.root.children) == 0

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
        """根据条件函数查找节点（DFS 顺序），返回第一个匹配节点或 None。"""

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
