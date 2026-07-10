#! /usr/bin/env python
# @Time    : 2026/1/8 15:00
# @Author  : afish
# @File    : tree.py
from collections.abc import Callable, Iterator
from typing import Any, Generic, Optional, TypeVar

from rich.tree import Tree as RichTree

from wordformat.rules.node import TreeNode


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


def print_tree(
    node_or_jsonpath: TreeNode | str,
    show_confidence: bool = False,
    show_index: bool = False,
    filter_categories: list[str] | None = None,
) -> None:
    """以树形结构打印多叉树（使用 rich 渲染）。

    支持直接传入 TreeNode 或 JSON 文件路径。
    """
    from rich.console import Console
    from rich.table import Table

    console = Console()

    if isinstance(node_or_jsonpath, str):
        node_or_jsonpath, paragraphs = _load_tree_from_json(node_or_jsonpath)

        # 统计各类别
        from collections import Counter

        cat_counter = Counter()
        for p in paragraphs:
            cat_counter[p.get("category", "unknown")] += 1

        console.print(
            f"\n:page_facing_up: 文档结构树 ([bold]{len(paragraphs)}[/] 个段落)"
        )
        table = Table(show_header=False, box=None, padding=(0, 4))
        for cat, count in cat_counter.most_common():
            table.add_row(f"  {cat}", f"{count:>4d}")
        console.print(table)

    rich_tree = _build_rich_tree(
        node_or_jsonpath, show_confidence, show_index, filter_categories
    )
    console.print(rich_tree)


def _load_tree_from_json(json_path: str) -> tuple[TreeNode, list[dict]]:
    """从 JSON 文件加载段落列表并构建简单树。"""
    import json

    with open(json_path, encoding="utf-8") as f:
        paragraphs = json.load(f)
    root_node = TreeNode({"category": "top", "paragraph": ""})
    _build_simple_tree(root_node, paragraphs)
    return root_node, paragraphs


def _build_simple_tree(root: TreeNode, items: list[dict]) -> None:
    """从扁平段落列表构建简单树（不依赖 config，仅用于展示）"""
    from wordformat.tree import Stack

    HEADING_CATEGORIES = {
        "heading_level_1",
        "heading_level_2",
        "heading_level_3",
        "heading_mulu",
        "heading_fulu",
        "references_title",
        "acknowledgements_title",
        "abstract_chinese_title",
        "abstract_english_title",
        "abstract_chinese_title_content",
        "abstract_english_title_content",
    }
    LEVEL_MAP = {
        "heading_level_1": 1,
        "heading_level_2": 2,
        "heading_level_3": 3,
        "heading_mulu": 1,
        "heading_fulu": 1,
        "references_title": 1,
        "acknowledgements_title": 1,
        "abstract_chinese_title": 1,
        "abstract_english_title": 1,
        "abstract_chinese_title_content": 1,
        "abstract_english_title_content": 1,
        "keywords_chinese": 3,
        "keywords_english": 3,
    }

    stack = Stack()
    stack.push(root)

    for item in items:
        category = item.get("category", "body_text")
        node = TreeNode(item)
        level = LEVEL_MAP.get(category, 999)

        if category in HEADING_CATEGORIES:
            # 标题节点：维护栈层级
            while not stack.is_empty():
                top = stack.peek()
                top_level = LEVEL_MAP.get(
                    top.value.get("category", "")
                    if isinstance(top.value, dict)
                    else "",
                    999,
                )
                if top_level >= level:
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


def _tree_node_category(value) -> str:
    """从节点 value 中提取类别字符串。"""
    if isinstance(value, dict):
        return value.get("category", "")
    if hasattr(value, "paragraph") and isinstance(value.paragraph, dict):
        return value.paragraph.get("category", "")
    return ""


def _format_node_label(
    value, show_confidence: bool, show_index: bool, counter: int
) -> str:
    """格式化单个节点的显示标签。"""
    idx_str = f"[{counter:>3d}] " if show_index else ""
    if isinstance(value, dict):
        cat = value.get("category", "unknown")
        para = str(value.get("paragraph", ""))[:50]
        conf = value.get("confidence", None)
        conf_str = f" ({conf:.0%})" if show_confidence and conf is not None else ""
        return f"{idx_str}[dim]{cat}[/dim]{conf_str} {para}"
    if hasattr(value, "paragraph") and isinstance(value.paragraph, dict):
        cat = value.paragraph.get("category", "unknown")
        para = str(value.paragraph.get("paragraph", ""))[:50]
        return f"{idx_str}[dim]{cat}[/dim] {para}"
    return f"{idx_str}{str(value)[:60]}"


def _build_rich_tree(
    node: TreeNode,
    show_confidence: bool,
    show_index: bool,
    filter_categories: list[str] | None,
    _counter: list[int] | None = None,
) -> RichTree | None:
    """递归构建 rich Tree。"""

    if _counter is None:
        _counter = [0]

    value = node.value
    category = _tree_node_category(value)

    if filter_categories is not None and category not in filter_categories:
        # 跳过当前节点但继续递归子节点
        for child in getattr(node, "children", []):
            subtree = _build_rich_tree(
                child, show_confidence, show_index, filter_categories, _counter
            )
            if subtree is not None:
                return subtree
        return None

    label = _format_node_label(value, show_confidence, show_index, _counter[0])
    _counter[0] += 1
    rich_node = RichTree(label)

    if hasattr(node, "children"):
        for child in node.children:
            child_tree = _build_rich_tree(
                child, show_confidence, show_index, filter_categories, _counter
            )
            if child_tree is not None:
                rich_node.add(child_tree)

    return rich_node


# ---------------------------------------------------------------------------
# 独立于 Tree 类的通用遍历工具（直接作用于 TreeNode / FormatNode）
# ---------------------------------------------------------------------------


def bfs_walk(root: TreeNode):
    """BFS 层序遍历生成器，按文档顺序 yield 每个节点（含根节点）。

    用法：
        for node in bfs_walk(root_node):
            if isinstance(node, SomeType):
                ...
    """
    from collections import deque

    queue = deque([root])
    while queue:
        node = queue.popleft()
        yield node
        queue.extend(node.children)


def dfs_walk(root: TreeNode):
    """DFS 前序遍历生成器，yield 每个子节点（不含根节点）。

    用法：
        for node in dfs_walk(root_node):
            ...
    """
    for child in root.children:
        yield child
        yield from dfs_walk(child)
