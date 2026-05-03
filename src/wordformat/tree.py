#! /usr/bin/env python
# @Time    : 2026/1/8 15:00
# @Author  : afish
# @File    : tree.py
from collections.abc import Callable, Iterator
from typing import Any, Generic, Optional, TypeVar

from wordformat.rules.node import TreeNode


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


def print_tree(
    node_or_jsonpath: TreeNode | str,
    prefix: str = "",
    is_last: bool = True,
    show_confidence: bool = False,
    show_index: bool = False,
    filter_categories: list[str] | None = None,
    _counter: list[int] | None = None,
) -> None:
    """
    以树形结构打印多叉树（类似 Linux 的 tree 命令）

    支持直接传入 TreeNode 或 JSON 文件路径。

    Args:
        node_or_jsonpath: 树节点或 JSON 文件路径
        prefix: 前缀（递归内部使用）
        is_last: 是否为同级最后一个节点
        show_confidence: 是否显示置信度
        show_index: 是否显示节点序号
        filter_categories: 仅显示指定类别（None 表示全部显示）
        _counter: 内部计数器（递归内部使用）
    """
    # 如果传入的是 JSON 文件路径，先构建树再打印
    if isinstance(node_or_jsonpath, str):
        import json
        from wordformat.rules.node import TreeNode

        with open(node_or_jsonpath, encoding="utf-8") as f:
            paragraphs = json.load(f)

        # 直接从 JSON 构建简单树（不依赖 config）
        root_node = TreeNode({"category": "top", "paragraph": ""})
        _build_simple_tree(root_node, paragraphs)

        # 统计各类别数量
        from collections import Counter
        cat_counter = Counter()
        for p in paragraphs:
            cat_counter[p.get("category", "unknown")] += 1

        print(f"\n📄 文档结构树 ({len(paragraphs)} 个段落)")
        print("=" * 60)
        for cat, count in cat_counter.most_common():
            print(f"  {cat:<30s} {count:>4d}")
        print("=" * 60)

        _print_tree_node(
            root_node, "", True,
            show_confidence, show_index,
            filter_categories,
        )
        return

    # 递归打印节点
    _print_tree_node(
        node_or_jsonpath, prefix, is_last,
        show_confidence, show_index,
        filter_categories,
    )


def _build_simple_tree(root: TreeNode, items: list[dict]) -> None:
    """从扁平段落列表构建简单树（不依赖 config，仅用于展示）"""
    from wordformat.tree import Stack

    HEADING_CATEGORIES = {
        "heading_level_1", "heading_level_2", "heading_level_3",
        "heading_mulu", "heading_fulu",
        "references_title", "acknowledgements_title",
        "abstract_chinese_title", "abstract_english_title",
        "abstract_chinese_title_content", "abstract_english_title_content",
    }
    LEVEL_MAP = {
        "heading_level_1": 1, "heading_level_2": 2, "heading_level_3": 3,
        "heading_mulu": 1, "heading_fulu": 1,
        "references_title": 1, "acknowledgements_title": 1,
        "abstract_chinese_title": 1, "abstract_english_title": 1,
        "abstract_chinese_title_content": 1, "abstract_english_title_content": 1,
        "keywords_chinese": 3, "keywords_english": 3,
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
                    top.value.get("category", "") if isinstance(top.value, dict) else "", 999
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


def _print_tree_node(
    node: TreeNode,
    prefix: str,
    is_last: bool,
    show_confidence: bool,
    show_index: bool,
    filter_categories: list[str] | None,
    _counter: list[int] | None = None,
) -> None:
    """递归打印单个节点及其子树"""
    if _counter is None:
        _counter = [0]

    value = node.value
    category = ""

    # 提取类别
    if isinstance(value, dict):
        category = value.get("category", "")
    elif hasattr(value, "paragraph") and isinstance(value.paragraph, dict):
        category = value.paragraph.get("category", "")

    # 过滤类别
    should_show = True
    if filter_categories is not None:
        should_show = category in filter_categories
        # 如果当前节点被过滤，但其子节点有匹配的，仍需遍历子节点
        if not should_show:
            if hasattr(node, "children"):
                for child in node.children:
                    _print_tree_node(
                        child, prefix, is_last,
                        show_confidence, show_index,
                        filter_categories, _counter,
                    )
            return

    # 打印当前节点
    connector = "└── " if is_last else "├── "

    if isinstance(value, dict):
        cat = value.get("category", "unknown")
        para = value.get("paragraph", "")[:50]
        idx_str = f"[{_counter[0]:>3d}] " if show_index else ""
        conf = value.get("confidence", None)
        conf_str = f" ({conf:.0%})" if show_confidence and conf is not None else ""
        display = f"{idx_str}【{cat}】{conf_str} {para}"
    elif hasattr(value, "paragraph") and isinstance(value.paragraph, dict):
        cat = value.paragraph.get("category", "unknown")
        para = value.paragraph.get("paragraph", "")[:50]
        idx_str = f"[{_counter[0]:>3d}] " if show_index else ""
        display = f"{idx_str}【{cat}】 {para}"
    else:
        idx_str = f"[{_counter[0]:>3d}] " if show_index else ""
        display = f"{idx_str}{str(value)[:60]}"

    print(prefix + connector + display)  # noqa t201
    _counter[0] += 1

    # 递归打印子节点
    if hasattr(node, "children"):
        children = node.children
        for i, child in enumerate(children):
            is_last_child = i == len(children) - 1
            extension = "    " if is_last else "│   "
            new_prefix = prefix + extension
            _print_tree_node(
                child, new_prefix, is_last_child,
                show_confidence, show_index,
                filter_categories, _counter,
            )
