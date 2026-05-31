#! /usr/bin/env python
# @Time    : 2026/1/8 15:00
# @Author  : afish
# @File    : tree.py
"""打印和展示用树工具。Tree/Stack/TreeNode 从 core 导入。"""

from wordformat.core.tree import Stack, Tree, TreeNode  # noqa: F401 — re-export


def print_tree(
    node_or_jsonpath: TreeNode | str,
    prefix: str = "",
    is_last: bool = True,
    show_confidence: bool = False,
    show_index: bool = False,
    filter_categories: list[str] | None = None,
    _counter: list[int] | None = None,
) -> None:
    """以树形结构打印多叉树（类似 Linux 的 tree 命令）

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
    if isinstance(node_or_jsonpath, str):
        import json

        from wordformat.core.tree import TreeNode as _TreeNode

        with open(node_or_jsonpath, encoding="utf-8") as f:
            paragraphs = json.load(f)

        root_node = _TreeNode({"category": "top", "paragraph": ""})
        _build_simple_tree(root_node, paragraphs)

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
            root_node,
            "",
            True,
            show_confidence,
            show_index,
            filter_categories,
        )
        return

    _print_tree_node(
        node_or_jsonpath,
        prefix,
        is_last,
        show_confidence,
        show_index,
        filter_categories,
    )


def _build_simple_tree(root: TreeNode, items: list[dict]) -> None:
    """从扁平段落列表构建简单树（委托给 core.tree_builder.build_tree）。"""
    from wordformat.core.tree_builder import build_tree

    built = build_tree(items, root_value=root.value)
    root.children = built.children


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

    if isinstance(value, dict):
        category = value.get("category", "")
    elif hasattr(value, "paragraph") and isinstance(value.paragraph, dict):
        category = value.paragraph.get("category", "")

    should_show = True
    if filter_categories is not None:
        should_show = category in filter_categories
        if not should_show:
            if hasattr(node, "children"):
                for child in node.children:
                    _print_tree_node(
                        child,
                        prefix,
                        is_last,
                        show_confidence,
                        show_index,
                        filter_categories,
                        _counter,
                    )
            return

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

    if hasattr(node, "children"):
        children = node.children
        for i, child in enumerate(children):
            is_last_child = i == len(children) - 1
            extension = "    " if is_last else "│   "
            new_prefix = prefix + extension
            _print_tree_node(
                child,
                new_prefix,
                is_last_child,
                show_confidence,
                show_index,
                filter_categories,
                _counter,
            )
