#! /usr/bin/env python
"""core 包 —— 纯数据层，零外部依赖。

提供：
  - TreeNode / Tree / Stack —— 树数据结构
  - CategoryRegistry —— 可配置的段落类别注册表
  - build_tree —— 统一的树构建器
"""

from wordformat.core.category import (
    CATEGORY_NAMES,
    HEADING_CATEGORIES,
    LEVEL_MAP,
    CategoryRegistry,
    get_registry,
    reset_registry,
)
from wordformat.core.tree import Stack, Tree, TreeNode
from wordformat.core.tree_builder import build_tree

__all__ = [
    "TreeNode",
    "Tree",
    "Stack",
    "build_tree",
    "CategoryRegistry",
    "get_registry",
    "reset_registry",
    "CATEGORY_NAMES",
    "HEADING_CATEGORIES",
    "LEVEL_MAP",
]
