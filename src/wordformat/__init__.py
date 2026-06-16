#! /usr/bin/env python
"""WordFormat —— 论文格式自动化处理工具。

公共 API：
  - 核心数据层：TreeNode, Tree, Stack, CategoryRegistry, build_tree
  - 编排层：    bind_and_sync, fix_all_style_definitions, format_table_content
  - 规则层：    FormatNode
  - 配置层：    NodeConfigRoot, init_config, get_config
  - 入口：      auto_format_thesis_document, set_tag_main
"""

from wordformat._version import __version__

# ---- 配置层 ----
from wordformat.config.config import get_config, init_config
from wordformat.config.datamodel import NodeConfigRoot

# ---- 核心数据层 ----
from wordformat.core import (
    CATEGORY_NAMES,
    HEADING_CATEGORIES,
    LEVEL_MAP,
    CategoryRegistry,
    Stack,
    Tree,
    TreeNode,
    build_tree,
    get_registry,
    reset_registry,
)

# ---- 公式渲染 ----
from wordformat.math import add_display_math, add_inline_math, latex_to_omath

# ---- 媒体层 ----
from wordformat.media import ImageRegistry, insert_element, merge_docx

# ---- 编排层 ----
from wordformat.orchestration import (
    bind_and_sync,
    fix_all_style_definitions,
    format_table_content,
)

# ---- 规则层 ----
from wordformat.rules.node import FormatNode

# ---- 入口 ----
from wordformat.set_style import auto_format_thesis_document
from wordformat.set_tag import set_tag_main

__all__ = [
    # 版本
    "__version__",
    # core — 树数据结构
    "TreeNode",
    "Tree",
    "Stack",
    "build_tree",
    # core — 类别注册
    "CategoryRegistry",
    "get_registry",
    "reset_registry",
    "CATEGORY_NAMES",
    "HEADING_CATEGORIES",
    "LEVEL_MAP",
    # rules
    "FormatNode",
    # orchestration
    "bind_and_sync",
    "fix_all_style_definitions",
    "format_table_content",
    # config
    "NodeConfigRoot",
    "init_config",
    "get_config",
    # media
    "ImageRegistry",
    "insert_element",
    "merge_docx",
    # math
    "latex_to_omath",
    "add_display_math",
    "add_inline_math",
    # 入口
    "auto_format_thesis_document",
    "set_tag_main",
]
