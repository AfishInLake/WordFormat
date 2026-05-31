#! /usr/bin/env python
"""orchestration 包 —— 编排层，引入 docx 依赖。

提供：
  - bind_and_sync:    虚拟节点树 ↔ docx 段落绑定同步
  - style_fixer:      样式定义修正
  - table_formatter:  表格内容格式化
"""

from wordformat.orchestration.binding import bind_and_sync
from wordformat.orchestration.style_fixer import fix_all_style_definitions
from wordformat.orchestration.table_formatter import format_table_content

__all__ = [
    "bind_and_sync",
    "fix_all_style_definitions",
    "format_table_content",
]
