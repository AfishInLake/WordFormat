#! /usr/bin/env python
"""WordFormat — 论文格式自动化处理工具。"""

from wordformat._version import __version__

# 显式装配：导入节点模块触发 @register 并补齐占位类别。
# 必须早于任何读取 CATEGORY_TO_CLASS / LEVEL_MAP 的模块导入。
from wordformat.assembly import register_nodes as _register_nodes

_register_nodes()

from wordformat.classify.tag import set_tag_main  # noqa: E402
from wordformat.pipeline.orchestrate import (  # noqa: E402
    auto_format_thesis_document,
    md_to_docx,
)

__all__ = [
    "__version__",
    "auto_format_thesis_document",
    "md_to_docx",
    "set_tag_main",
]
