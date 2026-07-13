#! /usr/bin/env python
"""WordFormat — 论文格式自动化处理工具。"""

from wordformat._version import __version__
from wordformat.classify.tag import set_tag_main
from wordformat.pipeline.orchestrate import auto_format_thesis_document, md_to_docx

__all__ = [
    "__version__",
    "auto_format_thesis_document",
    "md_to_docx",
    "set_tag_main",
]
