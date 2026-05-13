#! /usr/bin/env python
# @Time    : 2025/12/22 21:47
# @Author  : afish
# @File    : __init__.py.py
from wordformat._version import __version__
from wordformat.set_style import auto_format_thesis_document
from wordformat.set_tag import set_tag_main

__all__ = [
    "__version__",
    "auto_format_thesis_document",
    "set_tag_main",
]
