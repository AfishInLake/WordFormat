#! /usr/bin/env python
# @Time    : 2025/12/22 21:47
# @Author  : afish
# @File    : __init__.py.py
from wordformat._version import __version__

__all__ = [
    "__version__",
    "auto_format_thesis_document",
    "set_tag_main",
]


def __getattr__(name: str):
    """懒加载重导出，避免 import wordformat 时触发整个 ONNX / pipeline 依赖链。"""
    if name == "auto_format_thesis_document":
        from wordformat.pipeline.orchestrate import (
            auto_format_thesis_document as _fn,
        )

        return _fn
    if name == "set_tag_main":
        from wordformat.classify.tag import set_tag_main as _fn

        return _fn
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
