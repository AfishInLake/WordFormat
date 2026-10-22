#! /usr/bin/env python
# @Time    : 2026/1/9 20:08
# @Author  : afish
# @File    : __init__.py.py
"""框架通用格式规则节点。

仅含跨文体通用的结构节点与节点机制（FormatNode）。论文专属节点（摘要/
参考文献/致谢/关键词）已迁至 wordformat.domains.thesis.rules。
"""

from .body import BodyText
from .caption import CaptionFigure, CaptionTable
from .heading import HeadingLevel1Node, HeadingLevel2Node, HeadingLevel3Node
from .math import MathBlock
from .node import FormatNode
from .object import FigureImage, TableObject

__all__ = [
    "MathBlock",
    "BodyText",
    "CaptionFigure",
    "CaptionTable",
    "FigureImage",
    "TableObject",
    "HeadingLevel1Node",
    "HeadingLevel2Node",
    "HeadingLevel3Node",
    "FormatNode",
]
