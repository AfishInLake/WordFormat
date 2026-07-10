#! /usr/bin/env python
# @Time    : 2026/1/11 22:25
# @Author  : afish
# @File    : settings.py

from wordformat.rules import (  # noqa: F401 — 触发 @register 装饰器注册
    AbstractContentCN,
    AbstractContentEN,
    AbstractTitleCN,
    AbstractTitleContentCN,
    AbstractTitleContentEN,
    AbstractTitleEN,
    Acknowledgements,
    BodyText,
    CaptionFigure,
    CaptionTable,
    FigureImage,
    HeadingLevel1Node,
    HeadingLevel2Node,
    HeadingLevel3Node,
    KeywordsCN,
    KeywordsEN,
    References,
    TableObject,
)
from wordformat.structure.registry import _level_registry, _registry

CATEGORY_TO_CLASS = _registry
CATEGORY_TO_CLASS.setdefault("other", BodyText)  # 封面/声明等无需格式化的内容

LEVEL_MAP = _level_registry
# 无对应 FormatNode 的特殊 terminal 类别
LEVEL_MAP.setdefault("heading_mulu", 1)
LEVEL_MAP.setdefault("heading_fulu", 1)
