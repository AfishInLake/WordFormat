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
# 无需格式化的类别统一映射到 BodyText（再配合 settings.VOIDNODELIST 跳过格式化），
# 避免 create_node 因未知类别返回 None 丢弃节点，导致段落与树节点错位。
CATEGORY_TO_CLASS.setdefault("other", BodyText)  # 封面/声明等无需格式化的内容
CATEGORY_TO_CLASS.setdefault("document_title", BodyText)  # 文档标题（后处理扩展）
CATEGORY_TO_CLASS.setdefault("footer", BodyText)  # 页脚/AI 生成声明（后处理扩展）
# 目录/附录标题：作为 terminal 标题节点挂载以隔离子树，但不参与格式化。
# 必须在此注册，否则 create_node 会因未知类别返回 None 丢弃节点，导致段落与树节点错位
# （heading_fulu 是模型真实标签，附录段落必然触发；heading_mulu 为结构性 terminal 类别）。
CATEGORY_TO_CLASS.setdefault("heading_mulu", BodyText)  # 目录标题
CATEGORY_TO_CLASS.setdefault("heading_fulu", BodyText)  # 附录标题

LEVEL_MAP = _level_registry
# 无对应 FormatNode 的特殊 terminal 类别
LEVEL_MAP.setdefault("heading_mulu", 1)
LEVEL_MAP.setdefault("heading_fulu", 1)
