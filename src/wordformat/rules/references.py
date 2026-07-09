#! /usr/bin/env python
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : references.py

from wordformat.config.models import ReferencesContentConfig, ReferencesTitleConfig
from wordformat.rules.node import FormatNode


class References(FormatNode[ReferencesTitleConfig]):
    """参考文献节点"""

    NODE_TYPE = "references"
    NODE_LABEL = "参考文献标题"
    CONFIG_MODEL = ReferencesTitleConfig
    CONFIG_PATH = "references.title"


class ReferenceEntry(FormatNode[ReferencesContentConfig]):
    """参考文献条目节点"""

    NODE_LABEL = "参考文献条目"
    CONFIG_MODEL = ReferencesContentConfig
    CONFIG_PATH = "references.content"
