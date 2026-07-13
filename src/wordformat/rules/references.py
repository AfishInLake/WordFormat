#! /usr/bin/env python
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : references.py

from wordformat.config.dotdict import BASE_FORMAT, deep_merge
from wordformat.rules.node import FormatNode
from wordformat.structure.registry import register


@register("references_title", level=1)
class References(FormatNode):
    """参考文献节点"""

    NODE_TYPE = "references.title"
    NODE_LABEL = "参考文献标题"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            "paragraph": {
                "alignment": "居中对齐",
                "first_line_indent": "0字符",
            },
            "font": {"chinese_font_name": "黑体", "font_size": "三号", "bold": True},
        },
    )


@register("references_content")
class ReferenceEntry(FormatNode):
    """参考文献条目节点"""

    NODE_TYPE = "references.entry"
    NODE_LABEL = "参考文献条目"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            "paragraph": {"first_line_indent": "0字符"},
            "font": {"chinese_font_name": "宋体", "font_size": "五号"},
        },
    )
