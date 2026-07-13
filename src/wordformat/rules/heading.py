#! /usr/bin/env python
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : heading.py

from wordformat.config.dotdict import BASE_FORMAT, deep_merge
from wordformat.rules.node import FormatNode
from wordformat.structure.registry import register


@register("heading_level_1", level=1)
class HeadingLevel1Node(FormatNode):
    """一级标题节点"""

    NODE_TYPE = "headings.level_1"
    NODE_LABEL = "一级标题"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            "paragraph": {
                "alignment": "居中对齐",
                "first_line_indent": "0字符",
                "builtin_style_name": "Heading 1",
            },
            "font": {"chinese_font_name": "黑体", "font_size": "小二"},
        },
    )


@register("heading_level_2", level=2)
class HeadingLevel2Node(FormatNode):
    """二级标题节点"""

    NODE_TYPE = "headings.level_2"
    NODE_LABEL = "二级标题"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            "paragraph": {
                "first_line_indent": "0字符",
                "builtin_style_name": "Heading 2",
            },
            "font": {"chinese_font_name": "黑体", "font_size": "三号"},
        },
    )


@register("heading_level_3", level=3)
class HeadingLevel3Node(FormatNode):
    """三级标题节点"""

    NODE_TYPE = "headings.level_3"
    NODE_LABEL = "三级标题"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            "paragraph": {
                "first_line_indent": "0字符",
                "builtin_style_name": "Heading 3",
            },
            "font": {"chinese_font_name": "黑体", "font_size": "小四"},
        },
    )
