#! /usr/bin/env python
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : heading.py

from wordformat.rules.node import FormatNode
from wordformat.structure.registry import register

# 各级标题默认值
_H = {
    "alignment": "左对齐",
    "space_before": "0.5行",
    "space_after": "0.5行",
    "line_spacingrule": "单倍行距",
    "line_spacing": "1.5倍",
    "left_indent": "0字符",
    "right_indent": "0字符",
    "first_line_indent": "0字符",
    "chinese_font_name": "黑体",
    "english_font_name": "Times New Roman",
    "font_color": "黑色",
    "bold": False,
    "italic": False,
    "underline": False,
}


@register("heading_level_1", level=1)
class HeadingLevel1Node(FormatNode):
    """一级标题节点"""

    NODE_TYPE = "headings.level_1"
    NODE_LABEL = "一级标题"
    DEFAULTS = {
        **_H,
        "alignment": "居中对齐",
        "font_size": "小二",
        "builtin_style_name": "Heading 1",
    }


@register("heading_level_2", level=2)
class HeadingLevel2Node(FormatNode):
    """二级标题节点"""

    NODE_TYPE = "headings.level_2"
    NODE_LABEL = "二级标题"
    DEFAULTS = {**_H, "font_size": "三号", "builtin_style_name": "Heading 2"}


@register("heading_level_3", level=3)
class HeadingLevel3Node(FormatNode):
    """三级标题节点"""

    NODE_TYPE = "headings.level_3"
    NODE_LABEL = "三级标题"
    DEFAULTS = {**_H, "font_size": "小四", "builtin_style_name": "Heading 3"}
