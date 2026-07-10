#! /usr/bin/env python
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : references.py

from wordformat.rules.node import FormatNode
from wordformat.structure.registry import register

_REF_BASE = {
    "alignment": "左对齐",
    "space_before": "0.5行",
    "space_after": "0.5行",
    "line_spacingrule": "单倍行距",
    "line_spacing": "1.5倍",
    "left_indent": "0字符",
    "right_indent": "0字符",
}


@register("references_title", level=1)
class References(FormatNode):
    """参考文献节点"""

    NODE_TYPE = "references.title"
    NODE_LABEL = "参考文献标题"
    DEFAULTS = {
        **_REF_BASE,
        "alignment": "居中对齐",
        "first_line_indent": "0字符",
        "chinese_font_name": "黑体",
        "font_size": "三号",
        "bold": True,
        "builtin_style_name": "正文",
        "english_font_name": "Times New Roman",
        "font_color": "黑色",
        "italic": False,
        "underline": False,
    }


@register("references_content")
class ReferenceEntry(FormatNode):
    """参考文献条目节点"""

    NODE_TYPE = "references.content"
    NODE_LABEL = "参考文献条目"
    DEFAULTS = {
        **_REF_BASE,
        "first_line_indent": "0字符",
        "font_size": "五号",
        "chinese_font_name": "宋体",
        "builtin_style_name": "正文",
        "english_font_name": "Times New Roman",
        "font_color": "黑色",
        "bold": False,
        "italic": False,
        "underline": False,
    }
