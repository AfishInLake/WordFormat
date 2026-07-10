#! /usr/bin/env python
# @Time    : 2026/1/11 21:57
# @Author  : afish
# @File    : acknowledgement.py

from wordformat.rules.node import FormatNode
from wordformat.structure.registry import register

_ACK = {
    "alignment": "左对齐",
    "space_before": "0.5行",
    "space_after": "0.5行",
    "line_spacingrule": "单倍行距",
    "line_spacing": "1.5倍",
    "left_indent": "0字符",
    "right_indent": "0字符",
    "builtin_style_name": "正文",
    "english_font_name": "Times New Roman",
    "font_color": "黑色",
    "italic": False,
    "underline": False,
}


@register("acknowledgements_title", level=1)
class Acknowledgements(FormatNode):
    """致谢节点"""

    NODE_TYPE = "acknowledgements.title"
    NODE_LABEL = "致谢标题"
    DEFAULTS = {
        **_ACK,
        "alignment": "居中对齐",
        "first_line_indent": "0字符",
        "chinese_font_name": "黑体",
        "font_size": "小二",
        "bold": True,
    }


@register("acknowledgements_content")
class AcknowledgementsCN(FormatNode):
    """致谢内容"""

    NODE_TYPE = "acknowledgements.content"
    NODE_LABEL = "致谢内容"
    DEFAULTS = {
        **_ACK,
        "alignment": "两端对齐",
        "first_line_indent": "2字符",
        "chinese_font_name": "宋体",
        "font_size": "小四",
    }
