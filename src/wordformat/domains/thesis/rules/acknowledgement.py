#! /usr/bin/env python
# @Time    : 2026/1/11 21:57
# @Author  : afish
# @File    : acknowledgement.py

from wordformat.config.dotdict import BASE_FORMAT, deep_merge
from wordformat.rules.node import FormatNode
from wordformat.structure.registry import register


@register("acknowledgements_title", level=1)
class Acknowledgements(FormatNode):
    """致谢节点"""

    NODE_TYPE = "acknowledgements.title"
    NODE_LABEL = "致谢标题"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            "paragraph": {
                "alignment": "居中对齐",
                "first_line_indent": "0字符",
            },
            "font": {"chinese_font_name": "黑体", "font_size": "小二", "bold": True},
        },
    )


@register("acknowledgements_content")
class AcknowledgementsCN(FormatNode):
    """致谢内容"""

    NODE_TYPE = "acknowledgements.body"
    NODE_LABEL = "致谢内容"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            "paragraph": {"alignment": "两端对齐"},
            "font": {"chinese_font_name": "宋体"},
        },
    )
