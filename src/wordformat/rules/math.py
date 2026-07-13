#! /usr/bin/env python
"""数学公式节点。公式本身由 OMML 渲染，Node 仅占位确保不被跳过。"""

from wordformat.config.dotdict import BASE_FORMAT, deep_merge
from wordformat.rules.node import FormatNode
from wordformat.structure.registry import register


@register("math_block", level=5)
class MathBlock(FormatNode):
    """块级公式"""

    NODE_TYPE = "math.block"
    NODE_LABEL = "块级公式"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            "paragraph": {"alignment": "居中对齐"},
        },
    )
