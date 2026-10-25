#! /usr/bin/env python
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : heading.py

from wordformat.config.dotdict import BASE_FORMAT, deep_merge
from wordformat.rules.node import FormatNode
from wordformat.structure.registry import register


class _HeadingChecks:
    """标题结构性单行：首行缩进无视觉意义，不作为标题的判定项。"""

    def _para_checks_to_skip(self) -> set[str]:
        return {"first_line_indent"}


@register("heading_level_1", level=1)
class HeadingLevel1Node(_HeadingChecks, FormatNode):
    """一级标题节点"""

    NODE_TYPE = "headings.level_1"
    NODE_LABEL = "一级标题"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            "paragraph": {
                # 各校模板对齐/字号差异较大：居中/两端/左对齐、小二/三号/小三/小四均常见
                "alignment": ["居中对齐", "两端对齐", "左对齐"],
                "space_before": ["0.5行", "0行"],
                "space_after": ["0.5行", "0行"],
                "first_line_indent": "0字符",
                "line_spacingrule": ["1.5倍行距", "固定值"],
                "builtin_style_name": "Heading 1",
            },
            "font": {
                "chinese_font_name": "黑体",
                "font_size": ["二号", "小二", "三号", "小三", "小四"],
                # 黑体标题本身视觉已加粗，w:b 开/关各校不一
                "bold": None,
            },
        },
    )


@register("heading_level_2", level=2)
class HeadingLevel2Node(_HeadingChecks, FormatNode):
    """二级标题节点"""

    NODE_TYPE = "headings.level_2"
    NODE_LABEL = "二级标题"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            "paragraph": {
                "alignment": ["左对齐", "两端对齐"],
                # 段前段后 0/0.5 行两种排版都常见
                "space_before": ["0行", "0.5行"],
                "space_after": ["0行", "0.5行"],
                "first_line_indent": "0字符",
                "line_spacingrule": ["1.5倍行距", "固定值"],
                "builtin_style_name": "Heading 2",
            },
            "font": {
                "chinese_font_name": "黑体",
                "font_size": ["三号", "小四"],
                "bold": None,
            },
        },
    )


@register("heading_level_3", level=3)
class HeadingLevel3Node(_HeadingChecks, FormatNode):
    """三级标题节点"""

    NODE_TYPE = "headings.level_3"
    NODE_LABEL = "三级标题"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            "paragraph": {
                "alignment": ["左对齐", "两端对齐"],
                "space_before": ["0行", "0.5行"],
                "space_after": ["0行", "0.5行"],
                "first_line_indent": "0字符",
                # 部分模板用“固定值 20 磅”近似 1.5 倍行距
                "line_spacingrule": ["1.5倍行距", "固定值"],
                "builtin_style_name": "Heading 3",
            },
            "font": {
                "chinese_font_name": "黑体",
                "font_size": ["小四", "五号"],
                "bold": None,
            },
        },
    )
