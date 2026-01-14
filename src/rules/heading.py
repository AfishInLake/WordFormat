#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : heading.py
from typing import List, Dict, Any

from src.rules.node import FormatNode
from src.rules.style import *


class Heading(FormatNode):
    """标题节点"""

    def check_format(self, doc) -> List[Dict[str, Any]]:
        return []


class HeadingLevel1Node(Heading):
    """一级标题节点"""

    def check_format(self, doc) -> List[Dict[str, Any]]:
        ps = ParagraphStyle(
            alignment=Alignment.LEFT,
            space_before=Spacing.HALF_LINE,
            space_after=Spacing.HALF_LINE,
            line_spacing=LineSpacing.ONE_POINT_FIVE,
            first_line_indent=FirstLineIndent.NONE,
            builtin_style_name=BuiltInStyle.HEADING_1
        )
        ps.apply_to(self.paragraph)
        for index, run in enumerate(self.paragraph.runs):
            CharacterStyle(
                font_name=FontName.SIM_HEI,
                font_size=FontSize.XIAO_ER,
                font_color=FontColor.BLACK,
                bold=False,
                italic=False,
                underline=False
            ).apply_to(run)
        return []


class HeadingLevel2Node(Heading):
    """二级标题节点"""

    def check_format(self, doc) -> List[Dict[str, Any]]:
        ps = ParagraphStyle(
            alignment=Alignment.LEFT,
            space_before=Spacing.HALF_LINE,
            space_after=Spacing.NONE,
            line_spacing=LineSpacing.ONE_POINT_FIVE,
            first_line_indent=FirstLineIndent.NONE,
            builtin_style_name=BuiltInStyle.HEADING_2
        )
        ps.apply_to(self.paragraph)
        for index, run in enumerate(self.paragraph.runs):
            CharacterStyle(
                font_name=FontName.SIM_HEI,
                font_size=FontSize.SAN_HAO,
                font_color=FontColor.BLACK,
                bold=False,
                italic=False,
                underline=False
            ).apply_to(run)
        return []


class HeadingLevel3Node(Heading):
    """三级标题节点"""

    def check_format(self, doc) -> List[Dict[str, Any]]:
        ps = ParagraphStyle(
            alignment=Alignment.LEFT,
            space_before=Spacing.HALF_LINE,
            space_after=Spacing.NONE,
            line_spacing=LineSpacing.ONE_POINT_FIVE,
            first_line_indent=FirstLineIndent.NONE,
            builtin_style_name=BuiltInStyle.HEADING_3
        )
        ps.apply_to(self.paragraph)
        for index, run in enumerate(self.paragraph.runs):
            CharacterStyle(
                font_name=FontName.SIM_HEI,
                font_size=FontSize.XIAO_SI,
                font_color=FontColor.BLACK,
                bold=False,
                italic=False,
                underline=False
            ).apply_to(run)
        return []
