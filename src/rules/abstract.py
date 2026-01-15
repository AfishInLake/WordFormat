#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/9 20:08
# @Author  : afish
# @File    : abstract.py
import re
from typing import Dict, Any, List

from src.rules.node import FormatNode
from src.rules.style import *


class AbstractTitleCN(FormatNode):
    """摘要标题中文节点"""
    NODE_TYPE = 'abstract.chinese.chinese_title'

    def check_format(self, doc) -> List[Dict[str, Any]]:
        """
        设置 摘要 样式
        """
        ps = ParagraphStyle(
            alignment=Alignment.CENTER,
            space_before=Spacing.NONE,
            space_after=Spacing.NONE,
            line_spacing=LineSpacing.ONE_POINT_FIVE,
            first_line_indent=FirstLineIndent.NONE,
            builtin_style_name=BuiltInStyle.NORMAL
        )
        ps.apply_to(self.paragraph)

        for index, run in enumerate(self.paragraph.runs):
            # 对run对象设置样式
            CharacterStyle(
                font_name=FontName.SIM_HEI,
                font_size=FontSize.XIAO_SI,
                font_color=FontColor.BLACK,
                bold=True,
                italic=False,
                underline=False
            ).apply_to(run)

        return []


class AbstractTitleContentCN(FormatNode):
    """摘要标题正文混合中文节点"""
    NODE_TYPE = 'abstract.chinese'

    def check_title(self, run) -> bool:
        """检查标题是否包含在正文中"""
        pattern = r'摘[^a-zA-Z0-9\u4e00-\u9fff]*要'
        if re.search(pattern, run.text):
            return True

        return False

    def check_format(self, doc) -> List[Dict[str, Any]]:
        """
        设置 摘要 样式
        """
        ps = ParagraphStyle(
            alignment=Alignment.JUSTIFY,
            space_before=Spacing.NONE,
            space_after=Spacing.NONE,
            line_spacing=LineSpacing.ONE_POINT_FIVE,
            first_line_indent=FirstLineIndent.TWO_CHARS,
            builtin_style_name=BuiltInStyle.NORMAL
        )
        ps.apply_to(self.paragraph)

        for index, run in enumerate(self.paragraph.runs):
            # 检查标题是否包含在正文中
            if self.check_title(run):
                # 对run对象设置样式
                CharacterStyle(
                    font_name=FontName.SIM_HEI,
                    font_size=FontSize.XIAO_SI,
                    font_color=FontColor.BLACK,
                    bold=True,
                    italic=False,
                    underline=False
                ).apply_to(run)
                self.add_comment(doc=doc, runs=run, text=f"{run.text} 已经设置")
            else:
                # 对剩余部分的run设置样式
                CharacterStyle(
                    font_name=FontName.SIM_SUN,
                    font_size=FontSize.XIAO_SI,
                    font_color=FontColor.BLACK,
                    bold=False,
                    italic=False,
                    underline=False
                ).apply_to(run)
        return []


class AbstractContentCN(FormatNode):
    """摘要内容中文节点"""
    NODE_TYPE = 'abstract.chinese.chinese_content'

    def check_format(self, doc) -> List[Dict[str, Any]]:
        ps = ParagraphStyle(
            alignment=Alignment.JUSTIFY,
            space_before=Spacing.NONE,
            space_after=Spacing.NONE,
            line_spacing=LineSpacing.ONE_POINT_FIVE,
            first_line_indent=FirstLineIndent.TWO_CHARS,
            builtin_style_name=BuiltInStyle.NORMAL
        )
        ps.apply_to(self.paragraph)
        for index, run in enumerate(self.paragraph.runs):
            CharacterStyle(
                font_name=FontName.SIM_SUN,
                font_size=FontSize.XIAO_SI,
                font_color=FontColor.BLACK,
                bold=False,
                italic=False,
                underline=False
            ).apply_to(run)
        return []


class AbstractTitleEN(FormatNode):
    """摘要标题英文节点"""
    NODE_TYPE = 'abstract.english.english_title'

    def check_format(self, doc) -> List[Dict[str, Any]]:
        issues = []

        return issues


class AbstractTitleContentEN(FormatNode):
    """摘要标题正文混合英文节点"""
    NODE_TYPE = 'abstract.english'

    def check_title(self, run) -> bool:
        """检查标题是否包含在正文中"""
        pattern = r'Abstract'
        if re.search(pattern, run.text):
            return True

        return False

    def check_format(self, doc) -> List[Dict[str, Any]]:
        """
        设置 摘要 样式
        """
        ps = ParagraphStyle(
            alignment=Alignment.JUSTIFY,
            space_before=Spacing.NONE,
            space_after=Spacing.NONE,
            line_spacing=LineSpacing.ONE_POINT_FIVE,
            first_line_indent=FirstLineIndent.TWO
        )
        ps.apply_to(self.paragraph)
        for index, run in enumerate(self.paragraph.runs):
            # 检查标题是否包含在正文中
            if self.check_title(run):
                # 对run对象设置样式
                CharacterStyle(
                    font_name=FontName.SIM_HEI,
                    font_size=FontSize.XIAO_SI,
                    font_color=FontColor.BLACK,
                    bold=True,
                    italic=False,
                    underline=False
                ).apply_to(run)
                self.add_comment(doc=doc, runs=run, text=f"{run.text} 已经设置")
            else:
                # 对剩余部分的run设置样式
                CharacterStyle(
                    font_name=FontName.SIM_SUN,
                    font_size=FontSize.XIAO_SI,
                    font_color=FontColor.BLACK,
                    bold=False,
                    italic=False,
                    underline=False
                ).apply_to(run)
                self.add_comment(doc=doc, runs=run, text=f"{run.text} 已经设置")

        return []


class AbstractContentEN(FormatNode):
    """摘要内容英文节点"""
    NODE_TYPE = 'abstract.english.english_content'

    def check_format(self, doc) -> List[Dict[str, Any]]:
        issues = []
        rule = self.expected_rule.get("english", {})
        # 检查 content_font / content_size_pt
        return issues


class KeywordsEN(FormatNode):
    """关键词节点英文"""
    NODE_TYPE = 'abstract.keywords.english'

    def check_format(self, doc) -> List[Dict[str, Any]]:
        """解析关键词字符串为列表"""
        pass


class KeywordsCN(FormatNode):
    """关键词节点中文"""
    NODE_TYPE = 'abstract.keywords.chinese'

    def check_keword(self, run) -> bool:
        """检查标题是否包含在正文中"""
        pattern = r'关[^a-zA-Z0-9\u4e00-\u9fff]*键[^a-zA-Z0-9\u4e00-\u9fff]*词'
        if re.search(pattern, run.text):
            return True
        return False

    def check_format(self, doc) -> List[Dict[str, Any]]:
        """解析关键词字符串为列表"""
        ps = ParagraphStyle(
            alignment=Alignment.JUSTIFY,
            space_before=Spacing.NONE,
            space_after=Spacing.NONE,
            line_spacing=LineSpacing.ONE_POINT_FIVE,
            first_line_indent=FirstLineIndent.TWO_CHARS,
            builtin_style_name=BuiltInStyle.NORMAL
        )
        ps.apply_to(self.paragraph)
        for index, run in enumerate(self.paragraph.runs):
            if self.check_keword(run):
                CharacterStyle(
                    font_name=FontName.SIM_HEI,
                    font_size=FontSize.SI_HAO,
                    font_color=FontColor.BLACK,
                    bold=True,
                    italic=False,
                    underline=False
                ).apply_to(run)
                self.add_comment(doc=doc, runs=run, text=f"{run.text} 已经设置")
            else:
                CharacterStyle(
                    font_name=FontName.SIM_SUN,
                    font_size=FontSize.XIAO_SI,
                    font_color=FontColor.BLACK,
                    bold=False,
                    italic=False,
                    underline=False
                ).apply_to(run)
        return []
