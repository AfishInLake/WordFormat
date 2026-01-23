#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : heading.py
from typing import List, Dict, Any

from src.rules.node import FormatNode
from src.style.check_format import ParagraphStyle, CharacterStyle


class HeadingLevel1Node(FormatNode):
    """一级标题节点"""
    NODE_TYPE = 'headings.level_1'

    def check_format(self, doc) -> List[Dict[str, Any]]:
        cfg = self.config
        ps = ParagraphStyle(
            alignment=cfg.get('alignment', '居中对齐'),
            space_before=cfg.get('space_before', 'NONE'),
            space_after=cfg.get('space_after', 'NONE'),
            line_spacing=cfg.get('line_spacing', '单倍行距'),
            first_line_indent=cfg.get('first_line_indent', 'NONE'),
            builtin_style_name=cfg.get('builtin_style_name', '标题1')
        )
        issues = ps.diff_from_paragraph(self.paragraph)
        cstyle = CharacterStyle(
            font_name_cn=cfg.get('chinese_font_name', '宋体'),
            font_name_en=cfg.get('english_font_name', 'Times New Roman'),
            font_size=cfg.get('font_size', '二号'),
            font_color=cfg.get('font_color', 'BLACK'),
            bold=cfg.get('bold', True),
            italic=cfg.get('italic', False),
            underline=cfg.get('underline', False),
        )

        for run in self.paragraph.runs:
            diff_result = cstyle.diff_from_run(run)
            if diff_result:
                self.add_comment(doc=doc, runs=run, text=''.join(str(dr) for dr in diff_result))
        if issues:
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=''.join(str(dr) for dr in issues))
        return []


class HeadingLevel2Node(FormatNode):
    """二级标题节点"""
    NODE_TYPE = 'headings.level_2'
    def check_format(self, doc) -> List[Dict[str, Any]]:
        cfg = self.config
        ps = ParagraphStyle(
            alignment=cfg.get('alignment', '居中对齐'),
            space_before=cfg.get('space_before', 'NONE'),
            space_after=cfg.get('space_after', 'NONE'),
            line_spacing=cfg.get('line_spacing', '单倍行距'),
            first_line_indent=cfg.get('first_line_indent', 'NONE'),
            builtin_style_name=cfg.get('builtin_style_name', '标题2')
        )
        issues = ps.diff_from_paragraph(self.paragraph)
        cstyle = CharacterStyle(
            font_name_cn=cfg.get('chinese_font_name', '宋体'),
            font_name_en=cfg.get('english_font_name', 'Times New Roman'),
            font_size=cfg.get('font_size', '小四'),
            font_color=cfg.get('font_color', 'BLACK'),
            bold=cfg.get('bold', True),
            italic=cfg.get('italic', False),
            underline=cfg.get('underline', False),
        )

        for run in self.paragraph.runs:
            diff_result = cstyle.diff_from_run(run)
            if diff_result:
                self.add_comment(doc=doc, runs=run, text=''.join(str(dr) for dr in diff_result))
        if issues:
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=''.join(str(dr) for dr in issues))
        return []


class HeadingLevel3Node(FormatNode):
    """三级标题节点"""
    NODE_TYPE = 'headings.level_3'
    def check_format(self, doc) -> List[Dict[str, Any]]:
        cfg = self.config
        ps = ParagraphStyle(
            alignment=cfg.get('alignment', '居中对齐'),
            space_before=cfg.get('space_before', 'NONE'),
            space_after=cfg.get('space_after', 'NONE'),
            line_spacing=cfg.get('line_spacing', '单倍行距'),
            first_line_indent=cfg.get('first_line_indent', 'NONE'),
            builtin_style_name=cfg.get('builtin_style_name', '标题3')
        )
        issues = ps.diff_from_paragraph(self.paragraph)
        cstyle = CharacterStyle(
            font_name_cn=cfg.get('chinese_font_name', '宋体'),
            font_name_en=cfg.get('english_font_name', 'Times New Roman'),
            font_size=cfg.get('font_size', '小四'),
            font_color=cfg.get('font_color', 'BLACK'),
            bold=cfg.get('bold', True),
            italic=cfg.get('italic', False),
            underline=cfg.get('underline', False),
        )

        for run in self.paragraph.runs:
            diff_result = cstyle.diff_from_run(run)
            if diff_result:
                self.add_comment(doc=doc, runs=run, text=''.join(str(dr) for dr in diff_result))
        if issues:
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=''.join(str(dr) for dr in issues))
        return []
