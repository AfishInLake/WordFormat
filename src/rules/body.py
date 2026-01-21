#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : body.py
from typing import Dict

from src.rules.node import FormatNode
from src.rules.style import *


class BodyText(FormatNode):
    """正文节点"""
    NODE_TYPE = 'body_text'

    def check_format(self, doc) -> List[Dict[str, Any]]:
        """
        检查正文段落的字符与段落格式是否符合规范。
        """
        cfg = self.config
        # 段落样式
        ps = ParagraphStyle(
            alignment=cfg.get('alignment', '两端对齐'),
            space_before=cfg.get('space_before', '0磅'),
            space_after=cfg.get('space_after', '0磅'),
            line_spacing=cfg.get('line_spacing', '1.5倍'),
            first_line_indent=cfg.get('first_line_indent', '2字符'),
            builtin_style_name=cfg.get('builtin_style_name', '正文')
        )
        paragraph_issues = ps.diff_from_paragraph(self.paragraph)

        # 字符样式
        cstyle = CharacterStyle(
            font_name_cn=cfg.get('chinese_font_name', '宋体'),
            font_name_en=cfg.get('english_font_name', 'Times New Roman'),
            font_size=cfg.get('font_size', '小四'),
            font_color=cfg.get('font_color', 'BLACK'),
            bold=cfg.get('bold', False),
            italic=cfg.get('italic', False),
            underline=cfg.get('underline', False),
        )

        # 检查每个 run 的字符格式
        for run in self.paragraph.runs:
            diff_result = cstyle.diff_from_run(run)
            if diff_result:  # 仅当有差异时添加批注
                self.add_comment(doc=doc, runs=run, text=''.join(str(dr) for dr in diff_result))

        # 检查段落格式差异
        if paragraph_issues:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=''.join(str(issue) for issue in paragraph_issues)
            )

        return []
