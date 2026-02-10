#! /usr/bin/env python
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : references.py

from wordformat.config.datamodel import ReferencesContentConfig, ReferencesTitleConfig
from wordformat.rules.node import FormatNode
from wordformat.style.check_format import CharacterStyle, ParagraphStyle


class References(FormatNode[ReferencesTitleConfig]):
    """参考文献节点"""

    NODE_TYPE = "references"
    CONFIG_MODEL = ReferencesTitleConfig

    def _base(self, doc, p: bool, r: bool):
        cfg = self.pydantic_config
        # 段落样式
        ps = ParagraphStyle(
            alignment=cfg.alignment,
            space_before=cfg.space_before,
            space_after=cfg.space_after,
            line_spacing=cfg.line_spacing,
            line_spacingrule=cfg.line_spacingrule,
            first_line_indent=cfg.first_line_indent,
            builtin_style_name=cfg.builtin_style_name,
        )
        if p:
            paragraph_issues = ps.diff_from_paragraph(self.paragraph)
        else:
            paragraph_issues = ps.apply_to_paragraph(self.paragraph)
        # 字符样式
        cstyle = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )

        # 检查每个 run 的字符格式
        for run in self.paragraph.runs:
            if r:
                diff_result = cstyle.diff_from_run(run)
            else:
                diff_result = cstyle.apply_to_run(run)
            if diff_result:  # 仅当有差异时添加批注
                self.add_comment(
                    doc=doc, runs=run, text="".join(str(dr) for dr in diff_result)
                )

        # 检查段落格式差异
        if paragraph_issues:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text="".join(str(issue) for issue in paragraph_issues),
            )
        return []


class ReferenceEntry(FormatNode[ReferencesContentConfig]):
    """参考文献条目节点"""

    CONFIG_MODEL = ReferencesContentConfig

    def _base(self, doc, p: bool, r: bool):
        cfg = self.pydantic_config
        # 段落样式
        ps = ParagraphStyle(
            alignment=cfg.alignment,
            space_before=cfg.space_before,
            space_after=cfg.space_after,
            line_spacing=cfg.line_spacing,
            line_spacingrule=cfg.line_spacingrule,
            first_line_indent=cfg.first_line_indent,
            builtin_style_name=cfg.builtin_style_name,
        )
        if p:
            paragraph_issues = ps.diff_from_paragraph(self.paragraph)
        else:
            paragraph_issues = ps.apply_to_paragraph(self.paragraph)

        # 字符样式
        cstyle = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )

        # 检查每个 run 的字符格式
        for run in self.paragraph.runs:
            if r:
                diff_result = cstyle.diff_from_run(run)
            else:
                diff_result = cstyle.apply_to_run(run)
            if diff_result:  # 仅当有差异时添加批注
                self.add_comment(
                    doc=doc, runs=run, text=CharacterStyle.to_string(diff_result)
                )

        # 检查段落格式差异
        if paragraph_issues:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=ParagraphStyle.to_string(paragraph_issues),
            )
        return []
