#! /usr/bin/env python
# @Time    : 2026/1/9 20:08
# @Author  : afish
# @File    : abstract.py
import re

from wordformat.config.datamodel import (
    AbstractChineseConfig,
    AbstractEnglishConfig,
    AbstractTitleConfig,
)
from wordformat.rules.node import FormatNode
from wordformat.style.check_format import CharacterStyle, ParagraphStyle


class AbstractTitleCN(FormatNode[AbstractTitleConfig]):
    """摘要标题中文节点"""

    NODE_TYPE = "abstract.chinese.chinese_title"
    CONFIG_MODEL = AbstractTitleConfig

    def _base(self, doc, p: bool, r: bool):
        cfg = self.pydantic_config
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
            issues = ps.diff_from_paragraph(self.paragraph)
        else:
            issues = ps.apply_to_paragraph(self.paragraph)
        cstyle = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )

        for run in self.paragraph.runs:
            if r:
                diff_result = cstyle.diff_from_run(run)
            else:
                diff_result = cstyle.apply_to_run(run)
            self.add_comment(
                doc=doc, runs=run, text=CharacterStyle.to_string(diff_result)
            )
        self.add_comment(
            doc=doc, runs=self.paragraph.runs, text=ParagraphStyle.to_string(issues)
        )


class AbstractTitleContentCN(FormatNode[AbstractChineseConfig]):
    """摘要标题正文混合中文节点"""

    NODE_TYPE = "abstract.chinese"
    CONFIG_MODEL = AbstractChineseConfig

    def check_title(self, run) -> bool:
        """检查标题是否包含在正文中"""
        pattern = r"摘[^a-zA-Z0-9\u4e00-\u9fff]*要"
        if re.search(pattern, run.text):
            return True
        return False

    def _base(self, doc, p: bool, r: bool):
        cfg = self.pydantic_config
        # 移除段落的意外字符
        ps = ParagraphStyle(
            alignment=cfg.chinese_content.alignment,
            space_before=cfg.chinese_content.space_before,
            space_after=cfg.chinese_content.space_after,
            line_spacing=cfg.chinese_content.line_spacing,
            line_spacingrule=cfg.chinese_content.line_spacingrule,
            first_line_indent=cfg.chinese_content.first_line_indent,
            builtin_style_name=cfg.chinese_content.builtin_style_name,
        )
        if p:
            issues = ps.diff_from_paragraph(self.paragraph)
        else:
            issues = ps.apply_to_paragraph(self.paragraph)

        for run in self.paragraph.runs:  # 检查标题是否包含在正文中
            run.text = run.text.replace("\r", "").replace("\n", "")
            if self.check_title(run):
                # 对run对象设置样式
                C = CharacterStyle(
                    font_name_cn=cfg.chinese_title.chinese_font_name,
                    font_name_en=cfg.chinese_title.english_font_name,
                    font_size=cfg.chinese_title.font_size,
                    font_color=cfg.chinese_title.font_color,
                    bold=cfg.chinese_title.bold,
                    italic=cfg.chinese_title.italic,
                    underline=cfg.chinese_title.underline,
                )
                if r:
                    diff_result = C.diff_from_run(run)
                else:
                    diff_result = C.apply_to_run(run)
            else:
                # 对剩余部分的run设置样式
                C = CharacterStyle(
                    font_name_cn=cfg.chinese_content.chinese_font_name,
                    font_name_en=cfg.chinese_content.english_font_name,
                    font_size=cfg.chinese_content.font_size,
                    font_color=cfg.chinese_content.font_color,
                    bold=cfg.chinese_content.bold,
                    italic=cfg.chinese_content.italic,
                    underline=cfg.chinese_content.underline,
                )
                if r:
                    diff_result = C.diff_from_run(run)
                else:
                    diff_result = C.apply_to_run(run)
            self.add_comment(
                doc=doc, runs=run, text=CharacterStyle.to_string(diff_result)
            )
        self.add_comment(
            doc=doc, runs=self.paragraph.runs, text=ParagraphStyle.to_string(issues)
        )


class AbstractContentCN(FormatNode[AbstractChineseConfig]):
    """摘要内容中文节点"""

    NODE_TYPE = "abstract.chinese.chinese_content"
    CONFIG_MODEL = AbstractChineseConfig

    def _base(self, doc, p: bool, r: bool):
        cfg = self.pydantic_config.chinese_content
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
            issues = ps.diff_from_paragraph(self.paragraph)
        else:
            issues = ps.apply_to_paragraph(self.paragraph)
        cstyle = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )
        for _, run in enumerate(self.paragraph.runs):
            if r:
                diff_result = cstyle.diff_from_run(run)
            else:
                diff_result = cstyle.apply_to_run(run)
            self.add_comment(
                doc=doc, runs=run, text=CharacterStyle.to_string(diff_result)
            )
        self.add_comment(
            doc=doc, runs=self.paragraph.runs, text=ParagraphStyle.to_string(issues)
        )


class AbstractTitleEN(FormatNode[AbstractTitleConfig]):
    """摘要标题英文节点"""

    NODE_TYPE = "abstract.english.english_title"
    CONFIG_MODEL = AbstractTitleConfig

    def _base(self, doc, p: bool, r: bool):
        cfg = self.pydantic_config
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
            issues = ps.diff_from_paragraph(self.paragraph)
        else:
            issues = ps.apply_to_paragraph(self.paragraph)
        cstyle = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )

        for run in self.paragraph.runs:
            if r:
                diff_result = cstyle.diff_from_run(run)
            else:
                diff_result = cstyle.apply_to_run(run)
            if diff_result:
                self.add_comment(
                    doc=doc, runs=run, text=CharacterStyle.to_string(diff_result)
                )
        if issues:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text="".join(str(dr) for dr in issues),
            )
        return []


class AbstractTitleContentEN(FormatNode[AbstractEnglishConfig]):
    """摘要标题正文混合英文节点"""

    NODE_TYPE = "abstract.english"
    CONFIG_MODEL = AbstractEnglishConfig

    def check_title(self, run) -> bool:
        """检查标题是否包含在正文中"""
        pattern = r"Abstract"
        if re.search(pattern, run.text):
            return True

        return False

    def _base(self, doc, p: bool, r: bool):
        """
        设置 摘要 样式
        """
        cfg = self.pydantic_config
        ps = ParagraphStyle(
            alignment=cfg.english_content.alignment,
            space_before=cfg.english_content.space_before,
            space_after=cfg.english_content.space_after,
            line_spacing=cfg.english_content.line_spacing,
            line_spacingrule=cfg.english_content.line_spacingrule,
            first_line_indent=cfg.english_content.first_line_indent,
            builtin_style_name=cfg.english_content.builtin_style_name,
        )
        if p:
            issues = ps.diff_from_paragraph(self.paragraph)
        else:
            issues = ps.apply_to_paragraph(self.paragraph)

        for run in self.paragraph.runs:
            # 检查标题是否包含在正文中
            if self.check_title(run):
                # 对run对象设置样式
                c = CharacterStyle(
                    font_name_cn=cfg.english_title.chinese_font_name,
                    font_name_en=cfg.english_title.english_font_name,
                    font_size=cfg.english_title.font_size,
                    font_color=cfg.english_title.font_color,
                    bold=cfg.english_title.bold,
                    italic=cfg.english_title.italic,
                    underline=cfg.english_title.underline,
                )

            else:
                # 对剩余部分的run设置样式
                c = CharacterStyle(
                    font_name_cn=cfg.english_content.chinese_font_name,
                    font_name_en=cfg.english_content.english_font_name,
                    font_size=cfg.english_content.font_size,
                    font_color=cfg.english_content.font_color,
                    bold=cfg.english_content.bold,
                    italic=cfg.english_content.italic,
                    underline=cfg.english_content.underline,
                )
            if r:
                diff_result = c.diff_from_run(run)
            else:
                diff_result = c.apply_to_run(run)
            self.add_comment(
                doc=doc, runs=run, text=CharacterStyle.to_string(diff_result)
            )
        self.add_comment(
            doc=doc, runs=self.paragraph.runs, text=ParagraphStyle.to_string(issues)
        )


class AbstractContentEN(FormatNode[AbstractEnglishConfig]):
    """摘要内容英文节点"""

    NODE_TYPE = "abstract.english.english_content"
    CONFIG_MODEL = AbstractEnglishConfig

    def _base(self, doc, p: bool, r: bool):
        cfg = self.pydantic_config.english_content
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
            issues = ps.diff_from_paragraph(self.paragraph)
        else:
            issues = ps.apply_to_paragraph(self.paragraph)
        cstyle = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )
        for run in self.paragraph.runs:
            if r:
                diff_result = cstyle.diff_from_run(run)
            else:
                diff_result = cstyle.apply_to_run(run)
            if diff_result:  # 可选：仅当有差异时才添加批注
                self.add_comment(
                    doc=doc, runs=run, text=CharacterStyle.to_string(diff_result)
                )
        if issues:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text="".join(str(dr) for dr in issues),
            )
