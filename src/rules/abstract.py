#! /usr/bin/env python
# @Time    : 2026/1/9 20:08
# @Author  : afish
# @File    : abstract.py
import re
from typing import Any, cast

from src.config.datamodel import (
    AbstractChineseConfig,
    AbstractContentConfig,
    AbstractEnglishConfig,
    AbstractTitleConfig,
)
from src.rules.node import FormatNode
from src.style.check_format import CharacterStyle, ParagraphStyle


class AbstractTitleCN(FormatNode):
    """摘要标题中文节点"""

    NODE_TYPE = "abstract.chinese.chinese_title"
    CONFIG_MODEL = AbstractTitleConfig

    def check_format(self, doc) -> list[dict[str, Any]]:
        """
        检查 摘要 样式
        """
        cfg: AbstractTitleConfig = cast("AbstractTitleConfig", self.pydantic_config)
        ps = ParagraphStyle(
            alignment=cfg.alignment,
            space_before=cfg.space_before,
            space_after=cfg.space_after,
            line_spacing=cfg.line_spacing,
            first_line_indent=cfg.first_line_indent,
            builtin_style_name=cfg.builtin_style_name,
        )
        issues = ps.diff_from_paragraph(self.paragraph)
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
            diff_result = cstyle.diff_from_run(run)
            self.add_comment(
                doc=doc, runs=run, text="".join([str(dr) for dr in diff_result])
            )
        self.add_comment(
            doc=doc, runs=self.paragraph.runs, text="".join([str(dr) for dr in issues])
        )
        return []


class AbstractTitleContentCN(FormatNode):
    """摘要标题正文混合中文节点"""

    NODE_TYPE = "abstract.chinese"
    CONFIG_MODEL = AbstractChineseConfig

    def check_title(self, run) -> bool:
        """检查标题是否包含在正文中"""
        pattern = r"摘[^a-zA-Z0-9\u4e00-\u9fff]*要"
        if re.search(pattern, run.text):
            return True
        return False

    def check_format(self, doc) -> list[dict[str, Any]]:
        """
        设置 摘要 样式
        """
        # cfg = cls.config
        # ps = ParagraphStyle(
        #     alignment=cfg.get('alignment', '左对齐'),
        #     space_before=cfg.get('space_before', 'NONE'),
        #     space_after=cfg.get('space_after', 'NONE'),
        #     line_spacing=cfg.get('line_spacing', '1.5倍'),
        #     first_line_indent=cfg.get('first_line_indent', '2字符'),
        #     builtin_style_name=cfg.get('builtin_style_name', '正文')
        # )
        # ps.apply_to(cls.paragraph)
        #
        # for index, run in enumerate(cls.paragraph.runs):
        #     # 检查标题是否包含在正文中
        #     if cls.check_title(run):
        #         # 对run对象设置样式
        #         CharacterStyle(
        #             font_name_cn=cfg.get('chinese_font_name', '宋体'),
        #             font_name_en=cfg.get('english_font_name', 'Times New Roman'),
        #             font_size=cfg.get('font_size', '小四'),
        #             font_color=cfg.get('font_color', 'BLACK'),
        #             bold=cfg.get('bold', True),
        #             italic=cfg.get('italic', False),
        #             underline=cfg.get('underline', False),
        #         ).apply_to(run)
        #         cls.add_comment(doc=doc, runs=run, text=f"{run.text} 已经设置")
        #     else:
        #         # 对剩余部分的run设置样式
        #         CharacterStyle(
        #             font_name_cn=cfg.get('chinese_font_name', '宋体'),
        #             font_name_en=cfg.get('english_font_name', 'Times New Roman'),
        #             font_size=cfg.get('font_size', '小四'),
        #             font_color=cfg.get('font_color', 'BLACK'),
        #             bold=cfg.get('bold', True),
        #             italic=cfg.get('italic', False),
        #             underline=cfg.get('underline', False),
        #         ).apply_to(run)
        return []


class AbstractContentCN(FormatNode):
    """摘要内容中文节点"""

    NODE_TYPE = "abstract.chinese.chinese_content"
    CONFIG_MODEL = AbstractContentConfig

    def check_format(self, doc) -> list[dict[str, Any]]:
        cfg: AbstractContentConfig = cast("AbstractContentConfig", self.pydantic_config)
        ps = ParagraphStyle(
            alignment=cfg.alignment,
            space_before=cfg.space_before,
            space_after=cfg.space_after,
            line_spacing=cfg.line_spacing,
            first_line_indent=cfg.first_line_indent,
            builtin_style_name=cfg.builtin_style_name,
        )
        issues = ps.diff_from_paragraph(self.paragraph)
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
            diff_result = cstyle.diff_from_run(run)
            self.add_comment(
                doc=doc, runs=run, text="".join([str(dr) for dr in diff_result])
            )
        self.add_comment(
            doc=doc, runs=self.paragraph.runs, text="".join([str(dr) for dr in issues])
        )
        return []


class AbstractTitleEN(FormatNode):
    """摘要标题英文节点"""

    NODE_TYPE = "abstract.english.english_title"
    CONFIG_MODEL = AbstractTitleConfig

    def check_format(self, doc) -> list[dict[str, Any]]:
        cfg: AbstractTitleConfig = cast("AbstractTitleConfig", self.pydantic_config)
        ps = ParagraphStyle(
            alignment=cfg.alignment,
            space_before=cfg.space_before,
            space_after=cfg.space_after,
            line_spacing=cfg.line_spacing,
            first_line_indent=cfg.first_line_indent,
            builtin_style_name=cfg.builtin_style_name,
        )
        issues = ps.diff_from_paragraph(self.paragraph)
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
            diff_result = cstyle.diff_from_run(run)
            if diff_result:
                self.add_comment(
                    doc=doc, runs=run, text="".join(str(dr) for dr in diff_result)
                )
        if issues:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text="".join(str(dr) for dr in issues),
            )
        return []


class AbstractTitleContentEN(FormatNode):
    """摘要标题正文混合英文节点"""

    NODE_TYPE = "abstract.english"
    CONFIG_MODEL = AbstractEnglishConfig

    def check_title(self, run) -> bool:
        """检查标题是否包含在正文中"""
        pattern = r"Abstract"
        if re.search(pattern, run.text):
            return True

        return False

    def check_format(self, doc) -> list[dict[str, Any]]:
        """
        设置 摘要 样式
        """
        # cfg = cls.config
        # ps = ParagraphStyle(
        #     alignment=cfg.get('alignment', '左对齐'),
        #     space_before=cfg.get('space_before', 'NONE'),
        #     space_after=cfg.get('space_after', 'NONE'),
        #     line_spacing=cfg.get('line_spacing', '1.5倍'),
        #     first_line_indent=cfg.get('first_line_indent', '2字符'),
        #     builtin_style_name=cfg.get('builtin_style_name', '正文')
        # )
        # ps.apply_to(cls.paragraph)
        # for index, run in enumerate(cls.paragraph.runs):
        #     # 检查标题是否包含在正文中
        #     if cls.check_title(run):
        #         # 对run对象设置样式
        #         CharacterStyle(
        #             font_name_cn=cfg.get('chinese_font_name', '宋体'),
        #             font_name_en=cfg.get('english_font_name', 'Times New Roman'),
        #             font_size=cfg.get('font_size', '小四'),
        #             font_color=cfg.get('font_color', 'BLACK'),
        #             bold=cfg.get('bold', True),
        #             italic=cfg.get('italic', False),
        #             underline=cfg.get('underline', False),
        #         ).apply_to(run)
        #         cls.add_comment(doc=doc, runs=run, text=f"{run.text} 已经设置")
        #     else:
        #         # 对剩余部分的run设置样式
        #         CharacterStyle(
        #             font_name_cn=cfg.get('chinese_font_name', '宋体'),
        #             font_name_en=cfg.get('english_font_name', 'Times New Roman'),
        #             font_size=cfg.get('font_size', '小四'),
        #             font_color=cfg.get('font_color', 'BLACK'),
        #             bold=cfg.get('bold', True),
        #             italic=cfg.get('italic', False),
        #             underline=cfg.get('underline', False),
        #         ).apply_to(run)
        #         cls.add_comment(doc=doc, runs=run, text=f"{run.text} 已经设置")

        return []


class AbstractContentEN(FormatNode):
    """摘要内容英文节点"""

    NODE_TYPE = "abstract.english.english_content"
    CONFIG_MODEL = AbstractContentConfig

    def check_format(self, doc) -> list[dict[str, Any]]:
        cfg: AbstractContentConfig = cast("AbstractContentConfig", self.pydantic_config)
        ps = ParagraphStyle(
            alignment=cfg.alignment,
            space_before=cfg.space_before,
            space_after=cfg.space_after,
            line_spacing=cfg.line_spacing,
            first_line_indent=cfg.first_line_indent,
            builtin_style_name=cfg.builtin_style_name,
        )
        issues = ps.diff_from_paragraph(self.paragraph)
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
            diff_result = cstyle.diff_from_run(run)
            if diff_result:  # 可选：仅当有差异时才添加批注
                self.add_comment(
                    doc=doc, runs=run, text="".join(str(dr) for dr in diff_result)
                )
        if issues:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text="".join(str(dr) for dr in issues),
            )
        return []
