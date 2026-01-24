#! /usr/bin/env python
# @Time    : 2026/1/23 17:45
# @Author  : afish
# @File    : keywords.py
import re
from typing import Any

from src.rules.node import FormatNode
from src.style.check_format import CharacterStyle, ParagraphStyle


class KeywordsEN(FormatNode):
    """关键词节点英文"""

    NODE_TYPE = "abstract.keywords.english"

    def check_format(self, doc) -> list[dict[str, Any]]:
        cfg = self.config

        # --- 段落样式校验 ---
        ps = ParagraphStyle(
            alignment=cfg.get("alignment", "左对齐"),
            space_before=cfg.get("space_before", "NONE"),
            space_after=cfg.get("space_after", "NONE"),
            line_spacing=cfg.get("line_spacing", "单倍"),
            first_line_indent=cfg.get("first_line_indent", "NONE"),
            builtin_style_name=cfg.get("builtin_style_name", "正文"),
        )
        paragraph_issues = ps.diff_from_paragraph(self.paragraph)

        # --- 字符样式校验 ---
        cstyle = CharacterStyle(
            font_name_cn=cfg.get("chinese_font_name", "宋体"),
            font_name_en=cfg.get("english_font_name", "Times New Roman"),
            font_size=cfg.get("font_size", "小四"),
            font_color=cfg.get("font_color", "BLACK"),
            bold=cfg.get("bold", False),
            italic=cfg.get("italic", False),
            underline=cfg.get("underline", False),
        )

        for run in self.paragraph.runs:
            diff_result = cstyle.diff_from_run(run)
            if diff_result:
                self.add_comment(doc=doc, runs=run, text="".join(str(dr) for dr in diff_result))

        if paragraph_issues:
            self.add_comment(
                doc=doc, runs=self.paragraph.runs, text="".join(str(issue) for issue in paragraph_issues)
            )

        return []


class KeywordsCN(FormatNode):
    """关键词节点中文"""

    NODE_TYPE = "abstract.keywords.chinese"

    def check_keyword(self, run) -> bool:
        """检查 run 是否包含 '关键词' 字样（允许中间有非中英文字符）"""
        pattern = r"关[^a-zA-Z0-9\u4e00-\u9fff]*键[^a-zA-Z0-9\u4e00-\u9fff]*词"
        return bool(re.search(pattern, run.text))

    def check_format(self, doc) -> list[dict[str, Any]]:
        """
        校验中文关键词段落格式：
        - 段落整体格式（对齐、行距等）
        - “关键词：”部分应加粗，其余关键词内容不加粗
        """
        cfg = self.config

        # --- 1. 段落级格式校验 ---
        ps = ParagraphStyle(
            alignment=cfg.get("alignment", "左对齐"),
            space_before=cfg.get("space_before", "NONE"),
            space_after=cfg.get("space_after", "NONE"),
            line_spacing=cfg.get("line_spacing", "1.5倍"),
            first_line_indent=cfg.get("first_line_indent", "2字符"),
            builtin_style_name=cfg.get("builtin_style_name", "正文"),
        )
        paragraph_issues = ps.diff_from_paragraph(self.paragraph)
        if paragraph_issues:
            self.add_comment(
                doc=doc, runs=self.paragraph.runs, text="".join(str(issue) for issue in paragraph_issues)
            )

        # --- 2. 字符级格式校验：分两部分 ---
        # (a) "关键词：" 部分 —— 应加粗
        label_style = CharacterStyle(
            font_name_cn=cfg.get("chinese_font_name", "宋体"),
            font_name_en=cfg.get("english_font_name", "Times New Roman"),
            font_size=cfg.get("font_size", "小四"),
            font_color=cfg.get("font_color", "BLACK"),
            bold=cfg.get("kewords_bold", True),  # ← 关键词标签必须加粗
            italic=cfg.get("italic", False),
            underline=cfg.get("underline", False),
        )

        # (b) 关键词内容部分（如“人工智能；深度学习”）—— 不加粗
        content_style = CharacterStyle(
            font_name_cn=cfg.get("chinese_font_name", "宋体"),
            font_name_en=cfg.get("english_font_name", "Times New Roman"),
            font_size=cfg.get("font_size", "小四"),
            font_color=cfg.get("font_color", "BLACK"),
            bold=cfg.get("bold", False),  # ← 内容不加粗
            italic=cfg.get("italic", False),
            underline=cfg.get("underline", False),
        )

        # --- 3. 遍历每个 run 进行校验 ---
        for run in self.paragraph.runs:
            if not run.text.strip():
                continue  # 跳过空白 run

            if self.check_keyword(run):
                # 期望是 label_style（加粗）
                diff = label_style.diff_from_run(run)
                if diff:
                    self.add_comment(
                        doc=doc, runs=run, text=f"「关键词」标签格式不符: {''.join(str(d) for d in diff)}"
                    )
            else:
                # 期望是 content_style（不加粗）
                diff = content_style.diff_from_run(run)
                if diff:
                    self.add_comment(
                        doc=doc, runs=run, text=f"关键词内容格式不符: {''.join(str(d) for d in diff)}"
                    )
        return []
