#! /usr/bin/env python
# @Time    : 2026/1/11 19:43
# @Author  : afish
# @File    : caption.py

from src.config.datamodel import FiguresConfig, TablesConfig
from src.rules.node import FormatNode
from src.style.check_format import CharacterStyle, ParagraphStyle


class CaptionFigure(FormatNode[FiguresConfig]):
    """题注-图片"""

    NODE_TYPE = "figures"
    CONFIG_MODEL = FiguresConfig

    def _base(self, doc, p: bool, r: bool):
        cfg = self.pydantic_config
        # 段落样式
        ps = ParagraphStyle(
            alignment=cfg.alignment,
            space_before=cfg.space_before,
            space_after=cfg.space_after,
            line_spacing=cfg.line_spacing,
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
                text="".join(str(issue) for issue in paragraph_issues),
            )


class CaptionTable(FormatNode[TablesConfig]):
    """题注-表格"""

    NODE_TYPE = "tables"
    CONFIG_MODEL = TablesConfig

    def _base(self, doc, p: bool, r: bool):
        # TODO:暂时无法处理表格分页情况
        # cfg = self.config
        # # 段落样式
        # ps = ParagraphStyle(
        #     alignment=cfg.get("alignment", "两端对齐"),
        #     space_before=cfg.get("space_before", 0),
        #     space_after=cfg.get("space_after", 0),
        #     line_spacing=cfg.get("line_spacing", "1.5倍"),
        #     first_line_indent=cfg.get("first_line_indent", "2字符"),
        #     builtin_style_name=cfg.get("builtin_style_name", "正文"),
        # )
        # paragraph_issues = ps.diff_from_paragraph(self.paragraph)
        #
        # # 字符样式
        # cstyle = CharacterStyle(
        #     font_name_cn=cfg.get("chinese_font_name", "宋体"),
        #     font_name_en=cfg.get("english_font_name", "Times New Roman"),
        #     font_size=cfg.get("font_size", "小四"),
        #     font_color=cfg.get("font_color", "BLACK"),
        #     bold=cfg.get("bold", False),
        #     italic=cfg.get("italic", False),
        #     underline=cfg.get("underline", False),
        # )
        #
        # # 检查每个 run 的字符格式
        # for run in self.paragraph.runs:
        #     diff_result = cstyle.diff_from_run(run)
        #     if diff_result:  # 仅当有差异时添加批注
        #         self.add_comment(
        #             doc=doc, runs=run, text="".join(str(dr) for dr in diff_result)
        #         )
        #
        # # 检查段落格式差异
        # if paragraph_issues:
        #     self.add_comment(
        #         doc=doc,
        #         runs=self.paragraph.runs,
        #         text="".join(str(issue) for issue in paragraph_issues),
        #     )
        return []
