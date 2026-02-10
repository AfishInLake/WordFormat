#! /usr/bin/env python
# @Time    : 2026/1/12 10:46
# @Author  : afish
# @File    : style.py
from dataclasses import dataclass
from typing import Any

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from loguru import logger

from wordformat.config.config import get_config
from wordformat.config.datamodel import WarningFieldConfig
from wordformat.style.get_some import (
    paragraph_get_alignment,
    paragraph_get_builtin_style_name,
    paragraph_get_first_line_indent,
    paragraph_get_line_spacing,
    paragraph_get_space_after,
    paragraph_get_space_before,
    run_get_font_color,
    run_get_font_name,
    run_get_font_size_pt,
)

from .style_enum import (
    Alignment,
    BuiltInStyle,
    FirstLineIndent,
    FontColor,
    FontName,
    FontSize,
    LineSpacing,
    LineSpacingRule,
    Spacing,
)

style_checks_warning: WarningFieldConfig | None = None


@dataclass
class DIFFResult:
    """
    用来保存段落差异

    Attributes:
        diff_type: 不同类别
        expected_value: 期待值
        current_value: 当前值
        comment: 评论
    """

    diff_type: str = None
    expected_value: Any = None
    current_value: Any = None
    comment: str = None
    level: int = 0

    def __str__(self):
        return self.comment


class CharacterStyle:
    """字符样式类，用于定义 Word 文档中 Run 级别的文本格式。

    该类封装了常见的字符级格式属性，如字体名称、字号、颜色、加粗、斜体、下划线等，
    通常用于格式校验、自动修复或样式比对。所有字段均有默认值，符合中文文档常见排版规范。

    Attributes:
        font_name_cn (FontName): 字体名称（如黑体、宋体）。注意：中文字体需通过 `w:eastAsia` 属性设置。
        font_name_en (FontName): 字体名称（如Times New Roman）
        font_size (FontSize): 字号（如小四、四号等），内部以磅（pt）为单位存储。
        font_color (FontColor): 字体颜色，默认为黑色（RGB(0, 0, 0)）。
        bold (bool): 是否加粗。True 表示加粗，False 表示不加粗。
        italic (bool): 是否斜体。True 表示斜体，False 表示非斜体。
        underline (bool): 是否带下划线。True 表示有下划线，False 表示无下划线。
    """

    def __init__(
        self,
        font_name_cn: str = "宋体",
        font_name_en: str = "Times New Roman",
        font_size: str | float = "小四",
        font_color: str | tuple = "BLACK",
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
    ):
        self.font_name_cn: FontName = FontName(font_name_cn)
        self.font_name_en: FontName = FontName(font_name_en)
        self.font_size: FontSize = FontSize(font_size)
        self.font_color: FontColor = FontColor(font_color)
        self.bold: bool = bold
        self.italic: bool = italic
        self.underline: bool = underline
        if globals()["style_checks_warning"] is None:
            globals()["style_checks_warning"] = get_config().style_checks_warning

    def diff_from_run(self, run: Run) -> list[DIFFResult]:  # noqa c901
        """
        检查段落样式和指定样式是否一致
        """

        diffs = []

        # 1. 加粗
        bold = bool(run.font.bold)
        if bold != self.bold:
            diffs.append(
                DIFFResult(
                    "bold",
                    self.bold,
                    bold,
                    f"期待{'加粗' if self.bold else '不加粗'};",
                    1,
                )
            )
        # 2. 斜体
        italic = bool(run.font.italic)
        if italic != self.italic:
            diffs.append(
                DIFFResult(
                    "italic",
                    self.italic,
                    italic,
                    f"期待{'斜体' if self.italic else '非斜体'};",
                    1,
                )
            )

        # 3. 下划线
        underline = bool(run.font.underline)
        if underline != self.underline:
            diffs.append(
                DIFFResult(
                    "underline",
                    self.underline,
                    underline,
                    f"期待{'有下划线' if self.underline else '无下划线'};",
                    1,
                )
            )

        # 4. 字号
        current_size = run_get_font_size_pt(run)
        if current_size != self.font_size:
            diffs.append(
                DIFFResult(
                    "font_size",
                    self.font_size,
                    current_size,
                    f"期待字号{str(self.font_size)};",
                    1,
                )
            )

        # 5. 字体颜色
        current_color = run_get_font_color(run)
        if self.font_color != current_color:
            diffs.append(
                DIFFResult(
                    "font_color",
                    self.font_color,
                    current_color,
                    f"期待字体颜色{str(self.font_color)};",
                    1,
                )
            )

        # 6. 东亚字体
        font_name = run_get_font_name(run)
        if font_name != self.font_name_cn:
            diffs.append(
                DIFFResult(
                    "font_name_cn",
                    self.font_name_cn,
                    font_name,
                    f"期待的中文字体:{str(self.font_name_cn)}",
                    1,
                )
            )
        # 7. 非东亚字体
        ascii_font = run.font.name  # 注意：可能为 None
        if ascii_font != self.font_name_en:
            diffs.append(
                DIFFResult(
                    "font_name_en",
                    self.font_name_en,
                    ascii_font,
                    f"期待的英文字体:{str(self.font_name_en)};",
                    1,
                )
            )

        return sorted(diffs, key=lambda x: x.level)

    def apply_to_run(self, run: Run):
        """将字符样式应用到 docx.Run 对象"""
        diffs = self.diff_from_run(run)
        result = []
        for diff in diffs:
            tmp_str = ""
            match diff.diff_type:
                case "bold":
                    run.bold = diff.expected_value
                    tmp_str = (
                        f"加粗修正，原：{'加粗' if diff.current_value else '非加粗'};"
                    )
                case "italic":
                    run.italic = diff.expected_value
                    tmp_str = (
                        f"斜体修正，原：{'斜体' if diff.current_value else '非斜体'};"
                    )
                case "underline":
                    run.underline = diff.expected_value
                    tmp_str = f"下划线修正，原：{'有下划线' if diff.current_value else '无下划线'};"
                case "font_size":
                    self.font_size.format(docx_obj=run)
                    tmp_str = f"字号修正:{str(self.font_size)};"
                case "font_color":
                    self.font_color.format(docx_obj=run)
                    tmp_str = f"字体颜色修正:{str(self.font_color)};"
                case "font_name_cn":
                    self.font_name_cn.format(docx_obj=run)
                    tmp_str = f"中文字体修正：{str(self.font_name_cn)};"
                case "font_name_en":
                    self.font_name_en.format(docx_obj=run)
                    tmp_str = f"英文字体修正：{str(self.font_name_en)};"
                case _:
                    logger.warning(f"未知的 diff_type: {diff.diff_type}")
            diff.comment = tmp_str
            result.append(diff)

        return result

    @staticmethod
    def to_string(value: list[DIFFResult]) -> str:
        """
        将列表的DIFFResult转换为字符串
        Args:
            value:

        Returns:

        """
        t = []
        for diff in value:
            if style_checks_warning.bold and diff.diff_type == "bold":
                t.append(diff)
            if style_checks_warning.italic and diff.diff_type == "italic":
                t.append(diff)
            if style_checks_warning.underline and diff.diff_type == "underline":
                t.append(diff)
            if style_checks_warning.font_size and diff.diff_type == "font_size":
                t.append(diff)
            if style_checks_warning.font_color and diff.diff_type == "font_color":
                t.append(diff)
            if style_checks_warning.font_name and diff.diff_type in (
                "font_name_cn",
                "font_name_en",
            ):
                t.append(diff)

        return "\n".join([str(i) for i in t])


class ParagraphStyle:
    """段落样式类，用于定义 Word 文档中 Paragraph 级别的排版格式。

    该类封装了常见的段落级格式属性，包括对齐方式、段前/段后间距、行距、首行缩进等，
    常用于格式校验、自动修复或与标准模板进行比对。所有字段均提供合理的默认值，
    符合中文公文或学术论文的常见排版规范。

    Attributes:
        alignment (Alignment): 段落对齐方式。例如左对齐（LEFT）、居中（CENTER）、两端对齐（JUSTIFY）等。
        space_before (Spacing): 段前间距，表示当前段落与上一段之间的垂直距离（单位：磅）。
        space_after (Spacing): 段后间距，表示当前段落与下一段之间的垂直距离（单位：磅）。
        line_spacing (LineSpacing): 行距设置，支持固定值（如单倍、1.5 倍、双倍）或精确磅值。
        first_line_indent (FirstLineIndent): 首行缩进量，通常用于正文段落（如缩进两个汉字）。
        builtin_style_name ():预设样式
    """

    def __init__(
        self,
        alignment: str = "左对齐",
        space_before: str = "0.5行",
        space_after: str = "0.5行",
        line_spacing: str = "1.5倍",
        line_spacingrule: str = "单倍行距",
        first_line_indent: str = "0字符",
        builtin_style_name: str = "正文",
    ):
        self.alignment: Alignment = Alignment(alignment)
        self.space_before: Spacing = Spacing(space_before)
        self.space_after: Spacing = Spacing(space_after)
        self.line_spacing: LineSpacing | float = LineSpacing(line_spacing)
        self.line_spacingrule: LineSpacingRule = LineSpacingRule(line_spacingrule)
        self.first_line_indent: FirstLineIndent = FirstLineIndent(first_line_indent)
        self.builtin_style_name: BuiltInStyle = BuiltInStyle(builtin_style_name)
        if globals()["style_checks_warning"] is None:
            globals()["style_checks_warning"] = get_config().style_checks_warning

    def apply_to_paragraph(self, paragraph: Paragraph) -> list[DIFFResult]:
        """将段落样式应用到 docx.Paragraph 对象，返回样式修正结果"""
        # 先检测当前段落与目标样式的差异
        diffs = self.diff_from_paragraph(paragraph)
        result = []
        # 遍历差异项，逐一项修正并记录修正日志
        for diff in diffs:
            tmp_str = ""
            match diff.diff_type:
                case "alignment":
                    self.alignment.format(docx_obj=paragraph)
                    tmp_str = f"对齐方式修正：{str(self.alignment)};"
                case "space_before":
                    self.space_before.format(docx_obj=paragraph, spacing_type="before")
                    tmp_str = f"段前间距修正：{str(self.space_before)};"
                case "space_after":
                    self.space_after.format(docx_obj=paragraph, spacing_type="after")
                    tmp_str = f"段后间距修正：{str(self.space_after)};"
                case "line_spacing_rule":
                    self.line_spacingrule.format(docx_obj=paragraph)
                    tmp_str = f"间距修正：{str(self.line_spacingrule)};"
                case "line_spacing":
                    self.line_spacing.format(docx_obj=paragraph)
                    tmp_str = f"行距修正：{str(self.line_spacing)};"
                case "first_line_indent":
                    self.first_line_indent.format(docx_obj=paragraph)
                    tmp_str = f"首行缩进修正;{str(self.first_line_indent)};"  # noqa E501
                case "builtin_style_name":
                    self.builtin_style_name.format(docx_obj=paragraph)
                    tmp_str = f"内置样式修正：{str(self.builtin_style_name)};"
                case _:
                    # 替换原异常抛出，改用日志记录未知类型，避免程序中断
                    logger.warning(
                        f"未知的段落样式diff_type: {diff.diff_type}，跳过该样式修正"
                    )
                    continue
            # 更新差异项的评论为修正日志，加入结果列表
            diff.comment = tmp_str
            result.append(diff)
        # 返回所有修正结果，便于外部查看/记录
        return result

    def diff_from_paragraph(self, paragraph: Paragraph) -> list[DIFFResult]:  # noqa C901
        """检查当前段落样式与给定段落样式的差异"""
        if not paragraph:
            return []
        diffs = []
        # 对齐方式
        alignment = paragraph_get_alignment(paragraph)
        alignment = alignment if alignment else WD_ALIGN_PARAGRAPH.LEFT  # 默认左对齐
        if self.alignment != alignment:
            diffs.append(
                DIFFResult(
                    "alignment",
                    self.alignment,
                    alignment,
                    f"对齐方式期待{str(self.alignment)};",
                    0,
                )
            )
        # 段前间距
        unit = self.space_before.rel_unit
        match unit:
            case "hang":
                # 段前间距(行)
                space_before = paragraph_get_space_before(paragraph)
            case "pt":
                space_before = paragraph.paragraph_format.space_before.pt
            case "mm":
                space_before = paragraph.paragraph_format.space_before.mm
            case "cm":
                space_before = paragraph.paragraph_format.space_before.cm
            case "inch":
                space_before = paragraph.paragraph_format.space_before.inches
            case _:
                raise ValueError(f"未知的段前间距单位: {unit}")
        if self.space_before != Spacing(f"{space_before}{self.space_before.unit_ch}"):
            diffs.append(
                DIFFResult(
                    "space_before",
                    self.space_before,
                    space_before,
                    f"段前间距期待{str(self.space_before)};",
                    1,
                )
            )
        # 段后间距(行)
        unit = self.space_after.rel_unit
        match unit:
            case "hang":
                # 段前间距(行)
                space_after = paragraph_get_space_after(paragraph)
            case "pt":
                space_after = paragraph.paragraph_format.space_after.pt
            case "mm":
                space_after = paragraph.paragraph_format.space_after.mm
            case "cm":
                space_after = paragraph.paragraph_format.space_after.cm
            case "inch":
                space_after = paragraph.paragraph_format.space_after.inches
            case _:
                raise ValueError(f"未知的段后间距单位: {unit}")
        if self.space_after != Spacing(f"{space_after}{self.space_after.unit_ch}"):
            diffs.append(
                DIFFResult(
                    "space_after",
                    self.space_after,
                    space_after,
                    f"段后间距期待{str(self.space_after)};",
                    1,
                )
            )
        # 行距选项
        linespacingrule = paragraph.paragraph_format.line_spacing_rule
        linespacingrule = (
            linespacingrule if linespacingrule else WD_LINE_SPACING.MULTIPLE
        )  # 默认单倍行距
        if self.line_spacingrule != linespacingrule:
            diffs.append(
                DIFFResult(
                    "line_spacing_rule",
                    self.line_spacingrule,
                    linespacingrule,
                    f"行距选项期待{str(self.line_spacingrule)};",
                    2,
                )
            )
        # 行距
        line_spacing = paragraph_get_line_spacing(paragraph)
        if self.line_spacing != line_spacing:
            diffs.append(
                DIFFResult(
                    "line_spacing",
                    self.line_spacing,
                    line_spacing,
                    f"行距期待{str(self.line_spacing)};",
                    3,
                )
            )
        # 首行缩进
        unit = self.first_line_indent.rel_unit
        if paragraph.paragraph_format.first_line_indent is None:
            first_line_indent = paragraph_get_first_line_indent(paragraph)
            if first_line_indent is None:  # 无首行缩进
                first_line_indent = float("inf")
        else:
            match unit:
                case "char":
                    # 段前间距(行)
                    first_line_indent = paragraph_get_first_line_indent(paragraph)
                case "pt":
                    first_line_indent = paragraph.paragraph_format.first_line_indent.pt
                case "mm":
                    first_line_indent = paragraph.paragraph_format.first_line_indent.mm
                case "cm":
                    first_line_indent = paragraph.paragraph_format.first_line_indent.cm
                case "inch":
                    first_line_indent = (
                        paragraph.paragraph_format.first_line_indent.inches
                    )
                case _:
                    raise ValueError(f"未知的段前间距单位: {unit}")
        if self.first_line_indent != first_line_indent:
            diffs.append(
                DIFFResult(
                    "first_line_indent",
                    self.first_line_indent,
                    first_line_indent,
                    f"首行缩进期待{str(self.first_line_indent)};",
                    1,
                )
            )
        # 样式
        builtin_style_name = paragraph_get_builtin_style_name(paragraph)
        if self.builtin_style_name != builtin_style_name:
            diffs.append(
                DIFFResult(
                    "builtin_style_name",
                    self.builtin_style_name,
                    builtin_style_name,
                    f"样式期待{str(self.builtin_style_name)};",
                    0,
                )
            )
        return sorted(diffs, key=lambda x: x.level)

    @staticmethod
    def to_string(value: list[DIFFResult]):
        t = []
        for diff in value:
            if style_checks_warning.alignment and diff.diff_type == "alignment":
                t.append(diff)
            if style_checks_warning.space_before and diff.diff_type == "space_before":
                t.append(diff)
            if style_checks_warning.space_after and diff.diff_type == "space_after":
                t.append(diff)
            if style_checks_warning.line_spacing and diff.diff_type == "line_spacing":
                t.append(diff)
            if (
                style_checks_warning.line_spacingrule
                and diff.diff_type == "line_spacingrule"
            ):
                t.append(diff)
            if (
                style_checks_warning.first_line_indent
                and diff.diff_type == "first_line_indent"
            ):
                t.append(diff)
            if (
                style_checks_warning.builtin_style_name
                and diff.diff_type == "builtin_style_name"
            ):
                t.append(diff)
        return "\n".join([str(i) for i in t])
