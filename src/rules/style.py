#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/12 10:46
# @Author  : afish
# @File    : style.py
from dataclasses import dataclass
from enum import Enum, IntEnum
from typing import Union

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm
from docx.oxml import OxmlElement
from .utils import _is_chinese_font_name


class FontName(Enum):
    """
    常用中英文字体枚举。
    使用示例：
        font = FontName.SIM_SUN.value  # '宋体'
        style = ParagraphStyle(font_name=font)
    """
    # 中文字体
    SIM_SUN = "宋体"
    SIM_HEI = "黑体"
    KAI_TI = "楷体"
    FANG_SONG = "仿宋"
    MICROSOFT_YAHEI = "微软雅黑"
    HAN_YI_XIAO_BIAO_SONG = "汉仪小标宋"

    # 英文字体
    TIMES_NEW_ROMAN = "Times New Roman"
    ARIAL = "Arial"
    CALIBRI = "Calibri"
    COURIER_NEW = "Courier New"
    HELVETICA = "Helvetica"

    # 通用别名（可选）
    DEFAULT_CHINESE = "宋体"
    DEFAULT_ENGLISH = "Times New Roman"


class FontSize(IntEnum):
    """
    常用中文字档字号（单位：磅 / pt）。
    继承 IntEnum 以便直接用于数值比较或传递给 Pt()。

    示例：
        size = FontSize.XIAO_SI  # 12
        run.font.size = Pt(size)
    """
    BAO_BIAO = 22  # 报表/大标题
    YI_HAO = 26  # 一号
    XIAO_YI = 24  # 小一
    ER_HAO = 22  # 二号
    XIAO_ER = 18  # 小二
    SAN_HAO = 16  # 三号
    XIAO_SAN = 15  # 小三
    SI_HAO = 14  # 四号
    XIAO_SI = 12  # 小四（最常用正文）
    WU_HAO = 10.5  # 五号（注意：五号是 10.5pt）
    XIAO_WU = 9  # 小五
    LIU_HAO = 7.5  # 六号
    QI_HAO = 5.5  # 七号（极少用）

    # 别名（英文/通用）
    TITLE = 18  # 标题常用
    SUBTITLE = 16  # 副标题
    BODY = 12  # 正文（等同 XIAO_SI）
    FOOTNOTE = 9  # 脚注（等同 XIAO_WU）


class FontColor(Enum):
    """
    常用字体颜色（RGB 元组）。

    示例：
        color = FontColor.BLACK.value  # (0, 0, 0)
        run.font.color.rgb = RGBColor(*color)
    """
    BLACK = (0, 0, 0)
    WHITE = (255, 255, 255)
    RED = (255, 0, 0)
    GREEN = (0, 128, 0)  # 深绿（Word 默认绿色）
    BLUE = (0, 0, 255)
    GRAY = (128, 128, 128)
    DARK_GRAY = (64, 64, 64)
    LIGHT_GRAY = (192, 192, 192)
    ORANGE = (255, 165, 0)
    PURPLE = (128, 0, 128)
    BROWN = (165, 42, 42)

    # 中文公文常用色
    OFFICIAL_RED = (204, 0, 0)  # 公文红头常用色（如“中共中央文件”）
    LINK_BLUE = (0, 102, 204)  # 超链接蓝色


class Alignment(Enum):
    """
    段落对齐方式枚举，兼容 python-docx。

    使用示例：
        style = ParagraphStyle(alignment=Alignment.CENTER)
        # 或直接传给 paragraph.alignment
        paragraph.alignment = Alignment.LEFT.to_docx()
    """
    LEFT = WD_ALIGN_PARAGRAPH.LEFT  # 左侧对齐
    CENTER = WD_ALIGN_PARAGRAPH.CENTER  # 居中对齐
    RIGHT = WD_ALIGN_PARAGRAPH.RIGHT  # 右侧对齐
    JUSTIFY = WD_ALIGN_PARAGRAPH.JUSTIFY  # 两端对齐
    DISTRIBUTE = WD_ALIGN_PARAGRAPH.DISTRIBUTE  # 分散对齐（较少用）


class Spacing(Enum):
    """
       常用段落间距枚举（单位：磅 / pt）。

       适用于段前（space_before）或段后（space_after）。

       示例：
           style = ParagraphStyle(
               space_before=ParagraphSpacing.NONE,
               space_after=ParagraphSpacing.NORMAL
           )
       """
    NONE = 0  # 无间距
    TINY = 3  # 微小间距（如列表项）
    SMALL = 6  # 小间距（常见于正文段落之间）
    HALF_LINE = 9  # 半行间距
    NORMAL = 12  # 标准段后间距（中文公文常用）
    MEDIUM = 18  # 中等间距（章节分隔）
    LARGE = 24  # 大间距（标题前后）
    EXTRA_LARGE = 36  # 超大间距（封面、分章）


class LineSpacing(Enum):
    """
    常用行距枚举（倍数制），兼容 python-docx。

    注意：此枚举仅适用于“倍数行距”（single, 1.5, double），
    不包含固定值（如 exactly 20pt）——若需支持固定值，建议单独处理。

    使用示例：
        style = ParagraphStyle(line_spacing=LineSpacing.ONE_POINT_FIVE)
        paragraph.paragraph_format.line_spacing = style.line_spacing.value
    """
    SINGLE = 1.0  # 单倍行距 (1.0)
    ONE_POINT_FIVE = 1.5  # 1.5 倍
    DOUBLE = 2.0  # 双倍 (2.0)


class FirstLineIndent(Enum):
    """
    首行缩进枚举（单位：厘米 / cm），适用于中文排版。

    使用示例：
        style = ParagraphStyle(first_line_indent=FirstLineIndent.TWO_CHARS)
        paragraph.paragraph_format.first_line_indent = style.first_line_indent.to_pt()
    """
    NONE = 0  # 无缩进（用于标题、列表等）
    ONE_CHAR = 1  # 1 字符（约）
    TWO_CHARS = 2  # 2 字符（标准中文正文）
    THREE_CHARS = 3  # 3 字符（较少用）


class BuiltInStyle(Enum):
    """
    Word 内置段落样式名称（使用英文标准名称，跨语言兼容）。

    注意：这些名称是 python-docx 和 Word API 的标准名称，
    即使文档界面显示为“标题 1”，实际样式名仍是 "Heading 1"。
    """
    HEADING_1 = "Heading 1"
    HEADING_2 = "Heading 2"
    HEADING_3 = "Heading 3"
    HEADING_4 = "Heading 4"
    NORMAL = "Normal"  # 正文
    TITLE = "Title"
    SUBTITLE = "Subtitle"
    LIST_PARAGRAPH = "List Paragraph"


@dataclass
class CharacterStyle:
    """字符样式类，用于定义 Word 文档中 Run 级别的文本格式。

    该类封装了常见的字符级格式属性，如字体名称、字号、颜色、加粗、斜体、下划线等，
    通常用于格式校验、自动修复或样式比对。所有字段均有默认值，符合中文文档常见排版规范。

    Attributes:
        font_name (FontName): 字体名称（如黑体、宋体）。注意：中文字体需通过 `w:eastAsia` 属性设置。
        font_size (FontSize): 字号（如小四、四号等），内部以磅（pt）为单位存储。
        font_color (FontColor): 字体颜色，默认为黑色（RGB(0, 0, 0)）。
        bold (bool): 是否加粗。True 表示加粗，False 表示不加粗。
        italic (bool): 是否斜体。True 表示斜体，False 表示非斜体。
        underline (bool): 是否带下划线。True 表示有下划线，False 表示无下划线。
    """
    font_name: Union[FontName, str] = FontName.SIM_SUN
    font_size: FontSize = FontSize.XIAO_SI
    font_color: FontColor = FontColor.BLACK
    bold: bool = False
    italic: bool = False
    underline: bool = False

    def apply_to(self, run):
        """将字符样式应用到 docx.Run 对象"""
        run.bold = self.bold
        run.italic = self.italic
        run.underline = self.underline
        run.font.size = Pt(self.font_size.value)
        run.font.color.rgb = RGBColor(*self.font_color.value)

        # 获取实际字体名称字符串
        if isinstance(self.font_name, FontName):
            font_str = self.font_name.value
        else:
            font_str = str(self.font_name)
        # 判断是否为中文字体（根据是否包含中文字符）
        is_chinese = _is_chinese_font_name(font_str)
        # 设置中文字体（关键！）
        r = run._element
        rPr = r.rPr
        if rPr.rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        if is_chinese:
            # 中文字体：重点设置 eastAsia
            rPr.rFonts.set(qn('w:eastAsia'), font_str)
            # 西文部分建议使用标准英文字体（避免中文字体渲染英文难看）
            rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
            rPr.rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        else:
            # 纯英文字体
            rPr.rFonts.set(qn('w:ascii'), font_str)
            rPr.rFonts.set(qn('w:hAnsi'), font_str)
            # eastAsia 设为默认中文字体，用于混排时的中文显示
            rPr.rFonts.set(qn('w:eastAsia'), '宋体')


@dataclass
class ParagraphStyle:
    """段落样式类，用于定义 Word 文档中 Paragraph 级别的排版格式。

    该类封装了常见的段落级格式属性，包括对齐方式、段前/段后间距、行距、首行缩进等，
    常用于格式校验、自动修复或与标准模板进行比对。所有字段均提供合理的默认值，
    符合中文公文或学术论文的常见排版规范。

    Attributes:
        alignment (Alignment): 段落对齐方式。例如左对齐（LEFT）、居中（CENTER）、两端对齐（JUSTIFY）等。
        space_before (Spacing): 段前间距，表示当前段落与上一段之间的垂直距离（单位：磅）。默认为无间距（NONE）。
        space_after (Spacing): 段后间距，表示当前段落与下一段之间的垂直距离（单位：磅）。默认为无间距（NONE）。
        line_spacing (LineSpacing): 行距设置，支持固定值（如单倍、1.5 倍、双倍）或精确磅值。默认为 1.5 倍行距。
        first_line_indent (FirstLineIndent): 首行缩进量，通常用于正文段落（如缩进两个汉字）。标题类段落一般设为 NONE。
        builtin_style_name ():预设样式
    """
    alignment: Alignment = Alignment.LEFT
    space_before: Spacing = Spacing.NONE
    space_after: Spacing = Spacing.NONE
    line_spacing: LineSpacing = LineSpacing.ONE_POINT_FIVE
    first_line_indent: FirstLineIndent = FirstLineIndent.NONE
    builtin_style_name: BuiltInStyle = BuiltInStyle.NORMAL

    def apply_to(self, paragraph):
        """将段落样式应用到 docx.Paragraph 对象"""
        paragraph.style = self.builtin_style_name.value
        pf = paragraph.paragraph_format
        pf.alignment = self.alignment.value
        pf.space_before = Pt(self.space_before.value)
        pf.space_after = Pt(self.space_after.value)
        pf.line_spacing = self.line_spacing.value
        # TODO:首行缩进量不准确，待修复
        pf.first_line_indent = Cm(self.first_line_indent.value)
