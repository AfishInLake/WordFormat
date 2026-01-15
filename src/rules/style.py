#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/12 10:46
# @Author  : afish
# @File    : style.py
from dataclasses import dataclass
from typing import Union, Any, Optional, Tuple, List

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm
from docx.text.run import Run


class LabelEnum:
    _LABEL_MAP = {}

    @classmethod
    def from_label(cls, label: Any) -> Union[int, float, str, tuple]:
        # 检查配置是否有映射
        if label in cls._LABEL_MAP:
            return cls._LABEL_MAP[label]
        # 检查配置是否是类成员
        if isinstance(label, str):
            if label.isupper() and not label.startswith('_'):  # 只允许如 "BLACK"
                if hasattr(cls, label):
                    value = getattr(cls, label)
                    if not callable(value):  # 排除方法
                        return value
        # 检查是否是int, float, tuple三类数据结构
        if isinstance(label, int) or isinstance(label, float) or isinstance(label, tuple):
            return label
        raise ValueError(f"未知段落样式: '{label}'，支持的有: {list(cls._LABEL_MAP.keys())}")


class FontName(LabelEnum):
    """
    常用中英文字体枚举。
    使用示例：
        font = FontName.SIM_SUN  # '宋体'
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

    _LABEL_MAP = {
        "宋体": SIM_SUN,
        "黑体": SIM_HEI,
        "楷体": KAI_TI,
        "仿宋": FANG_SONG,
        "微软雅黑": MICROSOFT_YAHEI,
        "汉仪小标宋": HAN_YI_XIAO_BIAO_SONG,
        "Times New Roman": TIMES_NEW_ROMAN,
        "Arial": ARIAL,
        "Calibri": CALIBRI,
        "Courier New": COURIER_NEW,
        "Helvetica": HELVETICA,
    }

    def is_chinese(self, value: str):
        if value not in self._LABEL_MAP:
            raise ValueError(f"未知字体: '{value}'，支持的有: {list(self._LABEL_MAP.keys())}")
        if value in ["宋体", "黑体", "楷体", "仿宋", "微软雅黑", "汉仪小标宋"]:
            return True
        else:
            return False


class FontSize(LabelEnum):
    """
    常用中文字档字号（单位：磅 / pt）。
    继承 IntEnum 以便直接用于数值比较或传递给 Pt()。

    示例：
        size = FontSize.XIAO_SI  # 12
        run.font.size = Pt(size)
    """
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

    _LABEL_MAP = {
        '一号': YI_HAO,
        '小一': XIAO_YI,
        '二号': ER_HAO,
        '小二': XIAO_ER,
        '三号': SAN_HAO,
        '小三': XIAO_SAN,
        '四号': SI_HAO,
        '小四': XIAO_SI,
        '五号': WU_HAO,
        '小五': XIAO_WU,
        '六号': LIU_HAO,
        '七号': QI_HAO,
    }


class FontColor(LabelEnum):
    """
    常用字体颜色（RGB 元组）。

    示例：
        color = FontColor.BLACK  # (0, 0, 0)
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


class Alignment(LabelEnum):
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

    _LABEL_MAP = {
        '左对齐': LEFT,
        '居中对齐': CENTER,
        '右对齐': RIGHT,
        '两端对齐': JUSTIFY,
        '分散对齐': DISTRIBUTE,
    }


class Spacing(LabelEnum):
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

    _LABEL_MAP = {
        'NONE': NONE,
        'TINY': TINY,
        'SMALL': SMALL,
        'HALF_LINE': HALF_LINE,
        'NORMAL': NORMAL,
        'MEDIUM': MEDIUM,
        'LARGE': LARGE,
        'EXTRA_LARGE': EXTRA_LARGE,
    }


class LineSpacing(LabelEnum):
    """
    常用行距枚举（倍数制），兼容 python-docx。

    注意：此枚举仅适用于“倍数行距”（single, 1.5, double），
    不包含固定值（如 exactly 20pt）——若需支持固定值，建议单独处理。

    使用示例：
        style = ParagraphStyle(line_spacing=LineSpacing.ONE_POINT_FIVE)
        paragraph.paragraph_format.line_spacing = style.line_spacing
    """
    SINGLE = 1.0  # 单倍行距
    ONE_POINT_FIVE = 1.5  # 1.5 倍
    DOUBLE = 2.0  # 双倍

    _LABEL_MAP = {
        'SINGLE': SINGLE,
        'ONE_POINT_FIVE': ONE_POINT_FIVE,
        'DOUBLE': DOUBLE,
    }


class FirstLineIndent(LabelEnum):
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

    _LABEL_MAP = {
        'NONE': NONE,
        'ONE_CHAR': ONE_CHAR,
        'TWO_CHARS': TWO_CHARS,
        'THREE_CHARS': THREE_CHARS,
    }


class BuiltInStyle(LabelEnum):
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

    _LABEL_MAP = {
        'Heading 1': HEADING_1,
        "Heading 2": HEADING_2,
        "Heading 3": HEADING_3,
        "Heading 4": HEADING_4,
        '正文': NORMAL,
        '标题': TITLE,
        '副标题': SUBTITLE,
        '列表项': LIST_PARAGRAPH,
    }


@dataclass
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
            font_name_cn: str = '宋体',
            font_name_en: str = 'Times New Roman',
            font_size: Union[str, float] = '小四',
            font_color: Union[str, tuple] = 'BLACK',
            bold: bool = False,
            italic: bool = False,
            underline: bool = False
    ):
        self.font_name_cn: str = FontName.from_label(font_name_cn)
        self.font_name_en: str = FontName.from_label(font_name_en)
        self.font_size: float = FontSize.from_label(font_size)
        self.font_color: tuple = FontColor.from_label(font_color)
        self.bold: bool = bold
        self.italic: bool = italic
        self.underline: bool = underline

    def _get_run_font_name_cn(self, run) -> Optional[str]:
        """从 run 中提取中文字体 (w:eastAsia)"""
        rPr = run._element.rPr
        if rPr is not None and rPr.rFonts is not None:
            return rPr.rFonts.get(qn('w:eastAsia'))
        return None

    def _get_run_font_name_en(self, run) -> Optional[str]:
        """从 run 中提取英文字体 (w:ascii 或 w:hAnsi)"""
        rPr = run._element.rPr
        if rPr is not None and rPr.rFonts is not None:
            ascii_font = rPr.rFonts.get(qn('w:ascii'))
            hansi_font = rPr.rFonts.get(qn('w:hAnsi'))
            return ascii_font or hansi_font
        return None

    def _get_run_font_color(self, run) -> Optional[tuple]:
        """从 run 中提取字体颜色 RGB 元组"""
        color = run.font.color
        if color and color.rgb:
            rgb = color.rgb
            return (rgb[0], rgb[1], rgb[2])  # RGBColor 是 bytes-like，转为 tuple
        return None

    def _get_run_font_size_pt(self, run) -> Optional[float]:
        """从 run 中提取字号（单位：磅）"""
        size = run.font.size
        if size is not None:
            return size.pt  # Pt 对象有 .pt 属性
        return None

    def diff_from_run(self, run) -> List[Tuple[str, Any, Any]]:
        """
        比较当前 CharacterStyle 与给定 run 的实际格式。

        返回一个列表，每个元素为 (属性名, run当前值, 期望值)，
        仅包含不一致的属性。

        属性名使用内部字段名（如 'font_name_cn', 'bold' 等）。
        """
        diffs = []

        # 1. 加粗
        if run.bold != self.bold:
            diffs.append(('bold', run.bold, self.bold))

        # 2. 斜体
        if run.italic != self.italic:
            diffs.append(('italic', run.italic, self.italic))

        # 3. 下划线
        if run.underline != self.underline:
            diffs.append(('underline', run.underline, self.underline))

        # 4. 字号
        current_size = self._get_run_font_size_pt(run)
        if current_size != self.font_size:
            diffs.append(('font_size', current_size, self.font_size))

        # 5. 字体颜色
        current_color = self._get_run_font_color(run)
        if current_color != self.font_color:
            diffs.append(('font_color', current_color, self.font_color))

        # 6. 中文字体
        current_cn = self._get_run_font_name_cn(run)
        if current_cn != self.font_name_cn:
            diffs.append(('font_name_cn', current_cn, self.font_name_cn))

        # 7. 英文字体
        current_en = self._get_run_font_name_en(run)
        if current_en != self.font_name_en:
            diffs.append(('font_name_en', current_en, self.font_name_en))

        return diffs

    def apply_to(self, run: Run):
        """将字符样式应用到 docx.Run 对象"""
        diffs = self.diff_from_run(run)
        setters = {
            'bold': lambda v: setattr(run, 'bold', v),
            'italic': lambda v: setattr(run, 'italic', v),
            'underline': lambda v: setattr(run, 'underline', v),
            'font_size': lambda v: setattr(run.font, 'size', Pt(v)),
            'font_color': lambda v: setattr(run.font.color, 'rgb', RGBColor(*v)),
            'font_name_en': lambda v: self._set_run_font_en(run, v),
            'font_name_cn': lambda v: self._set_run_font_cn(run, v),
        }

        for attr, current, expected in diffs:
            if attr in setters:
                setters[attr](expected)

    def _set_run_font_en(self, run, font_name: str):
        """
        设置 英文 字体
        """
        rPr = run._element.rPr
        if rPr.rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        rPr.rFonts.set(qn('w:ascii'), font_name)
        rPr.rFonts.set(qn('w:hAnsi'), font_name)

    def _set_run_font_cn(self, run, font_name: str):
        """
        设置 中文 字体
        """
        rPr = run._element.rPr
        if rPr.rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        rPr.rFonts.set(qn('w:eastAsia'), font_name)


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

    def __init__(self,
                 alignment: str = '左对齐',
                 space_before: float = .0,
                 space_after: float = .0,
                 line_spacing: float = 1.5,
                 first_line_indent: float = .0,
                 builtin_style_name: str = '正文'
                 ):
        self.alignment: tuple = Alignment.from_label(alignment)
        self.space_before: float = float(Spacing.from_label(space_before))
        self.space_after: float = float(Spacing.from_label(space_after))
        self.line_spacing: Union[LineSpacing, float] = LineSpacing.from_label(line_spacing)
        self.first_line_indent: Union[FirstLineIndent, float] = FirstLineIndent.from_label(first_line_indent)
        self.builtin_style_name: str = BuiltInStyle.from_label(builtin_style_name)

    def apply_to(self, paragraph):
        """将段落样式应用到 docx.Paragraph 对象"""
        paragraph.style = self.builtin_style_name
        pf = paragraph.paragraph_format
        pf.alignment = self.alignment
        pf.space_before = Pt(self.space_before)
        pf.space_after = Pt(self.space_after)
        pf.line_spacing = self.line_spacing
        # TODO:首行缩进量不准确，待修复
        pf.first_line_indent = Cm(self.first_line_indent)
