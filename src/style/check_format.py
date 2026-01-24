#! /usr/bin/env python
# @Time    : 2026/1/12 10:46
# @Author  : afish
# @File    : style.py
from dataclasses import dataclass
from typing import Any

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from src.settings import CHARACTER_STYLE_CHECKS
from src.style.get_some import (
    paragraph_get_alignment,
    paragraph_get_builtin_style_name,
    paragraph_get_first_line_indent,
    paragraph_get_line_spacing,
    paragraph_get_space_after,
    paragraph_get_space_before,
    run_get_font_bold,
    run_get_font_color,
    run_get_font_italic,
    run_get_font_name,
    run_get_font_size,
    run_get_font_underline,
)

for check in ["bold", "italic", "underline", "font_size", "font_color", "font_name"]:
    if check not in CHARACTER_STYLE_CHECKS:
        raise ValueError(f" 未在 `settings.py` 的 CHARACTER_STYLE_CHECKS 配置中找到 {check}")


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

    def __str__(self):
        return self.comment


class LabelEnum:
    _LABEL_MAP = {}

    @classmethod
    def __init_subclass__(cls, **kwargs):
        """
        子类初始化钩子：自动生成【值->标签】的反向映射
        """
        super().__init_subclass__(**kwargs)
        if hasattr(cls, "_LABEL_MAP") and cls._LABEL_MAP:
            cls._VALUE_TO_LABEL = {}
            for label, value in cls._LABEL_MAP.items():
                # 若值是枚举成员（如FontSize.YI_HAO），取其值；否则直接用值
                real_value = value.value if hasattr(value, "value") else value
                if real_value not in cls._VALUE_TO_LABEL:
                    cls._VALUE_TO_LABEL[real_value] = label

    @classmethod
    def from_label(cls, label: Any) -> int | float | str | tuple:
        # 检查配置是否有映射
        if label in cls._LABEL_MAP:
            return cls._LABEL_MAP[label]
        # 检查配置是否是类成员
        if isinstance(label, str):
            if label.isupper() and not label.startswith("_"):  # 只允许如 "BLACK"
                if hasattr(cls, label):
                    value = getattr(cls, label)
                    if not callable(value):  # 排除方法
                        return value
        # 检查是否是int, float, tuple三类数据结构
        if isinstance(label, int) or isinstance(label, float) or isinstance(label, tuple):
            return label
        raise ValueError(f"未知段落样式: '{cls.__name__}:{label}'，支持的有: {list(cls._LABEL_MAP.keys())}")

    @classmethod
    def to_string(cls, value: Any) -> str:
        """通用to_string:优先用反向映射，无规则格式化输出"""
        real_value = value.value if hasattr(value, "value") else value

        # 1. 优先从反向映射找标签
        if real_value in cls._VALUE_TO_LABEL:
            return cls._VALUE_TO_LABEL[real_value]

        # 2. 特殊值格式化（如倍数、RGB元组）
        if isinstance(real_value, float) and cls.__name__ == "LineSpacing":
            return f"{real_value}倍"  # 行距特殊格式化
        if isinstance(real_value, tuple) and len(real_value) == 3:
            return cls._rgb_to_name(real_value)  # 颜色RGB转名称（可选）

        # 3. 默认返回字符串
        return str(real_value)

    @staticmethod
    def _rgb_to_name(rgb: tuple[int, int, int]) -> str:
        """辅助：RGB元组转颜色名称（可选扩展）"""
        color_map = {
            (0, 0, 0): "黑色",
            (255, 255, 255): "白色",
            (255, 0, 0): "红色",
            (0, 128, 0): "绿色",
            (0, 0, 255): "蓝色",
            (128, 128, 128): "灰色",
        }
        return color_map.get(rgb, f"RGB{rgb}")


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

    @classmethod
    def to_string(cls, value: Any) -> str:
        real_value = value.value if hasattr(value, "value") else value
        if isinstance(real_value, tuple) and len(real_value) == 3:
            cn, ascii_, high_ansi = real_value
            return f"中文字体: {cn}, ASCII: {ascii_}, High ANSI: {high_ansi}"
        return str(real_value)


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
        "一号": YI_HAO,
        "小一": XIAO_YI,
        "二号": ER_HAO,
        "小二": XIAO_ER,
        "三号": SAN_HAO,
        "小三": XIAO_SAN,
        "四号": SI_HAO,
        "小四": XIAO_SI,
        "五号": WU_HAO,
        "小五": XIAO_WU,
        "六号": LIU_HAO,
        "七号": QI_HAO,
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
        "左对齐": LEFT,
        "居中对齐": CENTER,
        "右对齐": RIGHT,
        "两端对齐": JUSTIFY,
        "分散对齐": DISTRIBUTE,
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
        "NONE": NONE,
        "TINY": TINY,
        "SMALL": SMALL,
        "HALF_LINE": HALF_LINE,
        "NORMAL": NORMAL,
        "MEDIUM": MEDIUM,
        "LARGE": LARGE,
        "EXTRA_LARGE": EXTRA_LARGE,
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
        "单倍行距": SINGLE,
        "1.5倍": ONE_POINT_FIVE,
        "双倍": DOUBLE,
    }


class FirstLineIndent(LabelEnum):
    """
    首行缩进枚举，适用于中文排版。
    """

    NONE = 0  # 无缩进（用于标题、列表等）
    ONE_CHAR = 1  # 1 字符（约）
    TWO_CHARS = 2  # 2 字符（标准中文正文）
    THREE_CHARS = 3  # 3 字符（较少用）

    _LABEL_MAP = {
        "无缩进": NONE,
        "1字符": ONE_CHAR,
        "2字符": TWO_CHARS,
        "3字符": THREE_CHARS,
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
        "Heading 1": HEADING_1,
        "Heading 2": HEADING_2,
        "Heading 3": HEADING_3,
        "Heading 4": HEADING_4,
        "正文": NORMAL,
        "标题": TITLE,
        "副标题": SUBTITLE,
        "列表项": LIST_PARAGRAPH,
    }


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
        self.font_name_cn: str = FontName.from_label(font_name_cn)
        self.font_name_en: str = FontName.from_label(font_name_en)
        self.font_size: float = FontSize.from_label(font_size)
        self.font_color: tuple = FontColor.from_label(font_color)
        self.bold: bool = bold
        self.italic: bool = italic
        self.underline: bool = underline

    def diff_from_run(self, run) -> list[DIFFResult]:
        """
        检查段落样式和指定样式是否一致
        """
        diffs = []

        # 1. 加粗
        bold = run_get_font_bold(run)
        if bold != self.bold:
            if CHARACTER_STYLE_CHECKS["bold"]:
                diffs.append(
                    DIFFResult(
                        "bold",
                        self.bold,
                        bold,
                        f"期待{'加粗' if self.bold else '不加粗'}，实际{'加粗' if bold else '不加粗'};",
                    )
                )
        # 2. 斜体
        italic = run_get_font_italic(run)
        if italic != self.italic:
            if CHARACTER_STYLE_CHECKS["italic"]:
                diffs.append(
                    DIFFResult(
                        "italic",
                        self.italic,
                        italic,
                        f"期待{'斜体' if self.italic else '非斜体'}，实际{'斜体' if italic else '非斜体'};",
                    )
                )

        # 3. 下划线
        underline = run_get_font_underline(run)
        if underline != self.underline:
            if CHARACTER_STYLE_CHECKS["underline"]:
                diffs.append(
                    DIFFResult(
                        "underline",
                        self.underline,
                        underline,
                        f"期待{'有下划线' if self.underline else '无下划线'}，实际{'有下划线' if underline else '无下划线'};",
                    )
                )

        # 4. 字号
        current_size = run_get_font_size(run)
        if current_size != self.font_size:
            if CHARACTER_STYLE_CHECKS["font_size"]:
                diffs.append(
                    DIFFResult(
                        "font_size",
                        self.font_size,
                        current_size,
                        f"期待字号{FontSize.to_string(self.font_size)}，实际字号{FontSize.to_string(current_size)};",
                    )
                )

        # 5. 字体颜色
        current_color = run_get_font_color(run)
        if current_color != self.font_color:
            if CHARACTER_STYLE_CHECKS["font_color"]:
                diffs.append(
                    DIFFResult(
                        "font_color",
                        self.font_color,
                        current_color,
                        f"期待字体颜色{FontColor.to_string(self.font_color)}, 实际字体颜色{FontColor.to_string(current_color)};",
                    )
                )

        # 6. 字体
        font_name = run_get_font_name(run)
        except_font = (self.font_name_cn, self.font_name_en, self.font_name_en)
        if font_name != except_font:
            if CHARACTER_STYLE_CHECKS["font_name"]:
                diffs.append(
                    DIFFResult(
                        "font_name_cn",
                        self.font_name_cn,
                        font_name,
                        f"期待的字体:{FontName.to_string(except_font)}，实际的字体:{FontName.to_string(font_name)}",
                    )
                )
        return diffs

    def apply_to(self, run: Run):
        """将字符样式应用到 docx.Run 对象"""
        diffs = self.diff_from_run(run)
        setters = {
            "bold": lambda v: setattr(run, "bold", v),
            "italic": lambda v: setattr(run, "italic", v),
            "underline": lambda v: setattr(run, "underline", v),
            "font_size": lambda v: setattr(run.font, "size", Pt(v)),
            "font_color": lambda v: setattr(run.font.color, "rgb", RGBColor(*v)),
            "font_name_en": lambda v: self._set_run_font_en(run, v),
            "font_name_cn": lambda v: self._set_run_font_cn(run, v),
        }

        for attr, current, expected in diffs:
            if attr in setters:
                setters[attr](expected)


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

    def __init__(
        self,
        alignment: str = "左对齐",
        space_before: float = 0.0,
        space_after: float = 0.0,
        line_spacing: float = 1.5,
        first_line_indent: float = 0.0,
        builtin_style_name: str = "正文",
    ):
        self.alignment: tuple = Alignment.from_label(alignment)
        self.space_before: float = float(Spacing.from_label(space_before))
        self.space_after: float = float(Spacing.from_label(space_after))
        self.line_spacing: LineSpacing | float = LineSpacing.from_label(line_spacing)
        self.first_line_indent: FirstLineIndent | float = FirstLineIndent.from_label(first_line_indent)
        self.builtin_style_name: str = BuiltInStyle.from_label(builtin_style_name)

    def apply_to(self, paragraph: Paragraph):
        """将段落样式应用到 docx.Paragraph 对象"""
        paragraph.style = self.builtin_style_name
        pf = paragraph.paragraph_format
        pf.alignment = self.alignment
        pf.space_before = Pt(self.space_before)
        pf.space_after = Pt(self.space_after)
        pf.line_spacing = self.line_spacing
        # TODO:首行缩进量不准确，待确认
        pf.first_line_indent = self.first_line_indent * paragraph.style.font.size

    def diff_from_paragraph(self, paragraph: Paragraph) -> list[DIFFResult]:
        """检查当前段落样式与给定段落样式的差异"""
        if not paragraph:
            return []
        diffs = []
        # 对齐方式
        alignment = paragraph_get_alignment(paragraph)
        if self.alignment != alignment:
            diffs.append(
                DIFFResult(
                    "alignment",
                    self.alignment,
                    alignment,
                    f"对齐方式期待{Alignment.to_string(self.alignment)}实际{Alignment.to_string(alignment)};",
                )
            )
        # 段前间距(行)
        space_before = paragraph_get_space_before(paragraph)
        if self.space_before != space_before:
            diffs.append(
                DIFFResult(
                    "space_before",
                    self.space_before,
                    space_before,
                    f"段前间距期待{self.space_before}行，实际{space_before}行;",
                )
            )
        # 段后间距(行)
        space_after = paragraph_get_space_after(paragraph)
        if self.space_after != space_after:
            diffs.append(
                DIFFResult(
                    "space_after",
                    self.space_after,
                    space_after,
                    f"段后间距期待{self.space_before}行，实际{space_before}行;",
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
                    f"行距期待{LineSpacing.to_string(self.line_spacing)}，实际{LineSpacing.to_string(line_spacing)};",
                )
            )
        # 首行缩进
        first_line_indent = paragraph_get_first_line_indent(paragraph)
        if self.first_line_indent != first_line_indent:
            diffs.append(
                DIFFResult(
                    "first_line_indent",
                    self.first_line_indent,
                    first_line_indent,
                    f"首行缩进期待{FirstLineIndent.to_string(self.first_line_indent)}字符，实际"
                    f"{FirstLineIndent.to_string(first_line_indent)}字符;",
                )
            )
        # 样式
        builtin_style_name = paragraph_get_builtin_style_name(paragraph)
        if self.builtin_style_name.lower() != builtin_style_name:
            diffs.append(
                DIFFResult(
                    "builtin_style_name",
                    self.builtin_style_name,
                    builtin_style_name,
                    f"样式期待{BuiltInStyle.to_string(self.builtin_style_name)}样式，实际{BuiltInStyle.to_string(builtin_style_name)}样式;",
                )
            )
        return diffs
