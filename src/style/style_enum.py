#! /usr/bin/env python
# @Time    : 2026/1/26 10:34
# @Author  : afish
# @File    : style_enmu.py
# from src.settings import CHARACTER_STYLE_CHECKS
from typing import Any

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor


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
        if (
            isinstance(label, int)
            or isinstance(label, float)
            or isinstance(label, tuple)
        ):
            return label
        raise ValueError(
            f"未知段落样式: '{cls.__name__}:{label}'，支持的有: {list(cls._LABEL_MAP.keys())}"
        )

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
            raise ValueError(
                f"未知字体: '{value}'，支持的有: {list(self._LABEL_MAP.keys())}"
            )
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

    @staticmethod
    def to_RGBObject(value: tuple):
        """
        辅助：将 RGB 元组转换为 RGBColor 对象。
        """
        return RGBColor(*value)


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
    CAPTION = "Caption"  # 题注

    _LABEL_MAP = {
        "Heading 1": HEADING_1,
        "Heading 2": HEADING_2,
        "Heading 3": HEADING_3,
        "Heading 4": HEADING_4,
        "正文": NORMAL,
        "标题": TITLE,
        "副标题": SUBTITLE,
        "列表项": LIST_PARAGRAPH,
        "题注": CAPTION,
    }
