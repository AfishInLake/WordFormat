#! /usr/bin/env python
# @Time    : 2026/1/26 10:34
# @Author  : afish
# @File    : style_enmu.py
# from src.settings import CHARACTER_STYLE_CHECKS
import re
from typing import Callable, Optional, Tuple

import webcolors
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from loguru import logger

from wordformat.style.set_some import (
    _SetFirstLineIndent,
    _SetLineSpacing,
    _SetSpacing,
    run_set_font_name,
)
from wordformat.style.utils import extract_unit_from_string


class UnitEnumMeta(type):
    """
    枚举元类：解析Meta类中的单位函数，绑定到枚举类
    """

    def __new__(cls, name: str, bases: tuple, attrs: dict):
        # 1. 提取Meta类中的单位函数映射
        meta_attrs = {}
        if "Meta" in attrs:
            meta_cls = attrs.pop("Meta")
            # 遍历Meta类的属性，收集单位函数（如chars=chars_format_func）
            for attr_name, attr_value in meta_cls.__dict__.items():
                if not attr_name.startswith("_") and callable(attr_value):
                    meta_attrs[attr_name] = attr_value

        # 2. 创建枚举类
        enum_cls = super().__new__(cls, name, bases, attrs)
        # 3. 将Meta中的单位函数绑定到枚举类
        enum_cls._meta_funcs = meta_attrs

        return enum_cls


class UnitLabelEnum(metaclass=UnitEnumMeta):
    """
    带有单位的枚举类
    可以实现自动处理单位问题
    """

    _LABEL_MAP = {}

    def __init__(self, value):
        self.value = value  # 获取的原始值
        self.original_unit = None  # 解析的单位
        self.unit_ch = None  # 中文单位
        self._rel_value = None  # 解析的真实值
        self._rel_unit = None  # 解析的标准单位
        self.extract_unit_result = None  # 解析的结果
        # 对于固定单位的枚举类，如行间距、对齐方式，不需要处理单位
        # 但需要确保rel_value正确设置
        class_name = self.__class__.__name__
        if class_name not in ["LineSpacingRule", "Alignment"]:
            self.split_unit()

    def split_unit(self):
        """
        将带单位的值拆分为数值和单位
        """
        result = self.extract_unit_result = extract_unit_from_string(str(self.value))
        self.original_unit = result.original_unit
        self.unit_ch = result.unit_ch
        self._rel_unit = result.standard_unit
        self._rel_value = result.value

    @property
    def rel_value(self):
        """
        真实值
        优先级：
            UnitResult
            _LABEL_MAP
            原始值
        Returns:
            返回枚举类型真实值
        """
        if self._rel_value is not None:
            return self._rel_value
        # 直接访问类的_LABEL_MAP属性
        if hasattr(self.__class__, "_LABEL_MAP"):
            label_map = self.__class__._LABEL_MAP
            if self.value in label_map:
                self._rel_value = label_map[self.value]
                return self._rel_value
        # 如果没有找到映射，返回原始值
        self._rel_value = self.value
        return self._rel_value

    @rel_value.setter
    def rel_value(self, value):
        self._rel_value = value

    @property
    def rel_unit(self):
        return self._rel_unit

    def base_set(self, docx_obj, **kwargs):
        """
        对于直接设置的属性，应该是直接设置
        示例：
            docx_obj.attr = self.value
            这里赋值原始值
        """
        logger.info(f"{self.__class__.__name__} 没有实现 base_set 方法")

    def function_map(self) -> Optional[Callable]:
        """
        需要子类根据unit返回指定函数
        Returns:
            function 返回一个可迭代对象，由 format 调用
        """
        return self._meta_funcs.get(self._rel_unit, None)

    def format(self, docx_obj: Paragraph | Run, **kwargs):
        """格式化"""
        # 先从meta_funcs中获取对应的函数
        fun = self.function_map()
        # 如果为空就调用子类继承的方法
        if fun is None:
            return self.base_set(docx_obj, **kwargs)
        if isinstance(docx_obj, Paragraph):
            return fun(paragraph=docx_obj, value=self.rel_value, **kwargs)
        else:
            return fun(run=docx_obj, value=self.rel_value, **kwargs)

    def __str__(self):
        return self.value

    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return self.rel_value == other.rel_value
        if isinstance(self.rel_value, str):
            return self.value.lower() == other
        return self.rel_value == other


class FontName(UnitLabelEnum):
    """
    常用中英文字体枚举。
    使用示例：
        font = FontName.SIM_SUN  # '宋体'
        style = ParagraphStyle(font_name=font)
    """

    def is_chinese(self, value: str):
        if value in ["宋体", "黑体", "楷体", "仿宋", "微软雅黑", "汉仪小标宋"]:
            return True
        else:
            return False

    def base_set(self, docx_obj: Run, **kwargs):
        """设置无单位属性"""
        if self.is_chinese(self.value):
            run_set_font_name(run=docx_obj, font_name=self.value)
        else:
            docx_obj.font.name = self.value


class FontSize(UnitLabelEnum):
    """
    常用中文字档字号（单位：磅 / pt）。
    继承 IntEnum 以便直接用于数值比较或传递给 Pt()。

    示例：
        size = FontSize.XIAO_SI  # 12
        run.font.size = Pt(size)
    """

    _LABEL_MAP = {
        "一号": 26,
        "小一": 24,
        "二号": 22,
        "小二": 18,
        "三号": 16,
        "小三": 15,
        "四号": 14,
        "小四": 12,
        "五号": 10.5,
        "小五": 9,
        "六号": 7.5,
        "七号": 5.5,
    }

    def base_set(self, docx_obj: Run, **kwargs):
        """仅作为将字符串转化为数值操作"""
        size = self._LABEL_MAP.get(self.value, None)
        if size:
            docx_obj.font.size = Pt(size)
        else:
            try:
                docx_obj.font.size = Pt(float(self.value))
            except ValueError as e:
                raise ValueError(
                    f"无效的字号: '{self.value}' font_size 必须为数字"
                ) from e


class FontColor(UnitLabelEnum):
    """
    字体颜色处理类（基于webcolors标准色）
    用法：FontColor().base_set(run, color_spec="red")
    支持color_spec类型：
    1. webcolors标准英文色名（red/black/lightblue）
    2. 中文色名（红色/黑色/浅蓝色）
    3. 十六进制色值（#FF0000/#f00/FF0000）
    """

    # 仅保留中文→英文映射（基于webcolors标准）
    _ZH_TO_EN = {
        "黑色": "black",
        "白色": "white",
        "红色": "red",
        "绿色": "green",
        "蓝色": "blue",
        "灰色": "gray",
        "浅灰色": "lightgray",
        "深灰色": "darkgray",
        "橙色": "orange",
        "紫色": "purple",
        "粉色": "pink",
        "棕色": "brown",
        "青色": "cyan",
        "黄色": "yellow",
        "品红": "magenta",
        "浅蓝色": "lightblue",
    }

    @property
    def rel_value(self):
        return self._parse_color(self.value)

    @staticmethod
    def _is_hex(color_spec: str) -> bool:
        """静态方法：校验是否为合法十六进制色值"""
        return bool(
            re.match(r"^#?([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$", color_spec.strip())
        )

    @staticmethod
    def _normalize_hex(hex_str: str) -> str:
        """静态方法：标准化十六进制为6位小写格式"""
        hex_clean = hex_str.strip().lstrip("#").lower()
        if len(hex_clean) == 3:
            hex_clean = "".join([c * 2 for c in hex_clean])
        return "#" + hex_clean.zfill(6).lower()

    @staticmethod
    def _parse_color(color_spec: str) -> Tuple[int, int, int]:
        """静态方法：核心解析逻辑，仅支持webcolors标准色，失败直接抛异常"""
        if not isinstance(color_spec, str):
            raise TypeError(f"颜色标识必须是字符串，当前类型：{type(color_spec)}")

        color_str = color_spec.strip()

        # 步骤1：处理中文名称
        if color_str in FontColor._ZH_TO_EN:
            color_str = FontColor._ZH_TO_EN[color_str]

        # 步骤2：处理十六进制
        if FontColor._is_hex(color_str):
            try:
                normalized_hex = FontColor._normalize_hex(color_str)
                rgb = webcolors.hex_to_rgb(normalized_hex)
                return (rgb.red, rgb.green, rgb.blue)
            except ValueError as e:
                raise ValueError(
                    f"非法十六进制色值：{color_spec}\n"
                    f"错误原因：{str(e)}\n"
                    "支持格式：#RRGGBB / #RGB / RRGGBB（如#FF0000 / #f00 / FF0000）"
                ) from e

        # 步骤3：处理webcolors标准英文名称
        try:
            # 优先CSS标准色，兼容X11扩展色
            rgb = webcolors.name_to_rgb(color_str.lower())
            return (rgb.red, rgb.green, rgb.blue)
        except ValueError:
            try:
                hex_val = webcolors.x11_color_name_to_hex(color_str.lower())
                rgb = webcolors.hex_to_rgb(hex_val)
                return (rgb.red, rgb.green, rgb.blue)
            except ValueError as e:
                raise ValueError(
                    f"不支持的颜色名称：{color_spec}（仅支持webcolors标准色）\n"
                    f"常用示例：{list(FontColor._ZH_TO_EN.keys())} / {list(FontColor._ZH_TO_EN.values())}"
                ) from e

    def base_set(self, docx_obj: Run, **kwargs):
        """
        唯一入口方法：设置字体颜色（符合UnitLabelEnum统一接口）
        :param docx_obj: Run对象（字体颜色载体）
        :raises ValueError/TypeError: 颜色解析失败直接抛出异常
        """

        # 核心逻辑：解析字符串类型的颜色标识（webcolors标准）
        if not isinstance(self.value, str):
            raise TypeError(f"颜色标识仅支持字符串，当前类型：{type(self.value)}")

        # 解析颜色并设置
        rgb_tuple = self._parse_color(self.value)
        docx_obj.font.color.rgb = RGBColor(*rgb_tuple)

    def __eq__(self, other):
        # 必须是元组，且长度必须为三，对应r,g,b
        if not (isinstance(other, tuple) or len(other) != 3):
            return False
        # 从rel_value 获取转化的rgb值
        color_rgb = self.rel_value
        for i in range(3):
            if color_rgb[i] != other[i]:
                return False
        return True


class Alignment(UnitLabelEnum):
    """
    段落对齐方式枚举，兼容 python-docx。

    使用示例：
        style = ParagraphStyle(alignment=Alignment.CENTER)
        # 或直接传给 paragraph.alignment
        paragraph.alignment = Alignment.LEFT.to_docx()
    """

    _LABEL_MAP = {
        "左对齐": WD_ALIGN_PARAGRAPH.LEFT,  # 左侧对齐,
        "居中对齐": WD_ALIGN_PARAGRAPH.CENTER,  # 居中对齐,
        "右对齐": WD_ALIGN_PARAGRAPH.RIGHT,  # 右侧对齐,
        "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY,  # 两端对齐,
        "分散对齐": WD_ALIGN_PARAGRAPH.DISTRIBUTE,  # 分散对齐（较少用）
    }

    def base_set(self, docx_obj: Paragraph, **kwargs):
        """仅把字符串转化为枚举类型操作"""
        alignment = self._LABEL_MAP.get(self.value, None)
        if alignment is not None:  # 必须为None 有可能是枚举类型
            docx_obj.alignment = alignment
        else:
            raise ValueError(f"无效的对齐方式: '{self.value}'")


class Spacing(UnitLabelEnum):
    """
    常用段落间距枚举（单位：磅 / pt）。

    适用于段前（space_before）或段后（space_after）。

    示例：
        style = ParagraphStyle(
            space_before=ParagraphSpacing.NONE,
            space_after=ParagraphSpacing.NORMAL
        )
    """

    class Meta:
        hang = _SetSpacing.set_hang
        pt = _SetSpacing.set_pt
        mm = _SetSpacing.set_mm
        cm = _SetSpacing.set_cm
        inch = _SetSpacing.set_inch


class LineSpacingRule(UnitLabelEnum):
    """
    设置行距选项
    """

    _LABEL_MAP = {
        "单倍行距": WD_LINE_SPACING.SINGLE,
        "1.5倍行距": WD_LINE_SPACING.ONE_POINT_FIVE,
        "2倍行距": WD_LINE_SPACING.DOUBLE,
        "最小值": WD_LINE_SPACING.AT_LEAST,
        "固定值": WD_LINE_SPACING.EXACTLY,
        "多倍行距": WD_LINE_SPACING.MULTIPLE,
    }

    def base_set(self, docx_obj: Paragraph, **kwargs):
        """仅设置倍为单位的数据"""
        line_spacing = self._LABEL_MAP.get(self.value, None)
        if line_spacing:
            docx_obj.paragraph_format.line_spacing_rule = line_spacing
        else:
            raise ValueError(f"无效的行距选项: '{self.value}'")


class LineSpacing(UnitLabelEnum):
    """
    常用行距值，兼容 python-docx。

    支持 倍、磅、英寸、厘米、毫米

    使用示例：
        style = ParagraphStyle(line_spacing=LineSpacing.ONE_POINT_FIVE)
        paragraph.paragraph_format.line_spacing = style.line_spacing
    """

    class Meta:
        pt = _SetLineSpacing.set_pt
        mm = _SetLineSpacing.set_mm
        cm = _SetLineSpacing.set_cm
        inch = _SetLineSpacing.set_inch

    def base_set(self, docx_obj: Paragraph, **kwargs):
        """仅设置倍为单位的数据"""
        line_spacing = self.rel_value
        if line_spacing is not None:  # 必须不为None 有可能是 0
            if line_spacing <= 0:
                line_spacing = 1  # 行距必须大于0 否则会消失
            docx_obj.paragraph_format.line_spacing = line_spacing
        else:
            raise ValueError(f"无效的行距: '{self.value}'")


class FirstLineIndent(UnitLabelEnum):
    """
    首行缩进枚举，适用于中文排版。
    """

    class Meta:
        char = _SetFirstLineIndent.set_char
        pt = _SetFirstLineIndent.set_pt
        mm = _SetFirstLineIndent.set_mm
        cm = _SetFirstLineIndent.set_cm
        inch = _SetFirstLineIndent.set_inch


class BuiltInStyle(UnitLabelEnum):
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

    def base_set(self, docx_obj: Paragraph, **kwargs):
        style = self._LABEL_MAP.get(self.value, None)
        if style:
            try:
                docx_obj.style = style
            except ValueError as e:
                raise ValueError(f"未使用的样式: '{self.value}'") from e
        else:
            try:
                docx_obj.style = self.value
            except ValueError as e:
                raise ValueError(f"未使用的样式: '{self.value}'") from e
