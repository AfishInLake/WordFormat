#! /usr/bin/env python
# @Time    : 2026/1/24 19:47
# @Author  : afish
# @File    : datamodel.py

from typing import Any, ClassVar, Literal, Optional, Tuple, Union

from pydantic import BaseModel, Field, field_validator

from wordformat.style.style_enum import FontColor, FontName, FontSize

# -------------------------- 基础类型定义 --------------------------
# 对齐方式类型
AlignmentType = Literal["左对齐", "居中对齐", "右对齐", "两端对齐", "分散对齐"]
# 行距类型
LineSpacingType = Literal["单倍行距", "1.5倍", "双倍"]
# 首行缩进类型
FirstLineIndentType = Literal["无缩进", "1字符", "2字符", "3字符"]
# 中文字体类型
ChineseFontType = Literal["宋体", "黑体", "楷体", "仿宋", "微软雅黑", "汉仪小标宋"]
# 英文字体类型
EnglishFontType = Literal[
    "Times New Roman", "Arial", "Calibri", "Courier New", "Helvetica"
]
# 字体颜色类型
FontColorType = Literal[
    "BLACK",
    "WHITE",
    "RED",
    "GREEN",
    "BLUE",
    "GRAY",
    "DARK_GRAY",
    "LIGHT_GRAY",
    "ORANGE",
    "PURPLE",
    "BROWN",
]
# 字号类型（兼容字符串和数值）
FontSizeType = Union[
    Literal[
        "一号",
        "小一",
        "二号",
        "小二",
        "三号",
        "小三",
        "四号",
        "小四",
        "五号",
        "小五",
        "六号",
        "七号",
    ],
    float,
    int,
]


# -------------------------- 预警字段配置模型 --------------------------


class WarningFieldConfig(BaseModel):
    """预警字段配置模型"""

    bold: bool = Field(default=True)
    italic: bool = Field(default=True)
    underline: bool = Field(default=True)
    font_size: bool = Field(default=True)
    font_name: bool = Field(default=False)
    font_color: bool = Field(default=False)


# -------------------------- 基础配置模型 --------------------------
class GlobalFormatConfig(BaseModel):
    """全局基础格式配置模型"""

    alignment: AlignmentType = Field(default="左对齐")
    space_before: float = Field(default=0.5, description="段前间距（行）")
    space_after: float = Field(default=0.5, description="段后间距（行）")
    line_spacing: LineSpacingType = Field(default="1.5倍")
    first_line_indent: FirstLineIndentType = Field(default="2字符")
    builtin_style_name: str = Field(default="正文")
    chinese_font_name: ChineseFontType = Field(default="宋体")
    english_font_name: EnglishFontType = Field(default="Times New Roman")
    font_size: FontSizeType = Field(default="小四")  # 修正类型和默认值
    font_color: Union[FontColorType, Tuple[int, int, int]] = Field(default="BLACK")
    bold: bool = Field(default=False)
    italic: bool = Field(default=False)
    underline: bool = Field(default=False)

    @field_validator("font_size")
    def validate_font_size(cls, v: str | float) -> str | float:
        """验证字号：兼容字符串（小四/五号）和数值，最终转换为磅值"""
        # 1. 如果是字符串（如"小四"），通过 FontSize 枚举转换为数值
        if isinstance(v, str):
            try:
                return FontSize.from_label(v)
            except ValueError as e:
                raise ValueError(
                    f"无效的字号：{v}，支持的有：{list(FontSize._LABEL_MAP.keys())}"
                ) from e
        # 2. 如果是数值，验证范围
        elif isinstance(v, (float, int)):
            if not (5.5 <= v <= 36):  # 七号(5.5pt) 到 一号(26pt)
                raise ValueError(f"字号数值 {v} 超出合理范围（5.5-36pt）")
            return float(v)
        # 3. 其他类型不支持
        else:
            raise ValueError(
                f"字号类型错误：{type(v)}，仅支持字符串（如小四）或数值（如12）"
            )

    @field_validator("font_color")
    def validate_font_color(cls, v: str | tuple) -> tuple:
        """验证字体颜色：兼容字符串（BLACK/RED）和RGB元组"""
        # 1. 如果是字符串（如"BLACK"），通过 FontColor 枚举转换为RGB元组
        if isinstance(v, str):
            try:
                return FontColor.from_label(v)
            except ValueError as e:
                raise ValueError(
                    f"无效的颜色：{v}，支持的有：{list(FontColor._LABEL_MAP.keys())}"
                ) from e
        # 2. 如果是RGB元组，验证范围
        elif isinstance(v, tuple):
            if len(v) != 3:
                raise ValueError(f"RGB颜色元组必须包含3个值，当前：{v}")
            for val in v:
                if not (0 <= val <= 255):
                    raise ValueError(f"RGB颜色值 {val} 超出范围（0-255）")
            return v
        # 3. 其他类型不支持
        else:
            raise ValueError(
                f"颜色类型错误：{type(v)}，仅支持字符串（如BLACK）或RGB元组（如(0,0,0)）"
            )

    @field_validator("chinese_font_name", "english_font_name")
    def validate_font_name(cls, v):
        """验证字体名称是否合法"""
        try:
            FontName.from_label(v)
            return v
        except ValueError as e:
            raise ValueError(
                f"无效的字体名称：{v}，支持的有：{list(FontName._LABEL_MAP.keys())}"
            ) from e


# -------------------------- 摘要配置模型 --------------------------
class KeywordsConfig(BaseModel):
    """关键词配置模型（继承全局格式）"""

    # 继承全局格式的所有字段
    alignment: AlignmentType = Field(default="左对齐")
    space_before: float = Field(default=0.5)
    space_after: float = Field(default=0.5)
    line_spacing: LineSpacingType = Field(default="1.5倍")
    first_line_indent: FirstLineIndentType = Field(default="2字符")
    builtin_style_name: str = Field(default="正文")
    chinese_font_name: ChineseFontType = Field(default="宋体")
    english_font_name: EnglishFontType = Field(default="Times New Roman")
    font_size: FontSizeType = Field(default="小四")  # 修正类型
    font_color: Union[FontColorType, Tuple[int, int, int]] = Field(default="BLACK")
    bold: bool = Field(default=False)
    italic: bool = Field(default=False)
    underline: bool = Field(default=False)
    section_title_re: str = Field(default=None, description="关键词匹配的正则表达式")

    # 关键词特有配置
    kewords_bold: bool = Field(default=True, description="关键字加粗")
    count_min: int = Field(default=4, description="最小关键字数")
    count_max: int = Field(default=4, description="最大关键字数")
    trailing_punct_forbidden: bool = Field(default=True, description="禁止最后有标点")

    # 复用字号/字体/颜色验证器
    validate_font_size: ClassVar[Any] = GlobalFormatConfig.validate_font_size
    validate_font_color: ClassVar[Any] = GlobalFormatConfig.validate_font_color
    validate_font_name: ClassVar[Any] = GlobalFormatConfig.validate_font_name

    @field_validator("count_min", "count_max")
    def validate_keyword_count(cls, v):
        """验证关键词数量为正整数"""
        if v <= 0:
            raise ValueError(f"关键词数量 {v} 必须大于0")
        return v


class AbstractTitleConfig(BaseModel):
    """摘要标题配置（继承全局格式）"""

    alignment: AlignmentType = Field(default="居中对齐")
    space_before: float = Field(default=0.5)
    space_after: float = Field(default=0.5)
    line_spacing: LineSpacingType = Field(default="1.5倍")
    first_line_indent: FirstLineIndentType = Field(default="无缩进")
    builtin_style_name: str = Field(default="正文")
    chinese_font_name: ChineseFontType = Field(default="黑体")
    english_font_name: EnglishFontType = Field(default="Times New Roman")
    font_size: FontSizeType = Field(default="小四")  # 修正类型
    font_color: Union[FontColorType, Tuple[int, int, int]] = Field(default="BLACK")
    bold: bool = Field(default=True)
    italic: bool = Field(default=False)
    underline: bool = Field(default=False)
    section_title_re: str = Field(description="标题正则表达式")

    # 复用验证器
    validate_font_size: ClassVar[Any] = GlobalFormatConfig.validate_font_size
    validate_font_color: ClassVar[Any] = GlobalFormatConfig.validate_font_color
    validate_font_name: ClassVar[Any] = GlobalFormatConfig.validate_font_name


class AbstractContentConfig(BaseModel):
    """摘要正文配置（继承全局格式）"""

    alignment: AlignmentType = Field(default="左对齐")
    space_before: float = Field(default=0.5)
    space_after: float = Field(default=0.5)
    line_spacing: LineSpacingType = Field(default="1.5倍")
    first_line_indent: FirstLineIndentType = Field(default="2字符")
    builtin_style_name: str = Field(default="正文")
    chinese_font_name: ChineseFontType = Field(default="宋体")
    english_font_name: EnglishFontType = Field(default="Times New Roman")
    font_size: FontSizeType = Field(default="小四")  # 修正类型
    font_color: Union[FontColorType, Tuple[int, int, int]] = Field(default="BLACK")
    bold: bool = Field(default=False)
    italic: bool = Field(default=False)
    underline: bool = Field(default=False)

    # 复用验证器
    validate_font_size: ClassVar[Any] = GlobalFormatConfig.validate_font_size
    validate_font_color: ClassVar[Any] = GlobalFormatConfig.validate_font_color
    validate_font_name: ClassVar[Any] = GlobalFormatConfig.validate_font_name


class AbstractChineseConfig(BaseModel):
    """中文摘要配置"""

    chinese_title: AbstractTitleConfig = Field(default_factory=AbstractTitleConfig)
    chinese_content: AbstractContentConfig = Field(
        default_factory=AbstractContentConfig
    )


class AbstractEnglishConfig(BaseModel):
    """英文摘要配置"""

    english_title: AbstractTitleConfig = Field(default_factory=AbstractTitleConfig)
    english_content: AbstractContentConfig = Field(
        default_factory=AbstractContentConfig
    )


class AbstractConfig(BaseModel):
    """摘要总配置"""

    chinese: AbstractChineseConfig = Field(default_factory=AbstractChineseConfig)
    english: AbstractEnglishConfig = Field(default_factory=AbstractEnglishConfig)
    keywords: dict[str, KeywordsConfig] = Field(
        default_factory=lambda: {
            "english": KeywordsConfig(),
            "chinese": KeywordsConfig(),
        }
    )


# -------------------------- 标题配置模型 --------------------------
class HeadingLevelConfig(BaseModel):
    """各级标题配置（继承全局格式）"""

    alignment: AlignmentType = Field(default="左对齐")
    space_before: float = Field(default=0.5)
    space_after: float = Field(default=0.5)
    line_spacing: LineSpacingType = Field(default="1.5倍")
    first_line_indent: FirstLineIndentType = Field(default="无缩进")
    builtin_style_name: str = Field(default="正文")
    chinese_font_name: ChineseFontType = Field(default="宋体")
    english_font_name: EnglishFontType = Field(default="Times New Roman")
    font_size: FontSizeType = Field(default="小四")  # 修正类型
    font_color: Union[FontColorType, Tuple[int, int, int]] = Field(default="BLACK")
    bold: bool = Field(default=False)
    italic: bool = Field(default=False)
    underline: bool = Field(default=False)
    section_title_re: str = Field(description="标题正则表达式")

    # 复用验证器
    validate_font_size: ClassVar[Any] = GlobalFormatConfig.validate_font_size
    validate_font_color: ClassVar[Any] = GlobalFormatConfig.validate_font_color
    validate_font_name: ClassVar[Any] = GlobalFormatConfig.validate_font_name


class HeadingsConfig(BaseModel):
    """标题总配置"""

    level_1: HeadingLevelConfig = Field(default_factory=HeadingLevelConfig)
    level_2: HeadingLevelConfig = Field(default_factory=HeadingLevelConfig)
    level_3: HeadingLevelConfig = Field(default_factory=HeadingLevelConfig)


# -------------------------- 正文配置模型 --------------------------
class BodyTextConfig(BaseModel):
    """正文配置（继承全局格式）"""

    alignment: AlignmentType = Field(default="两端对齐")
    space_before: float = Field(default=0.0)
    space_after: float = Field(default=0.0)
    line_spacing: LineSpacingType = Field(default="1.5倍")
    first_line_indent: FirstLineIndentType = Field(default="2字符")
    builtin_style_name: str = Field(default="正文")
    chinese_font_name: ChineseFontType = Field(default="宋体")
    english_font_name: EnglishFontType = Field(default="Times New Roman")
    font_size: FontSizeType = Field(default="小四")  # 修正类型
    font_color: Union[FontColorType, Tuple[int, int, int]] = Field(default="BLACK")
    bold: bool = Field(default=False)
    italic: bool = Field(default=False)
    underline: bool = Field(default=False)

    # 复用验证器
    validate_font_size: ClassVar[Any] = GlobalFormatConfig.validate_font_size
    validate_font_color: ClassVar[Any] = GlobalFormatConfig.validate_font_color
    validate_font_name: ClassVar[Any] = GlobalFormatConfig.validate_font_name


# -------------------------- 插图配置模型 --------------------------
class FiguresConfig(BaseModel):
    """插图配置"""

    alignment: AlignmentType = Field(default="左对齐")
    space_before: float = Field(default=0.5)
    space_after: float = Field(default=0.5)
    line_spacing: LineSpacingType = Field(default="1.5倍")
    first_line_indent: FirstLineIndentType = Field(default="无缩进")
    builtin_style_name: str = Field(default="正文")
    chinese_font_name: ChineseFontType = Field(default="宋体")
    english_font_name: EnglishFontType = Field(default="Times New Roman")
    font_size: FontSizeType = Field(default="小四")  # 修正类型
    font_color: Union[FontColorType, Tuple[int, int, int]] = Field(default="BLACK")
    bold: bool = Field(default=False)
    italic: bool = Field(default=False)
    underline: bool = Field(default=False)
    section_title_re: str = Field(description="图注正则表达式")

    caption_position: Literal["above", "below"] = Field(
        default="below", description="图注位置"
    )
    caption_prefix: Optional[str] = Field(default="图", description="图注编号前缀")

    # 复用字号验证器
    validate_font_size: ClassVar[Any] = GlobalFormatConfig.validate_font_size
    validate_font_color: ClassVar[Any] = GlobalFormatConfig.validate_font_color
    validate_font_name: ClassVar[Any] = GlobalFormatConfig.validate_font_name


# -------------------------- 表格配置模型 --------------------------
class TablesConfig(BaseModel):
    """表格配置"""

    alignment: AlignmentType = Field(default="左对齐")
    space_before: float = Field(default=0.5)
    space_after: float = Field(default=0.5)
    line_spacing: LineSpacingType = Field(default="1.5倍")
    first_line_indent: FirstLineIndentType = Field(default="无缩进")
    builtin_style_name: str = Field(default="正文")
    chinese_font_name: ChineseFontType = Field(default="宋体")
    english_font_name: EnglishFontType = Field(default="Times New Roman")
    font_size: FontSizeType = Field(default="小四")  # 修正类型
    font_color: Union[FontColorType, Tuple[int, int, int]] = Field(default="BLACK")
    bold: bool = Field(default=False)
    italic: bool = Field(default=False)
    underline: bool = Field(default=False)
    section_title_re: str = Field(description="图注正则表达式")

    caption_position: Literal["above", "below"] = Field(
        default="above", description="表注位置"
    )
    caption_prefix: Optional[str] = Field(default="表", description="表注编号前缀")

    # 复用验证器
    validate_font_size: ClassVar[Any] = GlobalFormatConfig.validate_font_size
    validate_font_color: ClassVar[Any] = GlobalFormatConfig.validate_font_color
    validate_font_name: ClassVar[Any] = GlobalFormatConfig.validate_font_name


# -------------------------- 参考文献配置模型 --------------------------
class ReferencesTitleConfig(BaseModel):
    """参考文献标题配置（继承全局格式）"""

    alignment: AlignmentType = Field(default="左对齐")
    space_before: float = Field(default=0.5)
    space_after: float = Field(default=0.5)
    line_spacing: LineSpacingType = Field(default="1.5倍")
    first_line_indent: FirstLineIndentType = Field(default="2字符")
    builtin_style_name: str = Field(default="正文")
    chinese_font_name: ChineseFontType = Field(default="宋体")
    english_font_name: EnglishFontType = Field(default="Times New Roman")
    font_size: FontSizeType = Field(default="小四")  # 修正类型
    font_color: Union[FontColorType, Tuple[int, int, int]] = Field(default="BLACK")
    bold: bool = Field(default=False)
    italic: bool = Field(default=False)
    underline: bool = Field(default=False)
    section_title_re: str = Field(description="参考文献正则表达式")

    section_title: Optional[str] = Field(
        default="参考文献", description="参考文献章节标题"
    )

    # 复用验证器
    validate_font_size: ClassVar[Any] = GlobalFormatConfig.validate_font_size
    validate_font_color: ClassVar[Any] = GlobalFormatConfig.validate_font_color
    validate_font_name: ClassVar[Any] = GlobalFormatConfig.validate_font_name


class ReferencesContentConfig(BaseModel):
    """参考文献内容配置（继承全局格式）"""

    alignment: AlignmentType = Field(default="左对齐")
    space_before: float = Field(default=0.5)
    space_after: float = Field(default=0.5)
    line_spacing: LineSpacingType = Field(default="1.5倍")
    first_line_indent: FirstLineIndentType = Field(default="2字符")
    builtin_style_name: str = Field(default="正文")
    chinese_font_name: ChineseFontType = Field(default="宋体")
    english_font_name: EnglishFontType = Field(default="Times New Roman")
    font_size: FontSizeType = Field(default="小四")  # 修正类型
    font_color: Union[FontColorType, Tuple[int, int, int]] = Field(default="BLACK")
    bold: bool = Field(default=False)
    italic: bool = Field(default=False)
    underline: bool = Field(default=False)

    numbering_format: Optional[str] = Field(default=None, description="编号格式")
    entry_indent: Optional[float] = Field(default=0.0, description="条目首行缩进量")
    entry_ending_punct: Optional[str] = Field(default=None, description="条目结束标点")

    # 复用验证器
    validate_font_size: ClassVar[Any] = GlobalFormatConfig.validate_font_size
    validate_font_color: ClassVar[Any] = GlobalFormatConfig.validate_font_color
    validate_font_name: ClassVar[Any] = GlobalFormatConfig.validate_font_name


class ReferencesConfig(BaseModel):
    """参考文献总配置"""

    title: ReferencesTitleConfig = Field(default_factory=ReferencesTitleConfig)
    content: ReferencesContentConfig = Field(default_factory=ReferencesContentConfig)


# -------------------------- 致谢配置模型 --------------------------
class AcknowledgementsTitleConfig(BaseModel):
    """致谢标题配置（继承全局格式）"""

    alignment: AlignmentType = Field(default="左对齐")
    space_before: float = Field(default=0.5)
    space_after: float = Field(default=0.5)
    line_spacing: LineSpacingType = Field(default="1.5倍")
    first_line_indent: FirstLineIndentType = Field(default="2字符")
    builtin_style_name: str = Field(default="正文")
    chinese_font_name: ChineseFontType = Field(default="宋体")
    english_font_name: EnglishFontType = Field(default="Times New Roman")
    font_size: FontSizeType = Field(default="小四")  # 修正类型
    font_color: Union[FontColorType, Tuple[int, int, int]] = Field(default="BLACK")
    bold: bool = Field(default=False)
    italic: bool = Field(default=False)
    underline: bool = Field(default=False)
    section_title_re: str = Field(description="致谢标题正则表达式")

    # 复用验证器
    validate_font_size: ClassVar[Any] = GlobalFormatConfig.validate_font_size
    validate_font_color: ClassVar[Any] = GlobalFormatConfig.validate_font_color
    validate_font_name: ClassVar[Any] = GlobalFormatConfig.validate_font_name


class AcknowledgementsContentConfig(BaseModel):
    """致谢内容配置（继承全局格式）"""

    alignment: AlignmentType = Field(default="左对齐")
    space_before: float = Field(default=0.5)
    space_after: float = Field(default=0.5)
    line_spacing: LineSpacingType = Field(default="1.5倍")
    first_line_indent: FirstLineIndentType = Field(default="2字符")
    builtin_style_name: str = Field(default="正文")
    chinese_font_name: ChineseFontType = Field(default="宋体")
    english_font_name: EnglishFontType = Field(default="Times New Roman")
    font_size: FontSizeType = Field(default="小四")  # 修正类型
    font_color: Union[FontColorType, Tuple[int, int, int]] = Field(default="BLACK")
    bold: bool = Field(default=False)
    italic: bool = Field(default=False)
    underline: bool = Field(default=False)

    # 复用验证器
    validate_font_size: ClassVar[Any] = GlobalFormatConfig.validate_font_size
    validate_font_color: ClassVar[Any] = GlobalFormatConfig.validate_font_color
    validate_font_name: ClassVar[Any] = GlobalFormatConfig.validate_font_name


class AcknowledgementsConfig(BaseModel):
    """致谢总配置"""

    title: AcknowledgementsTitleConfig = Field(
        default_factory=AcknowledgementsTitleConfig
    )
    content: AcknowledgementsContentConfig = Field(
        default_factory=AcknowledgementsContentConfig
    )


# -------------------------- 根配置模型 --------------------------
class NodeConfigRoot(BaseModel):
    """配置根节点模型"""

    style_checks_warning: WarningFieldConfig = Field(default_factory=WarningFieldConfig)
    global_format: GlobalFormatConfig = Field(default_factory=GlobalFormatConfig)
    abstract: AbstractConfig = Field(default_factory=AbstractConfig)
    headings: HeadingsConfig = Field(default_factory=HeadingsConfig)
    body_text: BodyTextConfig = Field(default_factory=BodyTextConfig)
    figures: FiguresConfig = Field(default_factory=FiguresConfig)
    tables: TablesConfig = Field(default_factory=TablesConfig)
    references: ReferencesConfig = Field(default_factory=ReferencesConfig)
    acknowledgements: AcknowledgementsConfig = Field(
        default_factory=AcknowledgementsConfig
    )
