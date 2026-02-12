#! /usr/bin/env python
# @Time    : 2026/1/24 19:47
# @Author  : afish
# @File    : datamodel.py

from typing import Literal, Optional, Union

from pydantic import BaseModel, Field, field_validator

# -------------------------- 基础类型定义 --------------------------
# 对齐方式类型
AlignmentType = Literal["左对齐", "居中对齐", "右对齐", "两端对齐", "分散对齐"]
# 行距类型
LineSpacingRuleType = Literal[
    "单倍行距", "1.5倍行距", "2倍行距", "最小值", "固定值", "多倍行距"
]
# 中文字体类型
ChineseFontType = Literal["宋体", "黑体", "楷体", "仿宋", "微软雅黑", "汉仪小标宋"]
# 英文字体类型
EnglishFontType = Literal[
    "Times New Roman", "Arial", "Calibri", "Courier New", "Helvetica"
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
    alignment: bool = Field(default=True)
    space_before: bool = Field(default=True)
    space_after: bool = Field(default=True)
    line_spacing: bool = Field(default=True)
    line_spacingrule: bool = Field(default=True)
    left_indent: bool = Field(default=True)
    right_indent: bool = Field(default=True)
    first_line_indent: bool = Field(default=True)
    builtin_style_name: bool = Field(default=True)


# -------------------------- 基础配置模型 --------------------------
class GlobalFormatConfig(BaseModel):
    """全局基础格式配置模型"""

    alignment: AlignmentType = Field(default="左对齐")
    space_before: str = Field(default="0.5行", description="段前间距（行）")
    space_after: str = Field(default="0.5行", description="段后间距（行）")
    line_spacingrule: LineSpacingRuleType = Field(default="单倍行距")
    line_spacing: str = Field(default="1.5倍")
    left_indent: str = Field(default="0字符")
    right_indent: str = Field(default="0字符")
    first_line_indent: str = Field(default="2字符")
    builtin_style_name: str = Field(default="正文")
    chinese_font_name: ChineseFontType | str = Field(default="宋体")
    english_font_name: EnglishFontType | str = Field(default="Times New Roman")
    font_size: FontSizeType = Field(default="小四")  # 修正类型和默认值
    font_color: str = Field(default="黑色")
    bold: bool = Field(default=False)
    italic: bool = Field(default=False)
    underline: bool = Field(default=False)


# -------------------------- 摘要配置模型 --------------------------
class KeywordsConfig(GlobalFormatConfig):
    """关键词配置模型（继承全局格式）"""

    section_title_re: str = Field(default=None, description="关键词匹配的正则表达式")

    # 关键词特有配置
    keywords_bold: bool = Field(default=True, description="关键字加粗")
    count_min: int = Field(default=4, description="最小关键字数")
    count_max: int = Field(default=4, description="最大关键字数")
    trailing_punct_forbidden: bool = Field(default=True, description="禁止最后有标点")

    @field_validator("count_min", "count_max")
    def validate_keyword_count(cls, v):
        """验证关键词数量为正整数"""
        if v <= 0:
            raise ValueError(f"关键词数量 {v} 必须大于0")
        return v


class AbstractTitleConfig(GlobalFormatConfig):
    """摘要标题配置（继承全局格式）"""

    section_title_re: str = Field(description="标题正则表达式")


class AbstractContentConfig(GlobalFormatConfig):
    """摘要正文配置（继承全局格式）"""


class AbstractChineseConfig(GlobalFormatConfig):
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
class HeadingLevelConfig(GlobalFormatConfig):
    """各级标题配置（继承全局格式）"""

    section_title_re: str = Field(description="标题正则表达式")


class HeadingsConfig(GlobalFormatConfig):
    """标题总配置"""

    level_1: HeadingLevelConfig = Field(default_factory=HeadingLevelConfig)
    level_2: HeadingLevelConfig = Field(default_factory=HeadingLevelConfig)
    level_3: HeadingLevelConfig = Field(default_factory=HeadingLevelConfig)


# -------------------------- 正文配置模型 --------------------------
class BodyTextConfig(GlobalFormatConfig):
    """正文配置（继承全局格式）"""


# -------------------------- 插图配置模型 --------------------------
class FiguresConfig(GlobalFormatConfig):
    """插图配置"""

    section_title_re: str = Field(description="图注正则表达式")

    caption_position: Literal["above", "below"] = Field(
        default="below", description="图注位置"
    )
    caption_prefix: Optional[str] = Field(default="图", description="图注编号前缀")


# -------------------------- 表格配置模型 --------------------------
class TablesConfig(GlobalFormatConfig):
    """表格配置"""

    section_title_re: str = Field(description="图注正则表达式")

    caption_position: Literal["above", "below"] = Field(
        default="above", description="表注位置"
    )
    caption_prefix: Optional[str] = Field(default="表", description="表注编号前缀")


# -------------------------- 参考文献配置模型 --------------------------
class ReferencesTitleConfig(GlobalFormatConfig):
    """参考文献标题配置（继承全局格式）"""

    section_title_re: str = Field(description="参考文献正则表达式")
    section_title: Optional[str] = Field(
        default="参考文献", description="参考文献章节标题"
    )


class ReferencesContentConfig(GlobalFormatConfig):
    """参考文献内容配置（继承全局格式）"""

    numbering_format: Optional[str] = Field(default=None, description="编号格式")
    entry_indent: Optional[float] = Field(default=0.0, description="条目首行缩进量")
    entry_ending_punct: Optional[str] = Field(default=None, description="条目结束标点")


class ReferencesConfig(BaseModel):
    """参考文献总配置"""

    title: ReferencesTitleConfig = Field(default_factory=ReferencesTitleConfig)
    content: ReferencesContentConfig = Field(default_factory=ReferencesContentConfig)


# -------------------------- 致谢配置模型 --------------------------
class AcknowledgementsTitleConfig(GlobalFormatConfig):
    """致谢标题配置（继承全局格式）"""

    section_title_re: str = Field(description="致谢标题正则表达式")


class AcknowledgementsContentConfig(GlobalFormatConfig):
    """致谢内容配置（继承全局格式）"""


class AcknowledgementsConfig(GlobalFormatConfig):
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
