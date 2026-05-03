#! /usr/bin/env python
# @Time    : 2026/1/24 19:47
# @Author  : afish
# @File    : datamodel.py

from typing import Literal, Optional, Union

from pydantic import BaseModel, Field, field_validator, model_validator

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


@field_validator("font_size")
@classmethod
def validate_font_size(cls, v):
    """验证字号为正数"""
    if isinstance(v, (int, float)) and v <= 0:
        raise ValueError(f"字号 {v} 必须大于0")
    return v


# -------------------------- 预警字段配置模型 --------------------------


class WarningFieldConfig(BaseModel):
    """是否显示预警字段配置模型"""

    bold: bool = Field(default=True, description="加粗")
    italic: bool = Field(default=True, description="斜体")
    underline: bool = Field(default=True, description="下划线")
    font_size: bool = Field(default=True, description="字号")
    font_name: bool = Field(default=False, description="字体名称")
    font_color: bool = Field(default=False, description="字体颜色")
    alignment: bool = Field(default=True, description="对齐方式")
    space_before: bool = Field(default=True, description="段前间距")
    space_after: bool = Field(default=True, description="段后间距")
    line_spacing: bool = Field(default=True, description="行距")
    line_spacingrule: bool = Field(default=True, description="行距类型")
    left_indent: bool = Field(default=True, description="文本之前")
    right_indent: bool = Field(default=True, description="文本之后")
    first_line_indent: bool = Field(default=True, description="段落首行缩进")
    builtin_style_name: bool = Field(default=True, description="内置样式名称")


# -------------------------- 基础配置模型 --------------------------
class GlobalFormatConfig(BaseModel):
    """全局基础格式配置模型"""

    alignment: AlignmentType = Field(default="左对齐", description="段落对齐方式")
    space_before: str = Field(default="0.5行", description="段前间距（行）")
    space_after: str = Field(default="0.5行", description="段后间距（行）")
    line_spacingrule: LineSpacingRuleType = Field(
        default="单倍行距", description="行距类型"
    )
    line_spacing: str = Field(default="1.5倍", description="行距参数")
    left_indent: str = Field(default="0字符", description="文本之前")
    right_indent: str = Field(default="0字符", description="文本之后")
    first_line_indent: str = Field(default="2字符", description="段落首行缩进")
    builtin_style_name: str = Field(default="正文", description="内置样式名称")
    chinese_font_name: ChineseFontType | str = Field(
        default="宋体", description="中文字体名称"
    )
    english_font_name: EnglishFontType | str = Field(
        default="Times New Roman", description="英文字体名称"
    )
    font_size: FontSizeType = Field(
        default="小四", description="字号"
    )  # 修正类型和默认值
    font_color: str = Field(default="黑色", description="字体颜色")

    @field_validator("font_color")
    @classmethod
    def validate_font_color(cls, v):
        """验证字体颜色为合法值"""
        if not v or not isinstance(v, str):
            raise ValueError(f"字体颜色不能为空")
        return v
    bold: bool = Field(default=False, description="加粗")
    italic: bool = Field(default=False, description="斜体")
    underline: bool = Field(default=False, description="下划线")


# -------------------------- 摘要配置模型 --------------------------
class KeywordsConfig(GlobalFormatConfig):
    """关键词配置模型（继承全局格式）"""

    # 关键词特有配置
    keywords_bold: bool = Field(default=True, description="关键字加粗")
    count_min: int = Field(default=4, description="最小关键字数")
    count_max: int = Field(default=4, description="最大关键字数")
    trailing_punct_forbidden: bool = Field(default=True, description="禁止最后有标点")

    @field_validator("count_min", "count_max")
    @classmethod
    def validate_keyword_count(cls, v, info):
        """验证关键词数量为正整数"""
        if v <= 0:
            raise ValueError(f"关键词数量 {v} 必须大于0")
        return v

    @model_validator(mode="after")
    def validate_count_range(self) -> "KeywordsConfig":
        """验证 count_min <= count_max"""
        if self.count_min > self.count_max:
            raise ValueError(
                f"count_min({self.count_min}) 不能大于 count_max({self.count_max})"
            )
        return self


class AbstractTitleConfig(GlobalFormatConfig):
    """摘要标题配置（继承全局格式）"""


class AbstractContentConfig(GlobalFormatConfig):
    """摘要正文配置（继承全局格式）"""


class AbstractChineseConfig(BaseModel):
    """中文摘要配置"""

    chinese_title: AbstractTitleConfig = Field(
        default_factory=AbstractTitleConfig, description="中文标题配置"
    )
    chinese_content: AbstractContentConfig = Field(
        default_factory=AbstractContentConfig,
        description="中文正文配置",
    )


class AbstractEnglishConfig(BaseModel):
    """英文摘要配置"""

    english_title: AbstractTitleConfig = Field(
        default_factory=AbstractTitleConfig, description="英文标题配置"
    )
    english_content: AbstractContentConfig = Field(
        default_factory=AbstractContentConfig,
        description="英文正文配置",
    )


class AbstractConfig(BaseModel):
    """摘要总配置"""

    chinese: AbstractChineseConfig = Field(
        default_factory=AbstractChineseConfig, description="中文摘要配置"
    )
    english: AbstractEnglishConfig = Field(
        default_factory=AbstractEnglishConfig, description="英文摘要配置"
    )
    keywords: dict[str, KeywordsConfig] = Field(
        default_factory=lambda: {
            "english": KeywordsConfig(),
            "chinese": KeywordsConfig(),
        },
        description="关键词配置",
    )


# -------------------------- 标题配置模型 --------------------------
class HeadingLevelConfig(GlobalFormatConfig):
    """各级标题配置（继承全局格式）"""


class HeadingsConfig(BaseModel):
    """标题总配置"""

    level_1: HeadingLevelConfig = Field(
        default_factory=HeadingLevelConfig, description="一级标题"
    )
    level_2: HeadingLevelConfig = Field(
        default_factory=HeadingLevelConfig, description="二级标题"
    )
    level_3: HeadingLevelConfig = Field(
        default_factory=HeadingLevelConfig, description="三级标题"
    )


# -------------------------- 正文配置模型 --------------------------
class BodyTextConfig(GlobalFormatConfig):
    """正文配置（继承全局格式）"""


# -------------------------- 插图配置模型 --------------------------
class FiguresConfig(GlobalFormatConfig):
    """插图配置"""

    caption_position: Literal["above", "below"] = Field(
        default="below", description="图注位置"
    )
    caption_prefix: Optional[str] = Field(default="图", description="图注编号前缀")


# -------------------------- 表格配置模型 --------------------------
class TablesConfig(GlobalFormatConfig):
    """表格配置"""

    caption_position: Literal["above", "below"] = Field(
        default="above", description="表注位置"
    )
    caption_prefix: Optional[str] = Field(default="表", description="表注编号前缀")


# -------------------------- 参考文献配置模型 --------------------------
class ReferencesTitleConfig(GlobalFormatConfig):
    """参考文献标题配置（继承全局格式）"""

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

    title: ReferencesTitleConfig = Field(
        default_factory=ReferencesTitleConfig, description="参考文献标题配置"
    )
    content: ReferencesContentConfig = Field(
        default_factory=ReferencesContentConfig, description="参考文献内容配置"
    )


# -------------------------- 致谢配置模型 --------------------------
class AcknowledgementsTitleConfig(GlobalFormatConfig):
    """致谢标题配置（继承全局格式）"""


class AcknowledgementsContentConfig(GlobalFormatConfig):
    """致谢内容配置（继承全局格式）"""


class AcknowledgementsConfig(BaseModel):
    """致谢总配置"""

    title: AcknowledgementsTitleConfig = Field(
        default_factory=AcknowledgementsTitleConfig,
        description="致谢标题配置",
    )
    content: AcknowledgementsContentConfig = Field(
        default_factory=AcknowledgementsContentConfig,
        description="致谢内容配置",
    )


# -------------------------- 编号配置模型 --------------------------
class NumberingLevelConfig(BaseModel):
    """单级标题编号配置"""

    enabled: bool = Field(default=False, description="是否启用该级别自动编号")
    # 编号模板，%1=本级序号，%2=上级序号，%3=上上级序号
    # 例如："第%1章"、"%1.%2"、"%1.%2.%3"、"%1)"
    template: Optional[str] = Field(default=None, description="编号模板")
    # 清除手动编号的正则表达式，仅作用于该级别标题段落的开头
    # 例如："^第[一二三四五六七八九十百千零]+章\\s*"
    strip_pattern: Optional[str] = Field(
        default=None, description="清除手动编号的正则表达式"
    )
    # 编号之后：tab（制表符）、space（空格）、nothing（无）
    suffix: Optional[str] = Field(
        default="space",
        description="编号之后的分隔符：tab/space/nothing",
    )
    # 编号缩进：编号距左边距的距离，如 "0.75cm"、"420磅"、"0字符"
    numbering_indent: Optional[str] = Field(
        default=None,
        description="编号缩进，如 '0.75cm'、'420磅'、'0字符'",
    )
    # 文本缩进：文本距左边距的距离（即悬挂缩进量），如 "0.75cm"、"420磅"、"0字符"
    text_indent: Optional[str] = Field(
        default=None,
        description="文本缩进（悬挂缩进），如 '0.75cm'、'420磅'、'0字符'",
    )


class NumberingConfig(BaseModel):
    """标题自动编号配置"""

    enabled: bool = Field(default=False, description="是否启用自动编号功能")
    level_1: NumberingLevelConfig = Field(
        default_factory=lambda: NumberingLevelConfig(
            enabled=False,
            template="第%1章",
            strip_pattern="^第[一二三四五六七八九十百千零]+章\\s*",
        ),
        description="一级标题编号配置",
    )
    level_2: NumberingLevelConfig = Field(
        default_factory=lambda: NumberingLevelConfig(
            enabled=False,
            template="%1.%2",
            strip_pattern="^\\d+(\\.\\d+)\\s*",
        ),
        description="二级标题编号配置",
    )
    level_3: NumberingLevelConfig = Field(
        default_factory=lambda: NumberingLevelConfig(
            enabled=False,
            template="%1.%2.%3",
            strip_pattern="^\\d+(\\.\\d+){2}\\s*",
        ),
        description="三级标题编号配置",
    )


# -------------------------- 根配置模型 --------------------------
class NodeConfigRoot(BaseModel):
    """配置根节点模型"""

    style_checks_warning: WarningFieldConfig = Field(
        default_factory=WarningFieldConfig, description="警告字段配置"
    )
    global_format: GlobalFormatConfig = Field(
        default_factory=GlobalFormatConfig, description="默认格式配置"
    )
    abstract: AbstractConfig = Field(
        default_factory=AbstractConfig, description="摘要总配置"
    )
    headings: HeadingsConfig = Field(
        default_factory=HeadingsConfig, description="标题配置"
    )
    body_text: BodyTextConfig = Field(
        default_factory=BodyTextConfig, description="正文配置"
    )
    figures: FiguresConfig = Field(
        default_factory=FiguresConfig, description="插图配置"
    )
    tables: TablesConfig = Field(default_factory=TablesConfig, description="表格配置")
    references: ReferencesConfig = Field(
        default_factory=ReferencesConfig, description="参考文献总配置"
    )
    acknowledgements: AcknowledgementsConfig = Field(
        default_factory=AcknowledgementsConfig,
        description="致谢总配置",
    )
    numbering: NumberingConfig = Field(
        default_factory=NumberingConfig, description="标题自动编号配置"
    )
