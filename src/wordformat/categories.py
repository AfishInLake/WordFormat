"""文档结构类别（category）的唯一权威定义。

所有 category 都应通过本枚举引用，禁止在业务代码中硬编码字面量：
- 拼错成员名立即 AttributeError（不再是静默不命中）；
- 拼接构造类别时用 ``Category(f"heading_level_{n}")``，非法值立即 ValueError；
- StrEnum 是 str 子类，可原样参与 dict 查找、JSON 序列化与日志输出。

新增类别时：在此添加成员 → 在对应领域 rules 模块用 @register 注册。
"""

from enum import Enum


class _StrEnum(str, Enum):
    """兼容 Python 3.10 的 StrEnum（项目 requires-python >=3.10）。

    str()/f-string 返回枚举值本身，保证日志、JSON、类别名拼接与普通
    字符串完全一致；与 str 的相等性/哈希也相同，可直接作 dict 键。
    """

    def __str__(self) -> str:
        return str(self.value)

    def __format__(self, format_spec: str) -> str:
        return format(str(self.value), format_spec)


class Category(_StrEnum):
    """文档结构类别。"""

    # ===== 结构/占位类别 =====
    TOP = "top"
    OTHER = "other"
    BODY_TEXT = "body_text"
    DOCUMENT_TITLE = "document_title"
    FOOTER = "footer"
    TOC_ENTRY = "toc_entry"
    HEADING_MULU = "heading_mulu"
    HEADING_FULU = "heading_fulu"

    # ===== 标题 =====
    HEADING_LEVEL_1 = "heading_level_1"
    HEADING_LEVEL_2 = "heading_level_2"
    HEADING_LEVEL_3 = "heading_level_3"

    # ===== 摘要 =====
    ABSTRACT_CHINESE_TITLE = "abstract_chinese_title"
    ABSTRACT_CHINESE_TITLE_CONTENT = "abstract_chinese_title_content"
    ABSTRACT_CHINESE_CONTENT = "abstract_chinese_content"
    ABSTRACT_ENGLISH_TITLE = "abstract_english_title"
    ABSTRACT_ENGLISH_TITLE_CONTENT = "abstract_english_title_content"
    ABSTRACT_ENGLISH_CONTENT = "abstract_english_content"

    # ===== 关键词/致谢/参考文献 =====
    KEYWORDS_CHINESE = "keywords_chinese"
    KEYWORDS_ENGLISH = "keywords_english"
    ACKNOWLEDGEMENTS_TITLE = "acknowledgements_title"
    ACKNOWLEDGEMENTS_CONTENT = "acknowledgements_content"
    REFERENCES_TITLE = "references_title"
    REFERENCES_CONTENT = "references_content"

    # ===== 图表/公式 =====
    CAPTION_FIGURE = "caption_figure"
    CAPTION_TABLE = "caption_table"
    FIGURE_IMAGE = "figure_image"
    TABLE_OBJECT = "table_object"
    MATH_BLOCK = "math_block"


#: 建树时视为正文的类别（挂到最近的标题节点下）
NON_HEADING_CATEGORIES: frozenset[Category] = frozenset(
    {
        Category.BODY_TEXT,
        Category.MATH_BLOCK,
        Category.FIGURE_IMAGE,
        Category.TABLE_OBJECT,
        Category.TOC_ENTRY,
    }
)

#: 展示树不视为可挂载标题的类别（排除式名单，含不参与格式化的占位类别）
DISPLAY_NON_HEADING_CATEGORIES: frozenset[Category] = (
    NON_HEADING_CATEGORIES
    | frozenset(
        {
            Category.OTHER,
            Category.DOCUMENT_TITLE,
            Category.FOOTER,
            Category.KEYWORDS_CHINESE,
            Category.KEYWORDS_ENGLISH,
        }
    )
)
