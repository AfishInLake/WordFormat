#! /usr/bin/env python
# @Time    : 2026/1/11 22:25
# @Author  : afish
# @File    : settings.py
from src.rules import (
    AbstractContentCN,
    AbstractContentEN,
    AbstractTitleCN,
    AbstractTitleContentCN,
    AbstractTitleContentEN,
    AbstractTitleEN,
    Acknowledgements,
    BodyText,
    CaptionFigure,
    CaptionTable,
    HeadingLevel1Node,
    HeadingLevel2Node,
    HeadingLevel3Node,
    KeywordsCN,
    KeywordsEN,
    References,
)

# 标签节点映射
CATEGORY_TO_CLASS = {
    "abstract_chinese_title": AbstractTitleCN,
    "abstract_english_title": AbstractTitleEN,
    "abstract_chinese_title_content": AbstractTitleContentCN,
    "abstract_english_title_content": AbstractTitleContentEN,
    "abstract_chinese_content": AbstractContentCN,
    "abstract_english_content": AbstractContentEN,
    "keywords_chinese": KeywordsCN,
    "keywords_english": KeywordsEN,
    "heading_level_1": HeadingLevel1Node,
    "heading_level_2": HeadingLevel2Node,
    "heading_level_3": HeadingLevel3Node,
    "references_title": References,
    "acknowledgements_title": Acknowledgements,
    "caption_figure": CaptionFigure,
    "caption_table": CaptionTable,
    "body_text": BodyText,
}
