#! /usr/bin/env python
# @Time    : 2026/2/11 11:46
# @Author  : afish
# @File    : test_rules.py
"""
测试rules模块的功能
"""

import pytest
from unittest.mock import MagicMock, patch
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from wordformat.rules.keywords import BaseKeywordsNode, KeywordsEN, KeywordsCN
from wordformat.rules.abstract import (
    AbstractTitleCN,
    AbstractTitleContentCN,
    AbstractContentCN,
    AbstractTitleEN,
    AbstractTitleContentEN,
    AbstractContentEN,
)
from wordformat.rules.node import FormatNode
from wordformat.rules.acknowledgement import Acknowledgements, AcknowledgementsCN
from wordformat.rules.body import BodyText
from wordformat.rules.caption import CaptionFigure, CaptionTable
from wordformat.rules.heading import BaseHeadingNode, HeadingLevel1Node, HeadingLevel2Node, HeadingLevel3Node
from wordformat.rules.references import References, ReferenceEntry
from wordformat.config.datamodel import (
    KeywordsConfig,
    NodeConfigRoot,
    AbstractChineseConfig,
    AbstractEnglishConfig,
    AbstractTitleConfig,
    HeadingLevelConfig,
    AcknowledgementsTitleConfig,
    AcknowledgementsContentConfig,
    BodyTextConfig,
    FiguresConfig,
    TablesConfig,
    ReferencesTitleConfig,
    ReferencesContentConfig,
)


# 测试keywords.py中的类
def test_base_keywords_node_load_config():
    """测试BaseKeywordsNode的load_config方法"""
    # 创建BaseKeywordsNode的子类实例
    class TestKeywordsNode(BaseKeywordsNode):
        LANG = "cn"
        NODE_TYPE = "test.keywords"
    
    mock_paragraph = MagicMock(spec=Paragraph)
    node = TestKeywordsNode(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 测试从字典加载配置
    config_dict = {
        "abstract": {
            "keywords": {
                "cn": {
                    "alignment": "左对齐",
                    "space_before": "0.5行",
                    "space_after": "0.5行",
                    "line_spacing": "1.5倍",
                    "line_spacingrule": "单倍行距",
                    "first_line_indent": "0字符",
                    "builtin_style_name": "正文",
                    "chinese_font_name": "宋体",
                    "english_font_name": "Times New Roman",
                    "font_size": "小四",
                    "font_color": "BLACK",
                    "bold": False,
                    "kewords_bold": True,
                    "italic": False,
                    "underline": False,
                    "count_min": 3,
                    "count_max": 8,
                    "trailing_punct_forbidden": True,
                }
            }
        }
    }
    
    node.load_config(config_dict)
    assert node.config is not None
    
    # 测试从NodeConfigRoot加载配置
    mock_config = MagicMock(spec=NodeConfigRoot)
    mock_keywords_config_cn = MagicMock(spec=KeywordsConfig)
    mock_keywords_config_en = MagicMock(spec=KeywordsConfig)
    
    # 设置mock对象的嵌套属性
    mock_abstract = MagicMock()
    mock_abstract.keywords = {"chinese": mock_keywords_config_cn, "english": mock_keywords_config_en}
    mock_config.abstract = mock_abstract
    
    node.load_config(mock_config)
    assert node.pydantic_config is not None


def test_keywords_en_check_keyword_label():
    """测试KeywordsEN的_check_keyword_label方法"""
    # 创建KeywordsEN实例
    mock_paragraph = MagicMock(spec=Paragraph)
    keywords_en = KeywordsEN(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 测试包含关键词标签的情况
    run = MagicMock(spec=Run)
    run.text = "Keywords:"
    assert keywords_en._check_keyword_label(run) is True
    
    # 测试包含KEY WORDS的情况
    run.text = "KEY WORDS:"
    assert keywords_en._check_keyword_label(run) is True
    
    # 测试不包含关键词标签的情况
    run.text = "test"
    assert keywords_en._check_keyword_label(run) is False


def test_keywords_cn_check_keyword_label():
    """测试KeywordsCN的_check_keyword_label方法"""
    # 创建KeywordsCN实例
    mock_paragraph = MagicMock(spec=Paragraph)
    keywords_cn = KeywordsCN(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 测试包含关键词标签的情况
    run = MagicMock(spec=Run)
    run.text = "关键词："
    assert keywords_cn._check_keyword_label(run) is True
    
    # 测试不包含关键词标签的情况
    run.text = "测试"
    assert keywords_cn._check_keyword_label(run) is False


def test_keywords_en_base():
    """测试KeywordsEN的_base方法"""
    # 创建KeywordsEN实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    mock_paragraph.runs[0].text = "Keywords: test1, test2"
    keywords_en = KeywordsEN(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=KeywordsConfig)
    mock_config.alignment = "左对齐"
    mock_config.space_before = "0.5行"
    mock_config.space_after = "0.5行"
    mock_config.line_spacing = "1.5倍"
    mock_config.line_spacingrule = "单倍行距"
    mock_config.first_line_indent = "0字符"
    mock_config.builtin_style_name = "正文"
    mock_config.chinese_font_name = "宋体"
    mock_config.english_font_name = "Times New Roman"
    mock_config.font_size = "小四"
    mock_config.font_color = "BLACK"
    mock_config.bold = False
    mock_config.kewords_bold = True
    mock_config.italic = False
    mock_config.underline = False
    mock_config.count_min = 3
    mock_config.count_max = 8
    
    keywords_en._pydantic_config = mock_config
    keywords_en.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.keywords.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = []
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.keywords.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = []
            mock_cs.apply_to_run.return_value = []
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            result = keywords_en._base(doc, True, True)
            
            # 验证结果
            assert result is None or isinstance(result, list)


def test_keywords_cn_base():
    """测试KeywordsCN的_base方法"""
    # 创建KeywordsCN实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    mock_paragraph.runs[0].text = "关键词：测试1；测试2"
    keywords_cn = KeywordsCN(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=KeywordsConfig)
    mock_config.alignment = "左对齐"
    mock_config.space_before = "0.5行"
    mock_config.space_after = "0.5行"
    mock_config.line_spacing = "1.5倍"
    mock_config.line_spacingrule = "单倍行距"
    mock_config.first_line_indent = "0字符"
    mock_config.builtin_style_name = "正文"
    mock_config.chinese_font_name = "宋体"
    mock_config.english_font_name = "Times New Roman"
    mock_config.font_size = "小四"
    mock_config.font_color = "BLACK"
    mock_config.bold = False
    mock_config.kewords_bold = True
    mock_config.italic = False
    mock_config.underline = False
    mock_config.count_min = 3
    mock_config.count_max = 8
    mock_config.trailing_punct_forbidden = True
    
    keywords_cn._pydantic_config = mock_config
    keywords_cn.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.keywords.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = []
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.keywords.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = []
            mock_cs.apply_to_run.return_value = []
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            result = keywords_cn._base(doc, True, True)
            
            # 验证结果
            assert result is None or isinstance(result, list)


def test_keywords_load_config_error():
    """测试BaseKeywordsNode的load_config方法错误场景"""
    # 创建BaseKeywordsNode的子类实例
    class TestKeywordsNode(BaseKeywordsNode):
        LANG = "cn"
        NODE_TYPE = "test.keywords"
    
    mock_paragraph = MagicMock(spec=Paragraph)
    node = TestKeywordsNode(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 测试不支持的配置类型
    with pytest.raises(TypeError):
        node.load_config("invalid_config")


def test_keywords_check_paragraph_style_apply():
    """测试BaseKeywordsNode的_check_paragraph_style方法（应用模式）"""
    # 创建BaseKeywordsNode的子类实例
    class TestKeywordsNode(BaseKeywordsNode):
        LANG = "cn"
        NODE_TYPE = "test.keywords"
    
    mock_paragraph = MagicMock(spec=Paragraph)
    node = TestKeywordsNode(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=KeywordsConfig)
    mock_config.alignment = "左对齐"
    mock_config.space_before = "0.5行"
    mock_config.space_after = "0.5行"
    mock_config.line_spacing = "1.5倍"
    mock_config.line_spacingrule = "单倍行距"
    mock_config.first_line_indent = "0字符"
    mock_config.builtin_style_name = "正文"
    
    # 模拟ParagraphStyle（返回带comment属性的对象）
    with patch('wordformat.rules.keywords.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        # 创建带comment属性的对象
        mock_obj = MagicMock()
        mock_obj.comment = "测试注释"
        mock_ps.apply_to_paragraph.return_value = [mock_obj]
        mock_paragraph_style.return_value = mock_ps
        
        # 执行方法（应用模式）
        result = node._check_paragraph_style(mock_config, False)
        
        # 验证结果
        assert isinstance(result, list)
        assert len(result) > 0


def test_keywords_en_base_style_errors():
    """测试KeywordsEN的_base方法（样式错误场景）"""
    # 创建KeywordsEN实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_run1 = MagicMock(spec=Run)
    mock_run1.text = "Keywords:"
    mock_run2 = MagicMock(spec=Run)
    mock_run2.text = " test1, test2"
    mock_paragraph.runs = [mock_run1, mock_run2]
    keywords_en = KeywordsEN(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=KeywordsConfig)
    mock_config.alignment = "左对齐"
    mock_config.space_before = "0.5行"
    mock_config.space_after = "0.5行"
    mock_config.line_spacing = "1.5倍"
    mock_config.line_spacingrule = "单倍行距"
    mock_config.first_line_indent = "0字符"
    mock_config.builtin_style_name = "正文"
    mock_config.chinese_font_name = "宋体"
    mock_config.english_font_name = "Times New Roman"
    mock_config.font_size = "小四"
    mock_config.font_color = "BLACK"
    mock_config.bold = False
    mock_config.kewords_bold = True
    mock_config.italic = False
    mock_config.underline = False
    mock_config.count_min = 3
    mock_config.count_max = 8
    
    keywords_en._pydantic_config = mock_config
    keywords_en.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle（返回带comment属性的对象）
    with patch('wordformat.rules.keywords.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        # 创建带comment属性的对象
        mock_obj1 = MagicMock()
        mock_obj1.comment = "段落样式错误"
        mock_ps.diff_from_paragraph.return_value = [mock_obj1]
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.keywords.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            # 创建带comment属性的对象
            mock_obj2 = MagicMock()
            mock_obj2.comment = "字符样式错误"
            mock_cs.diff_from_run.return_value = [mock_obj2]
            mock_cs.apply_to_run.return_value = [mock_obj2]
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            result = keywords_en._base(doc, True, True)
            
            # 验证结果
            assert result is None


def test_keywords_en_base_keyword_count_error():
    """测试KeywordsEN的_base方法（关键词数量错误场景）"""
    # 创建KeywordsEN实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    mock_paragraph.runs[0].text = "Keywords: test1"
    keywords_en = KeywordsEN(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置（设置最小3个，最大8个）
    mock_config = MagicMock(spec=KeywordsConfig)
    mock_config.alignment = "左对齐"
    mock_config.space_before = "0.5行"
    mock_config.space_after = "0.5行"
    mock_config.line_spacing = "1.5倍"
    mock_config.line_spacingrule = "单倍行距"
    mock_config.first_line_indent = "0字符"
    mock_config.builtin_style_name = "正文"
    mock_config.chinese_font_name = "宋体"
    mock_config.english_font_name = "Times New Roman"
    mock_config.font_size = "小四"
    mock_config.font_color = "BLACK"
    mock_config.bold = False
    mock_config.kewords_bold = True
    mock_config.italic = False
    mock_config.underline = False
    mock_config.count_min = 3
    mock_config.count_max = 8
    
    keywords_en._pydantic_config = mock_config
    keywords_en.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.keywords.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = []
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.keywords.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = []
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            result = keywords_en._base(doc, True, True)
            
            # 验证结果
            assert result is None


# 测试abstract.py中的类
def test_abstract_title_cn_base():
    """测试AbstractTitleCN的_base方法"""
    # 创建AbstractTitleCN实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    abstract_title_cn = AbstractTitleCN(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=AbstractTitleConfig)
    mock_config.alignment = "左对齐"
    mock_config.space_before = "0.5行"
    mock_config.space_after = "0.5行"
    mock_config.line_spacing = "1.5倍"
    mock_config.line_spacingrule = "单倍行距"
    mock_config.first_line_indent = "0字符"
    mock_config.builtin_style_name = "正文"
    mock_config.chinese_font_name = "宋体"
    mock_config.english_font_name = "Times New Roman"
    mock_config.font_size = "小四"
    mock_config.font_color = "BLACK"
    mock_config.bold = False
    mock_config.italic = False
    mock_config.underline = False
    
    abstract_title_cn._pydantic_config = mock_config
    abstract_title_cn.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.abstract.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = []
        mock_ps.apply_to_paragraph.return_value = []
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.abstract.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = []
            mock_cs.apply_to_run.return_value = []
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            abstract_title_cn._base(doc, True, True)
            
            # 验证调用
            assert abstract_title_cn.add_comment.called


def test_abstract_title_content_cn_check_title():
    """测试AbstractTitleContentCN的check_title方法"""
    # 创建AbstractTitleContentCN实例
    mock_paragraph = MagicMock(spec=Paragraph)
    abstract_title_content_cn = AbstractTitleContentCN(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 测试包含摘要标题的情况
    run = MagicMock(spec=Run)
    run.text = "摘要"
    assert abstract_title_content_cn.check_title(run) is True
    
    # 测试不包含摘要标题的情况
    run = MagicMock(spec=Run)
    run.text = "测试内容"
    assert abstract_title_content_cn.check_title(run) is False


def test_abstract_title_content_cn_base():
    """测试AbstractTitleContentCN的_base方法"""
    # 创建AbstractTitleContentCN实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    mock_paragraph.runs[0].text = "摘要：测试内容"
    abstract_title_content_cn = AbstractTitleContentCN(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=AbstractChineseConfig)
    mock_config.chinese_title = MagicMock(spec=AbstractTitleConfig)
    mock_config.chinese_title.alignment = "左对齐"
    mock_config.chinese_title.space_before = "0.5行"
    mock_config.chinese_title.space_after = "0.5行"
    mock_config.chinese_title.line_spacing = "1.5倍"
    mock_config.chinese_title.line_spacingrule = "单倍行距"
    mock_config.chinese_title.first_line_indent = "0字符"
    mock_config.chinese_title.builtin_style_name = "正文"
    mock_config.chinese_title.chinese_font_name = "宋体"
    mock_config.chinese_title.english_font_name = "Times New Roman"
    mock_config.chinese_title.font_size = "小四"
    mock_config.chinese_title.font_color = "BLACK"
    mock_config.chinese_title.bold = False
    mock_config.chinese_title.italic = False
    mock_config.chinese_title.underline = False
    
    mock_config.chinese_content = MagicMock(spec=AbstractTitleConfig)
    mock_config.chinese_content.alignment = "左对齐"
    mock_config.chinese_content.space_before = "0.5行"
    mock_config.chinese_content.space_after = "0.5行"
    mock_config.chinese_content.line_spacing = "1.5倍"
    mock_config.chinese_content.line_spacingrule = "单倍行距"
    mock_config.chinese_content.first_line_indent = "0字符"
    mock_config.chinese_content.builtin_style_name = "正文"
    mock_config.chinese_content.chinese_font_name = "宋体"
    mock_config.chinese_content.english_font_name = "Times New Roman"
    mock_config.chinese_content.font_size = "小四"
    mock_config.chinese_content.font_color = "BLACK"
    mock_config.chinese_content.bold = False
    mock_config.chinese_content.italic = False
    mock_config.chinese_content.underline = False
    
    abstract_title_content_cn._pydantic_config = mock_config
    abstract_title_content_cn.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.abstract.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = []
        mock_ps.apply_to_paragraph.return_value = []
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.abstract.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = []
            mock_cs.apply_to_run.return_value = []
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            abstract_title_content_cn._base(doc, True, True)
            
            # 验证调用
            assert abstract_title_content_cn.add_comment.called


def test_abstract_content_cn_base():
    """测试AbstractContentCN的_base方法"""
    # 创建AbstractContentCN实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    abstract_content_cn = AbstractContentCN(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=AbstractChineseConfig)
    mock_config.chinese_content = MagicMock(spec=AbstractTitleConfig)
    mock_config.chinese_content.alignment = "左对齐"
    mock_config.chinese_content.space_before = "0.5行"
    mock_config.chinese_content.space_after = "0.5行"
    mock_config.chinese_content.line_spacing = "1.5倍"
    mock_config.chinese_content.line_spacingrule = "单倍行距"
    mock_config.chinese_content.first_line_indent = "0字符"
    mock_config.chinese_content.builtin_style_name = "正文"
    mock_config.chinese_content.chinese_font_name = "宋体"
    mock_config.chinese_content.english_font_name = "Times New Roman"
    mock_config.chinese_content.font_size = "小四"
    mock_config.chinese_content.font_color = "BLACK"
    mock_config.chinese_content.bold = False
    mock_config.chinese_content.italic = False
    mock_config.chinese_content.underline = False
    
    abstract_content_cn._pydantic_config = mock_config
    abstract_content_cn.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.abstract.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = []
        mock_ps.apply_to_paragraph.return_value = []
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.abstract.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = []
            mock_cs.apply_to_run.return_value = []
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            abstract_content_cn._base(doc, True, True)
            
            # 验证调用
            assert abstract_content_cn.add_comment.called


def test_abstract_title_en_base():
    """测试AbstractTitleEN的_base方法"""
    # 创建AbstractTitleEN实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    abstract_title_en = AbstractTitleEN(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=AbstractTitleConfig)
    mock_config.alignment = "左对齐"
    mock_config.space_before = "0.5行"
    mock_config.space_after = "0.5行"
    mock_config.line_spacing = "1.5倍"
    mock_config.line_spacingrule = "单倍行距"
    mock_config.first_line_indent = "0字符"
    mock_config.builtin_style_name = "正文"
    mock_config.chinese_font_name = "宋体"
    mock_config.english_font_name = "Times New Roman"
    mock_config.font_size = "小四"
    mock_config.font_color = "BLACK"
    mock_config.bold = False
    mock_config.italic = False
    mock_config.underline = False
    
    abstract_title_en._pydantic_config = mock_config
    abstract_title_en.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.abstract.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = []
        mock_ps.apply_to_paragraph.return_value = []
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.abstract.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = []
            mock_cs.apply_to_run.return_value = []
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            result = abstract_title_en._base(doc, True, True)
            
            # 验证结果
            assert result == []


def test_abstract_title_content_en_check_title():
    """测试AbstractTitleContentEN的check_title方法"""
    # 创建AbstractTitleContentEN实例
    mock_paragraph = MagicMock(spec=Paragraph)
    abstract_title_content_en = AbstractTitleContentEN(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 测试包含Abstract标题的情况
    run = MagicMock(spec=Run)
    run.text = "Abstract"
    assert abstract_title_content_en.check_title(run) is True
    
    # 测试不包含Abstract标题的情况
    run = MagicMock(spec=Run)
    run.text = "Test content"
    assert abstract_title_content_en.check_title(run) is False


def test_abstract_title_content_en_base():
    """测试AbstractTitleContentEN的_base方法"""
    # 创建AbstractTitleContentEN实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    mock_paragraph.runs[0].text = "Abstract: Test content"
    abstract_title_content_en = AbstractTitleContentEN(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=AbstractEnglishConfig)
    mock_config.english_title = MagicMock(spec=AbstractTitleConfig)
    mock_config.english_title.alignment = "左对齐"
    mock_config.english_title.space_before = "0.5行"
    mock_config.english_title.space_after = "0.5行"
    mock_config.english_title.line_spacing = "1.5倍"
    mock_config.english_title.line_spacingrule = "单倍行距"
    mock_config.english_title.first_line_indent = "0字符"
    mock_config.english_title.builtin_style_name = "正文"
    mock_config.english_title.chinese_font_name = "宋体"
    mock_config.english_title.english_font_name = "Times New Roman"
    mock_config.english_title.font_size = "小四"
    mock_config.english_title.font_color = "BLACK"
    mock_config.english_title.bold = False
    mock_config.english_title.italic = False
    mock_config.english_title.underline = False
    
    mock_config.english_content = MagicMock(spec=AbstractTitleConfig)
    mock_config.english_content.alignment = "左对齐"
    mock_config.english_content.space_before = "0.5行"
    mock_config.english_content.space_after = "0.5行"
    mock_config.english_content.line_spacing = "1.5倍"
    mock_config.english_content.line_spacingrule = "单倍行距"
    mock_config.english_content.first_line_indent = "0字符"
    mock_config.english_content.builtin_style_name = "正文"
    mock_config.english_content.chinese_font_name = "宋体"
    mock_config.english_content.english_font_name = "Times New Roman"
    mock_config.english_content.font_size = "小四"
    mock_config.english_content.font_color = "BLACK"
    mock_config.english_content.bold = False
    mock_config.english_content.italic = False
    mock_config.english_content.underline = False
    
    abstract_title_content_en._pydantic_config = mock_config
    abstract_title_content_en.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.abstract.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = []
        mock_ps.apply_to_paragraph.return_value = []
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.abstract.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = []
            mock_cs.apply_to_run.return_value = []
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            abstract_title_content_en._base(doc, True, True)
            
            # 验证调用
            assert abstract_title_content_en.add_comment.called


def test_abstract_content_en_base():
    """测试AbstractContentEN的_base方法"""
    # 创建AbstractContentEN实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    abstract_content_en = AbstractContentEN(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=AbstractEnglishConfig)
    mock_config.english_content = MagicMock(spec=AbstractTitleConfig)
    mock_config.english_content.alignment = "左对齐"
    mock_config.english_content.space_before = "0.5行"
    mock_config.english_content.space_after = "0.5行"
    mock_config.english_content.line_spacing = "1.5倍"
    mock_config.english_content.line_spacingrule = "单倍行距"
    mock_config.english_content.first_line_indent = "0字符"
    mock_config.english_content.builtin_style_name = "正文"
    mock_config.english_content.chinese_font_name = "宋体"
    mock_config.english_content.english_font_name = "Times New Roman"
    mock_config.english_content.font_size = "小四"
    mock_config.english_content.font_color = "BLACK"
    mock_config.english_content.bold = False
    mock_config.english_content.italic = False
    mock_config.english_content.underline = False
    
    abstract_content_en._pydantic_config = mock_config
    abstract_content_en.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.abstract.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = [MagicMock()]
        mock_ps.apply_to_paragraph.return_value = [MagicMock()]
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.abstract.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = [MagicMock()]
            mock_cs.apply_to_run.return_value = [MagicMock()]
            mock_cs.to_string.return_value = "测试差异"
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            abstract_content_en._base(doc, True, True)
            
            # 验证调用
            assert abstract_content_en.add_comment.called


# 测试node.py中的类
def test_format_node_init():
    """测试FormatNode的初始化"""
    # 创建FormatNode实例
    mock_paragraph = MagicMock(spec=Paragraph)
    format_node = FormatNode(value={'category': 'test', 'fingerprint': 'test_fingerprint'}, level=0, paragraph=mock_paragraph)
    
    # 验证属性
    assert hasattr(format_node, 'value')
    assert hasattr(format_node, 'level')
    assert hasattr(format_node, 'paragraph')


# 测试acknowledgement.py中的类
def test_acknowledgements_base():
    """测试Acknowledgements的_base方法"""
    # 创建Acknowledgements实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    acknowledgements = Acknowledgements(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=AcknowledgementsTitleConfig)
    mock_config.alignment = "左对齐"
    mock_config.space_before = "0.5行"
    mock_config.space_after = "0.5行"
    mock_config.line_spacing = "1.5倍"
    mock_config.line_spacingrule = "单倍行距"
    mock_config.first_line_indent = "0字符"
    mock_config.builtin_style_name = "正文"
    mock_config.chinese_font_name = "宋体"
    mock_config.english_font_name = "Times New Roman"
    mock_config.font_size = "小四"
    mock_config.font_color = "BLACK"
    mock_config.bold = False
    mock_config.italic = False
    mock_config.underline = False
    
    acknowledgements._pydantic_config = mock_config
    acknowledgements.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.acknowledgement.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = [MagicMock()]
        mock_ps.apply_to_paragraph.return_value = [MagicMock()]
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.acknowledgement.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = [MagicMock()]
            mock_cs.apply_to_run.return_value = [MagicMock()]
            mock_cs.to_string.return_value = "测试差异"
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            result = acknowledgements._base(doc, True, True)
            
            # 验证结果
            assert result == []
            assert acknowledgements.add_comment.called


def test_acknowledgements_cn_base():
    """测试AcknowledgementsCN的_base方法"""
    # 创建AcknowledgementsCN实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    acknowledgements_cn = AcknowledgementsCN(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=AcknowledgementsContentConfig)
    mock_config.alignment = "左对齐"
    mock_config.space_before = "0.5行"
    mock_config.space_after = "0.5行"
    mock_config.line_spacing = "1.5倍"
    mock_config.line_spacingrule = "单倍行距"
    mock_config.first_line_indent = "0字符"
    mock_config.builtin_style_name = "正文"
    mock_config.chinese_font_name = "宋体"
    mock_config.english_font_name = "Times New Roman"
    mock_config.font_size = "小四"
    mock_config.font_color = "BLACK"
    mock_config.bold = False
    mock_config.italic = False
    mock_config.underline = False
    
    acknowledgements_cn._pydantic_config = mock_config
    acknowledgements_cn.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.acknowledgement.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = [MagicMock()]
        mock_ps.apply_to_paragraph.return_value = [MagicMock()]
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.acknowledgement.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = [MagicMock()]
            mock_cs.apply_to_run.return_value = [MagicMock()]
            mock_cs.to_string.return_value = "测试差异"
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            result = acknowledgements_cn._base(doc, True, True)
            
            # 验证结果
            assert result == []
            assert acknowledgements_cn.add_comment.called


# 测试body.py中的类
def test_body_text_base():
    """测试BodyText的_base方法"""
    # 创建BodyText实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    body_text = BodyText(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=BodyTextConfig)
    mock_config.alignment = "左对齐"
    mock_config.space_before = "0.5行"
    mock_config.space_after = "0.5行"
    mock_config.line_spacing = "1.5倍"
    mock_config.line_spacingrule = "单倍行距"
    mock_config.first_line_indent = "2字符"
    mock_config.builtin_style_name = "正文"
    mock_config.chinese_font_name = "宋体"
    mock_config.english_font_name = "Times New Roman"
    mock_config.font_size = "小四"
    mock_config.font_color = "BLACK"
    mock_config.bold = False
    mock_config.italic = False
    mock_config.underline = False
    
    body_text._pydantic_config = mock_config
    body_text.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.body.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = [MagicMock()]
        mock_ps.apply_to_paragraph.return_value = [MagicMock()]
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.body.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = [MagicMock()]
            mock_cs.apply_to_run.return_value = [MagicMock()]
            mock_cs.to_string.return_value = "测试差异"
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            result = body_text._base(doc, True, True)
            
            # 验证结果
            assert result == []
            assert body_text.add_comment.called


# 测试caption.py中的类
def test_caption_figure_base():
    """测试CaptionFigure的_base方法"""
    # 创建CaptionFigure实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    caption_figure = CaptionFigure(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=FiguresConfig)
    mock_config.alignment = "居中"
    mock_config.space_before = "0.5行"
    mock_config.space_after = "0.5行"
    mock_config.line_spacing = "1.5倍"
    mock_config.line_spacingrule = "单倍行距"
    mock_config.first_line_indent = "0字符"
    mock_config.builtin_style_name = "正文"
    mock_config.chinese_font_name = "宋体"
    mock_config.english_font_name = "Times New Roman"
    mock_config.font_size = "小四"
    mock_config.font_color = "BLACK"
    mock_config.bold = False
    mock_config.italic = False
    mock_config.underline = False
    
    caption_figure._pydantic_config = mock_config
    caption_figure.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.caption.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = [MagicMock()]
        mock_ps.apply_to_paragraph.return_value = [MagicMock()]
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.caption.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = [MagicMock()]
            mock_cs.apply_to_run.return_value = [MagicMock()]
            mock_cs.to_string.return_value = "测试差异"
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            caption_figure._base(doc, True, True)
            
            # 验证结果
            assert caption_figure.add_comment.called


def test_caption_table_base():
    """测试CaptionTable的_base方法"""
    # 创建CaptionTable实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    caption_table = CaptionTable(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=TablesConfig)
    mock_config.alignment = "居中"
    mock_config.space_before = "0.5行"
    mock_config.space_after = "0.5行"
    mock_config.line_spacing = "1.5倍"
    mock_config.line_spacingrule = "单倍行距"
    mock_config.first_line_indent = "0字符"
    mock_config.builtin_style_name = "正文"
    mock_config.chinese_font_name = "宋体"
    mock_config.english_font_name = "Times New Roman"
    mock_config.font_size = "小四"
    mock_config.font_color = "BLACK"
    mock_config.bold = False
    mock_config.italic = False
    mock_config.underline = False
    
    caption_table._pydantic_config = mock_config
    caption_table.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.caption.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = [MagicMock()]
        mock_ps.apply_to_paragraph.return_value = [MagicMock()]
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.caption.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = [MagicMock()]
            mock_cs.apply_to_run.return_value = [MagicMock()]
            mock_cs.to_string.return_value = "测试差异"
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            caption_table._base(doc, True, True)
            
            # 验证结果
            assert caption_table.add_comment.called


# 测试heading.py中的类
def test_heading_level1_node_base():
    """测试HeadingLevel1Node的_base方法"""
    # 创建HeadingLevel1Node实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    mock_paragraph.text = "测试标题"
    heading_level1 = HeadingLevel1Node(value=mock_paragraph, level=1, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=HeadingLevelConfig)
    mock_config.alignment = "左对齐"
    mock_config.space_before = "1行"
    mock_config.space_after = "0.5行"
    mock_config.line_spacing = "1.5倍"
    mock_config.line_spacingrule = "单倍行距"
    mock_config.first_line_indent = "0字符"
    mock_config.builtin_style_name = "标题1"
    mock_config.chinese_font_name = "黑体"
    mock_config.english_font_name = "Times New Roman"
    mock_config.font_size = "二号"
    mock_config.font_color = "BLACK"
    mock_config.bold = True
    mock_config.italic = False
    mock_config.underline = False
    mock_config.section_title_re = "^第[一二三四五六七八九十]+章"
    
    heading_level1._pydantic_config = mock_config
    heading_level1.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.heading.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = [MagicMock()]
        mock_ps.apply_to_paragraph.return_value = [MagicMock()]
        mock_ps.to_string.return_value = "测试段落差异"
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.heading.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = [MagicMock()]
            mock_cs.apply_to_run.return_value = [MagicMock()]
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            result = heading_level1._base(doc, True, True)
            
            # 验证结果
            assert isinstance(result, list)
            assert heading_level1.add_comment.called


def test_heading_base_no_config():
    """测试BaseHeadingNode的_base方法（无配置情况）"""
    # 创建HeadingLevel1Node实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    heading_level1 = HeadingLevel1Node(value=mock_paragraph, level=1, paragraph=mock_paragraph)
    
    # 不设置配置
    heading_level1._pydantic_config = None
    heading_level1.add_comment = MagicMock()
    
    # 执行方法
    doc = MagicMock()
    result = heading_level1._base(doc, True, True)
    
    # 验证结果
    assert isinstance(result, list)
    assert len(result) == 1
    assert "error" in result[0]
    assert heading_level1.add_comment.called


def test_heading_load_config_dict():
    """测试BaseHeadingNode的load_config方法（字典配置）"""
    # 创建HeadingLevel1Node实例
    mock_paragraph = MagicMock(spec=Paragraph)
    heading_level1 = HeadingLevel1Node(value=mock_paragraph, level=1, paragraph=mock_paragraph)
    
    # 准备字典配置，添加缺失的section_title_re字段
    config_dict = {
        "headings": {
            "level_1": {
                "alignment": "左对齐",
                "space_before": "1行",
                "space_after": "0.5行",
                "line_spacing": "1.5倍",
                "line_spacingrule": "单倍行距",
                "first_line_indent": "0字符",
                "builtin_style_name": "标题1",
                "chinese_font_name": "黑体",
                "english_font_name": "Times New Roman",
                "font_size": "二号",
                "font_color": "BLACK",
                "bold": True,
                "italic": False,
                "underline": False,
                "section_title_re": "^第[一二三四五六七八九十]+章"
            }
        }
    }
    
    # 执行方法
    heading_level1.load_config(config_dict)
    
    # 验证结果
    assert hasattr(heading_level1, '_config')
    assert hasattr(heading_level1, '_pydantic_config')


def test_heading_load_config_node_config_root():
    """测试BaseHeadingNode的load_config方法（NodeConfigRoot配置）"""
    # 创建HeadingLevel1Node实例
    mock_paragraph = MagicMock(spec=Paragraph)
    heading_level1 = HeadingLevel1Node(value=mock_paragraph, level=1, paragraph=mock_paragraph)
    
    # 准备字典配置，添加缺失的section_title_re字段
    config_dict = {
        "headings": {
            "level_1": {
                "alignment": "左对齐",
                "space_before": "1行",
                "space_after": "0.5行",
                "line_spacing": "1.5倍",
                "line_spacingrule": "单倍行距",
                "first_line_indent": "0字符",
                "builtin_style_name": "标题1",
                "chinese_font_name": "黑体",
                "english_font_name": "Times New Roman",
                "font_size": "二号",
                "font_color": "BLACK",
                "bold": True,
                "italic": False,
                "underline": False,
                "section_title_re": "^第[一二三四五六七八九十]+章"
            }
        }
    }
    
    # 执行方法
    heading_level1.load_config(config_dict)
    
    # 验证结果
    assert hasattr(heading_level1, '_config')
    assert hasattr(heading_level1, '_pydantic_config')


def test_heading_level2_node():
    """测试HeadingLevel2Node"""
    # 创建HeadingLevel2Node实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    heading_level2 = HeadingLevel2Node(value=mock_paragraph, level=2, paragraph=mock_paragraph)
    
    # 验证属性
    assert heading_level2.LEVEL == 2
    assert heading_level2.NODE_TYPE == "headings.level_2"


def test_heading_level3_node():
    """测试HeadingLevel3Node"""
    # 创建HeadingLevel3Node实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    heading_level3 = HeadingLevel3Node(value=mock_paragraph, level=3, paragraph=mock_paragraph)
    
    # 验证属性
    assert heading_level3.LEVEL == 3
    assert heading_level3.NODE_TYPE == "headings.level_3"


# 测试references.py中的类
def test_references_base():
    """测试References的_base方法"""
    # 创建References实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    references = References(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=ReferencesTitleConfig)
    mock_config.alignment = "左对齐"
    mock_config.space_before = "1行"
    mock_config.space_after = "0.5行"
    mock_config.line_spacing = "1.5倍"
    mock_config.line_spacingrule = "单倍行距"
    mock_config.first_line_indent = "0字符"
    mock_config.builtin_style_name = "标题1"
    mock_config.chinese_font_name = "黑体"
    mock_config.english_font_name = "Times New Roman"
    mock_config.font_size = "二号"
    mock_config.font_color = "BLACK"
    mock_config.bold = True
    mock_config.italic = False
    mock_config.underline = False
    
    references._pydantic_config = mock_config
    references.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.references.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = [MagicMock()]
        mock_ps.apply_to_paragraph.return_value = [MagicMock()]
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.references.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = [MagicMock()]
            mock_cs.apply_to_run.return_value = [MagicMock()]
            mock_cs.to_string.return_value = "测试差异"
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            result = references._base(doc, True, True)
            
            # 验证结果
            assert result == []
            assert references.add_comment.called


def test_reference_entry_base():
    """测试ReferenceEntry的_base方法"""
    # 创建ReferenceEntry实例
    mock_paragraph = MagicMock(spec=Paragraph)
    mock_paragraph.runs = [MagicMock(spec=Run)]
    reference_entry = ReferenceEntry(value=mock_paragraph, level=0, paragraph=mock_paragraph)
    
    # 模拟配置
    mock_config = MagicMock(spec=ReferencesContentConfig)
    mock_config.alignment = "左对齐"
    mock_config.space_before = "0.5行"
    mock_config.space_after = "0.5行"
    mock_config.line_spacing = "1.5倍"
    mock_config.line_spacingrule = "单倍行距"
    mock_config.first_line_indent = "2字符"
    mock_config.builtin_style_name = "正文"
    mock_config.chinese_font_name = "宋体"
    mock_config.english_font_name = "Times New Roman"
    mock_config.font_size = "小四"
    mock_config.font_color = "BLACK"
    mock_config.bold = False
    mock_config.italic = False
    mock_config.underline = False
    
    reference_entry._pydantic_config = mock_config
    reference_entry.add_comment = MagicMock()
    
    # 模拟ParagraphStyle和CharacterStyle
    with patch('wordformat.rules.references.ParagraphStyle') as mock_paragraph_style:
        mock_ps = MagicMock()
        mock_ps.diff_from_paragraph.return_value = [MagicMock()]
        mock_ps.apply_to_paragraph.return_value = [MagicMock()]
        mock_paragraph_style.return_value = mock_ps
        
        with patch('wordformat.rules.references.CharacterStyle') as mock_character_style:
            mock_cs = MagicMock()
            mock_cs.diff_from_run.return_value = [MagicMock()]
            mock_cs.apply_to_run.return_value = [MagicMock()]
            mock_cs.to_string.return_value = "测试差异"
            mock_character_style.return_value = mock_cs
            
            # 执行方法
            doc = MagicMock()
            result = reference_entry._base(doc, True, True)
            
            # 验证结果
            assert result == []
            assert reference_entry.add_comment.called
