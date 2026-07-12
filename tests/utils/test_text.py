"""utils/_text.py 测试 — 罗马数字、中文数字、题注解析。"""
import pytest
from wordformat.utils import _from_chinese_num, _from_roman, parse_caption_text

class TestFromChineseNum:
    def test_single_digit(self):
        assert _from_chinese_num("一") == 1
        assert _from_chinese_num("九") == 9

    def test_ten(self):
        assert _from_chinese_num("十") == 10

    def test_teens(self):
        assert _from_chinese_num("十一") == 11
        assert _from_chinese_num("十五") == 15

    def test_tens(self):
        assert _from_chinese_num("二十") == 20
        assert _from_chinese_num("九十九") == 99

    def test_hundred(self):
        assert _from_chinese_num("一百") == 100

    def test_hundreds_complex(self):
        assert _from_chinese_num("一百二十三") == 123

    def test_financial_digits(self):
        assert _from_chinese_num("壹") == 1
        assert _from_chinese_num("叁") == 3

    def test_empty_raises(self):
        with pytest.raises(ValueError):
            _from_chinese_num("")

    def test_invalid_raises(self):
        with pytest.raises(ValueError):
            _from_chinese_num("abc")


# ======================== parse_caption_text ========================

class TestFromRoman:
    def test_basic_singles(self):
        assert _from_roman("I") == 1
        assert _from_roman("V") == 5
        assert _from_roman("X") == 10

    def test_subtractive(self):
        assert _from_roman("IV") == 4
        assert _from_roman("IX") == 9
        assert _from_roman("XL") == 40
        assert _from_roman("CM") == 900

    def test_composite(self):
        assert _from_roman("XIV") == 14
        assert _from_roman("XXVII") == 27

    def test_case_insensitive(self):
        assert _from_roman("iv") == 4

    def test_empty_raises(self):
        with pytest.raises(ValueError):
            _from_roman("")

    def test_invalid_raises(self):
        with pytest.raises(ValueError):
            _from_roman("ABC")


# ======================== _from_chinese_num ========================

class TestParseCaptionText:
    def test_basic_figure_arabic(self):
        result = parse_caption_text("图1.1 系统架构图")
        assert result is not None
        assert result["label"] == "图"
        assert result["chapter_num"] == 1
        assert result["separator"] == "."
        assert result["number_num"] == 1
        assert result["name"] == "系统架构图"

    def test_hyphen_separator(self):
        result = parse_caption_text("图1-1 数据流程图")
        assert result is not None
        assert result["separator"] == "-"

    def test_colon_separator(self):
        result = parse_caption_text("表1:1 测试数据")
        assert result["separator"] == ":"

    def test_em_dash_separator(self):
        result = parse_caption_text("图1—1 架构设计")
        assert result is not None
        assert result["separator"] == "—"

    def test_en_dash_separator(self):
        result = parse_caption_text("图1–1 测试图")
        assert result["separator"] == "–"

    def test_chinese_chapter_number(self):
        result = parse_caption_text("图一.1 系统架构图")
        assert result is not None
        assert result["chapter_text"] == "一"
        assert result["chapter_num"] == 1

    def test_roman_chapter_number(self):
        result = parse_caption_text("图I.1 系统架构图")
        assert result["chapter_text"] == "I"
        assert result["chapter_num"] == 1

    def test_roman_chapter_lowercase(self):
        result = parse_caption_text("图ii.1 数据图")
        assert result["chapter_num"] == 2

    def test_fullwidth_space(self):
        result = parse_caption_text("图1.1　全角空格名称")
        assert result is not None
        assert result["name"] == "全角空格名称"

    def test_empty_returns_none(self):
        assert parse_caption_text("") is None

    def test_plain_text_returns_none(self):
        assert parse_caption_text("这是一段普通正文") is None

    def test_no_space_before_name_returns_none(self):
        """编号后无空格无法可靠提取名称，返回 None。"""
        result = parse_caption_text("图1.1测试")
        assert result is None

    def test_space_after_label(self):
        """标签后有空格：图 1.1 测试。"""
        result = parse_caption_text("图 1.1 系统架构图")
        assert result is not None
        assert result["label"] == "图"
        assert result["chapter_num"] == 1
        assert result["number_num"] == 1
        assert result["name"] == "系统架构图"

    def test_continued_table_caption(self):
        """续表 5.3 API接口测试结果 → 正确解析"""
        result = parse_caption_text("续表5.3 API接口测试结果")
        assert result is not None
        assert result["label"] == "表"
        assert result["chapter_num"] == 5
        assert result["separator"] == "."
        assert result["number_num"] == 3
        assert result["name"] == "API接口测试结果"
        assert result["is_continued"] is True

    def test_continued_figure_caption(self):
        """续图 2.1 系统架构图 → 正确解析"""
        result = parse_caption_text("续图2.1 系统架构图")
        assert result is not None
        assert result["label"] == "图"
        assert result["chapter_num"] == 2
        assert result["number_num"] == 1
        assert result["name"] == "系统架构图"
        assert result["is_continued"] is True

    def test_continued_caption_with_space_after_label(self):
        """续 表 5.3 xxx → 去掉续后正常匹配。"""
        result = parse_caption_text("续 表 5.3 测试表格")
        assert result is not None
        assert result["label"] == "表"
        assert result["is_continued"] is True
        assert result["name"] == "测试表格"

    def test_continued_table_with_hyphen(self):
        """续表5-3 测试 → 连字符分隔符"""
        result = parse_caption_text("续表5-3 测试")
        assert result is not None
        assert result["label"] == "表"
        assert result["chapter_num"] == 5
        assert result["separator"] == "-"
        assert result["number_num"] == 3
        assert result["is_continued"] is True

    def test_regular_caption_not_continued(self):
        """普通题注 is_continued 为 False"""
        result = parse_caption_text("表5.3 测试")
        assert result is not None
        assert result["is_continued"] is False


# ======================== _replace_paragraph_text ========================
