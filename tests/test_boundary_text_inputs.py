"""边界输入测试：文本解析、run 级改写、Markdown 解析。

这些用例不追求覆盖率，而是用空值、非字符串、超范围数值、异常分隔符、
结构残缺的段落等输入，检查函数在边界处是否给出可预期的结果。
"""

import base64

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from wordformat.markdown.parser import parse_markdown
from wordformat.rules.node import FormatNode
from wordformat.text_edit import _distribute, paragraph_text, set_paragraph_text
from wordformat.utils._text import (
    _format_number,
    _from_chinese_num,
    _from_roman,
    _to_chinese_num,
    _to_roman,
    count_chinese_chars,
    extract_chinese_chars,
    has_chinese,
    is_chinese_char,
    parse_caption_text,
    resolve_item_text,
)

_PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ"
    "AAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def _para(texts=()):
    para = Document().add_paragraph()
    for text in texts:
        para.add_run(text)
    return para


def _add_hyperlink(para, text):
    link = OxmlElement("w:hyperlink")
    run = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    run.append(t)
    link.append(run)
    para._p.append(link)
    return link


def _add_smart_tag(para, text):
    tag = OxmlElement("w:smartTag")
    run = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    run.append(t)
    tag.append(run)
    para._p.append(tag)
    return tag


def _add_simple_field(para, text):
    field = OxmlElement("w:fldSimple")
    run = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    run.append(t)
    field.append(run)
    para._p.append(field)
    return field


# ---------------------------------------------------------------------------
# resolve_item_text
# ---------------------------------------------------------------------------


class TestResolveItemTextBoundary:
    @pytest.mark.parametrize(
        "value",
        [None, "", "原文", 0, 1.5, [], (), set(), object()],
        ids=lambda v: type(v).__name__,
    )
    def test_non_dict_returns_empty(self, value):
        assert resolve_item_text(value) == ""

    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            ({}, ""),
            ({"paragraph": "正文"}, "正文"),
            ({"paragraph": ""}, ""),
            ({"paragraph": None}, ""),
            ({"paragraph": 123}, ""),
            ({"paragraph": ["正文"]}, ""),
            # replace 一存在且为 str 就以它为准，空串表示清空而非未提供
            ({"replace": ""}, ""),
            ({"replace": "", "paragraph": "原文"}, ""),
            ({"replace": None, "paragraph": "回退"}, "回退"),
            ({"replace": 0, "paragraph": "回退"}, "回退"),
            ({"replace": "替换", "paragraph": "原文"}, "替换"),
            ({"replace": " 保留空格 "}, " 保留空格 "),
            ({"replace": "替换", "paragraph": 123}, "替换"),
        ],
    )
    def test_replace_wins_over_paragraph(self, value, expected):
        assert resolve_item_text(value) == expected


# ---------------------------------------------------------------------------
# 罗马数字 / 中文数字 / 编号格式
# ---------------------------------------------------------------------------


class TestRomanBoundary:
    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (0, "0"),
            (-1, "0"),
            (-3999, "0"),
            (1, "i"),
            (4, "iv"),
            (9, "ix"),
            (49, "xlix"),
            (3999, "mmmcmxcix"),
            (4000, "mmmm"),
        ],
    )
    def test_to_roman(self, num, expected):
        assert _to_roman(num) == expected

    @pytest.mark.parametrize(
        ("text", "expected"),
        [
            ("i", 1),
            ("IIII", 4),
            ("VX", 5),
            ("  X  ", 10),
            ("MCMXCIV", 1994),
            ("mmxxvi", 2026),
        ],
    )
    def test_from_roman_lenient_inputs(self, text, expected):
        assert _from_roman(text) == expected

    @pytest.mark.parametrize("text", ["", "   ", "a", "i i", "Ⅳ", "i-v"])
    def test_from_roman_rejects_invalid(self, text):
        with pytest.raises(ValueError):
            _from_roman(text)

    @pytest.mark.parametrize("num", [1, 2, 4, 5, 9, 10, 14, 40, 90, 400, 1994, 3999])
    def test_roman_roundtrip(self, num):
        assert _from_roman(_to_roman(num)) == num


class TestChineseNumeralBoundary:
    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (0, "0"),
            (-5, "-5"),
            (1, "一"),
            (9, "九"),
            (10, "十"),
            (11, "十一"),
            (19, "十九"),
            (20, "二十"),
            (99, "九十九"),
            (100, "一百"),
            (101, "101"),
            (110, "110"),
            (1000, "1000"),
        ],
    )
    def test_to_chinese(self, num, expected):
        assert _to_chinese_num(num) == expected

    @pytest.mark.parametrize(
        ("text", "expected"),
        [
            ("一", 1),
            ("零", 0),
            ("十", 10),
            ("十一", 11),
            ("十二", 12),
            ("二十", 20),
            ("二十三", 23),
            ("一百", 100),
            ("一百零一", 101),
            ("一零一", 101),
            ("一百二十三", 123),
            ("壹佰贰拾叁", 123),
            (" 十二 ", 12),
            # 重复/跳跃单位按顺序累加，不抛错
            ("十十", 20),
            ("十百", 110),
        ],
    )
    def test_from_chinese_lenient_inputs(self, text, expected):
        assert _from_chinese_num(text) == expected

    @pytest.mark.parametrize("text", ["", "   ", "两", "一 二", "十二abc", "１２"])
    def test_from_chinese_rejects_invalid(self, text):
        with pytest.raises(ValueError):
            _from_chinese_num(text)


class TestFormatNumberBoundary:
    @pytest.mark.parametrize(
        ("fmt", "num", "expected"),
        [
            ("decimal", 0, "0"),
            ("decimal", -3, "-3"),
            ("upperRoman", 0, "0"),
            ("upperRoman", 4, "IV"),
            ("lowerRoman", 4, "iv"),
            ("upperLetter", 0, "0"),
            ("upperLetter", 1, "A"),
            ("upperLetter", 26, "Z"),
            ("upperLetter", 27, "27"),
            ("lowerLetter", 0, "0"),
            ("lowerLetter", 26, "z"),
            ("lowerLetter", 27, "27"),
            ("chineseCountingThousand", 12, "十二"),
            ("chineseCounting", 100, "一百"),
            ("ideographTraditional", 9, "九"),
            ("unknownFormat", 12, "12"),
            ("", 7, "7"),
        ],
    )
    def test_format_number(self, fmt, num, expected):
        assert _format_number(num, fmt) == expected


# ---------------------------------------------------------------------------
# 题注解析
# ---------------------------------------------------------------------------


class TestCaptionParserBoundary:
    @pytest.mark.parametrize(
        "text",
        [
            "",
            "   ",
            "续",
            "续 ",
            "续续图1.1 名称",
            "图1-1",  # 缺少名称
            "图 名称",  # 缺少编号
            "图片1.1 名称",  # 标签只支持“图/表”单字
            "图1.-1 名称",
            "1.1 名称",  # 缺少标签
        ],
    )
    def test_unparsable(self, text):
        assert parse_caption_text(text) is None

    @pytest.mark.parametrize(
        ("text", "expected"),
        [
            (
                "图1-1 架构图",
                {
                    "label": "图",
                    "chapter_text": "1",
                    "chapter_num": 1,
                    "separator": "-",
                    "number_text": "1",
                    "number_num": 1,
                    "name": "架构图",
                    "is_continued": False,
                },
            ),
            (
                "表2.1 参数",
                {
                    "label": "表",
                    "chapter_num": 2,
                    "number_num": 1,
                    "name": "参数",
                    "is_continued": False,
                },
            ),
            (
                "续表2.1 参数",
                {"label": "表", "chapter_num": 2, "is_continued": True},
            ),
            (
                "续 表2.1 参数",
                {"label": "表", "chapter_num": 2, "is_continued": True},
            ),
            (
                "图 1 . 1  系 统",
                {"chapter_num": 1, "number_num": 1, "name": "系 统"},
            ),
            (
                "图IV.2 系统",
                {"chapter_text": "IV", "chapter_num": 4, "number_num": 2},
            ),
            ("图iv.2 系统", {"chapter_num": 4, "number_num": 2}),
            ("图一.2 系统", {"chapter_text": "一", "chapter_num": 1}),
            ("表2.三 名称", {"number_text": "三", "number_num": 3}),
            ("表2.IV 名称", {"number_text": "IV", "number_num": 4}),
            ("表2.iv 名称", {"number_num": 4}),
            ("表2.十二 名称", {"number_text": "十二", "number_num": 12}),
            ("图0.0 空号", {"chapter_num": 0, "number_num": 0, "name": "空号"}),
            ("表1.1 多 名 称", {"name": "多 名 称"}),
            ("图1—1 破折号", {"separator": "—"}),
            ("图1–1 短破折号", {"separator": "–"}),
            ("图1:1 冒号", {"separator": ":"}),
        ],
    )
    def test_parsed_fields(self, text, expected):
        result = parse_caption_text(text)
        assert result is not None, text
        for key, value in expected.items():
            assert result[key] == value, f"{text}: {key}"

    def test_chinese_tens_number_is_parsed(self):
        assert parse_caption_text("表2.十 名称")["number_num"] == 10

    def test_chinese_chapter_and_chinese_number(self):
        parsed = parse_caption_text("图一.三 名称")
        assert parsed is not None
        assert parsed["chapter_num"] == 1
        assert parsed["number_num"] == 3


# ---------------------------------------------------------------------------
# CJK 字符检测
# ---------------------------------------------------------------------------


class TestCjkBoundary:
    @pytest.mark.parametrize(
        ("char", "expected"),
        [
            ("一", True),
            ("鿿", True),
            ("\u4dbf", False),
            ("a", False),
            (" ", False),
            ("〇", False),
            ("，", False),
        ],
    )
    def test_is_chinese_char_boundaries(self, char, expected):
        assert is_chinese_char(char) is expected

    def test_count_and_extract_mixed_text(self):
        text = "中文abc123，。测试"
        assert count_chinese_chars(text) == 4
        assert extract_chinese_chars(text) == "中文测试"

    @pytest.mark.parametrize(
        ("text", "expected"),
        [("", False), ("   ", False), ("abc 123", False), ("a中", True), ("㐀", False)],
    )
    def test_has_chinese(self, text, expected):
        assert has_chinese(text) is expected


# ---------------------------------------------------------------------------
# _distribute：新文本按字符对齐回 run 序列
# ---------------------------------------------------------------------------


class TestDistributeBoundary:
    @pytest.mark.parametrize(
        ("segments", "text", "expected"),
        [
            ([], "x", None),
            ([], "", None),
            ([""], "x", ["x"]),
            ([""], "", [""]),
            (["", ""], "x", ["x", ""]),
            (["", ""], "", ["", ""]),
            (["ab", "cd"], "abcdX", ["ab", "cdX"]),
            (["ab", "cd"], "Xabcd", ["Xab", "cd"]),
            (["ab", "cd"], "aXcd", ["aX", "cd"]),
            (["ab", "cd"], "ab", ["ab", ""]),
            (["ab", "cd"], "", ["", ""]),
            (["相同", "相同"], "相同", ["相同", ""]),
        ],
    )
    def test_distribute(self, segments, text, expected):
        assert _distribute(segments, text) == expected


# ---------------------------------------------------------------------------
# set_paragraph_text：无变化、清空、空段、受保护结构
# ---------------------------------------------------------------------------


class TestSetParagraphTextBoundary:
    def test_no_change_returns_false(self):
        para = _para(["甲", "乙"])
        xml_before = para._p.xml

        assert set_paragraph_text(para, "甲乙") is False
        assert para._p.xml == xml_before

    def test_clear_text_keeps_runs(self):
        para = _para(["要删除的文字"])

        assert set_paragraph_text(para, "") is True
        assert para.text == ""
        assert len(para.runs) == 1

    def test_empty_paragraph_with_empty_text_is_noop(self):
        para = _para([])

        assert set_paragraph_text(para, "") is False
        assert para.text == ""

    def test_empty_paragraph_gets_new_run(self):
        para = _para([])

        assert set_paragraph_text(para, "全新内容") is True
        assert para.text == "全新内容"

    def test_whitespace_only_text_can_be_cleared(self):
        para = _para([" ", "\u00a0"])

        assert set_paragraph_text(para, "") is True
        assert para.text == ""

    def test_image_only_paragraph_keeps_image(self, tmp_path):
        png = tmp_path / "fig.png"
        png.write_bytes(_PNG_1PX)
        para = Document().add_paragraph()
        para.add_run().add_picture(str(png))

        assert set_paragraph_text(para, "") is False
        assert len(para._p.xpath(".//w:drawing")) == 1

    def test_image_only_paragraph_appends_text_run(self, tmp_path):
        png = tmp_path / "fig.png"
        png.write_bytes(_PNG_1PX)
        para = Document().add_paragraph()
        para.add_run().add_picture(str(png))

        assert set_paragraph_text(para, "新题注") is True
        assert para.text == "新题注"
        assert len(para._p.xpath(".//w:drawing")) == 1

    def test_simple_field_only_paragraph_appends_text_run(self):
        para = _para([])
        _add_simple_field(para, "42")

        assert set_paragraph_text(para, "正文") is True
        assert para.text == "正文"
        assert len(para._p.xpath(".//w:fldSimple")) == 1

    def test_smart_tag_text_is_writable_but_not_in_paragraph_text(self):
        para = _para(["正文"])
        _add_smart_tag(para, "智能")

        assert paragraph_text(para) == "正文智能"
        assert para.text == "正文"

        assert set_paragraph_text(para, "正文智能2") is True
        assert (
            "".join(t.text or "" for t in para._p.xpath(".//w:smartTag//w:t"))
            == "智能2"
        )

    def test_hyperlink_text_writable_without_losing_link(self):
        para = _para(["见 "])
        _add_hyperlink(para, "文献1")

        assert set_paragraph_text(para, "参见 文献1") is True
        assert para.text == "参见 文献1"
        link = para._p.find(qn("w:hyperlink"))
        assert link is not None
        assert "".join(t.text or "" for t in link.xpath(".//w:t")) == "文献1"


# ---------------------------------------------------------------------------
# FormatNode.apply_replace：JSON replace 字段
# ---------------------------------------------------------------------------


class TestApplyReplaceBoundary:
    @staticmethod
    def _node(value, para):
        return FormatNode(value=value, level=0, paragraph=para)

    @pytest.mark.parametrize(
        "value",
        [
            None,
            "文本",
            123,
            [],
            {},
            {"replace": None},
            {"replace": 123},
            {"replace": ["x"]},
        ],
    )
    def test_invalid_value_returns_false(self, value):
        para = _para(["原文"])
        assert self._node(value, para).apply_replace(Document()) is False
        assert para.text == "原文"

    def test_missing_paragraph_returns_false(self):
        node = FormatNode(value={"replace": "新文本"}, level=0)
        assert node.apply_replace(Document()) is False

    def test_empty_replace_clears_text(self):
        para = _para(["原文", "续写"])
        assert self._node({"replace": ""}, para).apply_replace(Document()) is True
        assert para.text == ""

    def test_replace_strips_all_spaces(self):
        para = _para(["原文"])
        assert self._node({"replace": " 新 文 本 "}, para).apply_replace() is True
        assert para.text == "新文本"

    @pytest.mark.parametrize(
        ("old", "new", "expected"),
        [
            (["ab", "cd"], "abcdX", ["ab", "cdX"]),
            (["ab", "cd"], "Xabcd", ["Xa", "bcd"]),
            (["ab", "cd"], "ab", ["ab", ""]),
            (["ab", "cd"], "", ["", ""]),
            (["甲", "乙", "丙"], "甲乙丙丁", ["甲", "乙", "丙丁"]),
        ],
    )
    def test_multi_run_replacement_keeps_run_count(self, old, new, expected):
        para = _para(old)
        assert self._node({"replace": new}, para).apply_replace() is True
        assert [r.text for r in para.runs] == expected

    def test_replacing_paragraph_with_one_run(self):
        para = _para(["原文"])
        assert self._node({"replace": "新文"}, para).apply_replace() is True
        assert [r.text for r in para.runs] == ["新文"]

    def test_replace_on_runless_paragraph_inserts_text(self):
        para = _para([])
        node = self._node({"replace": "新增"}, para)

        assert node.apply_replace(Document()) is True
        assert para.text == "新增"

    def test_empty_replace_preserves_image_run(self, tmp_path):
        png = tmp_path / "fig.png"
        png.write_bytes(_PNG_1PX)
        para = Document().add_paragraph()
        para.add_run().add_picture(str(png))
        para.add_run("旧题注")

        self._node({"replace": ""}, para).apply_replace(Document())

        assert len(para._p.xpath(".//w:drawing")) == 1

    def test_non_empty_replace_preserves_image_run(self, tmp_path):
        png = tmp_path / "fig.png"
        png.write_bytes(_PNG_1PX)
        para = Document().add_paragraph()
        para.add_run().add_picture(str(png))
        para.add_run("旧题注")

        self._node({"replace": "新题注"}, para).apply_replace(Document())

        assert len(para._p.xpath(".//w:drawing")) == 1


# ---------------------------------------------------------------------------
# Markdown 解析
# ---------------------------------------------------------------------------


class TestMarkdownParserBoundary:
    @pytest.mark.parametrize("text", ["", "   ", "\n\n", "---", "***"])
    def test_empty_or_break_only(self, text):
        assert parse_markdown(text) == []

    @pytest.mark.parametrize(
        ("text", "category"),
        [
            ("# 一级", "heading_level_1"),
            ("## 二级", "heading_level_2"),
            ("### 三级", "heading_level_3"),
            ("#### 四级", "heading_level_3"),
            ("###### 六级", "heading_level_3"),
        ],
    )
    def test_heading_levels_clamp_to_three(self, text, category):
        result = parse_markdown(text)
        assert len(result) == 1
        assert result[0]["category"] == category

    def test_image_with_alt_creates_figure_then_caption(self):
        result = parse_markdown("![架构图](a.png)")

        assert [item["category"] for item in result] == [
            "figure_image",
            "caption_figure",
        ]
        assert result[0]["paragraph"] == "a.png"
        assert result[1]["paragraph"] == "架构图"

    def test_image_without_alt_creates_only_figure(self):
        result = parse_markdown("![](a.png)")

        assert [item["category"] for item in result] == ["figure_image"]

    def test_inline_image_with_text_stays_body_text(self):
        result = parse_markdown("前文 ![图](a.png) 后文")

        assert result[0]["category"] == "body_text"
        assert result[0]["paragraph"] == "前文 图 后文"
        assert any(
            mark.get("image_url") == "a.png" for mark in result[0]["inline_marks"]
        )

    def test_code_block_keeps_indentation_and_skips_blank_lines(self):
        result = parse_markdown("```python\n  keep me\n\n    indented\n```")

        assert [item["paragraph"] for item in result] == ["  keep me", "    indented"]
        assert all(item["category"] == "body_text" for item in result)
        assert result[0]["inline_marks"] == [{"text": "  keep me", "code": True}]

    def test_block_math(self):
        result = parse_markdown("$$\nx = 1\n$$")

        assert len(result) == 1
        assert result[0]["category"] == "math_block"
        assert "x = 1" in result[0]["paragraph"]
        assert result[0]["inline_marks"][0]["math_display"] is True

    def test_table_rows(self):
        result = parse_markdown("| 列A | 列B |\n| --- | --- |\n| 1 | 2 |")

        assert len(result) == 1
        assert result[0]["category"] == "table_object"
        assert result[0]["table_rows"] == [["列A", "列B"], ["1", "2"]]

    def test_list_items_become_paragraphs(self):
        result = parse_markdown("- 第一项\n- 第二项")

        assert [item["paragraph"] for item in result] == ["第一项", "第二项"]
        assert all(item["category"] == "body_text" for item in result)

    def test_blockquote_flattens_to_body_text(self):
        result = parse_markdown("> 引用内容")

        assert len(result) == 1
        assert result[0]["category"] == "body_text"
        assert result[0]["paragraph"] == "引用内容"

    def test_inline_marks_cover_all_format_types(self):
        result = parse_markdown(
            "**粗体** *斜体* ~~删除~~ `代码` [链接](https://e.com) $x+1$"
        )
        marks = result[0]["inline_marks"]

        assert any(m.get("bold") for m in marks)
        assert any(m.get("italic") for m in marks)
        assert any(m.get("strikethrough") for m in marks)
        assert any(m.get("code") for m in marks)
        assert any(m.get("link_url") == "https://e.com" for m in marks)
        assert any(m.get("math") for m in marks)

    def test_softbreak_becomes_space_and_linebreak_becomes_newline(self):
        soft = parse_markdown("相邻\n两行")
        hard = parse_markdown("硬换行  \n下一行")

        assert soft[0]["paragraph"] == "相邻 两行"
        assert hard[0]["paragraph"] == "硬换行\n下一行"

    def test_crlf_and_cr_are_normalized(self):
        result = parse_markdown("第一行\r\n第二行\r第三行")

        assert result[0]["paragraph"] == "第一行 第二行 第三行"
