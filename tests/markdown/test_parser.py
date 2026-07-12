#! /usr/bin/env python
"""Markdown 解析器单元测试。"""

import pytest

from wordformat.markdown.parser import parse_markdown


class TestParseMarkdown:
    """基础解析测试。"""

    def test_empty_text(self):
        result = parse_markdown("")
        assert result == []

    def test_blank_lines_only(self):
        result = parse_markdown("\n\n\n")
        assert result == []

    def test_single_heading_h1(self):
        result = parse_markdown("# 第一章 绪论")
        assert len(result) == 1
        assert result[0]["category"] == "heading_level_1"
        assert result[0]["paragraph"] == "第一章 绪论"
        assert result[0]["score"] == 1.0

    def test_single_heading_h2(self):
        result = parse_markdown("## 1.1 研究背景")
        assert len(result) == 1
        assert result[0]["category"] == "heading_level_2"
        assert result[0]["paragraph"] == "1.1 研究背景"

    def test_heading_h4_maps_to_level_3(self):
        result = parse_markdown("#### 四级标题")
        assert len(result) == 1
        assert result[0]["category"] == "heading_level_3"

    def test_paragraph(self):
        result = parse_markdown("这是一段正文。")
        assert len(result) == 1
        assert result[0]["category"] == "body_text"
        assert result[0]["paragraph"] == "这是一段正文。"

    def test_multiple_paragraphs(self):
        text = "第一段。\n\n第二段。"
        result = parse_markdown(text)
        assert len(result) == 2
        assert result[0]["category"] == "body_text"
        assert result[0]["paragraph"] == "第一段。"
        assert result[1]["category"] == "body_text"
        assert result[1]["paragraph"] == "第二段。"

    def test_heading_and_paragraphs(self):
        text = "# 标题\n\n正文内容。\n\n更多正文。"
        result = parse_markdown(text)
        assert len(result) == 3
        assert result[0]["category"] == "heading_level_1"
        assert result[1]["category"] == "body_text"
        assert result[2]["category"] == "body_text"


class TestHeadingHierarchy:
    """测试多级标题的层级映射。"""

    def test_deep_hierarchy(self):
        text = "# 一级\n\n## 二级\n\n### 三级\n\n正文。"
        result = parse_markdown(text)
        assert len(result) == 4
        assert result[0]["category"] == "heading_level_1"
        assert result[1]["category"] == "heading_level_2"
        assert result[2]["category"] == "heading_level_3"
        assert result[3]["category"] == "body_text"


class TestInlineFormatting:
    """测试行内格式标记（inline_marks）的提取。"""

    def test_bold_text(self):
        result = parse_markdown("这是**粗体**文字。")
        assert len(result) == 1
        item = result[0]
        marks = item["inline_marks"]
        # 应有三个片段：普通 + 粗体 + 普通
        assert len(marks) >= 2
        bold_segs = [m for m in marks if m.get("bold")]
        assert len(bold_segs) == 1
        assert bold_segs[0]["text"] == "粗体"

    def test_italic_text(self):
        result = parse_markdown("这是*斜体*文字。")
        marks = result[0]["inline_marks"]
        italic_segs = [m for m in marks if m.get("italic")]
        assert len(italic_segs) == 1
        assert italic_segs[0]["text"] == "斜体"

    def test_plain_text_has_no_marks(self):
        result = parse_markdown("普通正文。")
        marks = result[0]["inline_marks"]
        for m in marks:
            assert not m.get("bold")
            assert not m.get("italic")

    def test_heading_no_inline_marks_used(self):
        """标题中的 inline_marks 也存在但不影响使用。"""
        result = parse_markdown("# 标题")
        assert "inline_marks" in result[0]


class TestSpecialBlocks:
    """测试特殊块的处理。"""

    def test_code_block_becomes_body_text(self):
        result = parse_markdown('```\nprint("hello")\n```')
        assert len(result) == 1
        assert result[0]["category"] == "body_text"
        assert 'print("hello")' in result[0]["paragraph"]

    def test_list_flattens_to_paragraphs(self):
        text = "- 第一项\n- 第二项"
        result = parse_markdown(text)
        assert len(result) == 2
        assert all(r["category"] == "body_text" for r in result)

    def test_standalone_image_becomes_figure(self):
        result = parse_markdown("![](test.png)")
        assert len(result) == 1
        assert result[0]["category"] == "figure_image"
        assert result[0]["paragraph"] == "test.png"

    def test_thematic_break_skipped(self):
        result = parse_markdown("---")
        assert len(result) == 0

    def test_table_with_header_and_body(self):
        text = "| A | B |\n|----|----|\n| 1 | 2 |\n| 3 | 4 |"
        result = parse_markdown(text)
        assert len(result) == 1
        assert result[0]["category"] == "table_object"
        rows = result[0]["table_rows"]
        assert len(rows) == 3  # header + 2 data rows
        assert rows[0] == ["A", "B"]

    def test_block_quote_flattens(self):
        result = parse_markdown("> 引用文字")
        assert len(result) == 1
        assert result[0]["category"] == "body_text"
        assert "引用文字" in result[0]["paragraph"]

    def test_code_block_empty_lines_skipped(self):
        result = parse_markdown('```\n\na = 1\n\nb = 2\n\n```')
        assert len(result) == 2
        assert result[0]["paragraph"] == "a = 1"
        assert result[1]["paragraph"] == "b = 2"

    def test_paragraph_with_inline_image_kept_as_body(self):
        """段落中文字+图片混合时保持为 body_text"""
        result = parse_markdown("文字 ![](img.png) 更多文字")
        assert len(result) == 1
        assert result[0]["category"] == "body_text"

    def test_image_with_alt_creates_caption_and_figure(self):
        result = parse_markdown("![图1.1 系统架构](arch.png)")
        assert len(result) == 2
        assert result[0]["category"] == "figure_image"
        assert result[0]["paragraph"] == "arch.png"
        assert result[1]["category"] == "caption_figure"
        assert result[1]["paragraph"] == "图1.1 系统架构"

    def test_image_no_alt_creates_figure_only(self):
        result = parse_markdown("![](noalt.png)")
        assert len(result) == 1
        assert result[0]["category"] == "figure_image"
        assert result[0]["paragraph"] == "noalt.png"

    def test_strikethrough_text(self):
        result = parse_markdown("这是~~删除~~文字")
        assert len(result) == 1
        marks = result[0]["inline_marks"]
        strike_segs = [m for m in marks if m.get("strikethrough")]
        assert len(strike_segs) == 1
        assert strike_segs[0]["text"] == "删除"

    def test_link_text(self):
        result = parse_markdown("访问[链接](https://example.com)")
        marks = result[0]["inline_marks"]
        link_segs = [m for m in marks if m.get("link_url")]
        assert len(link_segs) == 1
        assert link_segs[0]["link_url"] == "https://example.com"

    def test_crlf_normalized(self):
        result = parse_markdown("第一行\r\n第二行\r第三行")
        assert len(result) >= 1
        text = result[0]["paragraph"]
        assert "\r" not in text

    def test_codespan_in_text(self):
        result = parse_markdown("使用 `wordf md` 命令")
        text = result[0]["paragraph"]
        assert "wordf md" in text
        assert "`" not in text

    def test_linebreak_in_paragraph(self):
        result = parse_markdown("第一行  \n第二行")
        text = result[0]["paragraph"]
        assert "\n" in text  # linebreak preserved
