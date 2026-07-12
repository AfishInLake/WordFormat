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
