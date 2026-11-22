#! /usr/bin/env python
"""Markdown 解析器单元测试。"""

from unittest.mock import patch

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
        result = parse_markdown("```\n\na = 1\n\nb = 2\n\n```")
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

    def test_image_no_alt_warns_about_missing_caption(self):
        """无 alt 时不会生成题注段，必须留下 WARNING（论文图无题注是硬伤）。"""
        with patch("wordformat.markdown.parser.logger") as log:
            result = parse_markdown("![](noalt.png)")

        assert [r["category"] for r in result] == ["figure_image"]
        warnings = " ".join(str(c) for c in log.warning.call_args_list)
        assert "题注" in warnings and "noalt.png" in warnings

    def test_image_with_alt_does_not_warn(self):
        with patch("wordformat.markdown.parser.logger") as log:
            parse_markdown("![图1 系统架构](arch.png)")

        assert not [c for c in log.warning.call_args_list if "题注" in str(c)]

    # -- 手写题注识别（Markdown 无题注语法，作者常写「表X-Y 名称」） --------

    def test_table_caption_above_table_promoted(self):
        text = "表1-1 功能需求与指标\n\n| 功能 | 指标 |\n|---|---|\n| 登录 | 3 秒 |"
        result = parse_markdown(text)
        assert result[0]["category"] == "caption_table"
        assert result[0]["paragraph"] == "表1-1 功能需求与指标"
        assert result[1]["category"] == "table_object"

    def test_table_caption_below_table_promoted(self):
        text = "| 功能 | 指标 |\n|---|---|\n| 登录 | 3 秒 |\n\n表1-1 功能需求与指标"
        result = parse_markdown(text)
        assert result[0]["category"] == "table_object"
        assert result[1]["category"] == "caption_table"

    def test_caption_between_two_tables_belongs_to_the_next_one(self):
        """两表之间的「表X-Y」行归属下一张表（表题在上）。"""
        text = (
            "| A | B |\n|---|---|\n| 1 | 2 |\n\n"
            "表1-2 第二张表\n\n"
            "| C | D |\n|---|---|\n| 3 | 4 |"
        )
        result = parse_markdown(text)
        categories = [r["category"] for r in result]
        assert categories == ["table_object", "caption_table", "table_object"]
        assert result[1]["paragraph"] == "表1-2 第二张表"

    def test_body_sentence_starting_with_table_ref_not_promoted(self):
        """正文句「表1-1 的数据表明……」不能被当成题注。"""
        text = (
            "表1-1 的数据表明，该方法有效。\n\n"
            "| 功能 | 指标 |\n|---|---|\n| 登录 | 3 秒 |"
        )
        result = parse_markdown(text)
        assert result[0]["category"] == "body_text"
        assert result[1]["category"] == "table_object"

    def test_manual_figure_caption_promoted_without_alt(self):
        """图片无 alt 但旁边写了「图X-Y 名称」时，识别为图题注且不告警。"""
        text = "![](arch.png)\n\n图1-1 系统总体框图"
        with patch("wordformat.markdown.parser.logger") as log:
            result = parse_markdown(text)

        assert [r["category"] for r in result] == ["figure_image", "caption_figure"]
        assert result[1]["paragraph"] == "图1-1 系统总体框图"
        assert not [c for c in log.warning.call_args_list if "题注" in str(c)]

    def test_alt_caption_not_duplicated_by_manual_line(self):
        """alt 已生成题注时，相邻手写「图X-Y」行不再认领，避免两条题注。"""
        text = "图1-1 系统总体框图\n\n![图1-1 系统总体框图](arch.png)"
        result = parse_markdown(text)

        categories = [r["category"] for r in result]
        assert categories.count("caption_figure") == 1
        assert categories[0] == "body_text"

    def test_previous_figure_caption_not_counted_for_next_figure(self):
        """上一张图的题注不能当成下一张图的题注，缺题注的图仍要告警。"""
        text = "![图1-1 第一张](a.png)\n\n![](b.png)"
        with patch("wordformat.markdown.parser.logger") as log:
            result = parse_markdown(text)

        assert [r["category"] for r in result] == [
            "figure_image",
            "caption_figure",
            "figure_image",
        ]
        warnings = [str(c) for c in log.warning.call_args_list]
        assert len([w for w in warnings if "b.png" in w]) == 1

    def test_next_figure_manual_caption_still_promoted(self):
        """两张图连排时，第二张图下方的手写题注仍应认领。"""
        text = "![图1-1 第一张](a.png)\n\n![](b.png)\n\n图1-2 第二张"
        result = parse_markdown(text)

        assert [r["category"] for r in result] == [
            "figure_image",
            "caption_figure",
            "figure_image",
            "caption_figure",
        ]
        assert result[3]["paragraph"] == "图1-2 第二张"

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


class TestMathParsing:
    """测试 Markdown 中数学公式的解析。"""

    def test_inline_math_segment(self):
        result = parse_markdown("质能方程 $E=mc^2$ 是著名的")
        assert len(result) == 1
        assert result[0]["category"] == "body_text"
        marks = result[0]["inline_marks"]
        math_segs = [m for m in marks if m.get("math")]
        assert len(math_segs) == 1
        assert math_segs[0]["text"] == "E=mc^2"

    def test_inline_math_preserved_in_text(self):
        result = parse_markdown("The value $x^2$ is squared.")
        text = result[0]["paragraph"]
        assert "$x^2$" in text

    def test_standalone_display_math_block(self):
        result = parse_markdown("$$\n\\frac{1}{2}\n$$")
        assert len(result) == 1
        assert result[0]["category"] == "math_block"
        assert result[0]["paragraph"] == "\\frac{1}{2}"

    def test_display_math_in_paragraph_creates_display_segment(self):
        """行内 $$...$$ 也应被识别"""
        result = parse_markdown("公式 $$a^2 + b^2 = c^2$$ 是勾股定理")
        marks = result[0]["inline_marks"]
        math_segs = [m for m in marks if m.get("math_display")]
        assert len(math_segs) >= 1

    def test_multiple_inline_math_in_one_paragraph(self):
        result = parse_markdown("$x_1$ 和 $x_2$ 都是变量")
        marks = result[0]["inline_marks"]
        math_segs = [m for m in marks if m.get("math")]
        assert len(math_segs) == 2

    def test_math_with_subscripts(self):
        result = parse_markdown("权重 $w_{ij}$ 的计算公式如下。")
        marks = result[0]["inline_marks"]
        math_segs = [m for m in marks if m.get("math")]
        assert len(math_segs) == 1
        assert math_segs[0]["text"] == "w_{ij}"
