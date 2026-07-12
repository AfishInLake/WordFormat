"""MD pipeline stages 单元测试。"""
from pathlib import Path
from unittest.mock import patch

import pytest

from wordformat.pipeline.context import FormatContext
from wordformat.pipeline.stages_md import (
    DocumentCreationStage,
    LoadMarkdownStage,
    MarkdownParseStage,
)


class TestLoadMarkdownStage:
    def test_loads_file_content(self, tmp_path):
        md_file = tmp_path / "test.md"
        md_file.write_text("# 标题\n\n正文")
        ctx = FormatContext(md_path=str(md_file))
        stage = LoadMarkdownStage()
        result = stage.process(ctx)
        assert "# 标题" in result.md_text
        assert "正文" in result.md_text

    def test_missing_file_raises(self):
        ctx = FormatContext(md_path="/nonexistent/path.md")
        stage = LoadMarkdownStage()
        with pytest.raises(FileNotFoundError):
            stage.process(ctx)


class TestMarkdownParseStage:
    def test_parses_markdown_to_paragraphs(self, tmp_path):
        md_file = tmp_path / "test.md"
        md_file.write_text("# 标题\n\n正文段落")
        ctx = FormatContext(md_path=str(md_file))
        ctx = LoadMarkdownStage().process(ctx)
        ctx = MarkdownParseStage().process(ctx)
        assert len(ctx.paragraphs) >= 2
        categories = [p["category"] for p in ctx.paragraphs]
        assert "heading_level_1" in categories
        assert "body_text" in categories


class TestDocumentCreationStage:
    def test_creates_document_from_tree(self, root_node_fixture):
        """用简单 tree 创建文档。"""
        ctx = FormatContext()
        ctx.root_node = root_node_fixture
        ctx = DocumentCreationStage().process(ctx)
        assert ctx.document is not None
        assert len(ctx.document.paragraphs) >= 1

    def test_figure_image_creates_empty_run(self):
        """figure_image 节点只建空 run，路径留给 FigureImage。"""
        from wordformat.rules.node import FormatNode

        root = FormatNode(value={}, level=0)
        node = FormatNode(
            value={"category": "figure_image", "paragraph": "test.png"},
            level=1,
        )
        root.add_child_node(node)
        ctx = FormatContext()
        ctx.root_node = root
        ctx = DocumentCreationStage().process(ctx)
        assert node.paragraph is not None
        assert node.paragraph.runs[0].text == ""


@pytest.fixture
def root_node_fixture():
    from wordformat.rules.node import FormatNode

    root = FormatNode(value={}, level=0)
    child = FormatNode(
        value={"category": "body_text", "paragraph": "测试正文"},
        level=1,
    )
    root.add_child_node(child)
    return root
