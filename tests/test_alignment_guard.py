"""对齐阶段防错位校验：节点数与段落数不一致时必须显式报错。

背景：前端曾过滤 figure_image 占位节点，导致节点数组与文档段落
失去 1:1 对应，ParagraphAlignmentStage 的 zip 静默错位，
keywords_chinese/caption_figure/heading_level_* 等规则被应用到错误段落
（表现为正文被加"图1.1"前缀、中文关键词批注出现在论文题目上）。
"""
import pytest
from docx import Document

from wordformat.pipeline.context import FormatContext
from wordformat.pipeline.stages import ParagraphAlignmentStage
from wordformat.rules.node import FormatNode


def _make_ctx(tree_children: int, para_count: int) -> FormatContext:
    doc = Document()
    for _ in range(para_count):
        doc.add_paragraph("测试段落")
    root = FormatNode(value={"category": "top"}, level=0)
    for _ in range(tree_children):
        root.add_child_node(FormatNode(value={"category": "body_text"}, level=1))
    return FormatContext(document=doc, root_node=root)


def test_alignment_raises_when_node_count_mismatch():
    """节点数(1) != 段落数(2) 时抛 ValueError，防止静默错位。"""
    ctx = _make_ctx(tree_children=1, para_count=2)
    with pytest.raises(ValueError, match="不一致"):
        ParagraphAlignmentStage().process(ctx)


def test_alignment_raises_when_nodes_missing_figure_image_placeholder():
    """模拟旧版前端：过滤 figure_image 占位后节点数少于段落数，必须报错。"""
    ctx = _make_ctx(tree_children=2, para_count=3)
    with pytest.raises(ValueError, match="不一致"):
        ParagraphAlignmentStage().process(ctx)


def test_alignment_passes_when_counts_match():
    """节点数 == 段落数时正常对齐，每个节点拿到对应段落。"""
    ctx = _make_ctx(tree_children=2, para_count=2)
    result = ParagraphAlignmentStage().process(ctx)
    aligned = result.root_node.children
    assert all(n.paragraph is not None for n in aligned)
    assert [n.paragraph.text for n in aligned] == ["测试段落", "测试段落"]