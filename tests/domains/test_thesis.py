#! /usr/bin/env python
# @Time    : 2026/9/28
# @Author  : afish
# @File    : test_thesis.py
"""论文领域（domains/thesis.py）测试。

覆盖：
- 领域注册表与装配（发现/幂等/未知领域报错）；
- 题注章节号注入（章节递增、图/表计数器、续表保留、on_format_begin 重置）；
- 论文子树提升（Abstract/References 下 body_text 升级）；
- 检测摘要统计（中英文摘要字数、关键词数、参考文献中英文条数）。
"""

import pytest

from wordformat.domains import list_domains, load_domain, reset_domains
from wordformat.hooks import HookRegistry


def _registry_with_thesis():
    """独立 HookRegistry + 论文领域注册（不污染全局 hooks）。"""
    from wordformat.domains.thesis import register_handlers

    registry = HookRegistry()
    register_handlers(registry)
    return registry


def _node(category, text=None, doc=None, cls=None):
    """构造 FormatNode（可选带段落文本）。"""
    from wordformat.rules.node import FormatNode

    cls = cls or FormatNode
    node = cls(value={"category": category, "paragraph": text or ""}, level=0)
    if text is not None and doc is not None:
        node.paragraph = doc.add_paragraph(text)
    return node


def _emit_node(registry, category, text=None, doc=None, cls=None):
    """按 before_node_format 事件语义触发一次节点格式化前回调。"""
    node = _node(category, text, doc, cls)
    registry.emit(
        "before_node_format",
        node=node,
        paragraph=node.paragraph,
        config=None,
        check=True,
    )
    return node


# ============================================================
# 领域注册表与装配
# ============================================================


class TestDomainRegistry:
    def test_thesis_discovered(self):
        assert "thesis" in list_domains()

    def test_load_unknown_domain_raises(self):
        with pytest.raises(ValueError, match="未知文档领域"):
            load_domain("gongwen")

    def test_load_domain_idempotent(self):
        reset_domains()
        try:
            registry = HookRegistry()
            load_domain("thesis", registry)
            first = len(registry.registered_callbacks())
            assert first >= 3
            load_domain("thesis", registry)
            assert len(registry.registered_callbacks()) == first
        finally:
            reset_domains()

    def test_register_handlers_idempotent(self):
        registry = _registry_with_thesis()
        first = len(registry.registered_callbacks())
        from wordformat.domains.thesis import register_handlers

        register_handlers(registry)
        assert len(registry.registered_callbacks()) == first


# ============================================================
# 题注编号注入（thesis.caption_numbering）
# ============================================================


class TestCaptionNumbering:
    def test_chapter_and_sequence_injection(self, doc):
        registry = _registry_with_thesis()
        registry.emit("on_format_begin")

        h1 = _emit_node(registry, "heading_level_1", "第一章 绪论", doc)
        fig1 = _emit_node(registry, "caption_figure", "图1.1 系统架构", doc)
        fig2 = _emit_node(registry, "caption_figure", "图1.2 处理流程", doc)
        h2 = _emit_node(registry, "heading_level_1", "第二章 相关技术", doc)
        tab = _emit_node(registry, "caption_table", "表2.1 测试用例", doc)
        body = _emit_node(registry, "body_text", "正文内容", doc)

        # 一级标题自身取到新章节号
        assert h1.value["chapter_number"] == 1
        assert h2.value["chapter_number"] == 2
        # 题注注入章节号 + 顺序号（图、表各自独立计数）
        assert fig1.value["chapter_number"] == 1
        assert fig1.value["sequence_number"] == 1
        assert fig2.value["sequence_number"] == 2
        assert tab.value["chapter_number"] == 2
        assert tab.value["sequence_number"] == 1
        # 普通节点注入当前章节号（BodyText 第一章引用上标需要）
        assert body.value["chapter_number"] == 2

    def test_continued_caption_keeps_existing_number(self, doc):
        registry = _registry_with_thesis()
        registry.emit("on_format_begin")

        _emit_node(registry, "heading_level_1", "第一章 绪论", doc)
        _emit_node(registry, "caption_figure", "图1.1 系统架构", doc)
        cont = _emit_node(registry, "caption_table", "续表1.2 数据统计", doc)
        # 续表保留原标题注编号（1.2），不递增表计数器
        assert cont.value["chapter_number"] == 1
        assert cont.value["sequence_number"] == 2
        # 下一个普通表仍从 1 开始（续表未占用序号）
        tab = _emit_node(registry, "caption_table", "表1.1 测试用例", doc)
        assert tab.value["sequence_number"] == 1

    def test_on_format_begin_resets_counters(self, doc):
        registry = _registry_with_thesis()

        registry.emit("on_format_begin")
        _emit_node(registry, "heading_level_1", "第一章 绪论", doc)
        fig1 = _emit_node(registry, "caption_figure", "图1.1 系统架构", doc)
        assert fig1.value["sequence_number"] == 1

        # 新一轮格式化：状态重置，图序回到 1
        registry.emit("on_format_begin")
        h1 = _emit_node(registry, "heading_level_1", "第一章 绪论", doc)
        fig2 = _emit_node(registry, "caption_figure", "图1.1 系统架构", doc)
        assert h1.value["chapter_number"] == 1
        assert fig2.value["sequence_number"] == 1
        assert fig2.value["chapter_number"] == 1

    def test_caption_without_paragraph_keeps_defaults(self):
        registry = _registry_with_thesis()
        registry.emit("on_format_begin")
        fig = _emit_node(registry, "caption_figure")
        # 无段落文本：按新编号计（chapter=0 时归零容器）
        assert fig.value["chapter_number"] == 0
        assert fig.value["sequence_number"] == 1


# ============================================================
# 论文子树提升（thesis.tree_normalize）
# ============================================================


class TestTreeNormalize:
    def test_promotes_bodytext_in_thesis_subtrees(self):
        from wordformat.rules.abstract import (
            AbstractContentCN,
            AbstractContentEN,
            AbstractTitleCN,
            AbstractTitleEN,
        )
        from wordformat.rules.acknowledgement import (
            Acknowledgements,
            AcknowledgementsCN,
        )
        from wordformat.rules.body import BodyText
        from wordformat.rules.references import ReferenceEntry, References

        registry = _registry_with_thesis()

        root = _node("top")
        abs_title = _node("abstract_chinese_title", cls=AbstractTitleCN)
        abs_body = _node("body_text", cls=BodyText)
        abs_title.children.append(abs_body)

        abs_title_en = _node("abstract_english_title", cls=AbstractTitleEN)
        abs_body_en = _node("body_text", cls=BodyText)
        abs_title_en.children.append(abs_body_en)

        ref_title = _node("references_title", cls=References)
        ref_body = _node("body_text", cls=BodyText)
        ref_title.children.append(ref_body)

        ack_title = _node("acknowledgements_title", cls=Acknowledgements)
        ack_body = _node("body_text", cls=BodyText)
        ack_title.children.append(ack_body)

        root.children = [abs_title, abs_title_en, ref_title, ack_title]

        registry.emit("on_tree_normalize", root_node=root, config=None, check=True)

        assert isinstance(abs_title.children[0], AbstractContentCN)
        assert isinstance(abs_title_en.children[0], AbstractContentEN)
        assert isinstance(ref_title.children[0], ReferenceEntry)
        assert isinstance(ack_title.children[0], AcknowledgementsCN)

    def test_leaves_other_subtrees_untouched(self):
        from wordformat.rules.body import BodyText
        from wordformat.rules.node import FormatNode

        registry = _registry_with_thesis()
        root = _node("top")
        body = _node("body_text", cls=BodyText)
        root.children.append(body)
        registry.emit("on_tree_normalize", root_node=root, config=None, check=True)
        assert type(body) is BodyText
        assert isinstance(root.children[0], FormatNode)


# ============================================================
# 检测摘要统计（thesis.summary_build）
# ============================================================


class TestSummaryBuild:
    @pytest.fixture(autouse=True)
    def _stash_error_stats(self):
        from wordformat.rules.node import FormatNode

        original = dict(FormatNode._error_stats)
        FormatNode._error_stats = {"total": 3, "错误": 2, "提醒": 1}
        yield
        FormatNode._error_stats = original

    def test_summary_counts_abstract_keywords_references(self, doc):
        from wordformat.rules.abstract import AbstractContentCN, AbstractContentEN
        from wordformat.rules.keywords import KeywordsCN, KeywordsEN
        from wordformat.rules.references import ReferenceEntry

        registry = _registry_with_thesis()
        root = _node("top")

        abs_cn = _node("abstract_chinese_content", cls=AbstractContentCN)
        abs_cn.paragraph = doc.add_paragraph("本文研究深度学习与机器学习技术。")
        abs_en = _node("abstract_english_content", cls=AbstractContentEN)
        abs_en.paragraph = doc.add_paragraph("This paper studies deep learning.")
        kw_cn = _node("abstract_chinese_keywords", cls=KeywordsCN)
        kw_cn.paragraph = doc.add_paragraph("关键词：深度学习；机器学习；神经网络")
        kw_en = _node("abstract_english_keywords", cls=KeywordsEN)
        kw_en.paragraph = doc.add_paragraph("Keywords: deep learning; machine learning")
        ref_cn = _node("references_content", "参考文献一", doc, ReferenceEntry)
        ref_en = _node("references_content", "Reference one", doc, ReferenceEntry)
        from wordformat.rules.object import TableObject

        table = TableObject(
            value={
                "category": "table_object",
                "paragraph": "",
                "table_rows": [["a", "b"]],
                "exists": True,
            },
            level=999,
        )
        root.children = [abs_cn, abs_en, kw_cn, kw_en, ref_cn, ref_en, table]

        result = registry.emit(
            "on_summary_build", root_node=root, document=doc, config=None, check=True
        )
        summary = result.get("summary", "")

        assert "检测结果：" in summary
        assert "检测错误数：3，万字差错率" in summary
        assert "错误：2，提醒：1" in summary
        assert "中文摘要：规范：300字左右，原文：" in summary
        assert "英文摘要：规范：300字左右，原文：" in summary
        assert "中文关键词：规范：3-5个，原文：3个" in summary
        assert "英文关键词：规范：3-5个，原文：2个" in summary
        assert "参考文献：规范：不少于15条，原文：中文1条;外文1条" in summary
        assert "表格：1 个（表格内容不参与检查）" in summary

    def test_summary_ignores_sync_skipped_table(self, doc):
        # af --sync 忽略的新增表格未写入文档，不应计入表格数量
        from wordformat.rules.object import TableObject

        registry = _registry_with_thesis()
        root = _node("top")
        root.children = [
            TableObject(
                value={
                    "category": "table_object",
                    "paragraph": "",
                    "table_rows": [["a"]],
                    "sync_skip": True,
                },
                level=999,
            )
        ]
        result = registry.emit(
            "on_summary_build", root_node=root, document=doc, config=None, check=True
        )
        assert "表格：" not in result.get("summary", "")

    def test_summary_skipped_on_apply_mode(self, doc):
        registry = _registry_with_thesis()
        result = registry.emit(
            "on_summary_build",
            root_node=_node("top"),
            document=doc,
            config=None,
            check=False,
        )
        assert result.get("summary") is None
