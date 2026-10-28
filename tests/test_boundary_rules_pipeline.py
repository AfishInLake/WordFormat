"""边界输入测试：XML 原子操作、Hook、页眉页脚、身份、同步与遍历。"""

import base64
from types import SimpleNamespace

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from wordformat.handlers import (
    _anchor_comments,
    _header_footer,
    _post_processing,
    register_builtin_handlers,
)
from wordformat.hooks import HookRegistry, hooks
from wordformat.para_ids import (
    W14_PARA_ID,
    ensure_para_ids,
    ensure_table_ids,
    get_para_id,
    get_table_id,
)
from wordformat.pipeline.sync_apply import (
    StructureSyncStage,
    _body_tables,
    _identity_ready,
    _item_key,
    _item_text,
    _normalize,
    _table_identity_ready,
    expand_multi_paragraph_items,
    sync_document_structure,
)
from wordformat.rules.node import FormatNode, TreeNode
from wordformat.style.xml_ops import (
    ensure_pPr,
    ensure_rPr,
    line_rule_to_xml,
    line_spacing_val_to_xml,
    pPr_set_alignment,
    rPr_set_bold,
    rPr_set_font,
    rPr_set_font_color,
    rPr_set_font_size,
    rPr_set_italic,
    rPr_set_underline,
)
from wordformat.tree import (
    Tree,
    _format_node_label,
    _tree_node_category,
    bfs_walk,
    dfs_walk,
    render_tree,
)
from wordformat.utils._docx import (
    ensure_directory_exists,
    para_contains_image,
    para_contains_math,
    remove_all_numbering,
)
from wordformat.utils._fs import ensure_is_directory, get_file_name

_PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ"
    "AAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def _para(texts=()):
    para = Document().add_paragraph()
    for text in texts:
        para.add_run(text)
    return para


def _png(tmp_path):
    path = tmp_path / "pixel.png"
    path.write_bytes(_PNG_1PX)
    return str(path)


# ---------------------------------------------------------------------------
# style/xml_ops.py
# ---------------------------------------------------------------------------


class TestXmlOpsBoundary:
    @pytest.mark.parametrize(
        ("pt", "expected"),
        [
            (0, "0"),
            (10.24, "20"),
            (10.25, "20"),  # 银行家舍入：20.5 → 20
            (10.26, "21"),
            (10.5, "21"),
            (10.75, "22"),
            (-1.5, "-3"),
        ],
    )
    def test_font_size_half_point_rounding(self, pt, expected):
        rPr = OxmlElement("w:rPr")

        rPr_set_font_size(rPr, pt)
        rPr_set_font_size(rPr, pt)  # 重复设置不产生重复元素

        assert rPr.find(qn("w:sz")).get(qn("w:val")) == expected
        assert rPr.find(qn("w:szCs")).get(qn("w:val")) == expected
        assert len(rPr.findall(qn("w:sz"))) == 1
        assert len(rPr.findall(qn("w:szCs"))) == 1

    def test_font_color_removes_all_old_colors(self):
        rPr = OxmlElement("w:rPr")
        for value in ("000000", "FF0000"):
            old = OxmlElement("w:color")
            old.set(qn("w:val"), value)
            rPr.append(old)

        rPr_set_font_color(rPr, (1, 2, 3))

        colors = rPr.findall(qn("w:color"))
        assert len(colors) == 1
        assert colors[0].get(qn("w:val")) == "010203"

    def test_font_color_extremes(self):
        rPr = OxmlElement("w:rPr")

        rPr_set_font_color(rPr, (0, 0, 0))
        assert rPr.find(qn("w:color")).get(qn("w:val")) == "000000"

        rPr_set_font_color(rPr, (255, 255, 255))
        assert rPr.find(qn("w:color")).get(qn("w:val")) == "FFFFFF"

    def test_font_setters_ignore_missing_or_empty_values(self):
        rPr = OxmlElement("w:rPr")

        rPr_set_font(rPr, None, None)
        rPr_set_font(rPr, "", "")
        fonts = rPr.find(qn("w:rFonts"))
        assert fonts is not None
        assert fonts.get(qn("w:eastAsia")) is None
        assert fonts.get(qn("w:ascii")) is None

        rPr_set_font(rPr, cn_name="宋体")
        assert rPr.find(qn("w:rFonts")).get(qn("w:eastAsia")) == "宋体"

        rPr_set_font(rPr, en_name="Arial")
        fonts = rPr.find(qn("w:rFonts"))
        assert fonts.get(qn("w:ascii")) == "Arial"
        assert fonts.get(qn("w:hAnsi")) == "Arial"

    def test_font_set_reuses_existing_rfonts(self):
        rPr = OxmlElement("w:rPr")
        fonts = OxmlElement("w:rFonts")
        rPr.append(fonts)

        rPr_set_font(rPr, "黑体", "Times New Roman")

        assert rPr.find(qn("w:rFonts")) is fonts
        assert len(rPr.findall(qn("w:rFonts"))) == 1

    def test_bold_toggle_is_idempotent(self):
        rPr = OxmlElement("w:rPr")

        rPr_set_bold(rPr, True)
        rPr_set_bold(rPr, True)
        assert len(rPr.findall(qn("w:b"))) == 1
        assert len(rPr.findall(qn("w:bCs"))) == 1

        rPr_set_bold(rPr, False)
        assert rPr.find(qn("w:b")) is None
        assert rPr.find(qn("w:bCs")) is None
        rPr_set_bold(rPr, False)  # 已不存在时不报错

    def test_italic_and_underline_toggle(self):
        rPr = OxmlElement("w:rPr")

        rPr_set_italic(rPr, True)
        rPr_set_italic(rPr, True)
        assert len(rPr.findall(qn("w:i"))) == 1
        rPr_set_italic(rPr, False)
        assert rPr.find(qn("w:i")) is None

        rPr_set_underline(rPr, True)
        rPr_set_underline(rPr, True)
        underlines = rPr.findall(qn("w:u"))
        assert len(underlines) == 1
        assert underlines[0].get(qn("w:val")) == "single"
        rPr_set_underline(rPr, False)
        assert rPr.find(qn("w:u")) is None

    @pytest.mark.parametrize(
        ("alignment", "expected"),
        [
            (WD_ALIGN_PARAGRAPH.CENTER, "center"),
            (WD_ALIGN_PARAGRAPH.LEFT, "left"),
            (WD_ALIGN_PARAGRAPH.JUSTIFY, "both"),
            (999, "left"),
            (None, "left"),
        ],
    )
    def test_alignment_mapping_and_unknown_fallback(self, alignment, expected):
        pPr = OxmlElement("w:pPr")
        pPr_set_alignment(pPr, alignment)
        pPr_set_alignment(pPr, alignment)

        assert len(pPr.findall(qn("w:jc"))) == 1
        assert pPr.find(qn("w:jc")).get(qn("w:val")) == expected

    @pytest.mark.parametrize(
        ("value", "unit", "expected"),
        [
            (1.5, "倍", 360),
            (0, "倍", 0),
            (-1.5, "倍", -360),
            (12, "pt", 240),
            (0, "pt", 0),
            (2, "未知单位", 480),
        ],
    )
    def test_line_spacing_conversion(self, value, unit, expected):
        assert line_spacing_val_to_xml(value, unit) == expected

    @pytest.mark.parametrize(
        ("rule", "expected"),
        [
            (WD_LINE_SPACING.SINGLE, "auto"),
            (WD_LINE_SPACING.ONE_POINT_FIVE, "auto"),
            (WD_LINE_SPACING.AT_LEAST, "atLeast"),
            (WD_LINE_SPACING.EXACTLY, "exact"),
            (999, "auto"),
            (None, "auto"),
        ],
    )
    def test_line_rule_mapping(self, rule, expected):
        assert line_rule_to_xml(rule) == expected

    def test_ensure_rpr_inserts_at_front_and_is_idempotent(self):
        style = OxmlElement("w:style")
        style.append(OxmlElement("w:name"))

        rPr = ensure_rPr(style)

        assert style[0] is rPr
        assert ensure_rPr(style) is rPr
        assert len(style.findall(qn("w:rPr"))) == 1

    def test_ensure_ppr_appends_and_is_idempotent(self):
        style = OxmlElement("w:style")

        pPr = ensure_pPr(style)

        assert style[-1] is pPr
        assert ensure_pPr(style) is pPr
        assert len(style.findall(qn("w:pPr"))) == 1


# ---------------------------------------------------------------------------
# hooks.py
# ---------------------------------------------------------------------------


class TestHookRegistryBoundary:
    def test_custom_event_is_allowed_with_warning(self):
        reg = HookRegistry()
        reg.register("custom_event", lambda d: {"seen": True}, name="cb")

        assert reg.is_registered("custom_event") is True
        assert reg.emit("custom_event") == {"seen": True}

    def test_callback_returning_non_dict_is_ignored(self):
        reg = HookRegistry()
        reg.register("on_comment", lambda d: "不是字典", name="cb")

        assert reg.emit("on_comment", text="原文") == {"text": "原文"}

    def test_mutation_and_override_apply_in_order(self):
        reg = HookRegistry()

        def first(data):
            data["issues"].append("a")

        reg.register("on_caption_check", first, name="first")
        reg.register("on_caption_check", lambda d: {"text": "覆盖"}, name="second")

        result = reg.emit("on_caption_check", issues=[], text="原")
        assert result["issues"] == ["a"]
        assert result["text"] == "覆盖"

    def test_emit_without_callbacks_returns_new_dict(self):
        reg = HookRegistry()
        source = {"a": 1}

        result = reg.emit("on_comment", **source)

        assert result == source
        assert result is not source

    def test_unregister_duplicate_names_removes_all(self):
        reg = HookRegistry()
        reg.register("on_comment", lambda d: None, name="dup")
        reg.register("on_comment", lambda d: None, name="dup")

        assert reg.unregister("on_comment", "dup") is True
        assert reg.is_registered("on_comment") is False
        assert reg.unregister("on_comment", "dup") is False

    def test_unregister_none_on_missing_event(self):
        assert HookRegistry().unregister("on_comment", None) is False

    def test_registered_callbacks_none_aggregates_all(self):
        reg = HookRegistry()
        reg.register("on_comment", lambda d: None, name="a")
        reg.register("on_caption_check", lambda d: None, name="b")

        assert len(reg.registered_callbacks()) == 2
        assert reg.registered_callbacks("missing") == []

    def test_has_handlers_after_partial_unregister(self):
        reg = HookRegistry()
        reg.register("on_comment", lambda d: None, name="a")
        reg.register("on_caption_check", lambda d: None, name="b")

        reg.unregister("on_comment", "a")
        assert reg.has_handlers is True

        reg.unregister("on_caption_check", "b")
        assert reg.has_handlers is False


# ---------------------------------------------------------------------------
# handlers.py
# ---------------------------------------------------------------------------


class TestHandlerBoundary:
    def test_post_processing_skips_on_check(self, monkeypatch):
        calls = []
        monkeypatch.setattr(
            "wordformat.hyperlinks.create_citation_hyperlinks",
            lambda *a: calls.append(("hyper", a)),
        )
        monkeypatch.setattr(
            "wordformat.numbering.process_heading_numbering",
            lambda *a: calls.append(("number", a)),
        )
        ctx = SimpleNamespace(
            root_node=object(),
            document=Document(),
            config_model=SimpleNamespace(numbering=SimpleNamespace(enabled=True)),
        )

        _post_processing({"check": True, "ctx": ctx})

        assert calls == []

    def test_post_processing_skips_without_ctx_or_root(self, monkeypatch):
        calls = []
        monkeypatch.setattr(
            "wordformat.hyperlinks.create_citation_hyperlinks",
            lambda *a: calls.append(("hyper", a)),
        )
        monkeypatch.setattr(
            "wordformat.numbering.process_heading_numbering",
            lambda *a: calls.append(("number", a)),
        )

        _post_processing({"ctx": None})
        _post_processing({"ctx": SimpleNamespace(root_node=None)})

        assert calls == []

    def test_post_processing_with_numbering_disabled_still_links_citations(
        self, monkeypatch
    ):
        ctx = SimpleNamespace(
            root_node=object(),
            document=Document(),
            config_model=SimpleNamespace(numbering=SimpleNamespace(enabled=False)),
        )
        calls = []
        monkeypatch.setattr(
            "wordformat.hyperlinks.create_citation_hyperlinks",
            lambda *a: calls.append(("hyper", a)),
        )
        monkeypatch.setattr(
            "wordformat.numbering.process_heading_numbering",
            lambda *a: calls.append(("number", a)),
        )

        _post_processing({"ctx": ctx, "check": False})

        assert [name for name, _ in calls] == ["hyper"]

    def test_post_processing_with_numbering_enabled(self, monkeypatch):
        ctx = SimpleNamespace(
            root_node=object(),
            document=Document(),
            config_model=SimpleNamespace(
                numbering=SimpleNamespace(enabled=True),
                headings="headings-config",
            ),
        )
        calls = []
        monkeypatch.setattr(
            "wordformat.hyperlinks.create_citation_hyperlinks",
            lambda *a: calls.append(("hyper", a)),
        )
        monkeypatch.setattr(
            "wordformat.numbering.process_heading_numbering",
            lambda *a: calls.append(("number", a)),
        )

        _post_processing({"ctx": ctx})

        assert [name for name, _ in calls] == ["number", "hyper"]
        assert calls[0][1][3] == "headings-config"

    def test_header_footer_without_document(self):
        _header_footer({})

    def test_header_footer_check_without_issues(self, monkeypatch):
        calls = []
        monkeypatch.setattr(
            "wordformat.header_footer.check_header_footer",
            lambda doc, cfg: calls.append((doc, cfg)) or [],
        )
        monkeypatch.setattr(
            "wordformat.handlers._anchor_comments",
            lambda *a: calls.append("anchor"),
        )

        _header_footer({"document": Document(), "config": "cfg", "check": True})

        assert len(calls) == 1
        assert "anchor" not in calls

    def test_header_footer_check_anchors_issues(self, monkeypatch):
        issues = ["页眉-错误：缺失"]
        monkeypatch.setattr(
            "wordformat.header_footer.check_header_footer",
            lambda doc, cfg: issues,
        )
        anchored = []
        monkeypatch.setattr(
            "wordformat.handlers._anchor_comments",
            lambda doc, got: anchored.append(got),
        )

        _header_footer({"document": Document(), "config": "cfg", "check": True})

        assert anchored == [issues]

    def test_header_footer_apply_anchors_remaining(self, monkeypatch):
        remaining = ["页脚-提醒：残留"]
        monkeypatch.setattr(
            "wordformat.header_footer.apply_header_footer",
            lambda doc, cfg: remaining,
        )
        anchored = []
        monkeypatch.setattr(
            "wordformat.handlers._anchor_comments",
            lambda doc, got: anchored.append(got),
        )

        _header_footer({"document": Document(), "config": "cfg", "check": False})

        assert anchored == [remaining]

    def test_anchor_comments_without_paragraphs_does_not_count(self):
        FormatNode.reset_stats()
        try:
            _anchor_comments(Document(), ["页眉-错误：缺失"])
            assert FormatNode._error_stats == {"total": 0, "错误": 0, "提醒": 0}
        finally:
            FormatNode.reset_stats()

    def test_anchor_comments_adds_run_and_counts_by_severity(self):
        doc = Document()
        para = doc.add_paragraph()
        assert para.runs == []

        FormatNode.reset_stats()
        try:
            _anchor_comments(
                doc,
                ["页眉-字号错误：缺失", "页脚-标点问题：半角"],
            )
            assert len(para.runs) >= 1
            assert para.runs[0].text == ""
            assert FormatNode._error_stats["total"] == 2
            assert FormatNode._error_stats["错误"] == 1
            assert FormatNode._error_stats["提醒"] == 1
        finally:
            FormatNode.reset_stats()

    def test_anchor_comments_unknown_severity_falls_back_to_error(self):
        doc = Document()
        doc.add_paragraph("正文")

        FormatNode.reset_stats()
        try:
            _anchor_comments(doc, ["页眉-提醒：未知类型"])
            assert FormatNode._error_stats["错误"] == 1
            assert FormatNode._error_stats["提醒"] == 0
        finally:
            FormatNode.reset_stats()

    def test_register_builtin_handlers_is_idempotent(self):
        reg = HookRegistry()
        register_builtin_handlers(reg)
        first = len(reg.registered_callbacks())
        register_builtin_handlers(reg)

        assert len(reg.registered_callbacks()) == first

    def test_register_builtin_handlers_force_duplicates(self):
        reg = HookRegistry()
        register_builtin_handlers(reg)
        first = len(reg.registered_callbacks())
        register_builtin_handlers(reg, force=True)

        assert len(reg.registered_callbacks()) == first * 2


# ---------------------------------------------------------------------------
# para_ids.py
# ---------------------------------------------------------------------------


class TestParaIdBoundary:
    def test_empty_document(self):
        doc = Document()

        assert ensure_para_ids(doc) == 0
        assert ensure_table_ids(doc) == 0

    def test_table_without_rows(self):
        doc = Document()
        table = doc.add_table(rows=1, cols=1)
        for tr in table._tbl.findall(qn("w:tr")):
            table._tbl.remove(tr)

        assert get_table_id(table._tbl) is None
        assert ensure_table_ids(doc) == 0

    def test_body_ids_do_not_touch_table_cells(self):
        doc = Document()
        doc.add_paragraph("正文")
        table = doc.add_table(rows=1, cols=1)
        cell_para = table.rows[0].cells[0].paragraphs[0]

        ensure_para_ids(doc)

        assert get_para_id(doc.paragraphs[0]._p) is not None
        assert get_para_id(cell_para._p) is None

    def test_table_anchor_is_first_cell_paragraph(self):
        doc = Document()
        table = doc.add_table(rows=1, cols=1)
        table.rows[0].cells[0].text = ""

        ensure_table_ids(doc)

        anchor = table.rows[0].cells[0].paragraphs[0]
        assert get_table_id(table._tbl) == get_para_id(anchor._p)


# ---------------------------------------------------------------------------
# sync_apply.py 辅助函数
# ---------------------------------------------------------------------------


class TestSyncHelperBoundary:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [(None, ""), ("", ""), ("  x  ", "x"), ("\u3000x\u3000", "x")],
    )
    def test_normalize(self, value, expected):
        assert _normalize(value) == expected

    @pytest.mark.parametrize(
        ("item", "expected"),
        [
            ({"paragraph": "  旧  "}, "旧"),
            ({"paragraph": "旧", "replace": " 新 "}, "新"),
            ({"paragraph": 123}, ""),
            ({"category": "figure_image", "paragraph": "/a.png"}, ""),
            ({"category": "table_object", "paragraph": "x"}, ""),
        ],
    )
    def test_item_text(self, item, expected):
        assert _item_text(item) == expected

    def test_item_key(self):
        assert _item_key({"paragraph": " 正文 "}) == "正文"
        assert (
            _item_key({"category": "figure_image", "paragraph": "/a.png"}) == "<image>"
        )
        assert _item_key({"category": "table_object"}) == ""


class TestExpandMultiParagraphBoundary:
    @pytest.mark.parametrize("item", ["文本", None, 123, ["a"]])
    def test_non_dict_items_are_kept(self, item):
        assert expand_multi_paragraph_items([item]) == [item]

    def test_missing_paragraphs_is_kept(self):
        item = {"category": "body_text", "paragraph": "x"}
        assert expand_multi_paragraph_items([item]) == [item]

    def test_none_paragraphs_is_kept(self):
        item = {"category": "body_text", "paragraphs": None}
        assert expand_multi_paragraph_items([item]) == [item]

    @pytest.mark.parametrize("paragraphs", [[], ["only"], "ab", {}, ("a", "b")])
    def test_invalid_paragraphs_raise(self, paragraphs):
        with pytest.raises(ValueError, match="至少 2 段"):
            expand_multi_paragraph_items(
                [{"category": "body_text", "paragraphs": paragraphs}]
            )

    @pytest.mark.parametrize("category", ["table_object", "figure_image"])
    def test_unsupported_categories_raise(self, category):
        with pytest.raises(ValueError, match="不支持 paragraphs"):
            expand_multi_paragraph_items(
                [{"category": category, "paragraphs": ["a", "b"]}]
            )

    def test_replace_conflict_raises(self):
        with pytest.raises(ValueError, match="只能选一"):
            expand_multi_paragraph_items(
                [
                    {
                        "category": "body_text",
                        "paragraphs": ["a", "b"],
                        "replace": "x",
                    }
                ]
            )

    def test_first_inherits_identity_siblings_drop_it(self):
        items = expand_multi_paragraph_items(
            [
                {
                    "category": "body_text",
                    "paragraphs": ["第一段", "第二段", "第三段"],
                    "para_id": "1234ABCD",
                    "score": 0.9,
                }
            ]
        )

        first, second, third = items
        assert first["paragraph"] == "第一段"
        assert first["para_id"] == "1234ABCD"
        assert first["score"] == 0.9
        assert "paragraphs" not in first
        assert "replace" not in first

        assert second["paragraph"] == "第二段"
        assert second["comment"] == "由多段落条目拆分"
        assert "para_id" not in second
        assert "replace" not in second
        assert third["paragraph"] == "第三段"

    def test_empty_texts_are_allowed(self):
        items = expand_multi_paragraph_items(
            [{"category": "body_text", "paragraphs": ["", ""]}]
        )

        assert [item["paragraph"] for item in items] == ["", ""]

    def test_non_string_elements_are_rejected(self):
        with pytest.raises(ValueError):
            expand_multi_paragraph_items(
                [{"category": "body_text", "paragraphs": [None, 123]}]
            )


class TestIdentityReadinessBoundary:
    def test_no_para_id_anywhere(self):
        doc = Document()
        doc.add_paragraph("A")

        assert _identity_ready(doc, [{"paragraph": "A"}]) is False

    def test_non_dict_items_are_ignored(self):
        doc = Document()
        doc.add_paragraph("A")

        assert _identity_ready(doc, ["x", None, 1]) is False

    def test_all_document_paragraphs_have_unique_ids(self):
        doc = Document()
        doc.add_paragraph("A")
        doc.add_paragraph("B")
        ensure_para_ids(doc)
        ids = [get_para_id(p._p) for p in doc.paragraphs]

        items = [{"paragraph": "A", "para_id": ids[0]}]
        assert _identity_ready(doc, items) is True

    def test_document_paragraph_missing_id_blocks_identity_mode(self):
        doc = Document()
        doc.add_paragraph("A")
        doc.add_paragraph("B")
        ensure_para_ids(doc)
        doc.paragraphs[1]._p.attrib.pop(W14_PARA_ID)

        items = [{"paragraph": "A", "para_id": get_para_id(doc.paragraphs[0]._p)}]
        assert _identity_ready(doc, items) is False

    def test_duplicate_document_ids_block_identity_mode(self):
        doc = Document()
        doc.add_paragraph("A")
        doc.add_paragraph("B")
        doc.paragraphs[0]._p.set(W14_PARA_ID, "11111111")
        doc.paragraphs[1]._p.set(W14_PARA_ID, "11111111")

        assert (
            _identity_ready(doc, [{"paragraph": "A", "para_id": "11111111"}]) is False
        )

    def test_empty_document_with_id_item_is_ready(self):
        # body_ids 为空时 all([]) 为真；条目都会走新增分支
        assert _identity_ready(Document(), [{"para_id": "11111111"}]) is True


class TestTableIdentityReadinessBoundary:
    def test_no_tables_and_no_table_items(self):
        assert _table_identity_ready([], []) is True

    def test_table_without_id_blocks_identity_mode(self):
        doc = Document()
        doc.add_table(rows=1, cols=1)

        assert _table_identity_ready([], _body_tables(doc.element.body)) is False

    def test_old_json_table_item_without_table_id_blocks(self):
        doc = Document()
        doc.add_table(rows=1, cols=1)
        ensure_table_ids(doc)

        items = [{"category": "table_object", "exists": True}]
        assert _table_identity_ready(items, _body_tables(doc.element.body)) is False

    def test_missing_table_items_need_para_id_evidence(self):
        doc = Document()
        doc.add_table(rows=1, cols=1)
        ensure_table_ids(doc)
        tables = _body_tables(doc.element.body)

        assert _table_identity_ready([{"paragraph": "A"}], tables) is False
        assert _table_identity_ready([{"para_id": "11111111"}], tables) is True

    def test_new_table_item_without_id_is_allowed(self):
        doc = Document()
        doc.add_table(rows=1, cols=1)
        ensure_table_ids(doc)

        items = [{"category": "table_object", "table_rows": [["x"]]}]
        assert _table_identity_ready(items, _body_tables(doc.element.body)) is True

    def test_existing_table_item_with_id_is_allowed(self):
        doc = Document()
        doc.add_table(rows=1, cols=1)
        ensure_table_ids(doc)
        table_id = get_table_id(_body_tables(doc.element.body)[0])

        items = [{"category": "table_object", "exists": True, "table_id": table_id}]
        assert _table_identity_ready(items, _body_tables(doc.element.body)) is True


class TestStructureSyncStageBoundary:
    def test_empty_document_with_only_table_item_gets_placeholder(self):
        ctx = SimpleNamespace(
            document=Document(),
            paragraphs=[{"category": "table_object", "paragraph": ""}],
            json_path=None,
        )

        result = StructureSyncStage().process(ctx)

        assert result is ctx
        assert ctx.paragraphs[-1] == {
            "category": "body_text",
            "paragraph": "",
            "score": 1.0,
        }
        assert len(ctx.document.paragraphs) == 1

    def test_normal_items_are_synced_without_placeholder(self):
        ctx = SimpleNamespace(
            document=Document(),
            paragraphs=[{"category": "body_text", "paragraph": "新段"}],
            json_path=None,
        )

        StructureSyncStage().process(ctx)

        assert [p.text for p in ctx.document.paragraphs] == ["新段"]
        assert len(ctx.paragraphs) == 1

    def test_identity_items_insert_into_empty_document(self):
        doc = Document()
        items = [{"category": "body_text", "paragraph": "A", "para_id": "11111111"}]

        sync_document_structure(doc, items)

        assert [p.text for p in doc.paragraphs] == ["A"]
        assert get_para_id(doc.paragraphs[0]._p) is not None


# ---------------------------------------------------------------------------
# rules/node.py：批注缓冲、边缘空格、样式 handler
# ---------------------------------------------------------------------------


class TestNodeCommentBoundary:
    def test_whitespace_only_comment_is_skipped(self):
        doc = Document()
        para = doc.add_paragraph("正文")
        node = FormatNode(value={}, level=0, paragraph=para)

        node.add_comment(doc, para.runs[0], "   ")

        assert node._comment_texts == []

    def test_flush_without_runs_clears_buffer(self):
        doc = Document()
        para = doc.add_paragraph()
        node = FormatNode(value={}, level=0, paragraph=para)
        node.add_comment(doc, [], "正文-字号错误：x，规范：y")

        node._flush_comments(doc)

        assert node._comment_texts == []

    def test_comments_group_by_run_and_paragraph(self):
        doc = Document()
        para = doc.add_paragraph()
        first = para.add_run("甲")
        second = para.add_run("乙")
        node = FormatNode(value={}, level=0, paragraph=para)
        node.add_comment(doc, first, "正文-字号错误：x，规范：y")
        node.add_comment(doc, first, "正文-加粗错误：x，规范：y")
        node.add_comment(doc, second, "正文-标点问题：x，规范：y")
        node.add_comment(doc, para.runs, "正文-对齐错误：x，规范：y")

        node._flush_comments(doc)

        assert len(list(doc.comments)) == 3

    def test_on_comment_hook_can_override_text(self):
        doc = Document()
        para = doc.add_paragraph("正文")
        node = FormatNode(value={}, level=0, paragraph=para)

        def rewrite(data):
            return {"text": data["text"] + "(hook)"}

        hooks.register("on_comment", rewrite, name="boundary.test")
        try:
            node.add_comment(doc, para.runs[0], "原批注")
        finally:
            hooks.unregister("on_comment", "boundary.test")

        assert node._comment_texts[0][1] == "原批注(hook)"


class TestEdgeSpaceCleanupBoundary:
    def test_no_paragraph_or_no_runs(self):
        node = FormatNode(value={}, level=0)
        node._clean_paragraph_edge_spaces()  # 不抛异常

        para = Document().add_paragraph()
        node = FormatNode(value={}, level=0, paragraph=para)
        node._clean_paragraph_edge_spaces()
        assert para.text == ""

    def test_strips_first_leading_and_last_trailing_only(self):
        para = _para(["  开头", "中间", "结尾  "])
        FormatNode(value={}, level=0, paragraph=para)._clean_paragraph_edge_spaces()

        assert [r.text for r in para.runs] == ["开头", "中间", "结尾"]

    def test_space_only_runs_are_skipped_when_finding_edges(self):
        # 纯空白 run 清空后继续找真正有内容的 run
        para = _para(["  ", " 中 ", "\u00a0尾\u00a0"])
        FormatNode(value={}, level=0, paragraph=para)._clean_paragraph_edge_spaces()

        assert [r.text for r in para.runs] == ["", "中 ", "\u00a0尾"]

    def test_manual_line_breaks_are_removed(self):
        para = _para(["a"])
        break_run = para.add_run()
        break_run.add_break()
        para.add_run("b")
        doc = Document()
        node = FormatNode(value={}, level=0, paragraph=para)

        node.apply_format(doc)

        assert len(para._p.findall(qn("w:br"))) == 0
        assert para.text == "ab"


class TestStyleHandlerBoundary:
    def test_empty_paragraph_skips_paragraph_checks(self):
        from wordformat.rules import BodyText

        doc = Document()
        para = doc.add_paragraph()
        node = BodyText(value={}, level=0, paragraph=para)
        node.load_config({})
        node.protected_inline = False

        node.check_format(doc)

        assert node._comment_texts == []

    def test_character_style_skips_whitespace_only_runs(self):
        from wordformat.rules import BodyText

        doc = Document()
        para = doc.add_paragraph()
        para.add_run("   ")
        node = BodyText(value={}, level=0, paragraph=para)
        node.load_config({})

        node._handle_character_style(doc, None, p=True)

        assert node._comment_texts == []

    def test_minority_font_size_is_not_reported(self):
        from wordformat.rules import BodyText

        doc = Document()
        para = doc.add_paragraph()

        major = para.add_run("长文本内容")
        major.font.size = Pt(12)
        major.font.name = "Times New Roman"
        major_rPr = major._element.get_or_add_rPr()
        major_rFonts = major_rPr.find(qn("w:rFonts"))
        if major_rFonts is None:
            major_rFonts = OxmlElement("w:rFonts")
            major_rPr.insert(0, major_rFonts)
        major_rFonts.set(qn("w:eastAsia"), "宋体")
        major.font.color.rgb = RGBColor(0, 0, 0)
        major.bold = False
        major.italic = False
        major.underline = False

        minor = para.add_run("短")
        minor.font.size = Pt(10)

        node = BodyText(value={}, level=0, paragraph=para)
        node.load_config({})
        node._handle_character_style(doc, None, p=True)

        assert not any("字号错误" in text for _runs, text in node._comment_texts)

    def test_filter_emphasis(self):
        diffs = [
            SimpleNamespace(diff_type="bold"),
            SimpleNamespace(diff_type="italic"),
            SimpleNamespace(diff_type="underline"),
            SimpleNamespace(diff_type="font_size"),
        ]

        assert FormatNode._filter_emphasis(diffs, False) == diffs
        assert [d.diff_type for d in FormatNode._filter_emphasis(diffs, True)] == [
            "font_size"
        ]


# ---------------------------------------------------------------------------
# utils/_fs.py 与 utils/_docx.py
# ---------------------------------------------------------------------------


class TestFsBoundary:
    @pytest.mark.parametrize(
        ("path", "expected"),
        [
            ("a.md", "a"),
            ("/tmp/dir/a.tar.gz", "a.tar"),
            ("noext", "noext"),
            (".hidden", ".hidden"),
            ("dir/", ""),
            ("", ""),
        ],
    )
    def test_get_file_name(self, path, expected):
        assert get_file_name(path) == expected

    def test_ensure_is_directory(self, tmp_path):
        ensure_is_directory(str(tmp_path))

        with pytest.raises(ValueError, match="路径不存在"):
            ensure_is_directory(str(tmp_path / "missing"))

        file_path = tmp_path / "a.txt"
        file_path.write_text("x", encoding="utf-8")
        with pytest.raises(ValueError, match="不是一个文件夹"):
            ensure_is_directory(str(file_path))

    def test_ensure_directory_exists(self, tmp_path):
        target = tmp_path / "a" / "b"
        ensure_directory_exists(str(target))
        assert target.is_dir()

        ensure_directory_exists(str(target))  # 已存在是幂等的

        file_path = tmp_path / "a.txt"
        file_path.write_text("x", encoding="utf-8")
        with pytest.raises(ValueError, match="不是文件夹"):
            ensure_directory_exists(str(file_path))


class TestDocxUtilsBoundary:
    def test_image_only_in_direct_run(self, tmp_path):
        para = Document().add_paragraph()
        para.add_run().add_picture(_png(tmp_path))

        assert para_contains_image(para) is True

    def test_image_inside_hyperlink_is_detected(self, tmp_path):
        para = Document().add_paragraph()
        run = para.add_run()
        run.add_picture(_png(tmp_path))
        link = OxmlElement("w:hyperlink")
        link.append(run._r)
        para._p.append(link)

        assert para_contains_image(para) is True

    def test_math_detection(self):
        doc = Document()

        with_math = doc.add_paragraph()
        run = with_math.add_run()
        run._r.append(OxmlElement("m:oMath"))
        assert para_contains_math(with_math) is True

        empty_math_para = doc.add_paragraph()
        run = empty_math_para.add_run()
        run._r.append(OxmlElement("m:oMathPara"))
        assert para_contains_math(empty_math_para) is False

    def test_remove_all_numbering_strips_style_numbering(self):
        doc = Document()
        style = doc.styles["Heading 1"]
        pPr = ensure_pPr(style._element)
        for el in pPr.findall(qn("w:outlineLvl")):
            pPr.remove(el)
        numPr = OxmlElement("w:numPr")
        numPr.append(OxmlElement("w:numId"))
        pPr.append(numPr)
        pPr.append(OxmlElement("w:outlineLvl"))

        remove_all_numbering(doc)

        assert pPr.find(qn("w:numPr")) is None
        assert pPr.find(qn("w:outlineLvl")) is None


# ---------------------------------------------------------------------------
# tree.py
# ---------------------------------------------------------------------------


class TestTreeWalkBoundary:
    @staticmethod
    def _tree():
        root = TreeNode("r")
        a = root.add_child("a")
        a.add_child("aa")
        root.add_child("b")
        return root

    def test_bfs_includes_root(self):
        assert [n.value for n in bfs_walk(self._tree())] == ["r", "a", "b", "aa"]

    def test_dfs_excludes_root(self):
        assert [n.value for n in dfs_walk(self._tree())] == ["a", "aa", "b"]

    def test_dfs_single_node_is_empty(self):
        assert list(dfs_walk(TreeNode("only"))) == []

    def test_find_by_condition_can_match_root(self):
        tree = Tree("root")
        tree.root.add_child("child")

        assert tree.find_by_condition(lambda v: v == "root") is tree.root


class TestRenderTreeBoundary:
    def test_empty_root_renders_value(self):
        assert render_tree(TreeNode("top")).strip() == "top"

    def test_filter_returns_first_matching_branch_only(self):
        root = TreeNode({"category": "top", "paragraph": ""})
        root.add_child({"category": "body_text", "paragraph": "正文"})
        root.add_child({"category": "heading_level_1", "paragraph": "标题"})

        output = render_tree(root, filter_categories=["heading_level_1"])

        assert "heading_level_1" in output
        assert "body_text" not in output

    def test_empty_json_file(self, tmp_path):
        path = tmp_path / "empty.json"
        path.write_text("[]", encoding="utf-8")

        output = render_tree(str(path))

        assert "0 个段落" in output

    def test_json_file_with_items(self, tmp_path):
        path = tmp_path / "doc.json"
        path.write_text(
            '[{"category": "heading_level_1", "paragraph": "标题", "score": 0.5},'
            ' {"category": "body_text", "paragraph": "正文", "score": 0.9}]',
            encoding="utf-8",
        )

        output = render_tree(str(path), show_confidence=True, show_index=True)

        assert "2 个段落" in output
        assert "heading_level_1" in output
        assert "(50%)" in output

    def test_node_label_truncation_and_confidence(self):
        label = _format_node_label(
            {"category": "body_text", "paragraph": "x" * 60, "score": 0.0},
            show_confidence=True,
            show_index=False,
            counter=0,
        )
        assert "(0%)" in label
        assert "x" * 50 in label
        assert "x" * 51 not in label

        assert _format_node_label("y" * 70, False, False, 0) == "y" * 60
        assert _format_node_label("z", False, True, 7).startswith("[  7] ")

    def test_tree_node_category_for_object_value(self):
        value = SimpleNamespace(paragraph={"category": "body_text"})

        assert _tree_node_category(value) == "body_text"
        assert _tree_node_category("plain") == ""
