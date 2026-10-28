#! /usr/bin/env python
# @Time    : 2026-10-28
# @Author  : afish
# @File    : test_table_sync.py
"""af --sync 的表格身份对齐：按 table_id 归位、删除、新增。"""

import json

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from wordformat.para_ids import (
    W14_PARA_ID,
    ensure_para_ids,
    ensure_table_ids,
    get_para_id,
    get_table_id,
)
from wordformat.pipeline.sync_apply import (
    _body_tables,
    _table_identity_ready,
    sync_document_structure,
)


def _doc_two_tables():
    doc = Document()
    doc.add_paragraph("段落A")
    for name in ("表一", "表二"):
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).text = name
    doc.add_paragraph("段落B")
    ensure_para_ids(doc)
    ensure_table_ids(doc)
    return doc


def _body_tags(doc):
    """body 直接子级的结构元素序列（p/tbl），用于断言位置与增删结果。"""
    return [
        "tbl" if el.tag == qn("w:tbl") else "p"
        for el in doc.element.body.iterchildren()
        if el.tag in (qn("w:p"), qn("w:tbl"))
    ]


def _items(doc):
    """按 body 文档序生成 gj 形态的条目列表（段落带 para_id，表格带 table_id）。"""
    items = []
    for el in doc.element.body.iterchildren():
        if el.tag == qn("w:p"):
            items.append(
                {
                    "category": "body_text",
                    "paragraph": Paragraph(el, doc._body).text,
                    "score": 1.0,
                    "para_id": get_para_id(el),
                }
            )
        elif el.tag == qn("w:tbl"):
            items.append(
                {
                    "category": "table_object",
                    "paragraph": "",
                    "score": 1.0,
                    "exists": True,
                    "table_rows": [
                        [c.text for c in r.cells] for r in Table(el, doc._body).rows
                    ],
                    "table_id": get_table_id(el),
                }
            )
    return items


def _tables_of(doc):
    return _body_tables(doc.element.body)


class TestTableIdentityReady:
    def test_ready_when_both_sides_carry_ids(self):
        doc = _doc_two_tables()
        assert _table_identity_ready(_items(doc), _tables_of(doc)) is True

    def test_not_ready_when_a_table_item_lacks_table_id(self):
        doc = _doc_two_tables()
        items = _items(doc)
        items[1].pop("table_id")
        assert _table_identity_ready(items, _tables_of(doc)) is False

    def test_not_ready_when_a_doc_table_lacks_id(self):
        doc = _doc_two_tables()
        tables = _tables_of(doc)
        anchor = tables[1].xpath("./w:tr[1]/w:tc[1]/w:p[1]")[0]
        anchor.attrib.pop(W14_PARA_ID)
        assert _table_identity_ready(_items(doc), tables) is False

    def test_no_table_items_is_deletion_when_paragraphs_have_ids(self):
        # JSON 出自同一文档的 gj（段落带 para_id）却不提表格 → 认定表格被删
        doc = _doc_two_tables()
        items = [it for it in _items(doc) if it["category"] != "table_object"]
        assert _table_identity_ready(items, _tables_of(doc)) is True

    def test_no_table_items_and_no_para_id_is_not_ready(self):
        # 手写 JSON 忽略表格 ≠ 要删表格：身份证据不足时保守不动
        doc = _doc_two_tables()
        items = [{"category": "body_text", "paragraph": "段落A"}]
        assert _table_identity_ready(items, _tables_of(doc)) is False

    def test_ready_with_new_table_item_and_no_doc_tables(self):
        doc = Document()
        doc.add_paragraph("段落A")
        ensure_para_ids(doc)
        items = _items(doc) + [
            {"category": "table_object", "paragraph": "", "table_rows": [["x"]]}
        ]
        assert _table_identity_ready(items, _tables_of(doc)) is True

    def test_new_table_item_alongside_identified_ones_is_still_ready(self):
        # agent 新增的表格条目本来就没有 table_id，不能因此整表退回收位
        doc = _doc_two_tables()
        items = _items(doc)
        items.insert(
            2, {"category": "table_object", "paragraph": "", "table_rows": [["新"]]}
        )
        assert _table_identity_ready(items, _tables_of(doc)) is True

    def test_exists_item_without_table_id_is_not_ready(self):
        # exists=True 指既有表格；它缺 table_id 说明 JSON 早于表格身份，删表不安全
        doc = _doc_two_tables()
        items = _items(doc)
        items[1].pop("table_id")
        assert _table_identity_ready(items, _tables_of(doc)) is False


class TestTableReorderAndDelete:
    def test_reorder_swaps_two_tables(self):
        doc = _doc_two_tables()
        items = _items(doc)
        p_a, tbl1, tbl2, p_b = items
        sync_document_structure(doc, [p_a, tbl2, tbl1, p_b])

        ids = [get_table_id(t._tbl) for t in doc.tables]
        assert ids == [tbl2["table_id"], tbl1["table_id"]]
        assert [t.cell(0, 0).text for t in doc.tables] == ["表二", "表一"]

    def test_dropping_a_table_entry_deletes_that_table(self):
        doc = _doc_two_tables()
        items = _items(doc)
        p_a, tbl1, tbl2, p_b = items
        sync_document_structure(doc, [p_a, tbl2, p_b])

        assert len(doc.tables) == 1
        assert get_table_id(doc.tables[0]._tbl) == tbl2["table_id"]
        assert _body_tags(doc) == ["p", "tbl", "p"]
        assert [p.text for p in doc.paragraphs] == ["段落A", "段落B"]

    def test_paragraph_identity_survives_table_deletion(self):
        doc = _doc_two_tables()
        items = _items(doc)
        p_a, tbl1, tbl2, p_b = items
        kept = p_b["para_id"]
        sync_document_structure(doc, [p_a, tbl2, p_b])

        assert [get_para_id(p._p) for p in doc.paragraphs].count(kept) == 1

    def test_new_table_item_inserts_a_placeholder_at_its_position(self):
        doc = _doc_two_tables()
        items = _items(doc)
        p_a, tbl1, tbl2, p_b = items
        new = {"category": "table_object", "paragraph": "", "table_rows": [["新"]]}

        sync_document_structure(doc, [p_a, tbl1, new, tbl2, p_b])

        # sync 只占位；真正建表发生在 FormattingExecutionStage
        assert _body_tags(doc) == ["p", "tbl", "p", "tbl", "p"]
        assert len(doc.tables) == 2
        assert [get_table_id(t._tbl) for t in doc.tables] == [
            tbl1["table_id"],
            tbl2["table_id"],
        ]

    def test_without_table_id_falls_back_to_positional_and_keeps_all(self):
        doc = _doc_two_tables()
        items = [dict(it) for it in _items(doc)]
        for it in items:
            it.pop("table_id", None)
        before = [get_table_id(t._tbl) for t in doc.tables]

        sync_document_structure(doc, items)

        assert [get_table_id(t._tbl) for t in doc.tables] == before

    def test_unreferenced_tables_survive_without_ids(self):
        # JSON 里段落也不带 para_id：走文本对齐，表格一律原地不动
        doc = _doc_two_tables()
        items = [
            {"category": "body_text", "paragraph": "段落A"},
            {"category": "body_text", "paragraph": "段落B"},
        ]

        sync_document_structure(doc, items)

        assert len(doc.tables) == 2


class TestTableInsert:
    def test_new_table_item_is_not_skipped_when_identity_is_ready(self):
        doc = _doc_two_tables()
        items = _items(doc)
        p_a, tbl1, tbl2, p_b = items
        new = {
            "category": "table_object",
            "paragraph": "",
            "score": 1.0,
            "table_rows": [["甲", "乙"], ["丙", "丁"]],
        }
        items.insert(2, new)

        sync_document_structure(doc, items)

        assert "sync_skip" not in new
        # 占位段落绑定该节点，建表阶段才能把表格插到它后面并删掉占位段
        assert _body_tags(doc) == ["p", "tbl", "p", "tbl", "p"]
        assert len(doc.paragraphs) == 3

    def test_new_table_item_is_skipped_when_identity_is_unavailable(self):
        doc = _doc_two_tables()
        items = [
            {"category": "body_text", "paragraph": "段落A"},
            {
                "category": "table_object",
                "paragraph": "",
                "score": 1.0,
                "table_rows": [["甲"]],
            },
        ]

        sync_document_structure(doc, items)

        assert items[1]["sync_skip"] is True
        # 不占段落槽位：占位段一旦落下就没有节点绑定，段落数与条目数会对不上
        assert _body_tags(doc) == ["p", "tbl", "tbl"]
        assert len(doc.paragraphs) == 1

    def test_full_sync_pipeline_builds_the_new_table(self, tmp_path):
        from wordformat.pipeline.orchestrate import auto_format_thesis_document

        src = tmp_path / "t.docx"
        _doc_two_tables().save(src)
        items = _items(Document(src))
        items.insert(
            2,
            {
                "category": "table_object",
                "paragraph": "",
                "score": 1.0,
                "table_rows": [["甲", "乙"], ["丙", "丁"]],
            },
        )
        jp = tmp_path / "nodes.json"
        jp.write_text(json.dumps(items, ensure_ascii=False), encoding="utf-8")

        out = auto_format_thesis_document(
            jsonpath=str(jp),
            docxpath=str(src),
            savepath=str(tmp_path / "out"),
            check=False,
            sync=True,
        )

        result = Document(out)
        assert len(result.tables) == 3
        assert [t.cell(0, 0).text for t in result.tables] == ["表一", "甲", "表二"]
        # 新表格当场可寻址，下一轮 sync 才能删改它
        assert get_table_id(result.tables[1]._tbl)
