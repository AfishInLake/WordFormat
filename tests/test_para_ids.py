#! /usr/bin/env python
# @Time    : 2026-10-28
# @Author  : afish
# @File    : test_para_ids.py
"""段落身份单元测试（para_ids.py）。"""

import re

from docx import Document

from wordformat.para_ids import (
    W14_PARA_ID,
    ensure_para_ids,
    ensure_table_ids,
    get_para_id,
    get_table_id,
    paragraph_para_id,
)

_HEX8 = re.compile(r"^[0-9A-F]{8}$")


def _make_doc(texts):
    doc = Document()
    for t in texts:
        doc.add_paragraph(t)
    return doc


class TestEnsureParaIds:
    def test_mints_unique_hex8_for_every_body_paragraph(self):
        doc = _make_doc(["A", "B", "C"])
        total = len(doc.paragraphs)
        assert ensure_para_ids(doc) == total
        ids = [get_para_id(p._p) for p in doc.paragraphs]
        assert all(pid and _HEX8.match(pid) for pid in ids)
        assert len(set(ids)) == total

    def test_idempotent(self):
        doc = _make_doc(["A", "B"])
        ensure_para_ids(doc)
        assert ensure_para_ids(doc) == 0

    def test_keeps_ids_written_by_word(self):
        doc = _make_doc(["A", "B"])
        doc.paragraphs[0]._p.set(W14_PARA_ID, "1234ABCD")
        assert ensure_para_ids(doc) == len(doc.paragraphs) - 1
        assert get_para_id(doc.paragraphs[0]._p) == "1234ABCD"

    def test_never_mints_zero(self):
        doc = _make_doc(["A"])
        assert get_para_id(doc.paragraphs[0]._p) is None
        ensure_para_ids(doc)
        assert get_para_id(doc.paragraphs[0]._p) != "00000000"

    def test_avoids_ids_already_used_by_table_cells(self):
        # 单元格段落也有 paraId，body 段落新铸造的取值不得与之撞号
        doc = _make_doc(["A"])
        table = doc.add_table(rows=1, cols=1)
        cell_para = table.rows[0].cells[0].paragraphs[0]
        cell_para._p.set(W14_PARA_ID, "00000002")

        ensure_para_ids(doc)

        body_id = get_para_id(doc.paragraphs[0]._p)
        assert body_id != "00000002"
        assert _HEX8.match(body_id)
        # 单元格段落自身的取值不动
        assert get_para_id(cell_para._p) == "00000002"


class TestReaders:
    def test_get_para_id_returns_none_when_absent(self):
        doc = _make_doc(["A"])
        assert get_para_id(doc.paragraphs[0]._p) is None

    def test_paragraph_helper_matches_element_reader(self):
        doc = _make_doc(["A"])
        ensure_para_ids(doc)
        para = doc.paragraphs[0]
        assert paragraph_para_id(para) == get_para_id(para._p)

    def test_identity_survives_save_and_reload(self, tmp_path):
        # 身份必须穿过进程边界（gj 写、af 读），靠 XML 属性而不是对象引用
        doc = _make_doc(["甲", "乙"])
        ensure_para_ids(doc)
        before = [get_para_id(p._p) for p in doc.paragraphs]
        path = tmp_path / "ids.docx"
        doc.save(path)

        reloaded = Document(str(path))
        assert [get_para_id(p._p) for p in reloaded.paragraphs] == before


def _doc_with_tables(n=2):
    doc = Document()
    doc.add_paragraph("表前")
    tables = []
    for k in range(n):
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = f"甲{k}"
        tables.append(table)
    doc.add_paragraph("表后")
    return doc, tables


class TestTableIds:
    def test_get_table_id_reads_first_cell_paragraph(self):
        doc, tables = _doc_with_tables(1)
        assert get_table_id(tables[0]._tbl) is None

        assert ensure_table_ids(doc) == 1

        table_id = get_table_id(tables[0]._tbl)
        assert table_id and _HEX8.match(table_id)
        anchor = tables[0].rows[0].cells[0].paragraphs[0]
        assert get_para_id(anchor._p) == table_id

    def test_two_tables_get_distinct_ids_and_idempotent(self):
        doc, tables = _doc_with_tables(2)
        assert ensure_table_ids(doc) == 2
        ids = {get_table_id(t._tbl) for t in tables}
        assert len(ids) == 2 and all(ids)
        assert ensure_table_ids(doc) == 0

    def test_keeps_table_id_written_by_word(self):
        doc, tables = _doc_with_tables(1)
        tables[0].rows[0].cells[0].paragraphs[0]._p.set(W14_PARA_ID, "11112222")

        assert ensure_table_ids(doc) == 0
        assert get_table_id(tables[0]._tbl) == "11112222"

    def test_does_not_mint_body_paragraph_ids(self):
        # 两个职责分开：补表格身份只动首单元格段落，不越权给 body 段落铸 id
        doc, _ = _doc_with_tables(1)
        ensure_table_ids(doc)
        assert [get_para_id(p._p) for p in doc.paragraphs] == [None, None]

    def test_only_body_level_tables_are_addressed(self):
        doc = Document()
        outer = doc.add_table(rows=1, cols=1)
        outer.cell(0, 0).add_table(rows=1, cols=1)

        assert ensure_table_ids(doc) == 1
        assert get_table_id(outer._tbl)

    def test_avoids_ids_already_in_use(self):
        doc, tables = _doc_with_tables(2)
        tables[0].rows[0].cells[0].paragraphs[0]._p.set(W14_PARA_ID, "00000003")

        ensure_table_ids(doc)

        assert get_table_id(tables[1]._tbl) != "00000003"

    def test_table_id_survives_save_and_reload(self, tmp_path):
        doc, tables = _doc_with_tables(2)
        ensure_table_ids(doc)
        before = [get_table_id(t._tbl) for t in tables]
        path = tmp_path / "tables.docx"
        doc.save(path)

        reloaded = Document(path)
        assert [get_table_id(t._tbl) for t in reloaded.tables] == before
