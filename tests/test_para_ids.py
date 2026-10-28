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
    get_para_id,
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
