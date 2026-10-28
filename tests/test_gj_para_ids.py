#! /usr/bin/env python
# @Time    : 2026-10-28
# @Author  : afish
# @File    : test_gj_para_ids.py
"""gj 段落身份：para_id 导出、w14:paraId 落盘、写不回时的降级。"""

from unittest.mock import patch

from docx import Document
from docx.document import Document as DocxDocument

from wordformat.base import DocxBase
from wordformat.para_ids import W14_PARA_ID, get_para_id


def _make_docx(tmp_path, body):
    """body: list，元素为 str（段落）、("table", rows) 或 ("id", text, pid)。"""
    doc = Document()
    for item in body:
        if isinstance(item, tuple):
            if item[0] == "table":
                payload = item[1]
                table = doc.add_table(rows=len(payload), cols=len(payload[0]))
                for i, row in enumerate(payload):
                    for j, text in enumerate(row):
                        table.cell(i, j).text = text
            else:
                _, text, pid = item
                para = doc.add_paragraph(text)
                para._p.set(W14_PARA_ID, pid)
        else:
            doc.add_paragraph(item)
    path = tmp_path / "src.docx"
    doc.save(path)
    return str(path)


def _parse(path, n_texts):
    results = [{"label": "body_text", "score": 0.9} for _ in range(n_texts)]
    with patch(
        "wordformat.base.onnx_batch_infer",
        return_value=(results, "body_text"),
    ):
        return DocxBase(path, "/fake/config.yaml").parse()


class TestGjParaIdExport:
    def test_every_paragraph_item_carries_the_document_id_in_order(self, tmp_path):
        path = _make_docx(tmp_path, ["甲段", "乙段", "丙段"])
        items = _parse(path, 3)
        doc = Document(path)

        ids = [get_para_id(p._p) for p in doc.paragraphs]
        assert all(ids)
        assert [it["para_id"] for it in items] == ids

    def test_ids_are_persisted_to_disk_for_the_next_process(self, tmp_path):
        # af 是另一次进程读取同一个文件，身份只能靠 XML 落盘，不能靠内存对象
        path = _make_docx(tmp_path, ["甲段"])
        _parse(path, 1)

        reloaded = Document(path)
        assert all(get_para_id(p._p) for p in reloaded.paragraphs)

    def test_ids_already_written_by_word_are_reused(self, tmp_path):
        path = _make_docx(tmp_path, [("id", "甲段", "11112222")])
        items = _parse(path, 1)

        assert items[0]["para_id"] == "11112222"
        assert get_para_id(Document(path).paragraphs[0]._p) == "11112222"

    def test_table_entries_are_not_paragraph_addressable(self, tmp_path):
        path = _make_docx(tmp_path, ["甲段", ("table", [["x"]]), "乙段"])
        items = _parse(path, 2)
        tables = [it for it in items if it["category"] == "table_object"]

        assert len(tables) == 1
        assert "para_id" not in tables[0]
        assert all("para_id" in it for it in items if it is not tables[0])

    def test_export_still_works_when_docx_cannot_be_written(self, tmp_path):
        # 文件被 Word 占用/只读：身份写不回，但流程不得中断（af 侧退回文本对齐）
        path = _make_docx(tmp_path, ["甲段", "乙段"])
        with patch.object(DocxDocument, "save", side_effect=OSError("文件被占用")):
            items = _parse(path, 2)

        assert len(items) == 2
        assert all("para_id" in it for it in items)

    def test_round_trip_through_sync_uses_the_exported_ids(self, tmp_path):
        # gj → af --sync 闭环：JSON 增删条目能精确落到对应段落上
        from wordformat.pipeline.sync_apply import sync_document_structure

        path = _make_docx(tmp_path, ["重复", "重复"])
        items = _parse(path, 2)
        first_id = items[0]["para_id"]

        doc = Document(path)
        edited = [
            items[1],
            {"category": "body_text", "paragraph": "插入段", "score": 1.0},
        ]
        sync_document_structure(doc, edited)

        assert [p.text for p in doc.paragraphs] == ["重复", "插入段"]
        assert get_para_id(doc.paragraphs[0]._p) == items[1]["para_id"]
        assert first_id not in [get_para_id(p._p) for p in doc.paragraphs]
