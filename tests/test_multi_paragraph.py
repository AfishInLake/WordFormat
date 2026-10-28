#! /usr/bin/env python
# @Time    : 2026-10-28
# @Author  : afish
# @File    : test_multi_paragraph.py
"""多段落条目（paragraphs）在 af --sync 前展开为多条同身份条目。"""

import json

import pytest
from docx import Document

from wordformat.para_ids import ensure_para_ids, get_para_id
from wordformat.pipeline.sync_apply import expand_multi_paragraph_items


def _body(text, **extra):
    return {"category": "body_text", "paragraph": text, "score": 1.0, **extra}


class TestExpand:
    def test_without_paragraphs_returns_same_list(self):
        items = [_body("甲", para_id="AAAAAAAA"), _body("乙", para_id="BBBBBBBB")]
        assert expand_multi_paragraph_items(items) == items

    def test_two_texts_split_with_id_on_first_only(self):
        items = [_body("甲", para_id="AAAAAAAA", paragraphs=["第一段", "第二段"])]
        out = expand_multi_paragraph_items(items)

        assert [it["paragraph"] for it in out] == ["第一段", "第二段"]
        assert out[0]["para_id"] == "AAAAAAAA"
        assert "para_id" not in out[1]
        assert "paragraphs" not in out[0] and "paragraphs" not in out[1]
        assert out[1]["comment"] == "由多段落条目拆分"

    def test_category_and_score_are_inherited(self):
        items = [
            {
                "category": "heading_level_1",
                "paragraph": "旧",
                "score": 0.42,
                "para_id": "AAAAAAAA",
                "paragraphs": ["第一章 绪论", "本章小结"],
            }
        ]
        out = expand_multi_paragraph_items(items)
        assert [it["category"] for it in out] == ["heading_level_1"] * 2
        assert out[1]["score"] == 0.42

    def test_single_element_list_is_rejected(self):
        with pytest.raises(ValueError, match="至少 2 段"):
            expand_multi_paragraph_items([_body("甲", paragraphs=["只有一段"])])

    def test_conflict_with_replace_is_rejected(self):
        with pytest.raises(ValueError, match="replace"):
            expand_multi_paragraph_items(
                [_body("甲", paragraphs=["a", "b"], replace="覆写")]
            )

    def test_table_and_image_entries_are_rejected(self):
        with pytest.raises(ValueError, match="table_object"):
            expand_multi_paragraph_items(
                [{"category": "table_object", "paragraph": "", "paragraphs": ["a", "b"]}]
            )
        with pytest.raises(ValueError, match="figure_image"):
            expand_multi_paragraph_items(
                [
                    {
                        "category": "figure_image",
                        "paragraph": "/a.png",
                        "paragraphs": ["a", "b"],
                    }
                ]
            )

    def test_non_dict_items_pass_through(self):
        assert expand_multi_paragraph_items(["裸字符串"]) == ["裸字符串"]


def _make_doc(texts):
    doc = Document()
    for t in texts:
        doc.add_paragraph(t)
    return doc


def _item(category, text, **extra):
    d = {"category": category, "paragraph": text, "score": 1.0}
    d.update(extra)
    return d


class TestSyncEndToEnd:
    def _run(self, tmp_path, doc, items):
        """把已备好身份的 doc 落盘，再跑 af --sync，返回输出文档。"""
        dp = tmp_path / "s.docx"
        doc.save(dp)
        jp = tmp_path / "s.json"
        jp.write_text(json.dumps(items, ensure_ascii=False), encoding="utf-8")
        from wordformat.pipeline.orchestrate import auto_format_thesis_document

        return auto_format_thesis_document(
            jsonpath=str(jp),
            docxpath=str(dp),
            check=False,
            savepath=str(tmp_path / "out"),
            sync=True,
        )

    def test_one_entry_becomes_two_paragraphs_at_right_place(self, tmp_path):
        doc = _make_doc(["甲", "乙"])
        ensure_para_ids(doc)
        ids = [get_para_id(p._p) for p in doc.paragraphs]
        items = [
            _item("body_text", "甲", para_id=ids[0], paragraphs=["甲改写", "插入的新段"]),
            _item("body_text", "乙", para_id=ids[1]),
        ]

        out = Document(self._run(tmp_path, doc, items))

        assert [p.text for p in out.paragraphs] == ["甲改写", "插入的新段", "乙"]
        assert get_para_id(out.paragraphs[0]._p) == ids[0]

    def test_jd_rejects_paragraphs_instead_of_dropping_them(self, tmp_path):
        from wordformat.pipeline.orchestrate import json_to_docx

        jp = tmp_path / "nodes.json"
        jp.write_text(
            json.dumps(
                [_item("body_text", "甲", paragraphs=["一", "二"])], ensure_ascii=False
            ),
            encoding="utf-8",
        )

        with pytest.raises(ValueError, match="af --sync"):
            json_to_docx(str(jp), save_dir=str(tmp_path / "out"))
