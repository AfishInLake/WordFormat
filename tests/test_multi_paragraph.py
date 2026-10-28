#! /usr/bin/env python
# @Time    : 2026-10-28
# @Author  : afish
# @File    : test_multi_paragraph.py
"""多段落条目（paragraphs）在 af --sync 前展开为多条同身份条目。"""

import pytest

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
