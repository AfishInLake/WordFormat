#! /usr/bin/env python
# @Time    : 2026/10/22
# @Author  : afish
# @File    : test_jd.py
"""jd（JSON → docx 重建）单元测试：json_to_docx / CLI 路由。"""

import base64
import json
from unittest import mock

import pytest
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

from wordformat.cli import main
from wordformat.pipeline.orchestrate import json_to_docx

# 1x1 透明 PNG
_PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ"
    "AAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def _item(category, text, **extra):
    d = {"category": category, "paragraph": text, "score": 1.0}
    d.update(extra)
    return d


@pytest.fixture
def json_file(tmp_path):
    def _write(items, name="thesis.json"):
        p = tmp_path / name
        p.write_text(json.dumps(items, ensure_ascii=False), encoding="utf-8")
        return str(p)

    return _write


@pytest.fixture
def png_file(tmp_path):
    p = tmp_path / "fig.png"
    p.write_bytes(_PNG_1PX)
    return str(p)


class TestJsonToDocx:
    def test_basic_build(self, tmp_path, json_file):
        jp = json_file(
            [
                _item("document_title", "一种方法研究"),
                _item("heading_level_1", "绪论"),
                _item("body_text", "本文研究……"),
            ]
        )
        out = json_to_docx(jp, save_dir=str(tmp_path))
        doc = Document(out)
        assert [p.text for p in doc.paragraphs] == [
            "一种方法研究",
            "绪论",
            "本文研究……",
        ]

    def test_replace_field_wins(self, tmp_path, json_file):
        jp = json_file([_item("body_text", "旧文本", replace="新文本")])
        out = json_to_docx(jp, save_dir=str(tmp_path))
        assert Document(out).paragraphs[0].text == "新文本"

    def test_default_config_when_none(self, tmp_path, json_file):
        jp = json_file([_item("body_text", "正文")])
        out = json_to_docx(jp, config_path=None, save_dir=str(tmp_path))
        assert Document(out).paragraphs[0].text == "正文"

    def test_figure_image_insert(self, tmp_path, json_file, png_file):
        jp = json_file(
            [
                _item("heading_level_1", "第一章"),
                _item("figure_image", png_file),
                _item("body_text", "如图1.1所示。"),
            ]
        )
        out = json_to_docx(jp, save_dir=str(tmp_path))
        doc = Document(out)
        drawing_paras = [p for p in doc.paragraphs if p._p.findall(qn("w:r"))]
        has_drawing = any(p._p.xpath(".//w:drawing") for p in doc.paragraphs)
        assert has_drawing, "figure_image 路径条目应真实嵌入图片"
        assert len(drawing_paras) >= 1

    def test_figure_image_missing_path_keeps_empty(self, tmp_path, json_file):
        jp = json_file([_item("figure_image", "/nonexist/fig.png")])
        out = json_to_docx(jp, save_dir=str(tmp_path))
        # 路径不存在：不报错，保留空段
        assert Document(out).paragraphs[0].text == ""

    def test_table_object_rows(self, tmp_path, json_file):
        rows = [["名称", "数值"], ["A", "1"], ["B", "2"]]
        jp = json_file(
            [
                _item("heading_level_1", "数据"),
                _item("table_object", "", table_rows=rows),
                _item("body_text", "见表1。"),
            ]
        )
        out = json_to_docx(jp, save_dir=str(tmp_path))
        doc = Document(out)
        assert len(doc.tables) == 1
        table = doc.tables[0]
        assert len(table.rows) == 3 and len(table.columns) == 2
        assert table.cell(0, 0).text == "名称"
        assert table.cell(2, 1).text == "2"
        # 空锚段已回收，表格前后不产生多余空段
        assert [p.text for p in doc.paragraphs] == ["数据", "见表1。"]

    def test_heading_renumber_after_delete(self, tmp_path, json_file):
        # 编号启用时：标题段写入段落级 numPr，重建按新树序套用——删掉首个
        # 标题后，剩余两个标题仍依次挂在同一编号定义上（Word 打开时展示为连续编号）
        cfg = tmp_path / "num.yaml"
        cfg.write_text(
            "numbering:\n"
            "  enabled: true\n"
            "  level_1:\n"
            "    enabled: true\n"
            "    template: '第%1章'\n"
            "    suffix: 'space'\n"
            "  level_2:\n"
            "    enabled: false\n"
            "  level_3:\n"
            "    enabled: false\n",
            encoding="utf-8",
        )
        items = [
            _item("heading_level_1", "绪论"),
            _item("body_text", "绪论内容"),
            _item("heading_level_1", "方法"),
            _item("body_text", "方法内容"),
            _item("heading_level_1", "结论"),
        ]
        out = json_to_docx(
            json_file(items), config_path=str(cfg), save_dir=str(tmp_path / "out1")
        )
        full = Document(out)
        numpr_of = {p.text: p._p.pPr.numPr for p in full.paragraphs if p.text}
        assert numpr_of["绪论"] is not None

        # 删第一章后重新生成：输出文档不再含该标题，且不含已删标题的正文
        deleted = items[1:]
        out2 = json_to_docx(
            json_file(deleted, name="deleted.json"),
            config_path=str(cfg),
            save_dir=str(tmp_path / "out2"),
        )
        doc2 = Document(out2)
        texts = [p.text for p in doc2.paragraphs]
        assert "绪论" not in texts
        remaining_nums = [
            (p._p.pPr.numPr.numId.val, p._p.pPr.numPr.ilvl.val)
            for p in doc2.paragraphs
            if p.text in ("方法", "结论")
        ]
        assert remaining_nums == [(n, 0) for n, _ in remaining_nums]
        assert len({n for n, _ in remaining_nums}) == 1  # 共用同一编号定义

    def test_empty_replace_clears_text(self, tmp_path, json_file):
        jp = json_file([_item("body_text", "旧文本", replace="")])
        out = json_to_docx(jp, save_dir=str(tmp_path))
        assert Document(out).paragraphs[0].text == ""

    def test_replace_overrides_stale_inline_marks(self, tmp_path, json_file):
        # agent 改了 replace 却残留旧 inline_marks：以 replace 为准，不回显旧文本
        jp = json_file(
            [
                _item(
                    "body_text",
                    "旧文本",
                    replace="新文本",
                    inline_marks=[{"text": "旧文本", "bold": True}],
                )
            ]
        )
        out = json_to_docx(jp, save_dir=str(tmp_path))
        para = Document(out).paragraphs[0]
        assert para.text == "新文本"
        assert all("旧文本" not in run.text for run in para.runs)

    def test_inline_marks_rich_runs(self, tmp_path, json_file):
        # md 风格 inline_marks → 多 run（粗/斜/代码/链接）
        jp = json_file(
            [
                {
                    "category": "body_text",
                    "paragraph": "加粗文本",
                    "score": 1.0,
                    "inline_marks": [{"text": "加粗文本", "bold": True}],
                },
                {
                    "category": "body_text",
                    "paragraph": "斜体文本",
                    "score": 1.0,
                    "inline_marks": [{"text": "斜体文本", "italic": True}],
                },
                {
                    "category": "body_text",
                    "paragraph": "code()",
                    "score": 1.0,
                    "inline_marks": [{"text": "code()", "code": True}],
                },
                {
                    "category": "body_text",
                    "paragraph": "链接",
                    "score": 1.0,
                    "inline_marks": [
                        {"text": "链接", "link_url": "https://example.com"}
                    ],
                },
            ]
        )
        out = json_to_docx(jp, save_dir=str(tmp_path))
        doc = Document(out)
        assert doc.paragraphs[0].runs[0].bold is True
        assert doc.paragraphs[1].runs[0].italic is True
        assert doc.paragraphs[2].runs[0].font.name == "Consolas"
        hyperlinks = doc.paragraphs[3]._p.xpath(".//w:hyperlink")
        assert len(hyperlinks) == 1
        texts = hyperlinks[0].xpath(".//w:t")
        assert texts and texts[0].text == "链接"

    def test_tree_normalize_promotes_abstract_body(self, tmp_path, json_file):
        # 摘要标题下的 body_text 会被论文领域提升为 AbstractContentCN，
        # 从而套用 abstract.chinese.body 配置（jd 流水线必须包含归一化阶段）
        cfg = tmp_path / "abstract.yaml"
        cfg.write_text(
            "abstract:\n"
            "  chinese:\n"
            "    body:\n"
            "      font_size: '三号'\n",
            encoding="utf-8",
        )
        items = [
            _item("abstract_chinese_title", "摘要"),
            _item("body_text", "本文研究了……"),
            _item("heading_level_1", "绪论"),
            _item("body_text", "正文内容"),
        ]
        out = json_to_docx(
            json_file(items),
            config_path=str(cfg),
            save_dir=str(tmp_path / "out"),
        )
        doc = Document(out)
        assert doc.paragraphs[1].text == "本文研究了……"
        assert doc.paragraphs[1].runs[0].font.size == Pt(16)  # 三号
        assert doc.paragraphs[3].runs[0].font.size == Pt(12)  # 默认小四

    def test_exists_table_not_rebuilt(self, tmp_path, json_file):
        # exists=True 的原文档表格：重建路线也不重复建表，且不留下占位空段
        jp = json_file(
            [
                _item("body_text", "A"),
                {
                    "category": "table_object",
                    "paragraph": "",
                    "score": 1.0,
                    "table_rows": [["a", "b"]],
                    "exists": True,
                },
                _item("body_text", "B"),
            ]
        )
        out = json_to_docx(jp, save_dir=str(tmp_path))
        doc = Document(out)
        assert doc.tables == []
        assert [p.text for p in doc.paragraphs] == ["A", "B"]

    def test_abstract_body_keeps_inline_bold(self, tmp_path, json_file):
        # 摘要正文走自定义 _base，也必须保留 inline_marks 的加粗
        jp = json_file(
            [
                _item("abstract_chinese_title", "摘要"),
                {
                    "category": "body_text",
                    "paragraph": "重点内容",
                    "score": 1.0,
                    "inline_marks": [
                        {"text": "重点", "bold": True},
                        {"text": "内容"},
                    ],
                },
                _item("heading_level_1", "绪论"),
            ]
        )
        out = json_to_docx(jp, save_dir=str(tmp_path))
        runs = {r.text: r for r in Document(out).paragraphs[1].runs if r.text}
        assert runs["重点"].bold is True
        assert not runs["内容"].bold

    def test_unknown_category_raises_with_index(self, tmp_path, json_file):
        jp = json_file(
            [
                _item("body_text", "A"),
                {"category": "body_txt", "paragraph": "B", "score": 1.0},
            ]
        )
        with pytest.raises(ValueError, match="第 2 条 'body_txt'"):
            json_to_docx(jp, save_dir=str(tmp_path))


class TestCliJd:
    def test_jd_mode_dispatch(self, tmp_path, json_file):
        jp = json_file([_item("body_text", "x")])
        with mock.patch("wordformat.cli.json_to_docx") as m:
            with mock.patch("sys.argv") as argv:
                argv.__getitem__.side_effect = lambda i: [
                    "wordf",
                    "jd",
                    "-f",
                    jp,
                    "-o",
                    str(tmp_path),
                ][i]
                argv.__len__.return_value = 5
                main()
            m.assert_called_once_with(
                json_path=jp, config_path=None, save_dir=str(tmp_path)
            )

    def test_af_sync_flag_dispatch(self, tmp_path):
        docx_p = tmp_path / "a.docx"
        Document().save(docx_p)
        cfg_p = tmp_path / "c.yaml"
        cfg_p.write_text("{}", encoding="utf-8")
        json_p = tmp_path / "n.json"
        json_p.write_text("[]", encoding="utf-8")
        with mock.patch("wordformat.cli.auto_format_thesis_document") as m:
            with mock.patch("sys.argv") as argv:
                argv.__getitem__.side_effect = lambda i: [
                    "wordf",
                    "af",
                    "-d",
                    str(docx_p),
                    "-c",
                    str(cfg_p),
                    "-f",
                    str(json_p),
                    "--sync",
                ][i]
                argv.__len__.return_value = 8
                main()
            kwargs = m.call_args.kwargs
            assert kwargs["sync"] is True
            assert kwargs["check"] is False
