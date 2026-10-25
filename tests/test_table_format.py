#! /usr/bin/env python
# @Time    : 2026/10/22
# @Author  : afish
# @File    : test_table_format.py
"""通用表格样式档案（rules/object.py）：任意边框/列宽/行属性/单元格格式。"""

import pytest
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

from wordformat.rules.object import TableObject

ROWS = [["名称", "数值"], ["A", "1"], ["B", "2"]]


def _node(profile_cfg, **value_extra):
    node = TableObject(
        value={"category": "table_object", "paragraph": "", **value_extra},
        level=999,
    )
    config = {"tables": profile_cfg} if profile_cfg is not None else {}
    node.load_config(config)
    return node


def _build(profile_cfg, **value_extra):
    node = _node(profile_cfg, table_rows=ROWS, **value_extra)
    doc = Document()
    node._try_insert_table(doc)
    return doc.tables[0] if doc.tables else None


def _three_line_profile():
    return {
        "default": "three_line",
        "styles": {
            "three_line": {
                "borders": {
                    "top": {"style": "single", "size_pt": 1.5},
                    "bottom": {"style": "single", "size_pt": 1.5},
                    "inside_h": {"style": "single", "size_pt": 0.75},
                    "left": {"style": "none"},
                    "right": {"style": "none"},
                    "inside_v": {"style": "none"},
                },
            }
        },
    }


class TestTableProfile:
    def test_default_profile_borders(self):
        table = _build(_three_line_profile())
        borders = table._tbl.tblPr.find(qn("w:tblBorders"))
        assert borders is not None
        edge = {child.tag: child for child in borders}
        assert edge[qn("w:top")].get(qn("w:val")) == "single"
        assert edge[qn("w:top")].get(qn("w:sz")) == "3"  # 1.5pt -> 3 half-points
        assert edge[qn("w:insideH")].get(qn("w:sz")) == "2"  # 0.75 -> 1.5 取整 2
        assert edge[qn("w:insideV")].get(qn("w:val")) == "none"
        assert edge[qn("w:left")].get(qn("w:val")) == "none"

    def test_item_level_table_style_overrides_default(self):
        cfg = {
            "default": "plain",
            "styles": {
                "plain": {"borders": {"top": {"style": "double"}}},
                "fancy": {"borders": {"bottom": {"style": "dashed"}}},
            },
        }
        table = _build(cfg, table_style="fancy")
        borders = table._tbl.tblPr.find(qn("w:tblBorders"))
        tags = {child.tag for child in borders}
        assert qn("w:bottom") in tags
        assert qn("w:top") not in tags

    def test_missing_profile_raises(self):
        cfg = {"default": "not_exist", "styles": {"grid": {}}}
        node = _node(cfg, table_rows=ROWS)
        with pytest.raises(ValueError, match="表格样式档案不存在"):
            node._try_insert_table(Document())

    def test_bad_border_edge_raises(self):
        cfg = {"styles": {"bad": {"borders": {"diagonal": {"style": "single"}}}}}
        node = _node(cfg, table_rows=ROWS, table_style="bad")
        with pytest.raises(ValueError, match="边名非法"):
            node._try_insert_table(Document())

    def test_border_shorthand_string(self):
        cfg = {"styles": {"s": {"borders": {"top": "double"}}}}
        table = _build(cfg, table_style="s")
        top = table._tbl.tblPr.find(qn("w:tblBorders")).find(qn("w:top"))
        assert top.get(qn("w:val")) == "double"

    def test_default_without_tables_section_keeps_table_grid(self):
        table = _build(None)
        assert table.style.name == "Table Grid"
        assert table._tbl.tblPr.find(qn("w:tblBorders")) is None


class TestTableProps:
    def test_custom_col_widths(self):
        cfg = {
            "default": "w",
            "styles": {"w": {"col_widths": {"mode": "custom", "custom": [4, 6]}}},
        }
        table = _build(cfg)
        grid = table._tbl.tblGrid
        widths = [int(col.get(qn("w:w"))) for col in grid.findall(qn("w:gridCol"))]
        assert widths[0] == int(round(4 * 567))
        assert widths[1] == int(round(6 * 567))
        # 单元格 tcW 也必须是自定义宽度，避免与 gridCol 冲突
        expected = [int(round(4 * 567)), int(round(6 * 567))]
        for row in table.rows:
            actual = [
                int(tc.find(qn("w:tcPr")).find(qn("w:tcW")).get(qn("w:w")))
                for tc in row._tr.findall(qn("w:tc"))
            ]
            assert actual == expected

    def test_equal_col_widths_keeps_default(self):
        cfg = {"default": "w", "styles": {"w": {"col_widths": {"mode": "equal"}}}}
        table = _build(cfg)
        # equal 模式沿用 python-docx 默认等宽，不写自定义宽度
        grid = table._tbl.tblGrid
        widths = {int(col.get(qn("w:w"))) for col in grid.findall(qn("w:gridCol"))}
        assert len(widths) == 1

    def test_layout_and_alignment_and_indent(self):
        cfg = {
            "default": "p",
            "styles": {
                "p": {
                    "layout": "fixed",
                    "alignment": "center",
                    "indent_cm": 0.5,
                    "width": {"type": "pct", "value": 100},
                }
            },
        }
        table = _build(cfg)
        tbl_pr = table._tbl.tblPr
        layout = tbl_pr.find(qn("w:tblLayout"))
        assert layout is not None and layout.get(qn("w:type")) == "fixed"
        jc = tbl_pr.find(qn("w:jc"))
        assert jc is not None and jc.get(qn("w:val")) == "center"
        ind = tbl_pr.find(qn("w:tblInd"))
        assert ind is not None and ind.get(qn("w:w")) == str(int(round(0.5 * 567)))
        width = tbl_pr.find(qn("w:tblW"))
        assert width.get(qn("w:type")) == "pct"
        assert width.get(qn("w:w")) == "5000"


class TestTableRowsAndCells:
    def test_cell_margins(self):
        cfg = {
            "default": "m",
            "styles": {
                "m": {"cell_margins_cm": {"left": 0.19, "right": 0.19, "top": 0}}
            },
        }
        table = _build(cfg)
        mar = table._tbl.tblPr.find(qn("w:tblCellMar"))
        assert mar is not None
        tags = {child.tag for child in mar}
        assert qn("w:left") in tags and qn("w:top") in tags

    def test_header_row_repeat_and_cant_split(self):
        cfg = {
            "default": "h",
            "styles": {"h": {"header_row": {"repeat": True, "cant_split": True}}},
        }
        table = _build(cfg)
        header_tr_pr = table.rows[0]._tr.trPr
        assert header_tr_pr.find(qn("w:tblHeader")) is not None
        for row in table.rows:
            assert row._tr.trPr.find(qn("w:cantSplit")) is not None

    def test_row_bands_fill_and_bold(self):
        cfg = {
            "default": "b",
            "styles": {
                "b": {"row_bands": [{"rows": [0], "fill": "EEEEEE", "bold": True}]}
            },
        }
        table = _build(cfg)
        cell = table.cell(0, 0)
        shd = cell._tc.tcPr.find(qn("w:shd"))
        assert shd is not None and shd.get(qn("w:fill")) == "EEEEEE"
        assert cell.paragraphs[0].runs[0].bold is True

    def test_cell_run_font(self):
        cfg = {
            "default": "c",
            "styles": {
                "c": {
                    "cell_run": {
                        "chinese_font_name": "黑体",
                        "english_font_name": "Arial",
                        "font_size": "五号",
                    }
                }
            },
        }
        table = _build(cfg)
        run = table.cell(0, 0).paragraphs[0].runs[0]
        assert run.font.name == "Arial"
        rfonts = run._element.rPr.find(qn("w:rFonts"))
        assert rfonts.get(qn("w:eastAsia")) == "黑体"

    def test_cell_run_does_not_reset_row_band_bold(self):
        # cell_run 未显式配置加粗时，不得把 row_bands 的 bold 重置为 False
        cfg = {
            "default": "mix",
            "styles": {
                "mix": {
                    "row_bands": [{"rows": "all", "bold": True}],
                    "cell_run": {"font_size": "五号"},
                }
            },
        }
        table = _build(cfg)
        for row in table.rows:
            for cell in row.cells:
                run = cell.paragraphs[0].runs[0]
                assert run.bold is True
                assert run.font.size == Pt(10.5)

    def test_row_bands_invalid_rows_raises(self):
        cfg = {"styles": {"bad": {"row_bands": [{"rows": "first"}]}}}
        node = _node(cfg, table_rows=ROWS, table_style="bad")
        with pytest.raises(ValueError, match="row_bands.rows"):
            node._try_insert_table(Document())

    def test_row_bands_negative_index_ignored(self):
        cfg = {"styles": {"b": {"row_bands": [{"rows": [-1], "bold": True}]}}}
        table = _build(cfg, table_style="b")
        assert table.cell(0, 0).paragraphs[0].runs[0].bold is not True

    def test_existing_table_not_reinserted(self):
        # exists=True 的原文档表格：不重复建表
        node = _node(None, table_rows=ROWS, exists=True)
        doc = Document()
        node._try_insert_table(doc)  # 直接调用建表方法仍会建表
        doc2 = Document()
        node._base(doc2, p=False, r=False)
        assert len(doc2.tables) == 0
