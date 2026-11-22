#! /usr/bin/env python
# @Time    : 2026-10-28
# @Author  : afish
# @File    : test_text_edit.py
"""run 级文本改写单元测试（text_edit.py）。"""

import base64

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from wordformat.text_edit import paragraph_text, set_paragraph_text

_PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ"
    "AAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def _para(texts):
    doc = Document()
    para = doc.add_paragraph()
    for t in texts:
        para.add_run(t)
    return para


def _add_hyperlink(para, text):
    link = OxmlElement("w:hyperlink")
    run = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    run.append(t)
    link.append(run)
    para._p.append(link)
    return link


def _add_page_field(para, cached="7"):
    """构造 PAGE 域：begin / instrText / separate / 结果 run / end。"""
    for kind, text in (
        ("begin", None),
        (None, " PAGE "),
        ("separate", None),
        (None, cached),
        ("end", None),
    ):
        run = OxmlElement("w:r")
        if kind is not None:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        elif text == " PAGE ":
            el = OxmlElement("w:instrText")
            el.text = text
        else:
            el = OxmlElement("w:t")
            el.text = text
        run.append(el)
        para._p.append(run)


class TestCharAlignment:
    def test_insert_between_runs_keeps_each_run_format(self):
        para = _para(["本文提出", "多智能体", "方法"])
        para.runs[1].bold = True

        assert set_paragraph_text(para, "本文提出一种多智能体协同方法") is True

        assert [r.text for r in para.runs] == ["本文提出一种", "多智能体协同", "方法"]
        assert [r.bold for r in para.runs] == [None, True, None]

    def test_replaced_characters_inherit_start_run_format(self):
        para = _para(["普通", "加粗"])
        para.runs[1].bold = True

        set_paragraph_text(para, "普通强调")

        assert [r.text for r in para.runs] == ["普通", "强调"]
        assert para.runs[1].bold is True

    def test_deletion_across_run_boundary(self):
        para = _para(["abcd", "efgh"])

        set_paragraph_text(para, "abgh")

        assert "".join(r.text for r in para.runs) == "abgh"

    def test_no_change_returns_false_and_touches_nothing(self):
        para = _para(["甲", "乙"])
        xml_before = para._p.xml

        assert set_paragraph_text(para, "甲乙") is False
        assert para._p.xml == xml_before

    def test_leading_trailing_whitespace_only_is_no_change(self):
        """gj 导出的 paragraph 已 strip：sync 不应抹掉原段落的定位空格。"""
        para = _para(["\u3000\u3000作者签名：      日期： "])
        xml_before = para._p.xml

        assert set_paragraph_text(para, "作者签名：      日期：") is False
        assert para._p.xml == xml_before

    def test_clear_text_empties_runs_without_removing_them(self):
        para = _para(["要删掉的文字"])

        assert set_paragraph_text(para, "") is True

        assert para.text == ""
        assert len(para.runs) == 1  # run 结构保留，只是文本清空

    def test_empty_paragraph_gets_a_new_run(self):
        para = _para([])

        set_paragraph_text(para, "全新内容")

        assert para.text == "全新内容"

    def test_tab_and_line_break_survive(self):
        para = _para(["a\tb"])

        set_paragraph_text(para, "a\tbc")

        assert para.runs[0].text == "a\tbc"


class TestStructureProtection:
    def test_hyperlink_content_untouched_when_plain_text_edited(self):
        para = _para(["见 "])
        _add_hyperlink(para, "文献1")

        set_paragraph_text(para, "参见 文献1")

        assert para.text == "参见 文献1"
        link = para._p.find(qn("w:hyperlink"))
        assert link is not None
        assert "".join(t.text or "" for t in link.xpath(".//w:t")) == "文献1"
        # 新增的字落在普通 run 上，没有污染超链接
        assert para.runs[0].text == "参见 "

    def test_image_run_survives_text_rewrite(self, tmp_path):
        png = tmp_path / "fig.png"
        png.write_bytes(_PNG_1PX)
        para = Document().add_paragraph()
        para.add_run().add_picture(str(png))
        para.add_run("旧题注")

        set_paragraph_text(para, "图 1-1 新题注")

        assert len(para._p.xpath(".//w:drawing")) == 1
        assert para.text == "图 1-1 新题注"

    def test_field_machinery_survives_text_rewrite(self):
        para = _para(["正文"])
        _add_page_field(para)

        set_paragraph_text(para, "正文改后")

        assert len(para._p.xpath(".//w:fldChar")) == 3
        assert len(para._p.xpath(".//w:instrText")) == 1
        assert " PAGE " in para._p.xml

    def test_math_paragraph_not_rewritten_into(self):
        # OMML 公式 run 不可改写；文字部分照常更新
        para = _para(["式 "])
        run = para.add_run()
        math_el = OxmlElement("m:oMath")
        run._r.append(math_el)

        set_paragraph_text(para, "式 （2-1）")

        assert len(para._p.xpath(".//m:oMath")) == 1
        assert para.runs[0].text == "式 （2-1）"
        assert para.runs[1].text == ""  # 公式 run 未被写入文字


class TestParagraphText:
    def test_matches_docx_text_including_hyperlink(self):
        para = _para(["普通"])
        _add_hyperlink(para, "链接")

        assert paragraph_text(para) == para.text == "普通链接"

    def test_excludes_field_result_like_docx(self):
        para = _para(["正文"])
        for tag, text in (("w:fldSimple", "42"),):
            el = OxmlElement(tag)
            run = OxmlElement("w:r")
            t = OxmlElement("w:t")
            t.text = text
            run.append(t)
            el.append(run)
            para._p.append(el)

        assert paragraph_text(para) == para.text == "正文"
