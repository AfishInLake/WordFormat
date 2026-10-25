#! /usr/bin/env python
# @Time    : 2026/9/28
# @Author  : afish
# @File    : test_header_footer.py
"""页眉页脚检查/修正（header_footer.py）单元测试（issue#93）。"""

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from wordformat.header_footer import (
    DEFAULT_HF_CONFIG,
    _resolve_length_cm,
    _resolve_page_format,
    apply_header_footer,
    check_header_footer,
    resolve_hf_config,
)


def _set_east_asia(run, font_name: str) -> None:
    """手动设置 run 的中文字体（python-docx 不支持 eastAsia）。"""
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), font_name)


def _header_text(
    doc,
    text,
    size_pt=9.0,
    alignment=WD_ALIGN_PARAGRAPH.CENTER,
    en_font="Times New Roman",
    cn_font="宋体",
):
    """构造带指定格式的页眉段落并返回其 run。"""
    section = doc.sections[0]
    para = section.header.paragraphs[0]
    run = para.add_run(text)
    run.font.size = Pt(size_pt)
    run.font.name = en_font
    _set_east_asia(run, cn_font)
    para.alignment = alignment
    return run


def _fld_simple(part, instr: str) -> None:
    """往部件首段插入 w:fldSimple 简单域（Word 保存页码域的常见写法之一）。"""
    para = part.paragraphs[0]
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "1"
    run.append(text)
    fld.append(run)
    para._p.append(fld)


def _complex_field(part, instr_pieces) -> None:
    """插入复杂域（fldChar begin + 多段 instrText + fldChar end）。"""
    para = part.paragraphs[0]
    for piece in instr_pieces:
        run = OxmlElement("w:r")
        instr = OxmlElement("w:instrText")
        instr.text = piece
        run.append(instr)
        para._p.append(run)


class TestNoPartCreation:
    """检查模式不得凭空创建页眉/页脚部件（python-docx 读 .paragraphs 会创建）。"""

    def test_check_leaves_header_linked(self):
        doc = Document()
        check_header_footer(
            doc, {"header_footer": {"header": {"require_content": True}}}
        )
        assert doc.sections[0].header.is_linked_to_previous is True

    def test_check_leaves_footer_linked(self):
        doc = Document()
        apply_header_footer(doc)
        assert doc.sections[0].footer.is_linked_to_previous is True

    def test_apply_page_number_creates_part(self):
        doc = Document()
        cfg = {"header_footer": {"footer": {"page_number": True}}}
        apply_header_footer(doc, cfg)
        assert doc.sections[0].footer.is_linked_to_previous is False


class TestPageGeometry:
    """page_size / orientation：写在 sectPr/pgSz 上。"""

    _A4 = {"header_footer": {"page_setup": {"page_size": "A4"}}}

    def test_letter_doc_reported_when_a4_required(self):
        doc = Document()  # python-docx 模板默认 Letter
        issues = check_header_footer(doc, self._A4)
        assert any("纸张尺寸错误" in t and "规范:A4" in t for t in issues)

    def test_apply_sets_a4_dims(self):
        doc = Document()
        assert apply_header_footer(doc, self._A4) == []
        section = doc.sections[0]
        assert round(section.page_width.cm, 1) == 21.0
        assert round(section.page_height.cm, 1) == 29.7

    def test_unknown_paper_size_skipped(self):
        doc = Document()
        cfg = {"header_footer": {"page_setup": {"page_size": "B6"}}}
        assert not any("纸张" in t for t in check_header_footer(doc, cfg))

    def test_orientation_mismatch_reported(self):
        doc = Document()
        cfg = {"header_footer": {"page_setup": {"orientation": "横向"}}}
        issues = check_header_footer(doc, cfg)
        assert any("纸张方向错误" in t for t in issues)

    def test_apply_orientation_swaps_dims(self):
        doc = Document()
        cfg = {"header_footer": {"page_setup": {"orientation": "横向"}}}
        before_w = doc.sections[0].page_width.cm
        assert apply_header_footer(doc, cfg) == []
        section = doc.sections[0]
        assert section.page_width.cm > before_w
        assert section._sectPr.find(qn("w:pgSz")).get(qn("w:orient")) == "landscape"

    def test_a4_landscape_expects_swapped_dims(self):
        doc = Document()
        cfg = {
            "header_footer": {"page_setup": {"page_size": "A4", "orientation": "横向"}}
        }
        apply_header_footer(doc, cfg)
        section = doc.sections[0]
        assert round(section.page_width.cm, 1) == 29.7
        assert round(section.page_height.cm, 1) == 21.0
        assert check_header_footer(doc, cfg) == []


class TestPageNumberType:
    """page_number_start / page_number_format：写在 sectPr/pgNumType 上。"""

    _CFG = {
        "header_footer": {
            "page_setup": {"page_number_start": 1, "page_number_format": "小写罗马数字"}
        }
    }

    def test_missing_pg_num_type_reported(self):
        doc = Document()
        issues = check_header_footer(doc, self._CFG)
        texts = "\n".join(issues)
        assert "起始页码错误" in texts
        assert "页码格式错误" in texts

    def test_apply_writes_pg_num_type(self):
        doc = Document()
        assert apply_header_footer(doc, self._CFG) == []
        node = doc.sections[0]._sectPr.find(qn("w:pgNumType"))
        assert node.get(qn("w:start")) == "1"
        assert node.get(qn("w:fmt")) == "lowerRoman"
        assert check_header_footer(doc, self._CFG) == []

    def test_pg_num_type_ordered_before_cols(self):
        doc = Document()
        apply_header_footer(doc, self._CFG)
        tags = [c.tag.split("}")[1] for c in doc.sections[0]._sectPr]
        assert tags.index("pgNumType") < tags.index("cols")

    def test_not_checked_without_config(self):
        doc = Document()
        assert not any("页码格式" in t for t in check_header_footer(doc))

    @staticmethod
    def test_format_labels_and_raw_values():
        assert _resolve_page_format("小写罗马数字") == "lowerRoman"
        assert _resolve_page_format("UpperRoman") == "upperRoman"
        assert _resolve_page_format(" nonsense") is None
        assert _resolve_page_format(None) is None


class TestFirstPageAndDocFlags:
    """首页不同（sectPr/titlePg）与文档级开关（settings.xml）。"""

    def test_different_first_page_check_and_apply(self):
        doc = Document()
        cfg = {"header_footer": {"page_setup": {"different_first_page": True}}}
        assert any("首页不同" in t for t in check_header_footer(doc, cfg))
        assert apply_header_footer(doc, cfg) == []
        assert doc.sections[0].different_first_page_header_footer is True

    def test_disabling_first_page_flag_removes_title_pg(self):
        doc = Document()
        doc.sections[0].different_first_page_header_footer = True
        cfg = {"header_footer": {"page_setup": {"different_first_page": False}}}
        assert any("首页不同" in t for t in check_header_footer(doc, cfg))
        assert apply_header_footer(doc, cfg) == []
        assert doc.sections[0].different_first_page_header_footer is False

    def test_odd_even_pages_doc_flag(self):
        doc = Document()
        cfg = {"header_footer": {"odd_even_pages": True}}
        assert any("奇偶页不同" in t for t in check_header_footer(doc, cfg))
        assert apply_header_footer(doc, cfg) == []
        element = doc.settings.element.find(qn("w:evenAndOddHeaders"))
        assert element is not None
        assert check_header_footer(doc, cfg) == []

    def test_update_fields_on_open_doc_flag(self):
        doc = Document()
        cfg = {"header_footer": {"update_fields_on_open": True}}
        assert any("打开时更新域" in t for t in check_header_footer(doc, cfg))
        assert apply_header_footer(doc, cfg) == []
        element = doc.settings.element.find(qn("w:updateFields"))
        assert element.get(qn("w:val")) == "true"

    def test_doc_flags_untouched_without_config(self):
        doc = Document()
        doc.settings.element.append(_mk_update_fields())
        assert check_header_footer(doc) == []
        assert apply_header_footer(doc) == []
        assert doc.settings.element.find(qn("w:updateFields")) is not None

    def test_explicit_false_removes_existing_update_fields(self):
        doc = Document()
        doc.settings.element.append(_mk_update_fields())
        cfg = {"header_footer": {"update_fields_on_open": False}}
        assert any("打开时更新域" in t for t in check_header_footer(doc, cfg))
        apply_header_footer(doc, cfg)
        assert doc.settings.element.find(qn("w:updateFields")) is None


def _mk_update_fields():
    element = OxmlElement("w:updateFields")
    element.set(qn("w:val"), "true")
    return element


class TestPageNumberField:
    """header/footer 的 page_number：页码域缺失检查与自动插入。"""

    _CFG = {"header_footer": {"footer": {"page_number": True}}}

    def test_missing_field_reported(self):
        doc = Document()
        issues = check_header_footer(doc, self._CFG)
        assert any("页脚(第1节)-页码缺失" in t for t in issues)

    def test_apply_inserts_page_field(self):
        doc = Document()
        assert apply_header_footer(doc, self._CFG) == []
        footer = doc.sections[0].footer
        assert footer._element.xpath(".//w:fldSimple/@w:instr") == [" PAGE "]
        assert check_header_footer(doc, self._CFG) == []

    def test_existing_simple_field_not_reported(self):
        doc = Document()
        _fld_simple(doc.sections[0].footer, " PAGE ")
        assert check_header_footer(doc, self._CFG) == []

    def test_split_instr_text_field_not_reported(self):
        doc = Document()
        _complex_field(doc.sections[0].footer, [" P", "AGE ", "\\* MERGEFORMAT"])
        assert check_header_footer(doc, self._CFG) == []

    def test_numpages_alone_still_missing(self):
        # NUMPAGES 是总页数域，不能当作页码
        doc = Document()
        _fld_simple(doc.sections[0].footer, " NUMPAGES ")
        issues = check_header_footer(doc, self._CFG)
        assert any("页码缺失" in t for t in issues)

    def test_apply_does_not_duplicate_existing_field(self):
        doc = Document()
        _fld_simple(doc.sections[0].footer, " PAGE ")
        apply_header_footer(doc, self._CFG)
        footer = doc.sections[0].footer
        assert len(footer._element.xpath(".//w:fldSimple")) == 1

    def test_page_number_satisfies_require_content(self):
        cfg = {
            "header_footer": {"footer": {"page_number": True, "require_content": True}}
        }
        doc = Document()
        issues = check_header_footer(doc, cfg)
        assert any("页码缺失" in t for t in issues)
        assert not any("内容缺失" in t for t in issues)

    def test_no_page_number_config_keeps_old_behaviour(self):
        doc = Document()
        assert check_header_footer(doc) == []

    def test_linked_section_footer_skipped(self):
        doc = Document()
        _fld_simple(doc.sections[0].footer, " PAGE ")
        doc.add_section()  # 沿用第 1 节页脚
        assert check_header_footer(doc, self._CFG) == []


class TestHeaderVariants:
    """首页/偶页变体：仅在对应开关生效时才参与检查。"""

    def test_first_page_variant_checked_when_enabled(self):
        doc = Document()
        section = doc.sections[0]
        section.different_first_page_header_footer = True
        run = section.first_page_header.paragraphs[0].add_run("首页页眉")
        run.font.size = Pt(12)
        issues = check_header_footer(doc)
        assert any("页眉(第1节·首页)-字号错误" in t for t in issues)

    def test_absent_first_page_variant_not_reported(self):
        doc = Document()
        cfg = {
            "header_footer": {
                "page_setup": {"different_first_page": True},
                "header": {"require_content": True},
            }
        }
        doc.sections[0].different_first_page_header_footer = True
        _header_text(doc, "某某大学毕业论文")
        # 首页变体没有自己的部件（Word 显示空白首页页眉）→ 不报缺失，避免误伤
        assert check_header_footer(doc, cfg) == []

    def test_even_page_variant_checked_when_doc_flag_on(self):
        doc = Document()
        cfg = {"header_footer": {"odd_even_pages": True}}
        apply_header_footer(doc, cfg)
        run = doc.sections[0].even_page_footer.paragraphs[0].add_run("偶页页脚")
        run.font.size = Pt(12)
        issues = check_header_footer(doc, cfg)
        assert any("页脚(第1节·偶页)-字号错误" in t for t in issues)


class TestResolveHfConfig:
    def test_none_returns_default(self):
        cfg = resolve_hf_config(None)
        assert cfg == DEFAULT_HF_CONFIG

    def test_dict_merges_user_config(self):
        cfg = resolve_hf_config({"header_footer": {"header": {"font_size": "五号"}}})
        assert cfg["header"]["font_size"] == "五号"
        assert cfg["footer"]["font_size"] == "小五"  # 深合并保留缺省

    def test_disabled_returns_empty(self):
        assert resolve_hf_config({"header_footer": {"enabled": False}}) == {}

    def test_config_object_attribute(self):
        class FakeConfig:
            header_footer = {"footer": {"chinese_font_name": "黑体"}}

        cfg = resolve_hf_config(FakeConfig())
        assert cfg["footer"]["chinese_font_name"] == "黑体"


class TestCheckHeaderFooter:
    def test_ok_header_no_issues(self):
        doc = Document()
        _header_text(doc, "某某大学毕业论文")
        assert check_header_footer(doc) == []

    def test_wrong_alignment_reported(self):
        doc = Document()
        _header_text(doc, "某某大学毕业论文", alignment=WD_ALIGN_PARAGRAPH.LEFT)
        issues = check_header_footer(doc)
        assert any("对齐错误" in t and "页眉(第1节)" in t for t in issues)

    def test_wrong_font_size_reported(self):
        doc = Document()
        _header_text(doc, "某某大学毕业论文", size_pt=12)  # 小四 ≫ 小五
        issues = check_header_footer(doc)
        assert any("字号错误" in t for t in issues)

    def test_wrong_fonts_reported(self):
        doc = Document()
        _header_text(doc, "某某大学毕业论文 Thesis", en_font="Arial", cn_font="黑体")
        issues = check_header_footer(doc)
        texts = "\n".join(issues)
        assert "中文字体错误" in texts
        assert "英文字体错误" in texts

    def test_chinese_only_header_skips_english_font(self):
        """纯中文 run 的 ascii 字体不影响显示，不报英文字体错误。"""
        doc = Document()
        _header_text(doc, "某某大学毕业论文", en_font="Arial", cn_font="宋体")
        issues = check_header_footer(doc)
        assert not any("英文字体错误" in t for t in issues)

    def test_empty_header_with_require_content(self):
        cfg = {"header_footer": {"header": {"require_content": True}}}
        doc = Document()  # 页眉为空
        issues = check_header_footer(doc, cfg)
        assert any("内容缺失" in t for t in issues)

    def test_empty_header_without_require_content(self):
        doc = Document()
        assert check_header_footer(doc) == []

    def test_linked_to_previous_section_skipped(self):
        doc = Document()
        _header_text(doc, "某某大学毕业论文", size_pt=12)
        doc.add_section()
        # 第二节沿用第一节（is_linked_to_previous=True），其页眉不参与检查
        assert doc.sections[1].header.is_linked_to_previous is True
        issues = check_header_footer(doc)
        # 只有第一节问题；第二节标题不带 (第2节) 前缀
        assert len(issues) == 1

    def test_disabled_config_skips(self):
        doc = Document()
        _header_text(doc, "某某大学毕业论文", size_pt=10.5)
        assert check_header_footer(doc, {"header_footer": {"enabled": False}}) == []


class TestApplyHeaderFooter:
    def test_apply_fixes_format_no_remaining(self):
        doc = Document()
        _header_text(
            doc,
            "某某大学毕业论文",
            size_pt=10.5,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            en_font="Arial",
            cn_font="黑体",
        )
        remaining = apply_header_footer(doc)
        assert remaining == []
        assert check_header_footer(doc) == []

    def test_apply_empty_header_leaves_remaining_empty(self):
        doc = Document()  # 页眉为空：不自动补内容，无残留
        assert apply_header_footer(doc) == []

    def test_apply_empty_required_header_reports(self):
        # 显式 require_content 时，apply 模式也必须报告空页眉，不能静默放过
        cfg = {"header_footer": {"header": {"require_content": True}}}
        doc = Document()
        remaining = apply_header_footer(doc, cfg)
        assert any("页眉(第1节)-内容缺失" in t for t in remaining)

    def test_apply_disabled_config_noop(self):
        doc = Document()
        run = _header_text(doc, "某某大学毕业论文", size_pt=10.5)
        assert apply_header_footer(doc, {"header_footer": {"enabled": False}}) == []
        assert run.font.size.pt == 10.5

    def test_apply_second_section_linked_skipped(self):
        doc = Document()
        _header_text(doc, "某某大学毕业论文", size_pt=10.5)
        doc.add_section()
        apply_header_footer(doc)
        # 第二节沿用第一节（未指认新内容），修正后第一节字号为 9pt
        assert doc.sections[0].header.paragraphs[0].runs[0].font.size.pt == 9.0


class TestResolveLengthCm:
    def test_units_and_defaults(self):
        assert _resolve_length_cm("2.54cm") == 2.54
        assert abs(_resolve_length_cm("31.8mm") - 3.18) < 1e-6
        assert _resolve_length_cm("1.5") == 1.5  # 缺省 cm
        assert _resolve_length_cm(3) == 3.0  # 数字按 cm
        assert _resolve_length_cm("1in") == 2.54

    def test_unparseable_returns_none(self):
        assert _resolve_length_cm("abc") is None
        assert _resolve_length_cm(None) is None


class TestPageMargin:
    """页边距为 opt-in：仅当配置 header_footer.page_setup 才校验/修正。"""

    _CFG = {
        "header_footer": {
            "page_setup": {
                "top_margin": "3.0cm",
                "bottom_margin": "3.0cm",
                "left_margin": "2.5cm",
                "right_margin": "2.5cm",
            }
        }
    }

    def test_not_checked_without_config(self):
        # 未配 page_setup 时，默认文档不报批边距问题（向后兼容）
        doc = Document()
        assert not any("页面" in t for t in check_header_footer(doc))

    def test_check_reports_margin_error(self):
        doc = Document()
        doc.sections[0].top_margin = Cm(2.0)  # 非 3.0cm
        issues = check_header_footer(doc, self._CFG)
        assert any("上边距错误" in t and "页面(第1节)" in t for t in issues)

    def test_check_ok_when_matches(self):
        doc = Document()
        s = doc.sections[0]
        s.top_margin = s.bottom_margin = Cm(3.0)
        s.left_margin = s.right_margin = Cm(2.5)
        assert not any("页面" in t for t in check_header_footer(doc, self._CFG))

    def test_apply_fixes_margins(self):
        doc = Document()
        doc.sections[0].top_margin = Cm(2.0)
        remaining = apply_header_footer(doc, self._CFG)
        assert remaining == []
        assert round(doc.sections[0].top_margin.cm, 2) == 3.0

    def test_partial_config_only_checks_given_fields(self):
        cfg = {"header_footer": {"page_setup": {"left_margin": "2.5cm"}}}
        doc = Document()
        doc.sections[0].top_margin = Cm(9.9)  # 未配 top_margin，不应报
        doc.sections[0].left_margin = Cm(5.0)  # 配了，应报
        issues = [t for t in check_header_footer(doc, cfg) if "页面" in t]
        assert any("左边距" in t for t in issues)
        assert not any("上边距" in t for t in issues)
