"""
rules 模块测试 —— 聚焦真实行为验证，无填充。
"""

import pytest
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.shared import Pt

from wordformat.rules import (
    BodyText,
)
from wordformat.rules.body import _looks_like_code

# ---------------------------------------------------------------------------
# 共享 fixtures / helpers
# ---------------------------------------------------------------------------


def _make_node(cls, text="测试文本"):
    """创建一个带 paragraph 的 FormatNode 实例。"""
    doc = Document()
    p = doc.add_paragraph(text)
    return cls(value=p, level=0, paragraph=p)


def _load_root_config(config_path):
    """从 YAML 路径加载配置 dict。"""
    return _load_yaml(config_path)


def _load_yaml(path):
    import yaml

    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


@pytest.fixture
def root_config(sample_yaml_config):
    """从 sample_yaml_config 加载 NodeConfigRoot，与示例文件解耦。"""
    from wordformat.config.loader import init_config

    init_config(sample_yaml_config)
    return _load_root_config(sample_yaml_config)


@pytest.fixture
def doc():
    return Document()


@pytest.fixture
def para(doc):
    return doc.add_paragraph("测试段落")


@pytest.fixture
def run_with_text(para):
    r = para.add_run("关键词")
    r.font.size = Pt(12)
    r.font.bold = True
    return r


# ---------------------------------------------------------------------------
# 1. FormatNode 基类行为
# ---------------------------------------------------------------------------


class TestBodyTextCitationSuperscript:
    """测试 BodyText.apply_format 中的引用上标自动设置。"""

    @staticmethod
    def _get_vertAlign(run):
        """返回 run 的 w:vertAlign 值，无则返回 None。"""
        rPr = run._element.find(qn("w:rPr"))
        if rPr is None:
            return None
        va = rPr.find(qn("w:vertAlign"))
        return va.get(qn("w:val")) if va is not None else None

    def test_single_citation_gets_superscript(self):
        """单引用 [1] 应被设为上标。"""
        doc = Document()
        p = doc.add_paragraph("这是一篇论文[1]的引用。")
        node = BodyText(
            value={"category": "body_text", "fingerprint": "fp1"}, level=0, paragraph=p
        )
        # 直接调用 apply_format（跳过 load_config，不依赖完整配置）
        node._apply_citation_superscript()
        # "[1]" 应在独立的上标 run 中
        runs = p.runs
        superscript_texts = [
            r.text for r in runs if self._get_vertAlign(r) == "superscript"
        ]
        assert "[1]" in superscript_texts

    def test_multiple_citations(self):
        """多个引用 [1] 和 [2,3] 都应被设为上标。"""
        doc = Document()
        p = doc.add_paragraph("参见文献[1]和[2,3]的讨论。")
        node = BodyText(
            value={"category": "body_text", "fingerprint": "fp2"}, level=0, paragraph=p
        )
        node._apply_citation_superscript()
        superscript_texts = [
            r.text for r in p.runs if self._get_vertAlign(r) == "superscript"
        ]
        assert "[1]" in superscript_texts
        assert "[2,3]" in superscript_texts

    def test_range_citation(self):
        """范围引用 [1-3] 应被设为上标。"""
        doc = Document()
        p = doc.add_paragraph("文献[1-3]提供了详细分析。")
        node = BodyText(
            value={"category": "body_text", "fingerprint": "fp3"}, level=0, paragraph=p
        )
        node._apply_citation_superscript()
        superscript_texts = [
            r.text for r in p.runs if self._get_vertAlign(r) == "superscript"
        ]
        assert "[1-3]" in superscript_texts

    def test_chinese_comma_citation(self):
        """中文逗号分隔的引用 [1，2] 应被设为上标。"""
        doc = Document()
        p = doc.add_paragraph("见[1，2，3]的研究。")
        node = BodyText(
            value={"category": "body_text", "fingerprint": "fp4"}, level=0, paragraph=p
        )
        node._apply_citation_superscript()
        superscript_texts = [
            r.text for r in p.runs if self._get_vertAlign(r) == "superscript"
        ]
        assert any("[1，2，3]" in t for t in superscript_texts)

    def test_no_citation_leaves_runs_unchanged(self):
        """无引用的段落应保持原样。"""
        doc = Document()
        p = doc.add_paragraph("这是一段没有引用的正文。")
        node = BodyText(
            value={"category": "body_text", "fingerprint": "fp5"}, level=0, paragraph=p
        )
        original_text = p.text
        node._apply_citation_superscript()
        # 文本不变，且无上标 run
        assert p.text == original_text
        assert all(self._get_vertAlign(r) is None for r in p.runs)

    def test_non_citation_brackets_not_affected(self):
        """非数字方括号如 [注] 不应被设为上标。"""
        doc = Document()
        p = doc.add_paragraph("这是一个[注]释说明。")
        node = BodyText(
            value={"category": "body_text", "fingerprint": "fp6"}, level=0, paragraph=p
        )
        node._apply_citation_superscript()
        assert all(self._get_vertAlign(r) is None for r in p.runs)

    def test_citation_split_across_runs(self):
        """引用跨 run 时先分割再设上标。"""
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("前面文字[1")
        p.add_run("2]后面文字")
        node = BodyText(
            value={"category": "body_text", "fingerprint": "fp7"}, level=0, paragraph=p
        )
        node._apply_citation_superscript()
        superscript_texts = [
            r.text for r in p.runs if self._get_vertAlign(r) == "superscript"
        ]
        # 分割后 [1 和 2] 分别在两个 run 中，但都是上标
        assert "[1" in superscript_texts
        assert "2]" in superscript_texts


# ---------------------------------------------------------------------------
# 8. AbstractTitleCN._base 完整 diff/apply 覆盖
# ---------------------------------------------------------------------------


class TestBodyTextEdges:
    """字数规则与引用上标的空段落/特殊 run 分支。"""

    def test_length_skips_missing_paragraph(self):
        node = BodyText(value={}, level=0, paragraph=None)
        node._check_paragraph_length(Document(), {"min_chars": 1})

    def test_superscript_skips_missing_paragraph(self):
        node = BodyText(value={}, level=0, paragraph=None)
        node._apply_citation_superscript()

    def test_apply_format_sets_superscript_in_chapter_one(self):
        doc = Document()
        p = doc.add_paragraph("这是一篇论文[1]的引用。")
        node = BodyText(
            value={"category": "body_text", "chapter_number": 1},
            level=0,
            paragraph=p,
        )
        node.load_config({})
        node.apply_format(doc)
        superscript_texts = [r.text for r in p.runs if r.font.superscript]
        assert "[1]" in superscript_texts

    def test_apply_format_skips_superscript_outside_chapter_one(self):
        doc = Document()
        p = doc.add_paragraph("这是一篇论文[1]的引用。")
        node = BodyText(
            value={"category": "body_text", "chapter_number": 2},
            level=0,
            paragraph=p,
        )
        node.load_config({})
        node.apply_format(doc)
        assert all(r.font.superscript is None for r in p.runs)

    def test_superscript_skips_run_without_text_element(self):
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("文献[1]")
        p.add_run().add_break()
        node = BodyText(value={}, level=0, paragraph=p)
        node._apply_citation_superscript()
        assert p.text.strip() == "文献[1]"

    def test_superscript_keeps_spaces_around_citation(self):
        doc = Document()
        p = doc.add_paragraph("文字 [1] 文字")
        node = BodyText(value={}, level=0, paragraph=p)
        node._apply_citation_superscript()
        assert p.text == "文字 [1] 文字"
        superscript = [
            run.text
            for run in p.runs
            if TestBodyTextCitationSuperscript._get_vertAlign(run) == "superscript"
        ]
        assert superscript == ["[1]"]

    def test_superscript_preserves_run_formatting(self):
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("文献[1]").bold = True
        node = BodyText(value={}, level=0, paragraph=p)
        node._apply_citation_superscript()
        citation = next(run for run in p.runs if run.text == "[1]")
        assert citation.bold is True
        assert TestBodyTextCitationSuperscript._get_vertAlign(citation) == "superscript"


# ---------------------------------------------------------------------------
# 源码粘贴行识别（含中文行尾注释）
# ---------------------------------------------------------------------------


M = "http://schemas.openxmlformats.org/officeDocument/2006/math"


class TestLooksLikeCode:
    @pytest.mark.parametrize(
        "text",
        [
            "Usart2_Init(115200);//串口2，ESP8266用",
            "    bh_data_send(BHReset);     //重置数据寄存器值在PowerOn模式下有效",
            "GPIO_InitStructure.GPIO_Mode=GPIO_Mode_IPD;//设置GPIO口为下拉输入模式",
            "Delay_us(15);//延时15微秒",
            "while(1) //循环等待",
            'printf("ok"); /* 注释 */',
            "#include <stdio.h>",
            "int main(void)",
            "t=DS18B20_DQ_IN;",
            "// 只有注释的一行",
            "import random",
            "from os import path",
            "def main():",
        ],
    )
    def test_detects_code_line(self, text):
        assert _looks_like_code(text) is True

    @pytest.mark.parametrize(
        "text",
        [
            "本文研究了单片机控制系统的设计方法。",
            "图2.1 系统总体框图",
            "",
            "该模块采用串口通信完成数据交换",
            "血压传感器的选择",
        ],
    )
    def test_rejects_natural_language(self, text):
        assert _looks_like_code(text) is False


class TestBodyTextParaSkips:
    def _node(self, text):
        doc = Document()
        p = doc.add_paragraph(text)
        node = BodyText(value={"category": "body_text"}, level=0, paragraph=p)
        node.load_config({})
        return node

    def test_code_line_skips_indent_and_alignment(self):
        node = self._node("Delay_us(15);//延时15微秒")
        assert {
            "first_line_indent",
            "alignment",
            "space_before",
            "space_after",
        } <= node._para_checks_to_skip()

    def test_short_label_line_skips_spacing(self):
        node = self._node("血压传感器的选择")
        assert node._is_short_label_line()
        assert {
            "first_line_indent",
            "alignment",
            "space_before",
            "space_after",
        } <= node._para_checks_to_skip()

    def test_normal_paragraph_skips_nothing(self):
        node = self._node("本文设计了一套基于单片机的控制系统，实现了数据采集与显示功能。")
        assert node._para_checks_to_skip() == set()

    def test_formula_explanation_skips_indent(self):
        doc = Document()
        p = doc.add_paragraph("式中，u为输入电压。")
        p._p.append(
            parse_xml(f'<m:oMath xmlns:m="{M}"><m:r><m:t>u</m:t></m:r></m:oMath>')
        )
        node = BodyText(value={"category": "body_text"}, level=0, paragraph=p)
        node.load_config({})
        assert node._is_formula_explanation()
        assert "first_line_indent" in node._para_checks_to_skip()
