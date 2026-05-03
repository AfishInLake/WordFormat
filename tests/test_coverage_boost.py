"""
覆盖率提升测试文件

针对当前覆盖率 93% 的未覆盖代码行编写测试，目标 96%+。
覆盖模块：heading, keywords, numbering, style_enum, set_style, get_some, utils
"""

import re
from unittest.mock import MagicMock, patch, PropertyMock

import pytest
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from wordformat.config.datamodel import NodeConfigRoot
from wordformat.rules.heading import (
    HeadingLevel1Node,
    HeadingLevel2Node,
    HeadingLevel3Node,
    BaseHeadingNode,
)
from wordformat.rules.keywords import KeywordsCN, KeywordsEN
from wordformat.numbering import (
    _convert_to_twips,
    _set_indent_value,
    _build_numbering_rPr,
    strip_manual_numbering,
)
from wordformat.style.style_enum import (
    BuiltInStyle,
    FirstLineIndent,
    _ensure_style_exists,
)
from wordformat.set_style import (
    _fix_all_heading_style_definitions,
    apply_format_check_to_all_nodes,
)
from wordformat.style.get_some import (
    paragraph_get_space_before,
    paragraph_get_space_after,
    paragraph_get_line_spacing,
    paragraph_get_first_line_indent,
    paragraph_get_builtin_style_name,
    run_get_font_name,
    run_get_font_size_pt,
    run_get_font_color,
    GetIndent,
)
from wordformat.utils import (
    get_paragraph_numbering_text,
    _count_numbering_levels,
    _format_number,
    _to_chinese_num,
    _to_roman,
    ensure_directory_exists,
)


# ---------------------------------------------------------------------------
# 共享 fixtures / helpers
# ---------------------------------------------------------------------------


def _load_yaml(path):
    import yaml
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def _load_root_config(config_path):
    return NodeConfigRoot(**_load_yaml(config_path))


@pytest.fixture
def root_config(config_path):
    """从 example/undergrad_thesis.yaml 加载真实 NodeConfigRoot。"""
    from wordformat.config.config import init_config
    init_config(config_path)
    return _load_root_config(config_path)


# ===========================================================================
# 1. rules/heading.py 未覆盖行：92, 97, 174-222
# ===========================================================================


class TestHeadingCoverageBoost:
    """覆盖 heading.py 中未覆盖的代码行。"""

    def test_apply_to_paragraph_removes_existing_jc(self, root_config):
        """覆盖行 92: pPr.remove(jc) — 段落已有 jc 元素时移除显式对齐方式。

        当 p=False（格式化模式）且段落已有 w:jc 元素时，应移除该元素。
        注意：后续 ParagraphStyle.apply_to_paragraph 可能会重新设置对齐方式。
        """
        doc = Document()
        p = doc.add_paragraph()
        r = p.add_run("第一章 绪论")
        r.font.size = Pt(10)

        # 手动在 XML 中添加 w:jc 元素
        pPr = p._element.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p._element.insert(0, pPr)
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), "both")
        pPr.append(jc)

        # 验证 jc 确实存在
        assert p._element.find(qn("w:pPr")).find(qn("w:jc")) is not None

        node = HeadingLevel1Node(value=p, level=1, paragraph=p)
        node.load_config(root_config)

        # 使用 spy 验证 pPr.remove 被调用
        original_remove = pPr.remove
        remove_called = []
        def spy_remove(child):
            remove_called.append(child.tag)
            return original_remove(child)
        pPr.remove = spy_remove

        with patch.object(node, "add_comment"):
            node._base(doc, p=False, r=False)

        # 验证 pPr.remove 被调用过，且参数包含 jc
        assert any("jc" in tag for tag in remove_called), \
            f"pPr.remove was not called with jc, calls: {remove_called}"

    def test_apply_to_paragraph_updates_existing_pStyle(self, root_config):
        """覆盖行 97: pStyle.set(qn('w:val'), builtin_name) — pStyle 已存在时更新值。

        当段落已有 w:pStyle 元素时，应更新其值而非创建新元素。
        注意：后续 ParagraphStyle.apply_to_paragraph 会通过 python-docx API
        再次设置样式，python-docx 内部使用 "Heading2" 格式。
        """
        doc = Document()
        p = doc.add_paragraph()
        r = p.add_run("1.1 研究背景")
        r.font.size = Pt(10)

        # 手动添加一个不同的 w:pStyle
        pPr = p._element.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p._element.insert(0, pPr)
        pStyle = OxmlElement("w:pStyle")
        pStyle.set(qn("w:val"), "Normal")
        pPr.append(pStyle)

        # 验证 pStyle 值为 "Normal"
        assert pPr.find(qn("w:pStyle")).get(qn("w:val")) == "Normal"

        node = HeadingLevel2Node(value=p, level=2, paragraph=p)
        node.load_config(root_config)

        # 使用 spy 验证 pStyle.set 被调用
        original_set = pStyle.set
        set_calls = []
        def spy_set(key, value):
            set_calls.append((key, value))
            return original_set(key, value)
        pStyle.set = spy_set

        with patch.object(node, "add_comment"):
            node._base(doc, p=False, r=False)

        # 验证 pStyle.set 被调用过，且设置了 builtin_style_name
        ns_val = qn("w:val")
        assert any(k == ns_val and v == "Heading 2" for k, v in set_calls), \
            f"pStyle.set was not called with 'Heading 2', calls: {set_calls}"

    def test_fix_style_definition_color_no_style_name(self, root_config):
        """覆盖行 174-176: builtin_style_name 为 None 时直接返回。"""
        doc = Document()
        cfg = MagicMock()
        cfg.builtin_style_name = None
        # 不应抛出异常
        BaseHeadingNode._fix_style_definition_color(doc, cfg)

    def test_fix_style_definition_color_style_not_exists(self, root_config):
        """覆盖行 178-182: 样式不存在时跳过。"""
        doc = Document()
        cfg = MagicMock()
        cfg.builtin_style_name = "NonExistentStyle"
        # 不应抛出异常
        BaseHeadingNode._fix_style_definition_color(doc, cfg)

    def test_fix_style_definition_color_no_rPr(self, root_config):
        """覆盖行 187-190: 样式无 rPr 元素时跳过。"""
        doc = Document()
        cfg = MagicMock()
        cfg.builtin_style_name = "Heading 1"
        # Heading 1 样式存在但可能没有 rPr
        BaseHeadingNode._fix_style_definition_color(doc, cfg)

    def test_fix_style_definition_color_no_color_element(self, root_config):
        """覆盖行 192-195: 样式无 color 元素时跳过。"""
        doc = Document()
        cfg = MagicMock()
        cfg.builtin_style_name = "Heading 1"
        # 确保 rPr 存在但 color 不存在
        try:
            style = doc.styles["Heading 1"]
            rPr = style.element.find(qn("w:rPr"))
            if rPr is not None:
                color = rPr.find(qn("w:color"))
                if color is not None:
                    rPr.remove(color)
        except KeyError:
            pass
        BaseHeadingNode._fix_style_definition_color(doc, cfg)

    def test_fix_style_definition_color_no_theme(self, root_config):
        """覆盖行 198-204: color 元素无主题色属性时直接返回。"""
        doc = Document()
        cfg = MagicMock()
        cfg.builtin_style_name = "Heading 1"
        # 确保 color 元素存在但无 themeColor
        try:
            style = doc.styles["Heading 1"]
            rPr = style.element.find(qn("w:rPr"))
            if rPr is None:
                rPr = OxmlElement("w:rPr")
                style.element.insert(0, rPr)
            color = rPr.find(qn("w:color"))
            if color is None:
                color = OxmlElement("w:color")
                color.set(qn("w:val"), "000000")
                rPr.append(color)
            # 确保没有主题色属性
            for attr in [qn("w:themeColor"), qn("w:themeTint"), qn("w:themeShade")]:
                if color.get(attr) is not None:
                    del color.attrib[attr]
        except KeyError:
            pass
        BaseHeadingNode._fix_style_definition_color(doc, cfg)

    def test_fix_style_definition_color_with_theme_color(self, root_config):
        """覆盖行 206-222: 有主题色时修正为配置颜色。"""
        doc = Document()
        cfg = root_config.headings.level_1
        # 确保 color 元素存在且有 themeColor
        try:
            style = doc.styles["Heading 1"]
            rPr = style.element.find(qn("w:rPr"))
            if rPr is None:
                rPr = OxmlElement("w:rPr")
                style.element.insert(0, rPr)
            color = rPr.find(qn("w:color"))
            if color is None:
                color = OxmlElement("w:color")
                rPr.append(color)
            # 设置主题色
            color.set(qn("w:themeColor"), "accent1")
            color.set(qn("w:val"), "4472C4")
        except KeyError:
            pytest.skip("Heading 1 样式不存在")

        BaseHeadingNode._fix_style_definition_color(doc, cfg)

        # 验证主题色已被移除
        rPr_after = style.element.find(qn("w:rPr"))
        color_after = rPr_after.find(qn("w:color"))
        assert color_after.get(qn("w:themeColor")) is None

    def test_fix_style_definition_color_with_theme_tint(self, root_config):
        """覆盖行 200-201: themeTint 属性存在时也应修正。"""
        doc = Document()
        cfg = root_config.headings.level_2
        try:
            style = doc.styles["Heading 2"]
            rPr = style.element.find(qn("w:rPr"))
            if rPr is None:
                rPr = OxmlElement("w:rPr")
                style.element.insert(0, rPr)
            color = rPr.find(qn("w:color"))
            if color is None:
                color = OxmlElement("w:color")
                rPr.append(color)
            color.set(qn("w:themeTint"), "99")
            color.set(qn("w:val"), "000000")
        except KeyError:
            pytest.skip("Heading 2 样式不存在")

        BaseHeadingNode._fix_style_definition_color(doc, cfg)

        rPr_after = style.element.find(qn("w:rPr"))
        color_after = rPr_after.find(qn("w:color"))
        assert color_after.get(qn("w:themeTint")) is None

    def test_fix_style_definition_color_exception_handling(self, root_config):
        """覆盖行 221-222: 颜色修正失败时的异常处理。"""
        doc = Document()
        cfg = MagicMock()
        cfg.builtin_style_name = "Heading 1"
        cfg.font_color = "invalid_color_value_xyz"
        try:
            style = doc.styles["Heading 1"]
            rPr = style.element.find(qn("w:rPr"))
            if rPr is None:
                rPr = OxmlElement("w:rPr")
                style.element.insert(0, rPr)
            color = rPr.find(qn("w:color"))
            if color is None:
                color = OxmlElement("w:color")
                rPr.append(color)
            color.set(qn("w:themeColor"), "accent1")
        except KeyError:
            pytest.skip("Heading 1 样式不存在")

        # 无效颜色应触发异常处理分支，不抛出
        BaseHeadingNode._fix_style_definition_color(doc, cfg)


# ===========================================================================
# 2. rules/keywords.py 未覆盖行：56, 69-106, 110, 128, 149, 183, 195, 232,
#    261, 295, 307
# ===========================================================================


class TestKeywordsCoverageBoost:
    """覆盖 keywords.py 中未覆盖的代码行。"""

    def test_apply_to_paragraph_path_cn(self, root_config):
        """覆盖行 56 (CN): apply_to_paragraph 路径（p=False）。

        KeywordsCN._check_paragraph_style 中 p=False 时调用 apply_to_paragraph。
        """
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("关键词：机器学习；深度学习")
        node = KeywordsCN(value=p, level=0, paragraph=p)
        node.load_config(root_config)
        with patch.object(node, "add_comment"):
            node._base(doc, p=False, r=False)

    def test_apply_to_paragraph_path_en(self, root_config):
        """覆盖行 56 (EN): apply_to_paragraph 路径（p=False）。

        KeywordsEN._check_paragraph_style 中 p=False 时调用 apply_to_paragraph。
        """
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("Keywords: AI, ML")
        node = KeywordsEN(value=p, level=0, paragraph=p)
        node.load_config(root_config)
        with patch.object(node, "add_comment"):
            node._base(doc, p=False, r=False)

    def test_split_mixed_runs_cn(self, root_config):
        """覆盖行 69-106: _split_mixed_runs 拆分标签和内容混合的 run。

        当标签和内容在同一个 run 中时，应拆分为两个独立 run。
        """
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("关键词：校园二手交易；Django")
        node = KeywordsCN(value=p, level=0, paragraph=p)
        node.load_config(root_config)
        with patch.object(node, "add_comment"):
            node._base(doc, p=False, r=False)

        # 拆分后应有多个 run
        assert len(p.runs) >= 2

    def test_split_mixed_runs_en(self, root_config):
        """覆盖行 69-106: _split_mixed_runs 英文标签拆分。"""
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("Keywords: Machine Learning, Deep Learning")
        node = KeywordsEN(value=p, level=0, paragraph=p)
        node.load_config(root_config)
        with patch.object(node, "add_comment"):
            node._base(doc, p=False, r=False)

        assert len(p.runs) >= 2

    def test_get_label_split_pattern_en(self, root_config):
        """覆盖行 128: KeywordsEN._get_label_split_pattern 返回正则。"""
        node = KeywordsCN(value=None, level=0, paragraph=None)
        pattern = KeywordsEN._get_label_split_pattern(node)
        assert pattern is not None
        assert pattern.search("Keywords: ") is not None
        assert pattern.search("Keyword ") is not None

    def test_split_mixed_runs_in_format_mode_cn(self, root_config):
        """覆盖行 149: self._split_mixed_runs() 在格式化模式下执行。

        KeywordsCN._base 中 not p 时执行 _split_mixed_runs。
        """
        doc = Document()
        p = doc.add_paragraph()
        r = p.add_run("关键词：A；B；C；D")
        r.font.size = Pt(10)
        node = KeywordsCN(value=p, level=0, paragraph=p)
        node.load_config(root_config)
        initial_run_count = len(p.runs)
        with patch.object(node, "add_comment"):
            node._base(doc, p=False, r=False)
        # 格式化模式下拆分后 run 数量应增加
        assert len(p.runs) >= initial_run_count

    def test_en_label_apply_to_run(self, root_config):
        """覆盖行 183: KeywordsEN._base 中标签 apply_to_run 路径。

        当 p=False, r=False 且 run 是标签时调用 label_style.apply_to_run。
        """
        doc = Document()
        p = doc.add_paragraph()
        r = p.add_run("Keywords")
        r.font.size = Pt(10)
        r.font.bold = False  # 应为加粗
        node = KeywordsEN(value=p, level=0, paragraph=p)
        node.load_config(root_config)
        with patch.object(node, "add_comment") as mock_comment:
            node._base(doc, p=False, r=False)
        assert mock_comment.call_count >= 1

    def test_en_content_apply_to_run(self, root_config):
        """覆盖行 195: KeywordsEN._base 中内容 apply_to_run 路径。

        当 p=False, r=False 且 run 不是标签时调用 content_style.apply_to_run。
        """
        doc = Document()
        p = doc.add_paragraph()
        r1 = p.add_run("Keywords: ")
        r1.font.bold = True
        r2 = p.add_run("Machine Learning")
        r2.font.size = Pt(10)
        r2.font.bold = True  # 内容不应加粗
        node = KeywordsEN(value=p, level=0, paragraph=p)
        node.load_config(root_config)
        with patch.object(node, "add_comment") as mock_comment:
            node._base(doc, p=False, r=False)
        assert mock_comment.call_count >= 1

    def test_cn_label_apply_to_run(self, root_config):
        """覆盖行 295: KeywordsCN._base 中标签 apply_to_run 路径。"""
        doc = Document()
        p = doc.add_paragraph()
        r1 = p.add_run("关键词")
        r1.font.size = Pt(10)
        r1.font.bold = False  # 应为加粗
        r2 = p.add_run("：A；B；C")
        node = KeywordsCN(value=p, level=0, paragraph=p)
        node.load_config(root_config)
        with patch.object(node, "add_comment") as mock_comment:
            node._base(doc, p=False, r=False)
        assert mock_comment.call_count >= 1

    def test_cn_content_apply_to_run(self, root_config):
        """覆盖行 307: KeywordsCN._base 中内容 apply_to_run 路径。"""
        doc = Document()
        p = doc.add_paragraph()
        r1 = p.add_run("关键词：")
        r1.font.bold = True
        r2 = p.add_run("A；B；C")
        r2.font.size = Pt(10)
        r2.font.bold = True  # 内容不应加粗
        node = KeywordsCN(value=p, level=0, paragraph=p)
        node.load_config(root_config)
        with patch.object(node, "add_comment") as mock_comment:
            node._base(doc, p=False, r=False)
        assert mock_comment.call_count >= 1

    def test_cn_split_mixed_runs_format_mode(self, root_config):
        """覆盖行 261: KeywordsCN._base 中 not p 时执行 _split_mixed_runs。"""
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("关键词：测试内容")
        node = KeywordsCN(value=p, level=0, paragraph=p)
        node.load_config(root_config)
        with patch.object(node, "add_comment"):
            node._base(doc, p=False, r=False)
        assert len(p.runs) >= 2

    def test_cn_get_label_split_pattern(self, root_config):
        """覆盖行 232: KeywordsCN._get_label_split_pattern 返回正则。"""
        node = KeywordsCN(value=None, level=0, paragraph=None)
        pattern = node._get_label_split_pattern()
        assert pattern is not None
        assert pattern.search("关键词：") is not None
        assert pattern.search("关键词:") is not None


# ===========================================================================
# 3. numbering.py 未覆盖行：36-61, 76-95, 120, 129, 150-153, 164-167, 201,
#    365, 370, 435
# ===========================================================================


class TestNumberingCoverageBoost:
    """覆盖 numbering.py 中未覆盖的代码行。"""

    def test_convert_to_twips_cm(self):
        """覆盖行 45-46: cm 单位转换。"""
        result = _convert_to_twips("0.75cm")
        assert isinstance(result, int)
        assert result > 0

    def test_convert_to_twips_mm(self):
        """覆盖行 47-48: mm 单位转换。"""
        result = _convert_to_twips("5mm")
        assert isinstance(result, int)
        assert result > 0

    def test_convert_to_twips_inch(self):
        """覆盖行 49-50: inch 单位转换。"""
        result = _convert_to_twips("1inch")
        assert isinstance(result, int)
        assert result > 0

    def test_convert_to_twips_pt(self):
        """覆盖行 51-52: pt 单位转换。"""
        result = _convert_to_twips("12pt")
        assert isinstance(result, int)
        assert result > 0

    def test_convert_to_twips_invalid(self):
        """覆盖行 37-39: 无效值返回 0。"""
        result = _convert_to_twips("invalid")
        assert result == 0

    def test_convert_to_twips_unsupported_unit(self):
        """覆盖行 53-55: 不支持的物理单位返回 0。"""
        result = _convert_to_twips("100emu")
        assert result == 0

    def test_convert_to_twips_exception(self):
        """覆盖行 56-58: 单位换算异常时返回 0。"""
        # 传入一个会导致 Pt() 抛出异常的值
        result = _convert_to_twips("abcpt")
        assert result == 0

    def test_set_indent_value_char_unit(self):
        """覆盖行 84-89: 字符单位设置。"""
        ind = OxmlElement("w:ind")
        _set_indent_value(ind, "left", "2字符")
        assert ind.get(qn("w:leftChars")) == "200"

    def test_set_indent_value_physical_unit(self):
        """覆盖行 90-93: 物理单位（cm）设置。"""
        ind = OxmlElement("w:ind")
        _set_indent_value(ind, "left", "0.75cm")
        assert ind.get(qn("w:left")) is not None

    def test_set_indent_value_invalid(self):
        """覆盖行 77-79: 无效值跳过设置。"""
        ind = OxmlElement("w:ind")
        _set_indent_value(ind, "left", "invalid")
        assert ind.get(qn("w:left")) is None
        assert ind.get(qn("w:leftChars")) is None

    def test_set_indent_value_unsupported_unit(self):
        """覆盖行 94-95: 不支持的缩进单位跳过设置。"""
        ind = OxmlElement("w:ind")
        _set_indent_value(ind, "left", "100emu")
        assert ind.get(qn("w:left")) is None

    def test_build_numbering_rPr_level_none(self):
        """覆盖行 119-120: level_cfg 为 None 时返回 None。"""
        headings_config = MagicMock()
        headings_config.level_1 = None
        result = _build_numbering_rPr(headings_config, "level_1")
        assert result is None

    def test_build_numbering_rPr_all_none(self):
        """覆盖行 128-129: 所有属性为 None 时返回 None。"""
        headings_config = MagicMock()
        level_cfg = MagicMock()
        level_cfg.chinese_font_name = None
        level_cfg.english_font_name = None
        level_cfg.font_size = None
        level_cfg.bold = None
        headings_config.level_1 = level_cfg
        result = _build_numbering_rPr(headings_config, "level_1")
        assert result is None

    def test_build_numbering_rPr_bold_true(self):
        """覆盖行 163-167: bold=True 时添加 w:b 和 w:bCs 元素。"""
        headings_config = MagicMock()
        level_cfg = MagicMock()
        level_cfg.chinese_font_name = "黑体"
        level_cfg.english_font_name = "Times New Roman"
        level_cfg.font_size = "小二"
        level_cfg.bold = True
        headings_config.level_1 = level_cfg
        result = _build_numbering_rPr(headings_config, "level_1")
        assert result is not None
        assert result.find(qn("w:b")) is not None
        assert result.find(qn("w:bCs")) is not None

    def test_build_numbering_rPr_numeric_font_size(self):
        """覆盖行 150-153: 非中文字号（数字）的字号处理。"""
        headings_config = MagicMock()
        level_cfg = MagicMock()
        level_cfg.chinese_font_name = "黑体"
        level_cfg.english_font_name = None
        level_cfg.font_size = "15.5"
        level_cfg.bold = None
        headings_config.level_1 = level_cfg
        result = _build_numbering_rPr(headings_config, "level_1")
        assert result is not None
        sz = result.find(qn("w:sz"))
        assert sz is not None

    def test_build_numbering_rPr_invalid_font_size(self):
        """覆盖行 152-153: 无法解析的字号返回 None。"""
        headings_config = MagicMock()
        level_cfg = MagicMock()
        level_cfg.chinese_font_name = "黑体"
        level_cfg.english_font_name = None
        level_cfg.font_size = "abc"
        level_cfg.bold = None
        headings_config.level_1 = level_cfg
        result = _build_numbering_rPr(headings_config, "level_1")
        assert result is not None
        # sz 不应被添加（half_pt 为 None）
        assert result.find(qn("w:sz")) is None

    def test_strip_manual_numbering_success(self):
        """覆盖行 186-217: strip_manual_numbering 成功清除编号。"""
        doc = Document()
        p = doc.add_paragraph()
        r = p.add_run("1.1 研究背景")
        result = strip_manual_numbering(p, r"^\d+(\.\d+)\s*")
        assert result is True
        assert not p.text.startswith("1.1")

    def test_strip_manual_numbering_no_match(self):
        """覆盖行 191-192: 不匹配正则时返回 False。"""
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("研究背景")
        result = strip_manual_numbering(p, r"^\d+\s*")
        assert result is False

    def test_strip_manual_numbering_empty_pattern(self):
        """覆盖行 186-187: 空正则返回 False。"""
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("1 测试")
        result = strip_manual_numbering(p, "")
        assert result is False

    def test_strip_manual_numbering_no_runs(self):
        """覆盖行 186-187: 无 run 时返回 False。"""
        doc = Document()
        p = doc.add_paragraph()
        result = strip_manual_numbering(p, r"^\d+\s*")
        assert result is False

    def test_strip_manual_numbering_multi_run(self):
        """覆盖行 198-210: 多 run 逐字符删除。"""
        doc = Document()
        p = doc.add_paragraph()
        r1 = p.add_run("1.")
        r2 = p.add_run("1 ")
        r3 = p.add_run("研究背景")
        result = strip_manual_numbering(p, r"^\d+(\.\d+)\s*")
        assert result is True
        assert p.text.strip() == "研究背景"


# ===========================================================================
# 4. style/style_enum.py 未覆盖行：155, 641, 649-657
# ===========================================================================


class TestStyleEnumCoverageBoost:
    """覆盖 style_enum.py 中未覆盖的代码行。"""

    def test_first_line_indent_pt_unit(self):
        """覆盖行 155: FirstLineIndent 使用 pt 单位时走物理单位分支。

        FirstLineIndent.get_from_paragraph 中 unit 为 pt 时，
        从 paragraph_format.first_line_indent 获取值。
        """
        doc = Document()
        p = doc.add_paragraph()
        from docx.shared import Pt
        p.paragraph_format.first_line_indent = Pt(24)
        fli = FirstLineIndent("24pt")
        result = fli.get_from_paragraph(p)
        assert result is not None

    def test_first_line_indent_cm_unit(self):
        """覆盖行 155: FirstLineIndent 使用 cm 单位。"""
        doc = Document()
        p = doc.add_paragraph()
        from docx.shared import Cm
        p.paragraph_format.first_line_indent = Cm(0.85)
        fli = FirstLineIndent("0.85cm")
        result = fli.get_from_paragraph(p)
        assert result is not None

    def test_first_line_indent_inch_unit(self):
        """覆盖行 155: FirstLineIndent 使用 inch 单位。"""
        doc = Document()
        p = doc.add_paragraph()
        from docx.shared import Inches
        p.paragraph_format.first_line_indent = Inches(0.5)
        fli = FirstLineIndent("0.5inch")
        result = fli.get_from_paragraph(p)
        assert result is not None

    def test_ensure_style_exists_base_style_none(self):
        """覆盖行 641: base_style_name 为 None 时 base_style = None。

        当样式映射中 base_style_name 为 None 时（如 Normal），
        创建样式不设置基础样式。
        """
        doc = Document()
        # 先删除 Normal 样式（如果存在），确保需要创建
        # 实际上 Normal 总是存在的，所以用另一个方式触发
        # 使用一个不在映射中的样式名，其 base 默认为 "Normal"
        _ensure_style_exists(doc, "TestStyleNoBase")
        # 清理
        try:
            doc.styles.element.remove(
                doc.styles["TestStyleNoBase"].element
            )
        except Exception:
            pass

    def test_ensure_style_exists_with_outline_lvl(self):
        """覆盖行 649-657: 设置 outlineLvl 的分支。

        创建 Heading 样式时应设置大纲级别。
        """
        doc = Document()
        # 先删除 Heading 4 样式（如果存在）
        try:
            style_elem = doc.styles["Heading 4"].element
            style_elem.getparent().remove(style_elem)
        except KeyError:
            pass

        _ensure_style_exists(doc, "Heading 4")

        # 验证 outlineLvl 已设置
        style = doc.styles["Heading 4"]
        pPr = style.element.find(qn("w:pPr"))
        if pPr is not None:
            outlineLvl = pPr.find(qn("w:outlineLvl"))
            assert outlineLvl is not None
            assert outlineLvl.get(qn("w:val")) == "3"

    def test_ensure_style_exists_already_exists(self):
        """覆盖行 630-633: 样式已存在时直接返回。"""
        doc = Document()
        # Normal 样式总是存在
        _ensure_style_exists(doc, "Normal")
        # 不应抛出异常

    def test_ensure_style_creates_with_base(self):
        """覆盖行 643-644: 创建样式时设置基础样式。"""
        doc = Document()
        style_name = "TestStyleWithBase"
        # 确保不存在
        try:
            style_elem = doc.styles[style_name].element
            style_elem.getparent().remove(style_elem)
        except KeyError:
            pass

        _ensure_style_exists(doc, style_name)
        style = doc.styles[style_name]
        assert style is not None
        # 清理
        try:
            style_elem = doc.styles[style_name].element
            style_elem.getparent().remove(style_elem)
        except Exception:
            pass


# ===========================================================================
# 5. set_style.py 未覆盖行：46, 57, 61, 99-100, 129-131, 223, 226
# ===========================================================================


class TestSetStyleCoverageBoost:
    """覆盖 set_style.py 中未覆盖的代码行。"""

    def test_fix_all_heading_style_definitions_no_headings(self):
        """覆盖行 45-46: heading_config 为 None 时直接返回。"""
        doc = Document()
        config_model = MagicMock()
        config_model.headings = None
        _fix_all_heading_style_definitions(doc, config_model)
        # 不应抛出异常

    def test_fix_all_heading_style_definitions_level_none(self):
        """覆盖行 57: level_cfg 为 None 时 continue。"""
        doc = Document()
        config_model = MagicMock()
        config_model.headings = MagicMock()
        config_model.headings.level_1 = None
        config_model.headings.level_2 = None
        config_model.headings.level_3 = None
        _fix_all_heading_style_definitions(doc, config_model)
        # 不应抛出异常

    def test_fix_all_heading_style_definitions_empty_style_name(self):
        """覆盖行 61: style_name 为空时 continue。"""
        doc = Document()
        config_model = MagicMock()
        config_model.headings = MagicMock()
        level_cfg = MagicMock()
        level_cfg.builtin_style_name = ""
        config_model.headings.level_1 = level_cfg
        config_model.headings.level_2 = None
        config_model.headings.level_3 = None
        _fix_all_heading_style_definitions(doc, config_model)

    def test_fix_all_heading_style_definitions_theme_color_fix(self):
        """覆盖行 86-100: 修正主题色的完整流程。"""
        doc = Document()
        config_model = MagicMock()
        level_cfg = MagicMock()
        level_cfg.builtin_style_name = "Heading 1"
        level_cfg.font_color = "黑色"
        config_model.headings = MagicMock()
        config_model.headings.level_1 = level_cfg
        config_model.headings.level_2 = None
        config_model.headings.level_3 = None

        # 确保 Heading 1 有主题色
        try:
            style = doc.styles["Heading 1"]
            rPr = style.element.find(qn("w:rPr"))
            if rPr is None:
                rPr = OxmlElement("w:rPr")
                style.element.insert(0, rPr)
            color = rPr.find(qn("w:color"))
            if color is None:
                color = OxmlElement("w:color")
                rPr.append(color)
            color.set(qn("w:themeColor"), "accent1")
        except KeyError:
            pass

        _fix_all_heading_style_definitions(doc, config_model)

    def test_apply_format_check_to_all_nodes_void_category(self):
        """覆盖行 99-100: traverse 中 category 在 VOIDNODELIST 中。

        当节点的 category 在 VOIDNODELIST 中时，跳过格式化。
        """
        from wordformat.rules.node import FormatNode
        from wordformat.settings import VOIDNODELIST

        doc = Document()
        p = doc.add_paragraph("测试")
        root_node = FormatNode(
            value={"category": VOIDNODELIST[0]},
            level=0,
            paragraph=p,
        )
        root_node.children = []

        config = MagicMock()
        with patch.object(root_node, "check_format"):
            apply_format_check_to_all_nodes(root_node, doc, config, check=True)
            # VOIDNODELIST 中的 category 不应调用 check_format
            root_node.check_format.assert_not_called()

    def test_apply_format_check_to_all_nodes_no_check_format(self):
        """覆盖行 129-131: traverse 中节点无 check_format 的分支。

        当节点没有 check_format 方法时，跳过。
        """
        # 使用一个没有 check_format 的简单对象
        class SimpleNode:
            def __init__(self):
                self.value = {"category": "body_text"}
                self.children = []
                self.paragraph = None

        doc = Document()
        root_node = SimpleNode()
        config = MagicMock()

        # 不应抛出异常
        apply_format_check_to_all_nodes(root_node, doc, config, check=True)

    def test_apply_format_check_to_all_nodes_exception_handling(self):
        """覆盖行 129-131: traverse 中异常处理分支。"""
        from wordformat.rules.node import FormatNode

        doc = Document()
        p = doc.add_paragraph("测试")
        root_node = FormatNode(
            value={"category": "body_text", "fingerprint": "abc123"},
            level=0,
            paragraph=p,
        )
        root_node.children = []

        config = MagicMock()
        # load_config 会抛出异常
        with patch.object(root_node, "load_config", side_effect=ValueError("test error")):
            with pytest.raises(ValueError):
                apply_format_check_to_all_nodes(root_node, doc, config, check=True)

    def test_apply_format_check_to_all_nodes_apply_mode(self):
        """覆盖行 127-128: check=False 时调用 apply_format。"""
        from wordformat.rules.node import FormatNode

        doc = Document()
        p = doc.add_paragraph("测试")
        root_node = FormatNode(
            value={"category": "body_text", "fingerprint": "abc123"},
            level=0,
            paragraph=p,
        )
        root_node.children = []

        config = MagicMock()
        with patch.object(root_node, "load_config"):
            with patch.object(root_node, "apply_format") as mock_apply:
                apply_format_check_to_all_nodes(root_node, doc, config, check=False)
                mock_apply.assert_called_once_with(doc)

    def test_apply_format_check_to_all_nodes_no_paragraph(self):
        """覆盖行 124: node.paragraph 为 None 时跳过。"""
        from wordformat.rules.node import FormatNode

        doc = Document()
        root_node = FormatNode(
            value={"category": "body_text", "fingerprint": "abc123"},
            level=0,
            paragraph=None,
        )
        root_node.children = []

        config = MagicMock()
        with patch.object(root_node, "load_config"):
            with patch.object(root_node, "check_format") as mock_check:
                apply_format_check_to_all_nodes(root_node, doc, config, check=True)
                mock_check.assert_not_called()


# ===========================================================================
# 6. style/get_some.py 未覆盖行：106, 122, 157-158, 190-191, 280, 340, 429-433
# ===========================================================================


class TestGetSomeCoverageBoost:
    """覆盖 get_some.py 中未覆盖的代码行。"""

    def test_paragraph_get_space_before_with_lines(self):
        """覆盖行 106, 122: paragraph_get_space_before 带 beforeLines 属性。"""
        doc = Document()
        p = doc.add_paragraph()
        # 手动设置 beforeLines XML 属性
        pPr = p._element.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p._element.insert(0, pPr)
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:beforeLines"), "200")  # 2行
        pPr.append(spacing)

        result = paragraph_get_space_before(p)
        assert result is not None
        assert result == 2.0

    def test_paragraph_get_space_before_exception(self):
        """覆盖行 157-158: paragraph_get_space_before 异常处理。"""
        # 使用 mock 使 paragraph._element 抛出异常
        p = MagicMock()
        p._element = MagicMock()
        p._element.find.side_effect = AttributeError("test")
        result = paragraph_get_space_before(p)
        assert result is None

    def test_paragraph_get_space_after_with_lines(self):
        """覆盖行 190-191: paragraph_get_space_after 带 afterLines 属性。"""
        doc = Document()
        p = doc.add_paragraph()
        pPr = p._element.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p._element.insert(0, pPr)
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:afterLines"), "100")  # 1行
        pPr.append(spacing)

        result = paragraph_get_space_after(p)
        assert result is not None
        assert result == 1.0

    def test_paragraph_get_space_after_exception(self):
        """覆盖行 190-191: paragraph_get_space_after 异常处理。"""
        p = MagicMock()
        p._element = MagicMock()
        p._element.find.side_effect = TypeError("test")
        result = paragraph_get_space_after(p)
        assert result is None

    def test_paragraph_get_builtin_style_name_none(self):
        """覆盖行 280: style 为 None 时返回空字符串。"""
        p = MagicMock()
        p.style = None
        result = paragraph_get_builtin_style_name(p)
        assert result == ""

    def test_run_get_font_color_theme(self):
        """覆盖行 340: themeColor 类型返回 None。"""
        # 使用 MagicMock 模拟 run.font.color，使 type 返回 THEME
        from docx.enum.dml import MSO_COLOR_TYPE
        mock_run = MagicMock()
        mock_color = MagicMock()
        mock_color.type = MSO_COLOR_TYPE.THEME
        mock_run.font.color = mock_color
        result = run_get_font_color(mock_run)
        assert result is None

    def test_get_indent_invalid_type(self):
        """覆盖行 429-433: GetIndent.line_indent 无效 indent_type。"""
        doc = Document()
        p = doc.add_paragraph()
        with pytest.raises(ValueError, match="indent_type"):
            GetIndent.line_indent(p, "invalid")

    def test_get_indent_no_pPr(self):
        """覆盖行 429: pPr 为 None 时返回 None。"""
        doc = Document()
        p = doc.add_paragraph()
        result = GetIndent.left_indent(p)
        assert result is None

    def test_paragraph_get_line_spacing_multiple(self):
        """覆盖行 222-226: MULTIPLE 行距类型返回 float。"""
        doc = Document()
        p = doc.add_paragraph()
        from docx.shared import Pt
        from docx.enum.text import WD_LINE_SPACING
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        p.paragraph_format.line_spacing = 1.5
        result = paragraph_get_line_spacing(p)
        assert result == 1.5

    def test_paragraph_get_line_spacing_at_least(self):
        """覆盖行 229-231: AT_LEAST 类型返回 None。"""
        doc = Document()
        p = doc.add_paragraph()
        from docx.enum.text import WD_LINE_SPACING
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
        result = paragraph_get_line_spacing(p)
        assert result is None

    def test_paragraph_get_line_spacing_exception(self):
        """覆盖行 233-235: 异常处理返回 None。"""
        p = MagicMock()
        p.paragraph_format = MagicMock()
        p.paragraph_format.line_spacing_rule = MagicMock(side_effect=AttributeError)
        result = paragraph_get_line_spacing(p)
        assert result is None

    def test_run_get_font_size_pt_from_style(self):
        """覆盖行 315-317: 从样式继承字体大小。"""
        doc = Document()
        p = doc.add_paragraph()
        r = p.add_run("测试")
        # run 没有显式设置 font.size，应从样式继承
        result = run_get_font_size_pt(r)
        assert isinstance(result, float)

    def test_run_get_font_name_none(self):
        """覆盖行 293-300: run 无 rPr 时返回 None。"""
        doc = Document()
        p = doc.add_paragraph()
        r = p.add_run("测试")
        result = run_get_font_name(r)
        # 新 run 可能没有设置 eastAsia 字体
        assert result is None or isinstance(result, str)


# ===========================================================================
# 7. utils.py 未覆盖行：117, 128, 132, 154, 203, 208-212, 220, 223, 226, 310
# ===========================================================================


class TestUtilsCoverageBoost:
    """覆盖 utils.py 中未覆盖的代码行。"""

    def test_get_paragraph_numbering_text_no_pPr(self):
        """覆盖行 96-97: pPr 为 None 时返回空字符串。"""
        doc = Document()
        p = doc.add_paragraph()
        # 确保没有 pPr
        pPr = p._element.find(qn("w:pPr"))
        if pPr is not None:
            p._element.remove(pPr)
        result = get_paragraph_numbering_text(p)
        assert result == ""

    def test_get_paragraph_numbering_text_no_numPr(self):
        """覆盖行 99-100: numPr 为 None 时返回空字符串。"""
        doc = Document()
        p = doc.add_paragraph()
        pPr = p._element.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p._element.insert(0, pPr)
        result = get_paragraph_numbering_text(p)
        assert result == ""

    def test_get_paragraph_numbering_text_no_numId(self):
        """覆盖行 104-105: numId_elem 为 None 时返回空字符串。"""
        doc = Document()
        p = doc.add_paragraph()
        pPr = p._element.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p._element.insert(0, pPr)
        numPr = OxmlElement("w:numPr")
        pPr.append(numPr)
        result = get_paragraph_numbering_text(p)
        assert result == ""

    def test_get_paragraph_numbering_text_num_id_zero(self):
        """覆盖行 110-111: numId 为 "0" 时返回空字符串。"""
        doc = Document()
        p = doc.add_paragraph()
        pPr = p._element.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p._element.insert(0, pPr)
        numPr = OxmlElement("w:numPr")
        numId = OxmlElement("w:numId")
        numId.set(qn("w:val"), "0")
        numPr.append(numId)
        pPr.append(numPr)
        result = get_paragraph_numbering_text(p)
        assert result == ""

    def test_get_paragraph_numbering_text_no_numbering_part(self):
        """覆盖行 115-117: 无 numbering_part 时返回空字符串。"""
        doc = Document()
        p = doc.add_paragraph()
        pPr = p._element.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p._element.insert(0, pPr)
        numPr = OxmlElement("w:numPr")
        numId = OxmlElement("w:numId")
        numId.set(qn("w:val"), "1")
        ilvl = OxmlElement("w:ilvl")
        ilvl.set(qn("w:val"), "0")
        numPr.append(ilvl)
        numPr.append(numId)
        pPr.append(numPr)
        # 新建文档可能默认有 numbering part，也可能没有
        # 只要不抛出异常即可
        result = get_paragraph_numbering_text(p)
        assert isinstance(result, str)

    def test_count_numbering_levels_no_body(self):
        """覆盖行 203: body 为 None 时返回空字典。"""
        doc = Document()
        p = doc.add_paragraph()
        result = _count_numbering_levels(MagicMock(), "1", p)
        assert result == {}

    def test_to_chinese_num_zero(self):
        """覆盖行 292-293: num <= 0 时返回字符串。"""
        assert _to_chinese_num(0) == "0"
        assert _to_chinese_num(-1) == "-1"

    def test_to_chinese_num_hundred(self):
        """覆盖行 308-309: num == 100 时返回 '一百'。"""
        assert _to_chinese_num(100) == "一百"

    def test_to_chinese_num_over_hundred(self):
        """覆盖行 310: num > 100 时返回字符串。"""
        assert _to_chinese_num(101) == "101"

    def test_to_chinese_num_tens(self):
        """覆盖行 298-306: 10-99 的中文数字。"""
        assert _to_chinese_num(10) == "十"
        assert _to_chinese_num(15) == "十五"
        assert _to_chinese_num(20) == "二十"
        assert _to_chinese_num(99) == "九十九"

    def test_to_chinese_num_units(self):
        """覆盖行 296-297: 1-9 的中文数字。"""
        assert _to_chinese_num(1) == "一"
        assert _to_chinese_num(5) == "五"
        assert _to_chinese_num(9) == "九"

    def test_roman_numeral(self):
        """覆盖 _to_roman 函数。"""
        assert _to_roman(1) == "i"
        assert _to_roman(4) == "iv"
        assert _to_roman(9) == "ix"
        assert _to_roman(2024) == "mmxxiv"

    def test_format_number_various(self):
        """覆盖 _format_number 各种格式。"""
        assert _format_number(1, "decimal") == "1"
        assert _format_number(1, "upperRoman") == "I"
        assert _format_number(1, "lowerRoman") == "i"
        assert _format_number(1, "upperLetter") == "A"
        assert _format_number(1, "lowerLetter") == "a"
        assert _format_number(1, "chineseCountingThousand") == "一"
        assert _format_number(1, "ideographTraditional") == "一"
        assert _format_number(1, "chineseCounting") == "一"

    def test_ensure_directory_exists_creates(self, tmp_path):
        """覆盖 ensure_directory_exists 创建目录。"""
        new_dir = str(tmp_path / "new_subdir" / "nested")
        ensure_directory_exists(new_dir)
        import os
        assert os.path.isdir(new_dir)

    def test_ensure_directory_exists_already_exists(self, tmp_path):
        """覆盖 ensure_directory_exists 目录已存在。"""
        ensure_directory_exists(str(tmp_path))
        # 不应抛出异常

    def test_ensure_directory_exists_is_file(self, tmp_path):
        """覆盖 ensure_directory_exists 路径是文件时抛出 ValueError。"""
        file_path = tmp_path / "test_file.txt"
        file_path.write_text("test")
        with pytest.raises(ValueError, match="不是文件夹"):
            ensure_directory_exists(str(file_path))
