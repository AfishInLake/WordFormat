#! /usr/bin/env python
# @Time    : 2024/10/12 15:18
# @Author  : afish
# @File    : test_get_some.py
"""
测试获取段落/字体属性的函数
"""

import unittest
from unittest.mock import Mock, MagicMock, patch

from docx.enum.text import WD_LINE_SPACING
from docx.oxml.shared import qn

from wordformat.style.get_some import (
    paragraph_get_alignment,
    _get_effective_line_height,
    _get_style_spacing,
    paragraph_get_space_before,
    paragraph_get_space_after,
    _get_space_from_style,
    paragraph_get_line_spacing,
    _get_font_size_pt,
    paragraph_get_first_line_indent,
    paragraph_get_builtin_style_name,
    run_get_font_name,
    run_get_font_size_pt,
    run_get_font_color,
    run_get_font_bold,
    run_get_font_italic,
    run_get_font_underline,
)


class TestGetSome(unittest.TestCase):
    """测试获取段落/字体属性的函数"""

    def test_paragraph_get_alignment(self):
        """测试获取段落的有效对齐方式"""
        # 测试直接设置了对齐方式的情况
        paragraph = Mock()
        paragraph.paragraph_format.alignment = Mock()
        result = paragraph_get_alignment(paragraph)
        self.assertEqual(result, paragraph.paragraph_format.alignment)

        # 测试从样式中获取对齐方式的情况
        paragraph.paragraph_format.alignment = None
        style = Mock()
        style.paragraph_format.alignment = Mock()
        style._base_style = None
        paragraph.style = style
        result = paragraph_get_alignment(paragraph)
        self.assertEqual(result, style.paragraph_format.alignment)

        # 测试样式链中获取对齐方式的情况
        base_style = Mock()
        base_style.paragraph_format.alignment = Mock()
        base_style._base_style = None
        style._base_style = base_style
        style.paragraph_format.alignment = None
        result = paragraph_get_alignment(paragraph)
        self.assertEqual(result, base_style.paragraph_format.alignment)

        # 测试无对齐方式设置的情况
        base_style.paragraph_format.alignment = None
        result = paragraph_get_alignment(paragraph)
        self.assertIsNone(result)

    def test_get_effective_line_height(self):
        """测试计算段落的有效行高"""
        # 测试从段落样式获取字体大小的情况
        paragraph = Mock()
        paragraph.paragraph_format.line_spacing = 1.5
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        style = Mock()
        style.font.size = Mock()
        style.font.size.pt = 12.0
        paragraph.style = style
        paragraph.runs = []
        result = _get_effective_line_height(paragraph)
        self.assertEqual(result, 18.0)

        # 测试从runs获取字体大小的情况
        paragraph.style.font.size = None
        run = Mock()
        run.font.size = Mock()
        run.font.size.pt = 14.0
        paragraph.runs = [run]
        result = _get_effective_line_height(paragraph)
        self.assertEqual(result, 21.0)

        # 测试固定行高的情况
        paragraph.paragraph_format.line_spacing = Mock()
        paragraph.paragraph_format.line_spacing.pt = 24.0
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        result = _get_effective_line_height(paragraph)
        self.assertEqual(result, 24.0)

        # 测试默认单倍行距的情况
        paragraph.paragraph_format.line_spacing = None
        paragraph.paragraph_format.line_spacing_rule = None
        result = _get_effective_line_height(paragraph)
        self.assertEqual(result, 14.0)

    def test_get_style_spacing(self):
        """测试递归查找样式中的段前/段后间距"""
        # 测试直接从样式中获取间距的情况
        style = Mock()
        style_elem = Mock()
        style.element = style_elem
        style_pPr = Mock()
        style_elem.find.return_value = style_pPr
        spacing = Mock()
        spacing.get.side_effect = lambda attr: {
            qn("w:beforeLines"): "200",
            qn("w:before"): "240",
        }[attr]
        style_pPr.find.return_value = spacing
        style.base_style = None
        result = _get_style_spacing(style, "before")
        self.assertEqual(result, (2.0, 240))

        # 测试从基样式中获取间距的情况
        base_style = Mock()
        base_style_elem = Mock()
        base_style.element = base_style_elem
        base_style_pPr = Mock()
        base_style_elem.find.return_value = base_style_pPr
        base_spacing = Mock()
        base_spacing.get.side_effect = lambda attr: {
            qn("w:beforeLines"): "100",
            qn("w:before"): "120",
        }[attr]
        base_style_pPr.find.return_value = base_spacing
        base_style.base_style = None
        style.base_style = base_style
        spacing.get.side_effect = lambda attr: {
            qn("w:beforeLines"): None,
            qn("w:before"): None,
        }[attr]
        result = _get_style_spacing(style, "before")
        self.assertEqual(result, (1.0, 120))

        # 测试无样式的情况
        result = _get_style_spacing(None, "before")
        self.assertEqual(result, (0.0, 0))

    def test_paragraph_get_space_before(self):
        """测试精准获取段前间距"""
        # 测试段落自身设置了Lines值的情况
        paragraph = Mock()
        p = Mock()
        paragraph._element = p
        pPr = Mock()
        p.find.return_value = pPr
        spacing = Mock()
        spacing.get.side_effect = lambda attr: {
            qn("w:beforeLines"): "200",
            qn("w:before"): "240",
        }[attr]
        pPr.find.return_value = spacing
        style = Mock()
        paragraph.style = style
        with patch('wordformat.style.get_some._get_font_size_pt', return_value=12.0):
            with patch('wordformat.style.get_some._get_style_spacing', return_value=(1.0, 120)):
                result = paragraph_get_space_before(paragraph)
                self.assertEqual(result, 2.0)

        # 测试从样式中获取Lines值的情况
        spacing.get.side_effect = lambda attr: {
            qn("w:beforeLines"): None,
            qn("w:before"): "240",
        }[attr]
        with patch('wordformat.style.get_some._get_font_size_pt', return_value=12.0):
            with patch('wordformat.style.get_some._get_style_spacing', return_value=(1.5, 180)):
                result = paragraph_get_space_before(paragraph)
                self.assertEqual(result, 1.5)

        # 测试从样式中获取twips值的情况
        with patch('wordformat.style.get_some._get_font_size_pt', return_value=12.0):
            with patch('wordformat.style.get_some._get_style_spacing', return_value=(0.0, 240)):
                result = paragraph_get_space_before(paragraph)
                self.assertEqual(result, 1.0)

        # 测试小数值twips视为0行的情况
        spacing.get.side_effect = lambda attr: {
            qn("w:beforeLines"): None,
            qn("w:before"): "100",
        }[attr]
        with patch('wordformat.style.get_some._get_font_size_pt', return_value=12.0):
            with patch('wordformat.style.get_some._get_style_spacing', return_value=(0.0, 100)):
                result = paragraph_get_space_before(paragraph)
                self.assertEqual(result, 0.0)

    def test_paragraph_get_space_after(self):
        """测试获取段落段后间距"""
        # 测试段落自身设置了Lines值的情况
        paragraph = Mock()
        p = Mock()
        paragraph._element = p
        pPr = Mock()
        p.find.return_value = pPr
        spacing = Mock()
        spacing.get.side_effect = lambda attr: {
            qn("w:afterLines"): "200",
            qn("w:after"): "240",
        }[attr]
        pPr.find.return_value = spacing
        style = Mock()
        style.name = "正文"
        paragraph.style = style
        with patch('wordformat.style.get_some._get_font_size_pt', return_value=12.0):
            with patch('wordformat.style.get_some._get_style_spacing', return_value=(1.0, 120)):
                result = paragraph_get_space_after(paragraph)
                self.assertEqual(result, 2.0)

        # 测试从样式中获取Lines值的情况
        spacing.get.side_effect = lambda attr: {
            qn("w:afterLines"): None,
            qn("w:after"): "240",
        }[attr]
        with patch('wordformat.style.get_some._get_font_size_pt', return_value=12.0):
            with patch('wordformat.style.get_some._get_style_spacing', return_value=(1.5, 180)):
                result = paragraph_get_space_after(paragraph)
                self.assertEqual(result, 1.5)

        # 测试从样式中获取twips值的情况
        with patch('wordformat.style.get_some._get_font_size_pt', return_value=12.0):
            with patch('wordformat.style.get_some._get_style_spacing', return_value=(0.0, 240)):
                result = paragraph_get_space_after(paragraph)
                self.assertEqual(result, 1.0)

        # 测试小数值twips视为0行的情况
        spacing.get.side_effect = lambda attr: {
            qn("w:afterLines"): None,
            qn("w:after"): "100",
        }[attr]
        with patch('wordformat.style.get_some._get_font_size_pt', return_value=12.0):
            with patch('wordformat.style.get_some._get_style_spacing', return_value=(0.0, 100)):
                result = paragraph_get_space_after(paragraph)
                self.assertEqual(result, 0.0)

        # 测试标题样式的默认值
        style.name = "标题1"
        spacing.get.side_effect = lambda attr: {
            qn("w:afterLines"): None,
            qn("w:after"): None,
        }[attr]
        with patch('wordformat.style.get_some._get_font_size_pt', return_value=12.0):
            with patch('wordformat.style.get_some._get_style_spacing', return_value=(0.0, 0)):
                result = paragraph_get_space_after(paragraph)
                self.assertEqual(result, 0.5)

    def test_get_space_from_style(self):
        """测试从样式中获取间距设置"""
        # 测试直接从样式中获取间距的情况
        paragraph = Mock()
        style = Mock()
        paragraph.style = style
        style_fmt = Mock()
        style.paragraph_format = style_fmt
        space_before = Mock()
        space_before.pt = 12.0
        style_fmt.space_before = space_before
        style.base_style = None
        result = _get_space_from_style(paragraph, "before")
        self.assertEqual(result, 12.0)

        # 测试从基样式中获取间距的情况
        base_style = Mock()
        base_style_fmt = Mock()
        base_style.paragraph_format = base_style_fmt
        base_space_before = Mock()
        base_space_before.pt = 6.0
        base_style_fmt.space_before = base_space_before
        base_style.base_style = None
        style.base_style = base_style
        style_fmt.space_before = None
        result = _get_space_from_style(paragraph, "before")
        self.assertEqual(result, 6.0)

        # 测试无样式的情况
        paragraph.style = None
        result = _get_space_from_style(paragraph, "before")
        self.assertEqual(result, 0.0)

    def test_paragraph_get_line_spacing(self):
        """测试获取段落行间距"""
        # 测试直接设置了多倍行距的情况
        paragraph = Mock()
        fmt = Mock()
        paragraph.paragraph_format = fmt
        fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        fmt.line_spacing = 1.5
        result = paragraph_get_line_spacing(paragraph)
        self.assertEqual(result, 1.5)

        # 测试直接设置了1.5倍行距的情况
        fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        result = paragraph_get_line_spacing(paragraph)
        self.assertEqual(result, 1.5)

        # 测试直接设置了双倍行距的情况
        fmt.line_spacing_rule = WD_LINE_SPACING.DOUBLE
        result = paragraph_get_line_spacing(paragraph)
        self.assertEqual(result, 2.0)

        # 测试直接设置了单倍行距的情况
        fmt.line_spacing_rule = WD_LINE_SPACING.SINGLE
        result = paragraph_get_line_spacing(paragraph)
        self.assertEqual(result, 1.0)

        # 测试从样式中获取行距的情况
        fmt.line_spacing_rule = None
        fmt.line_spacing = None
        style = Mock()
        paragraph.style = style
        style_fmt = Mock()
        style.paragraph_format = style_fmt
        style_fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        style_fmt.line_spacing = 1.25
        result = paragraph_get_line_spacing(paragraph)
        self.assertEqual(result, 1.25)

        # 测试固定行高的情况
        fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        fmt.line_spacing = Mock()
        fmt.line_spacing.pt = 18.0
        with patch('wordformat.style.get_some._get_font_size_pt', return_value=12.0):
            result = paragraph_get_line_spacing(paragraph)
            self.assertEqual(result, 1.5)

        # 测试默认情况
        fmt.line_spacing_rule = None
        fmt.line_spacing = None
        style.paragraph_format = None
        paragraph.runs = []
        result = paragraph_get_line_spacing(paragraph)
        self.assertEqual(result, 1.0)

    def test_get_font_size_pt(self):
        """测试获取段落的字体大小"""
        # 测试从runs中获取字体大小的情况
        paragraph = Mock()
        run1 = Mock()
        run1.font.size = Mock()
        run1.font.size.pt = 12.0
        run2 = Mock()
        run2.font.size = Mock()
        run2.font.size.pt = 14.0
        paragraph.runs = [run1, run2]
        result = _get_font_size_pt(paragraph)
        self.assertEqual(result, 14.0)

        # 测试无runs的情况
        paragraph.runs = []
        result = _get_font_size_pt(paragraph)
        self.assertEqual(result, 12.0)

    def test_paragraph_get_first_line_indent(self):
        """测试精准获取首行缩进"""
        # 测试字符单位的情况
        paragraph = Mock()
        p = Mock()
        paragraph._element = p
        pPr = Mock()
        p.find.return_value = pPr
        ind = Mock()
        ind.get.side_effect = lambda attr: {
            qn("w:firstLineChars"): "200",
            qn("w:firstLine"): None,
        }[attr]
        pPr.find.return_value = ind
        result = paragraph_get_first_line_indent(paragraph)
        self.assertEqual(result, 2.0)

        # 测试物理单位的情况
        ind.get.side_effect = lambda attr: {
            qn("w:firstLineChars"): None,
            qn("w:firstLine"): "240",
        }[attr]
        with patch('wordformat.style.get_some._get_font_size_pt', return_value=12.0):
            result = paragraph_get_first_line_indent(paragraph)
            self.assertEqual(result, 1.0)

        # 测试无缩进设置的情况
        pPr.find.return_value = None
        result = paragraph_get_first_line_indent(paragraph)
        self.assertIsNone(result)

        # 测试异常情况
        paragraph._element = None
        result = paragraph_get_first_line_indent(paragraph)
        self.assertIsNone(result)

    def test_paragraph_get_builtin_style_name(self):
        """测试获取段落样式名称"""
        # 测试有样式的情况
        paragraph = Mock()
        style = Mock()
        style.name = "Heading 1"
        paragraph.style = style
        result = paragraph_get_builtin_style_name(paragraph)
        self.assertEqual(result, "heading 1")

        # 测试无样式的情况
        paragraph.style = None
        result = paragraph_get_builtin_style_name(paragraph)
        self.assertEqual(result, "")

    def test_run_get_font_name(self):
        """测试获取 Run 对象的东亚字体名称"""
        # 测试有东亚字体的情况
        run = Mock()
        rPr = Mock()
        run._element.rPr = rPr
        rFonts = Mock()
        rPr.rFonts = rFonts
        rFonts.get.return_value = "Microsoft YaHei"
        result = run_get_font_name(run)
        self.assertEqual(result, "Microsoft YaHei")

        # 测试无东亚字体的情况
        rFonts.get.return_value = None
        result = run_get_font_name(run)
        self.assertIsNone(result)

        # 测试无rFonts的情况
        rPr.rFonts = None
        result = run_get_font_name(run)
        self.assertIsNone(result)

        # 测试无rPr的情况
        run._element.rPr = None
        result = run_get_font_name(run)
        self.assertIsNone(result)

    def test_run_get_font_size_pt(self):
        """测试获取run的字体大小"""
        # 测试直接设置了字体大小的情况
        run = Mock()
        font = Mock()
        run.font = font
        font.size = Mock()
        font.size.pt = 12.0
        result = run_get_font_size_pt(run)
        self.assertEqual(result, 12.0)

        # 测试从样式中获取字体大小的情况
        font.size = None
        parent = Mock()
        run._parent = parent
        style = Mock()
        parent.style = style
        style_font = Mock()
        style.font = style_font
        style_font.size = Mock()
        style_font.size.pt = 14.0
        result = run_get_font_size_pt(run)
        self.assertEqual(result, 14.0)

        # 测试默认情况
        style.font = None
        result = run_get_font_size_pt(run)
        self.assertEqual(result, 12.0)

    def test_run_get_font_color(self):
        """测试获取run的字体颜色"""
        # 测试有颜色设置的情况
        run = Mock()
        font = Mock()
        run.font = font
        color = Mock()
        font.color = color
        color.rgb = (255, 0, 0)
        result = run_get_font_color(run)
        self.assertEqual(result, (255, 0, 0))

        # 测试无颜色设置的情况
        color.rgb = None
        result = run_get_font_color(run)
        self.assertEqual(result, (0, 0, 0))

        # 测试无color对象的情况
        font.color = None
        result = run_get_font_color(run)
        self.assertEqual(result, (0, 0, 0))

    def test_run_get_font_bold(self):
        """测试获取run的字体是否加粗"""
        # 测试加粗的情况
        run = Mock()
        font = Mock()
        run.font = font
        font.bold = True
        result = run_get_font_bold(run)
        self.assertTrue(result)

        # 测试未加粗的情况
        font.bold = False
        result = run_get_font_bold(run)
        self.assertFalse(result)

    def test_run_get_font_italic(self):
        """测试获取run的字体是否斜体"""
        # 测试斜体的情况
        run = Mock()
        font = Mock()
        run.font = font
        font.italic = True
        result = run_get_font_italic(run)
        self.assertTrue(result)

        # 测试未斜体的情况
        font.italic = False
        result = run_get_font_italic(run)
        self.assertFalse(result)

    def test_run_get_font_underline(self):
        """测试获取run的字体是否下划线"""
        # 测试有下划线的情况
        run = Mock()
        font = Mock()
        run.font = font
        font.underline = True
        result = run_get_font_underline(run)
        self.assertTrue(result)

        # 测试无下划线的情况
        font.underline = False
        result = run_get_font_underline(run)
        self.assertFalse(result)


if __name__ == '__main__':
    unittest.main()
