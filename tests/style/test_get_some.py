#! /usr/bin/env python
# @Time    : 2024/10/12 15:18
# @Author  : afish
# @File    : test_get_some.py
"""
测试获取段落/字体属性的函数
"""

import unittest
from unittest.mock import Mock, patch

from docx.enum.text import WD_LINE_SPACING
from docx.oxml.shared import qn

from wordformat.style.get_some import (
    paragraph_get_alignment,
    _get_style_spacing,
    paragraph_get_space_before,
    paragraph_get_space_after,
    paragraph_get_line_spacing,
    paragraph_get_first_line_indent,
    paragraph_get_builtin_style_name,
    run_get_font_name,
    run_get_font_size_pt,
    run_get_font_color,
    run_get_font_bold,
    run_get_font_italic,
    run_get_font_underline, GetIndent,
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
        self.assertEqual(result, 2.0)

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
        self.assertEqual(result, 1.0)

        # 测试无样式的情况
        result = _get_style_spacing(None, "before")
        self.assertEqual(result, None)

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
        with patch('wordformat.style.get_some._get_style_spacing', return_value=1.0):
            result = paragraph_get_space_before(paragraph)
            self.assertEqual(result, 2.0)

        # 测试从样式中获取Lines值的情况
        spacing.get.side_effect = lambda attr: {
            qn("w:beforeLines"): None,
            qn("w:before"): "240",
        }[attr]
        with patch('wordformat.style.get_some._get_style_spacing', return_value=1.5):
            result = paragraph_get_space_before(paragraph)
            self.assertEqual(result, 1.5)

        # 测试从样式中获取twips值的情况
        with patch('wordformat.style.get_some._get_style_spacing', return_value=0.0):
            result = paragraph_get_space_before(paragraph)
            self.assertEqual(result, None)

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


        # 测试从样式中获取Lines值的情况
        spacing.get.side_effect = lambda attr: {
            qn("w:afterLines"): None,
            qn("w:after"): "240",
        }[attr]
        with patch('wordformat.style.get_some._get_style_spacing', return_value=1.5):
            result = paragraph_get_space_after(paragraph)
            self.assertEqual(result, 1.5)

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

        # 测试固定行高的情况
        fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        fmt.line_spacing = Mock()
        fmt.line_spacing.pt = 18.0

        # 测试默认情况
        fmt.line_spacing_rule = None
        fmt.line_spacing = None
        style.paragraph_format = None
        paragraph.runs = []
        result = paragraph_get_line_spacing(paragraph)
        self.assertEqual(result, None)

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

    def test_left_indent_valid_char_value(self):
        """测试 w:leftChars="200" 返回 2.0"""
        paragraph = Mock()
        mock_ind = Mock()
        mock_ind.get.side_effect = lambda attr: "200" if attr == qn('w:leftChars') else None

        mock_pPr = Mock()
        mock_pPr.find.return_value = mock_ind
        paragraph._element.pPr = mock_pPr

        result = GetIndent.left_indent(paragraph)
        self.assertEqual(result, 2.0)

    def test_right_indent_valid_char_value(self):
        """测试 w:rightChars="150" 返回 1.5"""
        paragraph = Mock()
        mock_ind = Mock()
        mock_ind.get.side_effect = lambda attr: "150" if attr == qn('w:rightChars') else None

        mock_pPr = Mock()
        mock_pPr.find.return_value = mock_ind
        paragraph._element.pPr = mock_pPr

        result = GetIndent.right_indent(paragraph)
        self.assertEqual(result, 1.5)

    def test_no_ind_element_returns_none(self):
        """当 <w:ind> 不存在时返回 None"""
        paragraph = Mock()
        mock_pPr = Mock()
        mock_pPr.find.return_value = None
        paragraph._element.pPr = mock_pPr

        self.assertIsNone(GetIndent.left_indent(paragraph))

    def test_pPr_is_none_returns_none(self):
        """当 pPr 为 None 时返回 None"""
        paragraph = Mock()
        paragraph._element.pPr = None

        self.assertIsNone(GetIndent.left_indent(paragraph))

    def test_missing_leftchars_attr_returns_none(self):
        """当 w:leftChars 未设置时返回 None"""
        paragraph = Mock()
        mock_ind = Mock()
        mock_ind.get.return_value = None  # 所有属性都为 None
        mock_pPr = Mock()
        mock_pPr.find.return_value = mock_ind
        paragraph._element.pPr = mock_pPr

        self.assertIsNone(GetIndent.left_indent(paragraph))

    def test_invalid_char_value_returns_none(self):
        """当 w:leftChars="abc" 时返回 None"""
        paragraph = Mock()
        mock_ind = Mock()
        mock_ind.get.side_effect = lambda attr: "abc" if attr == qn('w:leftChars') else None
        mock_pPr = Mock()
        mock_pPr.find.return_value = mock_ind
        paragraph._element.pPr = mock_pPr

        result = GetIndent.left_indent(paragraph)
        self.assertIsNone(result)

    def test_zero_char_indent(self):
        """w:leftChars="0" 应返回 0.0"""
        paragraph = Mock()
        mock_ind = Mock()
        mock_ind.get.side_effect = lambda attr: "0" if attr == qn('w:leftChars') else None
        mock_pPr = Mock()
        mock_pPr.find.return_value = mock_ind
        paragraph._element.pPr = mock_pPr

        result = GetIndent.left_indent(paragraph)
        self.assertEqual(result, 0.0)

    def test_only_right_set_left_is_none(self):
        """只设置了 rightChars，left 应为 None"""
        paragraph = Mock()
        mock_ind = Mock()

        def get_attr(attr):
            if attr == qn('w:rightChars'):
                return "100"
            return None

        mock_ind.get.side_effect = get_attr

        mock_pPr = Mock()
        mock_pPr.find.return_value = mock_ind
        paragraph._element.pPr = mock_pPr

        self.assertIsNone(GetIndent.left_indent(paragraph))
        self.assertEqual(GetIndent.right_indent(paragraph), 1.0)

    def test_invalid_indent_type_raises_value_error(self):
        """无效 indent_type 应抛出 ValueError"""
        paragraph = Mock()
        with self.assertRaises(ValueError) as cm:
            GetIndent.line_indent(paragraph, 'invalid')
        self.assertIn("必须是 'left' 或 'right'", str(cm.exception))


if __name__ == '__main__':
    unittest.main()
