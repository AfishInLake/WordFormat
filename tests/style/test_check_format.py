#! /usr/bin/env python
# @Time    : 2024/10/12 15:18
# @Author  : afish
# @File    : test_check_format.py
"""
测试检查格式的函数和类
"""

import unittest
from unittest.mock import Mock, patch

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from wordformat.style.check_format import DIFFResult, CharacterStyle, ParagraphStyle


class TestDIFFResult(unittest.TestCase):
    """测试 DIFFResult 类"""

    def test_init(self):
        """测试初始化"""
        diff_result = DIFFResult(
            diff_type="bold",
            expected_value=True,
            current_value=False,
            comment="期待加粗;",
            level=1
        )
        self.assertEqual(diff_result.diff_type, "bold")
        self.assertEqual(diff_result.expected_value, True)
        self.assertEqual(diff_result.current_value, False)
        self.assertEqual(diff_result.comment, "期待加粗;")
        self.assertEqual(diff_result.level, 1)

    def test_str(self):
        """测试 __str__ 方法"""
        diff_result = DIFFResult(comment="期待加粗;")
        self.assertEqual(str(diff_result), "期待加粗;")


class TestCharacterStyle(unittest.TestCase):
    """测试 CharacterStyle 类"""

    def setUp(self):
        """设置测试环境"""
        # 模拟配置
        with patch('wordformat.style.check_format.get_config') as mock_get_config:
            mock_config = Mock()
            mock_warning = Mock()
            mock_warning.bold = True
            mock_warning.italic = True
            mock_warning.underline = True
            mock_warning.font_size = True
            mock_warning.font_color = True
            mock_warning.font_name = True
            mock_config.style_checks_warning = mock_warning
            mock_get_config.return_value = mock_config
            self.character_style = CharacterStyle()

    def test_init(self):
        """测试初始化"""
        self.assertIsNotNone(self.character_style.font_name_cn)
        self.assertIsNotNone(self.character_style.font_name_en)
        self.assertIsNotNone(self.character_style.font_size)
        self.assertIsNotNone(self.character_style.font_color)
        self.assertFalse(self.character_style.bold)
        self.assertFalse(self.character_style.italic)
        self.assertFalse(self.character_style.underline)

    def test_diff_from_run(self):
        """测试 diff_from_run 方法"""
        # 模拟 Run 对象
        run = Mock()
        run.font.bold = False
        run.font.italic = False
        run.font.underline = False
        run.font.name = "Times New Roman"

        # 模拟获取字体大小
        with patch('wordformat.style.check_format.run_get_font_size_pt', return_value=12.0):
            # 模拟获取字体颜色
            with patch('wordformat.style.check_format.run_get_font_color', return_value=(0, 0, 0)):
                # 模拟获取字体名称
                with patch('wordformat.style.check_format.run_get_font_name', return_value="宋体"):
                    diffs = self.character_style.diff_from_run(run)
                    # 由于 font_name_en 可能与默认值不一致，所以会有 1 个差异
                    self.assertEqual(len(diffs), 1)

        # 测试有差异的情况
        run.font.bold = True
        with patch('wordformat.style.check_format.run_get_font_size_pt', return_value=14.0):
            with patch('wordformat.style.check_format.run_get_font_color', return_value=(255, 0, 0)):
                with patch('wordformat.style.check_format.run_get_font_name', return_value="黑体"):
                    diffs = self.character_style.diff_from_run(run)
                    # 应该有5个差异：bold, font_size, font_color, font_name_cn, font_name_en
                    self.assertEqual(len(diffs), 5)

    def test_apply_to_run(self):
        """测试 apply_to_run 方法"""
        # 模拟 Run 对象
        run = Mock()
        run.font.bold = True
        run.font.italic = False
        run.font.underline = False

        # 模拟获取字体大小
        with patch('wordformat.style.check_format.run_get_font_size_pt', return_value=14.0):
            # 模拟获取字体颜色
            with patch('wordformat.style.check_format.run_get_font_color', return_value=(255, 0, 0)):
                # 模拟获取字体名称
                with patch('wordformat.style.check_format.run_get_font_name', return_value="黑体"):
                    # 模拟 format 方法
                    self.character_style.font_size.format = Mock()
                    self.character_style.font_color.format = Mock()
                    self.character_style.font_name_cn.format = Mock()
                    self.character_style.font_name_en.format = Mock()

                    result = self.character_style.apply_to_run(run)
                    # 应该有5个修正：bold, font_size, font_color, font_name_cn, font_name_en
                    self.assertEqual(len(result), 5)

    def test_to_string(self):
        """测试 to_string 方法"""
        # 创建 DIFFResult 对象
        diffs = [
            DIFFResult(diff_type="bold", comment="期待不加粗;"),
            DIFFResult(diff_type="italic", comment="期待非斜体;"),
            DIFFResult(diff_type="underline", comment="期待无下划线;"),
        ]

        result = CharacterStyle.to_string(diffs)
        self.assertIn("期待不加粗;", result)
        self.assertIn("期待非斜体;", result)
        self.assertIn("期待无下划线;", result)


class TestParagraphStyle(unittest.TestCase):
    """测试 ParagraphStyle 类"""

    def setUp(self):
        """设置测试环境"""
        # 模拟配置
        with patch('wordformat.style.check_format.get_config') as mock_get_config:
            mock_config = Mock()
            mock_warning = Mock()
            mock_warning.alignment = True
            mock_warning.space_before = True
            mock_warning.space_after = True
            mock_warning.line_spacing = True
            mock_warning.line_spacingrule = True
            mock_warning.first_line_indent = True
            mock_warning.builtin_style_name = True
            mock_config.style_checks_warning = mock_warning
            mock_get_config.return_value = mock_config
            self.paragraph_style = ParagraphStyle()

    def test_init(self):
        """测试初始化"""
        self.assertIsNotNone(self.paragraph_style.alignment)
        self.assertIsNotNone(self.paragraph_style.space_before)
        self.assertIsNotNone(self.paragraph_style.space_after)
        self.assertIsNotNone(self.paragraph_style.line_spacing)
        self.assertIsNotNone(self.paragraph_style.line_spacingrule)
        self.assertIsNotNone(self.paragraph_style.first_line_indent)
        self.assertIsNotNone(self.paragraph_style.builtin_style_name)

    def test_apply_to_paragraph(self):
        """测试 apply_to_paragraph 方法"""
        # 模拟 Paragraph 对象
        paragraph = Mock()

        # 模拟 diff_from_paragraph 方法
        diffs = [
            DIFFResult(diff_type="alignment"),
            DIFFResult(diff_type="space_before"),
            DIFFResult(diff_type="space_after"),
        ]

        # 模拟 format 方法
        self.paragraph_style.alignment.format = Mock()
        self.paragraph_style.space_before.format = Mock()
        self.paragraph_style.space_after.format = Mock()

        with patch.object(self.paragraph_style, 'diff_from_paragraph', return_value=diffs):
            result = self.paragraph_style.apply_to_paragraph(paragraph)
            # 应该有3个修正
            self.assertEqual(len(result), 3)

    def test_diff_from_paragraph(self):
        """测试 diff_from_paragraph 方法"""
        # 模拟 Paragraph 对象
        paragraph = Mock()
        paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.space_before = Mock()
        paragraph.paragraph_format.space_before.pt = 6.0
        paragraph.paragraph_format.space_after = Mock()
        paragraph.paragraph_format.space_after.pt = 6.0
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        paragraph.paragraph_format.line_spacing = 1.0
        paragraph.paragraph_format.first_line_indent = None

        # 模拟获取段落属性的函数
        with patch('wordformat.style.check_format.paragraph_get_alignment', return_value=WD_ALIGN_PARAGRAPH.LEFT):
            with patch('wordformat.style.check_format.paragraph_get_space_before', return_value=0.5):
                with patch('wordformat.style.check_format.paragraph_get_space_after', return_value=0.5):
                    with patch('wordformat.style.check_format.paragraph_get_line_spacing', return_value=1.5):
                        with patch('wordformat.style.check_format.paragraph_get_first_line_indent', return_value=0.0):
                            with patch('wordformat.style.check_format.paragraph_get_builtin_style_name', return_value="normal"):
                                # 模拟 Spacing 类
                                with patch('wordformat.style.check_format.Spacing'):
                                    diffs = self.paragraph_style.diff_from_paragraph(paragraph)
                                    # 检查是否返回了差异列表
                                    self.assertIsInstance(diffs, list)

    def test_to_string(self):
        """测试 to_string 方法"""
        # 创建 DIFFResult 对象
        diffs = [
            DIFFResult(diff_type="alignment", comment="对齐方式期待左对齐;"),
            DIFFResult(diff_type="space_before", comment="段前间距期待0.5行;"),
            DIFFResult(diff_type="space_after", comment="段后间距期待0.5行;"),
        ]

        result = ParagraphStyle.to_string(diffs)
        self.assertIn("对齐方式期待左对齐;", result)
        self.assertIn("段前间距期待0.5行;", result)
        self.assertIn("段后间距期待0.5行;", result)


if __name__ == '__main__':
    unittest.main()
