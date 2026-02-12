#! /usr/bin/env python
# @Time    : 2026/2/12 22:45
# @Author  : afish
# @File    : test_utils.py
"""
测试 style/utils.py 中的函数和类
"""

import unittest

from wordformat.style.utils import UnitResult, extract_unit_from_string


class TestUnitResult(unittest.TestCase):
    """测试 UnitResult 类"""

    def test_init(self):
        """测试初始化"""
        unit_result = UnitResult()
        self.assertIsNone(unit_result.original_unit)
        self.assertIsNone(unit_result.standard_unit)
        self.assertIsNone(unit_result.value)
        self.assertFalse(unit_result.is_valid)

    def test_to_dict(self):
        """测试 to_dict 方法"""
        unit_result = UnitResult(
            original_unit="pt",
            standard_unit="pt",
            value=12.0,
            is_valid=True
        )
        result_dict = unit_result.to_dict()
        self.assertEqual(result_dict["original_unit"], "pt")
        self.assertEqual(result_dict["standard_unit"], "pt")
        self.assertEqual(result_dict["value"], 12.0)
        self.assertTrue(result_dict["is_valid"])

    def test_convert_to_emu(self):
        """测试 convert_to_emu 方法"""
        # 测试有效情况
        unit_result1 = UnitResult(
            original_unit="pt",
            standard_unit="pt",
            value=12.0,
            is_valid=True
        )
        self.assertEqual(unit_result1.convert_to_emu(), 152400)  # 12 * 12700

        # 测试无效情况
        unit_result2 = UnitResult()
        self.assertIsNone(unit_result2.convert_to_emu())

        # 测试非国际标准单位
        unit_result3 = UnitResult(
            original_unit="行",
            standard_unit="行",
            value=1.0,
            is_valid=True
        )
        self.assertIsNone(unit_result3.convert_to_emu())

    def test_unit_ch_property(self):
        """测试 unit_ch 属性"""
        # 测试各种单位
        test_cases = [
            ("pt", "磅"),
            ("cm", "厘米"),
            ("inch", "英寸"),
            ("mm", "毫米"),
            ("emu", "emu"),
            ("hang", "行"),
            ("char", "字符"),
            ("bei", "倍"),
            (None, "")
        ]

        for standard_unit, expected_unit_ch in test_cases:
            with self.subTest(standard_unit=standard_unit):
                unit_result = UnitResult(standard_unit=standard_unit)
                if standard_unit not in ["pt", "cm", "inch", "mm", "emu", "hang", "char", "bei", None]:
                    with self.assertRaises(ValueError):
                        _ = unit_result.unit_ch
                else:
                    self.assertEqual(unit_result.unit_ch, expected_unit_ch)


class TestExtractUnitFromString(unittest.TestCase):
    """测试 extract_unit_from_string 函数"""

    def test_extract_with_pt(self):
        """测试提取 pt 单位"""
        result = extract_unit_from_string("12pt")
        self.assertEqual(result.original_unit, "pt")
        self.assertEqual(result.standard_unit, "pt")
        self.assertEqual(result.value, 12.0)
        self.assertTrue(result.is_valid)

        result2 = extract_unit_from_string("12 磅")
        self.assertEqual(result2.original_unit, "磅")
        self.assertEqual(result2.standard_unit, "pt")
        self.assertEqual(result2.value, 12.0)
        self.assertTrue(result2.is_valid)

    def test_extract_with_cm(self):
        """测试提取 cm 单位"""
        result = extract_unit_from_string("2.5cm")
        self.assertEqual(result.original_unit, "cm")
        self.assertEqual(result.standard_unit, "cm")
        self.assertEqual(result.value, 2.5)
        self.assertTrue(result.is_valid)

        result2 = extract_unit_from_string("2.5 厘米")
        self.assertEqual(result2.original_unit, "厘米")
        self.assertEqual(result2.standard_unit, "cm")
        self.assertEqual(result2.value, 2.5)
        self.assertTrue(result2.is_valid)

    def test_extract_with_inch(self):
        """测试提取 inch 单位"""
        result = extract_unit_from_string("1inch")
        self.assertEqual(result.original_unit, "inch")
        self.assertEqual(result.standard_unit, "inch")
        self.assertEqual(result.value, 1.0)
        self.assertTrue(result.is_valid)

        result2 = extract_unit_from_string("1 英寸")
        self.assertEqual(result2.original_unit, "英寸")
        self.assertEqual(result2.standard_unit, "inch")
        self.assertEqual(result2.value, 1.0)
        self.assertTrue(result2.is_valid)

    def test_extract_with_mm(self):
        """测试提取 mm 单位"""
        result = extract_unit_from_string("25mm")
        self.assertEqual(result.original_unit, "mm")
        self.assertEqual(result.standard_unit, "mm")
        self.assertEqual(result.value, 25.0)
        self.assertTrue(result.is_valid)

        result2 = extract_unit_from_string("25 毫米")
        self.assertEqual(result2.original_unit, "毫米")
        self.assertEqual(result2.standard_unit, "mm")
        self.assertEqual(result2.value, 25.0)
        self.assertTrue(result2.is_valid)

    def test_extract_with_emu(self):
        """测试提取 emu 单位"""
        result = extract_unit_from_string("914400emu")
        self.assertEqual(result.original_unit, "emu")
        self.assertEqual(result.standard_unit, "emu")
        self.assertEqual(result.value, 914400.0)
        self.assertTrue(result.is_valid)

    def test_extract_with_hang(self):
        """测试提取 行 单位"""
        result = extract_unit_from_string("1行")
        self.assertEqual(result.original_unit, "行")
        self.assertEqual(result.standard_unit, "hang")
        self.assertEqual(result.value, 1.0)
        self.assertTrue(result.is_valid)

    def test_extract_with_char(self):
        """测试提取 字符 单位"""
        result = extract_unit_from_string("2字符")
        self.assertEqual(result.original_unit, "字符")
        self.assertEqual(result.standard_unit, "char")
        self.assertEqual(result.value, 2.0)
        self.assertTrue(result.is_valid)

    def test_extract_with_bei(self):
        """测试提取 倍 单位"""
        result = extract_unit_from_string("1.5倍")
        self.assertEqual(result.original_unit, "倍")
        self.assertEqual(result.standard_unit, "bei")
        self.assertEqual(result.value, 1.5)
        self.assertTrue(result.is_valid)

    def test_extract_without_unit(self):
        """测试无单位情况"""
        result = extract_unit_from_string("123")
        self.assertIsNone(result.original_unit)
        self.assertIsNone(result.standard_unit)
        self.assertIsNone(result.value)
        self.assertFalse(result.is_valid)

    def test_extract_with_invalid_input(self):
        """测试无效输入"""
        result = extract_unit_from_string("")
        self.assertIsNone(result.original_unit)
        self.assertIsNone(result.standard_unit)
        self.assertIsNone(result.value)
        self.assertFalse(result.is_valid)


if __name__ == '__main__':
    unittest.main()
