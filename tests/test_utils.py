"""
Core 模块综合测试

覆盖 tree.py, utils.py, rules/node.py, numbering.py, settings.py
"""

import os
import pytest
from io import StringIO
from unittest.mock import MagicMock, patch

from docx import Document
from docx.oxml.ns import qn

from wordformat.tree import Tree, Stack, print_tree
from wordformat.rules.node import TreeNode, FormatNode
from wordformat.numbering import (
    _auto_strip_numbering,
    _strip_reference_numbering,
    apply_auto_numbering,
    create_numbering_definition,
    process_heading_numbering,
)
from wordformat.utils import (
    get_file_name,
    ensure_is_directory,
    ensure_directory_exists,
    _to_roman,
    _to_chinese_num,
    load_yaml_with_merge,
    get_paragraph_numbering_text,
    remove_all_numbering,
    _format_number,
    _get_level_fmt,
    _count_numbering_levels,
)
from wordformat.base import DocxBase
from wordformat import settings


# ============================================================
# tree.py — Tree
# ============================================================


# ============================================================
# utils.py — get_file_name
# ============================================================


class TestGetFileName:
    @pytest.mark.parametrize(
        "path, expected",
        [
            ("/home/user/doc.docx", "doc"),
            ("simple.txt", "simple"),
            ("/path/to/my.file.name.pdf", "my.file.name"),
        ],
    )
    def test_extracts_name(self, path, expected):
        assert get_file_name(path) == expected


# ============================================================
# utils.py — ensure_is_directory
# ============================================================


class TestEnsureIsDirectory:
    def test_valid_directory(self, tmp_path):
        # tmp_path 本身就是目录
        ensure_is_directory(str(tmp_path))

    def test_nonexistent_path_raises(self):
        with pytest.raises(ValueError, match="路径不存在"):
            ensure_is_directory("/nonexistent/path/that/does/not/exist")

    def test_file_instead_of_directory_raises(self, tmp_path):
        file_path = tmp_path / "afile.txt"
        file_path.write_text("hello")
        with pytest.raises(ValueError, match="不是一个文件夹"):
            ensure_is_directory(str(file_path))


# ============================================================
# utils.py — ensure_directory_exists
# ============================================================


class TestEnsureDirectoryExists:
    def test_create_new_directory(self, tmp_path):
        new_dir = str(tmp_path / "sub" / "dir")
        ensure_directory_exists(new_dir)
        assert os.path.isdir(new_dir)

    def test_existing_directory_no_error(self, tmp_path):
        ensure_directory_exists(str(tmp_path))

    def test_existing_file_raises(self, tmp_path):
        file_path = tmp_path / "blocked.txt"
        file_path.write_text("data")
        with pytest.raises(ValueError, match="不是文件夹"):
            ensure_directory_exists(str(file_path))


# ============================================================
# utils.py — _to_roman
# ============================================================


class TestToRoman:
    @pytest.mark.parametrize(
        "num, expected",
        [
            (1, "i"),
            (4, "iv"),
            (5, "v"),
            (9, "ix"),
            (10, "x"),
            (40, "xl"),
            (50, "l"),
            (90, "xc"),
            (100, "c"),
            (400, "cd"),
            (500, "d"),
            (900, "cm"),
            (1000, "m"),
            (1984, "mcmlxxxiv"),
            (3999, "mmmcmxcix"),
        ],
    )
    def test_valid_numbers(self, num, expected):
        assert _to_roman(num) == expected

    def test_zero_returns_zero_string(self):
        """_to_roman(0) 返回 "0" 而非空字符串"""
        assert _to_roman(0) == "0"

    def test_negative_returns_zero_string(self):
        """_to_roman(-5) 返回 "0" 而非空字符串"""
        assert _to_roman(-5) == "0"


# ============================================================
# utils.py — _to_chinese_num
# ============================================================


class TestToChineseNum:
    @pytest.mark.parametrize(
        "num, expected",
        [
            (1, "一"),
            (5, "五"),
            (9, "九"),
            (10, "十"),
            (11, "十一"),
            (20, "二十"),
            (21, "二十一"),
            (99, "九十九"),
        ],
    )
    def test_valid_numbers(self, num, expected):
        assert _to_chinese_num(num) == expected

    def test_zero(self):
        assert _to_chinese_num(0) == "0"

    def test_negative(self):
        assert _to_chinese_num(-3) == "-3"

    def test_hundred_should_convert(self):
        assert _to_chinese_num(100) == "一百"

    def test_hundred_fallback(self):
        # _to_chinese_num(100) 返回 "一百"
        assert _to_chinese_num(100) == "一百"


# ============================================================
# utils.py — load_yaml_with_merge
# ============================================================


class TestLoadYamlWithMerge:
    def test_load_valid_yaml(self, tmp_path):
        yaml_file = tmp_path / "config.yaml"
        yaml_file.write_text("key: value\nnested:\n  a: 1\n", encoding="utf-8")
        result = load_yaml_with_merge(str(yaml_file))
        assert result["key"] == "value"
        assert result["nested"]["a"] == 1

    def test_nonexistent_file_raises(self):
        with pytest.raises(FileNotFoundError):
            load_yaml_with_merge("/nonexistent/file.yaml")


# ============================================================
# rules/node.py — TreeNode
# ============================================================


# ============================================================
# settings.py
# ============================================================


class TestSettings:
    def test_batch_size_is_int(self):
        assert isinstance(settings.BATCH_SIZE, int)

    def test_voidnodelist_contains_top(self):
        assert "top" in settings.VOIDNODELIST

    def test_voidnodelist_is_list(self):
        assert isinstance(settings.VOIDNODELIST, list)

    def test_host_is_string(self):
        assert isinstance(settings.HOST, str)

    def test_port_is_int(self):
        assert isinstance(settings.PORT, int)

    def test_server_host_format(self):
        assert settings.SERVER_HOST.startswith("http://")
        assert str(settings.PORT) in settings.SERVER_HOST


# ============================================================
# utils.py — get_paragraph_numbering_text
# ============================================================


# ============================================================
# utils.py — remove_all_numbering
# ============================================================


class TestRemoveAllNumbering:
    """测试 remove_all_numbering 移除标题样式的编号绑定"""

    def test_removes_numPr_from_heading_styles(self, doc):
        """从 Heading 1/2/3 样式中移除 numPr"""
        from docx.oxml import OxmlElement

        for style_name in ["Heading 1", "Heading 2", "Heading 3"]:
            style = doc.styles[style_name]
            style_element = style._element

            # 确保 pPr 存在
            pPr = style_element.find(qn("w:pPr"))
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                style_element.insert(0, pPr)

            # 添加 numPr
            numPr = OxmlElement("w:numPr")
            pPr.append(numPr)

        remove_all_numbering(doc)

        for style_name in ["Heading 1", "Heading 2", "Heading 3"]:
            style = doc.styles[style_name]
            pPr = style._element.find(qn("w:pPr"))
            if pPr is not None:
                numPr = pPr.find(qn("w:numPr"))
                assert numPr is None, f"{style_name} 的 numPr 未被移除"

    def test_removes_outlineLvl_from_heading_styles(self, doc):
        """从 Heading 1/2/3 样式中移除 outlineLvl"""
        from docx.oxml import OxmlElement

        for style_name in ["Heading 1", "Heading 2", "Heading 3"]:
            style = doc.styles[style_name]
            style_element = style._element

            pPr = style_element.find(qn("w:pPr"))
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                style_element.insert(0, pPr)

            # 先移除已有的 outlineLvl
            existing = pPr.find(qn("w:outlineLvl"))
            if existing is not None:
                pPr.remove(existing)

            outlineLvl = OxmlElement("w:outlineLvl")
            outlineLvl.set(qn("w:val"), "0")
            pPr.append(outlineLvl)

        remove_all_numbering(doc)

        for style_name in ["Heading 1", "Heading 2", "Heading 3"]:
            style = doc.styles[style_name]
            pPr = style._element.find(qn("w:pPr"))
            if pPr is not None:
                outlineLvl = pPr.find(qn("w:outlineLvl"))
                assert outlineLvl is None, f"{style_name} 的 outlineLvl 未被移除"

    def test_no_pPr_in_style_no_error(self, doc):
        """样式中没有 pPr 时不报错"""
        # Heading 1 的 pPr 可能存在也可能不存在，直接调用不应报错
        remove_all_numbering(doc)

    def test_style_not_in_doc_no_error(self, doc):
        """如果样式不存在（不太可能），不报错"""
        # Document() 默认包含 Heading 1/2/3，所以这个测试主要验证不会抛异常
        remove_all_numbering(doc)


# ============================================================
# numbering.py — auto_strip_numbering (empty result)
# ============================================================
