#! /usr/bin/env python
# -*- coding: utf-8 -*-
import os
import tempfile
import pytest
from unittest import mock

from wordformat.cli import validate_file, create_common_parser, validate_json_path, main


def test_validate_file():
    """测试validate_file函数"""
    # 测试有效的文件
    with tempfile.NamedTemporaryFile(delete=False) as f:
        temp_file = f.name
    try:
        result = validate_file(temp_file, "测试文件")
        assert result == os.path.abspath(temp_file)
    finally:
        os.unlink(temp_file)
    
    # 测试不存在的文件
    with pytest.raises(Exception):
        validate_file("nonexistent_file_12345.txt", "测试文件")
    
    # 测试文件夹路径
    with tempfile.TemporaryDirectory() as temp_dir:
        with pytest.raises(Exception):
            validate_file(temp_dir, "测试文件")


def test_validate_json_path():
    """测试validate_json_path函数"""
    # 测试generate-json模式（文件不存在但路径合法）
    nonexistent_json = "nonexistent_json_12345.json"
    try:
        result = validate_json_path(nonexistent_json, "generate-json")
        assert result == os.path.abspath(nonexistent_json)
    finally:
        if os.path.exists(nonexistent_json):
            os.unlink(nonexistent_json)
    
    # 测试check模式（文件必须存在）
    with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as f:
        temp_json = f.name
    try:
        result = validate_json_path(temp_json, "check-format")
        assert result == os.path.abspath(temp_json)
        
        # 测试check模式下文件不存在的情况
        os.unlink(temp_json)
        with pytest.raises(Exception):
            validate_json_path(temp_json, "check-format")
    finally:
        if os.path.exists(temp_json):
            os.unlink(temp_json)
    
    # 测试文件夹路径
    with tempfile.TemporaryDirectory() as temp_dir:
        with pytest.raises(Exception):
            validate_json_path(temp_dir, "generate-json")


def test_create_common_parser():
    """测试create_common_parser函数"""
    import argparse
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="mode")
    
    # 创建公共解析器
    common_parser = create_common_parser(subparsers, "test-command", "测试命令")
    assert common_parser is not None
    assert hasattr(common_parser, "add_argument")


@mock.patch('sys.argv')
def test_main_generate_json(mock_argv):
    """测试main函数的generate-json模式"""
    # 创建临时文件
    from docx import Document
    doc = Document()
    doc.add_paragraph("测试段落")
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
        temp_docx = f.name
    doc.save(temp_docx)
    
    # 创建临时配置文件
    with tempfile.NamedTemporaryFile(suffix='.yaml', delete=False, mode='w', encoding='utf-8') as f:
        f.write("# 测试配置\n")
        f.write("style_checks_warning:\n")
        f.write("  bold: true\n")
        f.write("global_format:\n")
        f.write("  alignment: 左对齐\n")
        f.write("abstract:\n")
        f.write("  chinese:\n")
        f.write("    chinese_title:\n")
        f.write("      section_title_re: ^摘要$\n")
        f.write("  english:\n")
        f.write("    english_title:\n")
        f.write("      section_title_re: ^Abstract$\n")
        f.write("  keywords:\n")
        f.write("    chinese:\n")
        f.write("      section_title_re: ^关键词$\n")
        f.write("    english:\n")
        f.write("      section_title_re: ^Keywords$\n")
        f.write("headings:\n")
        f.write("  level_1:\n")
        f.write("    section_title_re: ^第[一二三四五六七八九十]+章$\n")
        f.write("  level_2:\n")
        f.write("    section_title_re: ^[0-9]+\.[0-9]+\s*$\n")
        f.write("  level_3:\n")
        f.write("    section_title_re: ^[0-9]+\.[0-9]+\.[0-9]+\s*$\n")
        f.write("figures:\n")
        f.write("  section_title_re: ^图[0-9]+\.[0-9]+\s*$\n")
        f.write("tables:\n")
        f.write("  section_title_re: ^表[0-9]+\.[0-9]+\s*$\n")
        f.write("references:\n")
        f.write("  title:\n")
        f.write("    section_title_re: ^参考文献$\n")
        f.write("acknowledgements:\n")
        f.write("  title:\n")
        f.write("    section_title_re: ^致谢$")
        temp_config = f.name
    
    # 创建临时JSON文件路径（不存在）
    temp_json = tempfile.mktemp(suffix='.json')
    
    try:
        # 模拟命令行参数
        mock_argv.__getitem__.side_effect = lambda i: [
            "wordformat",
            "--docx", temp_docx,
            "--json", temp_json,
            "generate-json",
            "--config", temp_config
        ][i]
        mock_argv.__len__.return_value = 8
        
        # 调用main函数
        main()
        
        # 验证JSON文件被创建
        assert os.path.exists(temp_json)
        
    finally:
        # 清理临时文件
        if os.path.exists(temp_docx):
            os.unlink(temp_docx)
        if os.path.exists(temp_config):
            os.unlink(temp_config)
        if os.path.exists(temp_json):
            os.unlink(temp_json)


@mock.patch('sys.argv')
def test_main_check_format(mock_argv):
    """测试main函数的check-format模式"""
    # 创建临时文件
    from docx import Document
    doc = Document()
    doc.add_paragraph("测试段落")
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
        temp_docx = f.name
    doc.save(temp_docx)
    
    # 创建临时配置文件
    with tempfile.NamedTemporaryFile(suffix='.yaml', delete=False, mode='w', encoding='utf-8') as f:
        f.write("# 测试配置\n")
        f.write("style_checks_warning:\n")
        f.write("  bold: true\n")
        f.write("global_format:\n")
        f.write("  alignment: 左对齐\n")
        f.write("abstract:\n")
        f.write("  chinese:\n")
        f.write("    chinese_title:\n")
        f.write("      section_title_re: ^摘要$\n")
        f.write("  english:\n")
        f.write("    english_title:\n")
        f.write("      section_title_re: ^Abstract$\n")
        f.write("  keywords:\n")
        f.write("    chinese:\n")
        f.write("      section_title_re: ^关键词$\n")
        f.write("    english:\n")
        f.write("      section_title_re: ^Keywords$\n")
        f.write("headings:\n")
        f.write("  level_1:\n")
        f.write("    section_title_re: ^第[一二三四五六七八九十]+章$\n")
        f.write("  level_2:\n")
        f.write("    section_title_re: ^[0-9]+\.[0-9]+\s*$\n")
        f.write("  level_3:\n")
        f.write("    section_title_re: ^[0-9]+\.[0-9]+\.[0-9]+\s*$\n")
        f.write("figures:\n")
        f.write("  section_title_re: ^图[0-9]+\.[0-9]+\s*$\n")
        f.write("tables:\n")
        f.write("  section_title_re: ^表[0-9]+\.[0-9]+\s*$\n")
        f.write("references:\n")
        f.write("  title:\n")
        f.write("    section_title_re: ^参考文献$\n")
        f.write("acknowledgements:\n")
        f.write("  title:\n")
        f.write("    section_title_re: ^致谢$")
        temp_config = f.name
    
    # 创建临时JSON文件
    with tempfile.NamedTemporaryFile(suffix='.json', delete=False, mode='w', encoding='utf-8') as f:
        f.write('[{"fingerprint": "test", "category": "heading", "name": "标题1", "children": []}]')
        temp_json = f.name
    
    try:
        # 模拟命令行参数
        mock_argv.__getitem__.side_effect = lambda i: [
            "wordformat",
            "--docx", temp_docx,
            "--json", temp_json,
            "check-format",
            "--config", temp_config,
            "--output", "output/"
        ][i]
        mock_argv.__len__.return_value = 10
        
        # 调用main函数
        main()
        
    finally:
        # 清理临时文件
        if os.path.exists(temp_docx):
            os.unlink(temp_docx)
        if os.path.exists(temp_config):
            os.unlink(temp_config)
        if os.path.exists(temp_json):
            os.unlink(temp_json)


@mock.patch('sys.argv')
def test_main_apply_format(mock_argv):
    """测试main函数的apply-format模式"""
    # 创建临时文件
    from docx import Document
    doc = Document()
    doc.add_paragraph("测试段落")
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
        temp_docx = f.name
    doc.save(temp_docx)
    
    # 创建临时配置文件
    with tempfile.NamedTemporaryFile(suffix='.yaml', delete=False, mode='w', encoding='utf-8') as f:
        f.write("# 测试配置\n")
        f.write("style_checks_warning:\n")
        f.write("  bold: true\n")
        f.write("global_format:\n")
        f.write("  alignment: 左对齐\n")
        f.write("abstract:\n")
        f.write("  chinese:\n")
        f.write("    chinese_title:\n")
        f.write("      section_title_re: ^摘要$\n")
        f.write("  english:\n")
        f.write("    english_title:\n")
        f.write("      section_title_re: ^Abstract$\n")
        f.write("  keywords:\n")
        f.write("    chinese:\n")
        f.write("      section_title_re: ^关键词$\n")
        f.write("    english:\n")
        f.write("      section_title_re: ^Keywords$\n")
        f.write("headings:\n")
        f.write("  level_1:\n")
        f.write("    section_title_re: ^第[一二三四五六七八九十]+章$\n")
        f.write("  level_2:\n")
        f.write("    section_title_re: ^[0-9]+\.[0-9]+\s*$\n")
        f.write("  level_3:\n")
        f.write("    section_title_re: ^[0-9]+\.[0-9]+\.[0-9]+\s*$\n")
        f.write("figures:\n")
        f.write("  section_title_re: ^图[0-9]+\.[0-9]+\s*$\n")
        f.write("tables:\n")
        f.write("  section_title_re: ^表[0-9]+\.[0-9]+\s*$\n")
        f.write("references:\n")
        f.write("  title:\n")
        f.write("    section_title_re: ^参考文献$\n")
        f.write("acknowledgements:\n")
        f.write("  title:\n")
        f.write("    section_title_re: ^致谢$")
        temp_config = f.name
    
    # 创建临时JSON文件
    with tempfile.NamedTemporaryFile(suffix='.json', delete=False, mode='w', encoding='utf-8') as f:
        f.write('[{"fingerprint": "test", "category": "heading", "name": "标题1", "children": []}]')
        temp_json = f.name
    
    try:
        # 模拟命令行参数
        mock_argv.__getitem__.side_effect = lambda i: [
            "wordformat",
            "--docx", temp_docx,
            "--json", temp_json,
            "apply-format",
            "--config", temp_config,
            "--output", "output/"
        ][i]
        mock_argv.__len__.return_value = 10
        
        # 调用main函数
        main()
        
    finally:
        # 清理临时文件
        if os.path.exists(temp_docx):
            os.unlink(temp_docx)
        if os.path.exists(temp_config):
            os.unlink(temp_config)
        if os.path.exists(temp_json):
            os.unlink(temp_json)
