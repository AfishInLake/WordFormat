#! /usr/bin/env python
# -*- coding: utf-8 -*-
import os
import tempfile
import pytest
from docx import Document

from wordformat.utils import (
    check_duplicate_fingerprints,
    get_paragraph_xml_fingerprint,
    load_yaml_with_merge,
    ensure_is_directory,
    get_file_name,
    remove_all_numbering
)


def test_check_duplicate_fingerprints():
    """测试检查重复指纹的功能"""
    # 测试无重复的情况
    data_no_duplicates = [
        {"fingerprint": "123"},
        {"fingerprint": "456"},
        {"fingerprint": "789"}
    ]
    check_duplicate_fingerprints(data_no_duplicates)
    
    # 测试有重复的情况
    data_with_duplicates = [
        {"fingerprint": "123"},
        {"fingerprint": "456"},
        {"fingerprint": "123"}  # 重复
    ]
    check_duplicate_fingerprints(data_with_duplicates)
    
    # 测试缺少fingerprint字段的情况
    data_missing_fingerprint = [
        {"fingerprint": "123"},
        {"other_field": "456"}  # 缺少fingerprint
    ]
    with pytest.raises(ValueError):
        check_duplicate_fingerprints(data_missing_fingerprint)


def test_get_paragraph_xml_fingerprint():
    """测试获取段落XML指纹的功能"""
    doc = Document()
    paragraph = doc.add_paragraph("测试段落")
    fingerprint = get_paragraph_xml_fingerprint(paragraph)
    assert isinstance(fingerprint, str)
    assert len(fingerprint) == 64  # SHA-256的长度


def test_load_yaml_with_merge():
    """测试加载YAML文件并处理合并语法的功能"""
    # 创建临时YAML文件
    with tempfile.NamedTemporaryFile(mode='w', suffix='.yaml', delete=False, encoding='utf-8') as f:
        f.write("""
        defaults: &defaults
          font: Arial
          size: 12
          
        heading:
          <<: *defaults
          bold: true
        """)
        temp_file_path = f.name
    
    try:
        config = load_yaml_with_merge(temp_file_path)
        assert 'defaults' in config
        assert 'heading' in config
        assert config['heading']['font'] == 'Arial'  # 测试合并功能
    finally:
        os.unlink(temp_file_path)


def test_ensure_is_directory():
    """测试确保路径是目录的功能"""
    # 测试有效的目录
    with tempfile.TemporaryDirectory() as temp_dir:
        ensure_is_directory(temp_dir)
    
    # 测试不存在的路径
    with tempfile.NamedTemporaryFile(delete=False) as f:
        temp_file_path = f.name
    try:
        with pytest.raises(ValueError):
            ensure_is_directory("nonexistent_directory_12345")
        # 测试文件路径（不是目录）
        with pytest.raises(ValueError):
            ensure_is_directory(temp_file_path)
    finally:
        os.unlink(temp_file_path)


def test_get_file_name():
    """测试获取文件名（不含扩展名）的功能"""
    assert get_file_name("test.docx") == "test"
    assert get_file_name("path/to/test.docx") == "test"
    assert get_file_name("test") == "test"


def test_remove_all_numbering():
    """测试移除所有编号的功能"""
    doc = Document()
    # 添加一些标题段落
    doc.add_heading("标题1", level=1)
    doc.add_heading("标题2", level=2)
    # 调用函数
    remove_all_numbering(doc)
    # 验证文档仍然可以操作
    assert len(doc.paragraphs) > 0
