#! /usr/bin/env python
# -*- coding: utf-8 -*-
import json
import os
import tempfile

from wordformat.set_tag import set_tag_main, run


def test_run_function():
    """测试run函数的功能"""
    # 创建临时docx文件
    from docx import Document
    doc = Document()
    doc.add_paragraph("测试段落")
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
        temp_docx_path = f.name
    doc.save(temp_docx_path)
    
    # 创建临时json保存路径
    with tempfile.NamedTemporaryFile(suffix='.json', delete=False) as f:
        temp_json_path = f.name
    
    try:
        # 使用示例配置文件
        config_path = "example/undergrad_thesis.yaml"
        
        # 调用run函数
        result = run(temp_docx_path, temp_json_path, config_path)
        
        # 验证结果是一个列表
        assert isinstance(result, list)
        
        # 验证JSON文件被创建并包含数据
        assert os.path.exists(temp_json_path)
        with open(temp_json_path, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
        assert isinstance(json_data, list)
        
    finally:
        # 清理临时文件
        if os.path.exists(temp_docx_path):
            os.unlink(temp_docx_path)
        if os.path.exists(temp_json_path):
            os.unlink(temp_json_path)


def test_set_tag_main():
    """测试set_tag_main函数的功能"""
    # 创建临时docx文件
    from docx import Document
    doc = Document()
    doc.add_paragraph("测试段落")
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
        temp_docx_path = f.name
    doc.save(temp_docx_path)
    
    # 创建临时json保存路径
    with tempfile.NamedTemporaryFile(suffix='.json', delete=False) as f:
        temp_json_path = f.name
    
    try:
        # 使用示例配置文件
        config_path = "example/undergrad_thesis.yaml"
        
        # 调用set_tag_main函数
        result = set_tag_main(temp_docx_path, temp_json_path, config_path)
        
        # 验证结果是一个列表
        assert isinstance(result, list)
        
        # 验证JSON文件被创建并包含数据
        assert os.path.exists(temp_json_path)
        with open(temp_json_path, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
        assert isinstance(json_data, list)
        
    finally:
        # 清理临时文件
        if os.path.exists(temp_docx_path):
            os.unlink(temp_docx_path)
        if os.path.exists(temp_json_path):
            os.unlink(temp_json_path)
