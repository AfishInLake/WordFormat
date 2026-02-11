#! /usr/bin/env python
# -*- coding: utf-8 -*-
import json
import os
import tempfile
from pathlib import Path

from wordformat.set_style import auto_format_thesis_document


def test_auto_format_thesis_document_check_mode():
    """测试auto_format_thesis_document函数的检查模式"""
    # 创建临时docx文件
    from docx import Document
    doc = Document()
    doc.add_paragraph("测试段落")
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
        temp_docx_path = f.name
    doc.save(temp_docx_path)
    
    # 创建临时json文件
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as f:
        # 写入简单的JSON结构
        json_data = [
            {
                "fingerprint": "test_fingerprint_1",
                "category": "heading",
                "name": "标题1",
                "children": []
            }
        ]
        json.dump(json_data, f, ensure_ascii=False)
        temp_json_path = f.name
    
    # 创建临时输出目录
    with tempfile.TemporaryDirectory() as temp_output_dir:
        try:
            # 使用示例配置文件
            config_path = "example/undergrad_thesis.yaml"
            
            # 调用auto_format_thesis_document函数（检查模式）
            result_path = auto_format_thesis_document(
                jsonpath=temp_json_path,
                docxpath=temp_docx_path,
                configpath=config_path,
                savepath=temp_output_dir,
                check=True
            )
            
            # 验证结果文件被创建
            assert os.path.exists(result_path)
            assert "--标注版.docx" in result_path
            
        finally:
            # 清理临时文件
            if os.path.exists(temp_docx_path):
                os.unlink(temp_docx_path)
            if os.path.exists(temp_json_path):
                os.unlink(temp_json_path)


def test_auto_format_thesis_document_format_mode():
    """测试auto_format_thesis_document函数的格式化模式"""
    # 创建临时docx文件
    from docx import Document
    doc = Document()
    doc.add_paragraph("测试段落")
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
        temp_docx_path = f.name
    doc.save(temp_docx_path)
    
    # 创建临时json文件
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as f:
        # 写入简单的JSON结构
        json_data = [
            {
                "fingerprint": "test_fingerprint_1",
                "category": "heading",
                "name": "标题1",
                "children": []
            }
        ]
        json.dump(json_data, f, ensure_ascii=False)
        temp_json_path = f.name
    
    # 创建临时输出目录
    with tempfile.TemporaryDirectory() as temp_output_dir:
        try:
            # 使用示例配置文件
            config_path = "example/undergrad_thesis.yaml"
            
            # 调用auto_format_thesis_document函数（格式化模式）
            result_path = auto_format_thesis_document(
                jsonpath=temp_json_path,
                docxpath=temp_docx_path,
                configpath=config_path,
                savepath=temp_output_dir,
                check=False
            )
            
            # 验证结果文件被创建
            assert os.path.exists(result_path)
            assert "--修改版.docx" in result_path
            
        finally:
            # 清理临时文件
            if os.path.exists(temp_docx_path):
                os.unlink(temp_docx_path)
            if os.path.exists(temp_json_path):
                os.unlink(temp_json_path)
