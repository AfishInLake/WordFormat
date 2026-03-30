#! /usr/bin/env python
# -*- coding: utf-8 -*-
import json
import os
import tempfile

from wordformat.set_tag import set_tag_main


def test_set_tag_main():
    """测试set_tag_main函数的功能"""
    # 创建临时docx文件
    from docx import Document
    doc = Document()
    doc.add_paragraph("测试段落")
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
        temp_docx_path = f.name
    doc.save(temp_docx_path)

    try:
        # 使用示例配置文件
        config_path = "example/undergrad_thesis.yaml"
        
        # 调用set_tag_main函数
        result = set_tag_main(temp_docx_path, config_path)
        
        # 验证结果是一个列表
        assert isinstance(result, list)

    finally:
        # 清理临时文件
        if os.path.exists(temp_docx_path):
            os.unlink(temp_docx_path)
