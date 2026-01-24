#! /usr/bin/env python
# @Time    : 2026/1/18 13:22
# @Author  : afish
# @File    : main.py
from src.set_style import auto_format_thesis_document

auto_format_thesis_document(
    jsonpath="tmp/毕业设计说明书.json",
    docxpath="tmp/毕业设计说明书.docx",
    savepath="test/毕业设计说明书-修改版.docx",
    configpath="test/undergrad_thesis.yaml",
)
