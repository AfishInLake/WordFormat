#! /usr/bin/env python
# @Time    : 2026/1/18 13:22
# @Author  : afish
# @File    : main.py
from set_style import auto_format_thesis_document

# docxpath = "tmp/毕业设计说明书.docx"
# main(docxpath)

auto_format_thesis_document(
    jsonpath="tmp/毕业设计说明书.json",
    docxpath="tmp/毕业设计说明书.docx",
    configpath="test/undergrad_thesis.yaml",
)
