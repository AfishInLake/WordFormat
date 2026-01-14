#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 19:42
# @Author  : afish
# @File    : document_builder.py
# src/document_builder.py
import json
from typing import List, Dict, Any

from src.rules.node import FormatNode
from src.word_structure.tree_builder import DocumentTreeBuilder


class DocumentBuilder:
    """对外统一接口：加载 JSON 并构建文档树"""

    @staticmethod
    def load_paragraphs(json_path: str) -> List[Dict[str, Any]]:
        with open(json_path, 'r', encoding='utf-8') as f:
            return json.load(f)

    @classmethod
    def build_from_json(cls, json_path: str) -> FormatNode:
        paragraphs = cls.load_paragraphs(json_path)
        builder = DocumentTreeBuilder()
        return builder.build_tree(paragraphs)
