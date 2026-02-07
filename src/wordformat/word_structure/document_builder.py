#! /usr/bin/env python
# @Time    : 2026/1/11 19:42
# @Author  : afish
# @File    : document_builder.py
# wordformat/document_builder.py
import json
from typing import Any

from loguru import logger

from wordformat.rules.node import FormatNode
from wordformat.utils import check_duplicate_fingerprints
from wordformat.word_structure.tree_builder import DocumentTreeBuilder


class DocumentBuilder:
    """对外统一接口：加载 JSON 并构建文档树"""

    @staticmethod
    def load_paragraphs(json_path: str | list) -> list[dict[str, Any]]:
        try:
            if isinstance(json_path, list):
                return json_path
            else:
                return json.loads(json_path)
        except Exception:
            logger.warning("加载json文件...")
            with open(json_path, encoding="utf-8") as f:
                return json.load(f)

    @classmethod
    def build_from_json(cls, json_path: str | list, config) -> FormatNode:
        paragraphs = cls.load_paragraphs(json_path)
        logger.debug(f"共有 {len(paragraphs)} 条语料")
        check_duplicate_fingerprints(paragraphs)  # 检查重复的指纹
        DocumentTreeBuilder.CONFIG = config
        builder = DocumentTreeBuilder()
        return builder.build_tree(paragraphs)
