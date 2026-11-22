#! /usr/bin/env python
# @Time    : 2026/1/11 19:42
# @Author  : afish
# @File    : document_builder.py
# wordformat/document_builder.py
import json
from typing import Any

from loguru import logger

from wordformat.rules.node import FormatNode
from wordformat.structure.tree_builder import DocumentTreeBuilder


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
            # 非 JSON 文本（典型是文件路径）：转为读文件，这里是正常分支不是告警
            logger.debug(f"按路径加载 json 文件: {json_path}")
            with open(json_path, encoding="utf-8") as f:
                return json.load(f)

    @classmethod
    def build_from_json(
        cls, json_path: str | list, config, registry=None
    ) -> FormatNode:
        """加载 JSON 构建文档树；registry 为领域注册表实例（可空，默认论文域）。"""
        paragraphs = cls.load_paragraphs(json_path)
        logger.debug(f"共有 {len(paragraphs)} 条语料")
        DocumentTreeBuilder.CONFIG = config
        builder = DocumentTreeBuilder(registry=registry)
        builder._config = config
        return builder.build_tree(paragraphs)
