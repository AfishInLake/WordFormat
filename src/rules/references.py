#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : references.py
from typing import Dict, List, Any

from src.rules.node import FormatNode


class References(FormatNode):
    """参考文献节点"""
    NODE_TYPE = 'references'

    def check_format(self, doc) -> List[Dict[str, Any]]:
        pass


class ReferenceEntry(FormatNode):
    """参考文献条目节点"""

    def check_format(self, doc) -> List[Dict[str, Any]]:
        pass
