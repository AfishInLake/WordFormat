#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : heading.py
from typing import List, Dict, Any

from src.rules.node import FormatNode
from src.rules.style import *


class HeadingLevel1Node(FormatNode):
    """一级标题节点"""
    NODE_TYPE = 'headings.level_1'

    def check_format(self, doc) -> List[Dict[str, Any]]:

        return []


class HeadingLevel2Node(FormatNode):
    """二级标题节点"""
    NODE_TYPE = 'headings.level_2'
    def check_format(self, doc) -> List[Dict[str, Any]]:

        return []


class HeadingLevel3Node(FormatNode):
    """三级标题节点"""
    NODE_TYPE = 'headings.level_3'
    def check_format(self, doc) -> List[Dict[str, Any]]:

        return []
