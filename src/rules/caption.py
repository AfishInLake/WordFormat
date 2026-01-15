#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 19:43
# @Author  : afish
# @File    : caption.py
from typing import List, Dict, Any

from src.rules.node import FormatNode


class CaptionFigure(FormatNode):
    """题注-图片"""
    NODE_TYPE = 'figures'

    def check_format(self, doc) -> List[Dict[str, Any]]:
        pass


class CaptionTable(FormatNode):
    """题注-表格"""
    NODE_TYPE = 'tables'

    def check_format(self, doc) -> List[Dict[str, Any]]:
        pass
