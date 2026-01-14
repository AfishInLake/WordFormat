#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 19:43
# @Author  : afish
# @File    : caption.py
from typing import List, Dict, Any

from src.rules.node import FormatNode


class Caption(FormatNode):
    """题注"""

    def check_format(self, doc) -> List[Dict[str, Any]]:
        pass


class CaptionFigure(Caption):
    """题注-图片"""

    def check_format(self, doc) -> List[Dict[str, Any]]:
        pass


class CaptionTable(Caption):
    """题注-表格"""

    def check_format(self, doc) -> List[Dict[str, Any]]:
        pass
