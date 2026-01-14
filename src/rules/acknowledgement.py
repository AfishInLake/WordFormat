#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 21:57
# @Author  : afish
# @File    : acknowledgement.py
from typing import List, Dict, Any

from src.rules.node import FormatNode


class Acknowledgements(FormatNode):
    """致谢节点"""

    def check_format(self, doc) -> List[Dict[str, Any]]:
        pass
