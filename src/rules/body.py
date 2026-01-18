#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : body.py
from typing import List, Dict, Any

from src.rules.node import FormatNode
from src.rules.style import *


class BodyText(FormatNode):
    """正文节点"""
    NODE_TYPE = 'body_text'
    def check_format(self, doc) -> List[Dict[str, Any]]:
        """
        设置 正文 样式
        """
        for index, run in enumerate(self.paragraph.runs):
            CharacterStyle(
                font_name_en=FontName.SIM_SUN,
                font_name_cn=self.config.get(''),
                font_size=FontSize.XIAO_SI,
                font_color=FontColor.BLACK,
                bold=False,
                italic=False,
                underline=False
            ).apply_to(run)
        return []
