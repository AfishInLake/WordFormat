#! /usr/bin/env python
# @Time    : 2026/1/11 21:57
# @Author  : afish
# @File    : acknowledgement.py
from typing import Any

from src.rules.node import FormatNode


class Acknowledgements(FormatNode):
    """致谢节点"""

    NODE_TYPE = "acknowledgements"

    def check_format(self, doc) -> list[dict[str, Any]]:
        pass


class AcknowledgementsCN(FormatNode):
    """致谢节点"""

    NODE_TYPE = "acknowledgements.content"

    def check_format(self, doc) -> list[dict[str, Any]]:
        pass
