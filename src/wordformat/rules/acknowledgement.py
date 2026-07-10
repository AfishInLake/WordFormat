#! /usr/bin/env python
# @Time    : 2026/1/11 21:57
# @Author  : afish
# @File    : acknowledgement.py

from wordformat.config.models import (
    AcknowledgementsContentConfig,
    AcknowledgementsTitleConfig,
)
from wordformat.rules.node import FormatNode


class Acknowledgements(FormatNode[AcknowledgementsTitleConfig]):
    """致谢节点"""

    NODE_TYPE = "acknowledgements"
    NODE_LABEL = "致谢标题"
    CONFIG_MODEL = AcknowledgementsTitleConfig
    CONFIG_PATH = "acknowledgements.title"


class AcknowledgementsCN(FormatNode[AcknowledgementsContentConfig]):
    """致谢内容"""

    NODE_TYPE = "acknowledgements.content"
    NODE_LABEL = "致谢内容"
    CONFIG_MODEL = AcknowledgementsContentConfig
    CONFIG_PATH = "acknowledgements.content"
