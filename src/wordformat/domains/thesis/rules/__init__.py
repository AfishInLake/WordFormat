#! /usr/bin/env python
"""论文域格式规则节点（自 wordformat/rules 迁入，缝 A/B 归属收编）。

论文专属的 FormatNode 子类集中在论文领域包内，由 @register 注册到默认
领域注册表（structure/registry.default_registry）。框架核心的 rules/ 只保留
通用结构节点（正文/标题/题注/图/表/公式）与 FormatNode 机制。

新领域（公文、计划书……）在 domains/<name>/rules/ 下定义自己的节点，
DEFAULTS 与 NODE_TYPE 按各自文体声明即可，不进入本包。
"""

from .abstract import (
    AbstractContentCN,
    AbstractContentEN,
    AbstractTitleCN,
    AbstractTitleContentCN,
    AbstractTitleContentEN,
    AbstractTitleEN,
)
from .acknowledgement import Acknowledgements, AcknowledgementsCN
from .keywords import KeywordsCN, KeywordsEN
from .references import ReferenceEntry, References

__all__ = [
    "AbstractTitleCN",
    "AbstractTitleContentCN",
    "AbstractTitleEN",
    "AbstractTitleContentEN",
    "AbstractContentCN",
    "AbstractContentEN",
    "KeywordsEN",
    "KeywordsCN",
    "Acknowledgements",
    "AcknowledgementsCN",
    "References",
    "ReferenceEntry",
]
