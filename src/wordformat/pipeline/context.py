#! /usr/bin/env python
# @Time    : 2026/7/7 12:30
# @Author  : afish
# @File    : Structure.py

from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol

from docx.document import Document as DocumentObject

from wordformat.config.models import NodeConfigRoot
from wordformat.rules.node import FormatNode


@dataclass
class FormatContext:
    json_path: str
    docx_path: str
    check: bool
    config_path: str
    save_dir: str = "/output"
    # 运行时对象（由各阶段填充）
    document: DocumentObject | None = None
    root_node: FormatNode = None
    config_model: NodeConfigRoot = field(default_factory=NodeConfigRoot)
    output_path: Path | str = ""


class PipelineStage(Protocol):
    """pipeline stage"""

    def process(self, ctx: FormatContext) -> FormatContext: ...
