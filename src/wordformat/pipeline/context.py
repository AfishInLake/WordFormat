#! /usr/bin/env python
# @Time    : 2026/7/7 12:30
# @Author  : afish
# @File    : Structure.py

from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol

from docx.document import Document as DocumentObject

from wordformat.rules.node import FormatNode


@dataclass
class FormatContext:
    json_path: str = ""
    docx_path: str = ""
    check: bool = False
    config_path: str = ""
    save_dir: str = "/output"
    # MD → Docx 专用
    md_path: str = ""
    md_text: str = ""
    paragraphs: list = field(default_factory=list)
    # 运行时对象（由各阶段填充）
    document: DocumentObject | None = None
    root_node: FormatNode = None
    config_model: dict = field(default_factory=dict)
    output_path: Path | str = ""


class PipelineStage(Protocol):
    """pipeline stage"""

    def process(self, ctx: FormatContext) -> FormatContext: ...
