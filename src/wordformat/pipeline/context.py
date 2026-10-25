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
    # af --sync：建树前先按编辑后的 JSON 对原 docx 做结构同步（增删改段落）
    sync: bool = False
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
    # 预设（preset）系统：manifest 列表，加载顺序与 CLI --preset 一致
    preset_manifests: list = field(default_factory=list)
    # 结构类别策略（由 LoadConfigStage 从 YAML structure 段解析）
    # None = 未配置（用内置默认）；[] = 显式声明空（不排除任何类别）
    void_categories: list | None = None
    skip_children_categories: list | None = None


class PipelineStage(Protocol):
    """pipeline stage"""

    def process(self, ctx: FormatContext) -> FormatContext: ...
