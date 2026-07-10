#! /usr/bin/env python
# @Time    : 2026/1/11 19:51
# @Author  : afish
# @File    : set_style.py
from __future__ import annotations

from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from wordformat.pipeline import PipelineStage

from wordformat.pipeline.context import FormatContext
from wordformat.pipeline.stages import (
    DocumentSavingStage,
    FormattingExecutionStage,
    LoadConfigStage,
    LoadDocxStage,
    ParagraphAlignmentStage,
    PostProcessingStage,
    StyleDefinitionFixStage,
    SummaryGenerationStage,
    TreeBuildingStage,
    TreeNormalizationStage,
)


def auto_format_thesis_document(
    jsonpath: str | list,
    docxpath: str,
    configpath: Optional[str] = None,
    savepath: str = "output/",
    check=True,
):
    """自动对学位论文文档进行格式校验与批注。

    该函数根据结构化 JSON 描述和 YAML 格式配置，对指定的 Word 文档进行格式合规性检查，
    并在不符合规范的位置插入批注（comments）。主要用于学术论文（如本科/硕士/博士论文）
    的自动化格式审查。

    流程说明：
        1. 从 JSON 文件加载文档逻辑结构树；
        2. 加载 Word 文档，并将每个非空段落匹配到对应的结构节点；
        4. 对特定子树（如中英文摘要、参考文献）执行节点提升操作，确保内容节点正确挂载；
        5. 遍历所有结构节点，依据配置文件中的格式规则进行校验，并在文档中添加批注；
        6. 保存带批注的文档到指定路径。

    Args:
        check (bool): 用来控制是仅检查还是仅修改
        jsonpath (str): 文档逻辑结构的 JSON 文件路径 或 json 数据，描述各章节/段落的语义类型。
        docxpath (str): 待处理的原始 Word (.docx) 文档路径。
        savepath (str): 处理完成后带批注的文档保存路径。
        configpath (Optional[str]): 格式规范配置文件（YAML）路径，支持继承与合并。
                                 为 None 时使用内置默认配置。

    Side Effects:
        - 读取 jsonpath、docxpath 和 configpath 指定的文件；
        - 在 docx 文档中插入批注（不修改原文内容，仅添加审阅意见）；
        - 将结果文档写入 savepath。

    Example:
        >>> auto_format_thesis_document(
        ...     docxpath="draft.docx",
        ...     jsonpath="thesis_structure.json",
        ...     configpath="format_rules.yaml",
        ...     savepath="output/",
        ...     check=True
        ... )
    """

    ctx = FormatContext(
        docx_path=docxpath,
        json_path=jsonpath,
        config_path=configpath,
        save_dir=savepath,
        check=check,
    )
    # 2. 组装流水线
    pipeline: list[PipelineStage] = [
        LoadConfigStage(),
        LoadDocxStage(),
        TreeBuildingStage(),
        ParagraphAlignmentStage(),
        TreeNormalizationStage(),
        StyleDefinitionFixStage(),
        FormattingExecutionStage(),
        SummaryGenerationStage(),
        PostProcessingStage(),
        DocumentSavingStage(),
    ]
    for stage in pipeline:
        ctx = stage.process(ctx)
    return ctx.output_path
