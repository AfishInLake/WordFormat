#! /usr/bin/env python
# @Time    : 2026/1/11 19:51
# @Author  : afish
# @File    : set_style.py
from __future__ import annotations

from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from wordformat.pipeline import PipelineStage

from wordformat.log_config import logger
from wordformat.pipeline.context import FormatContext
from wordformat.pipeline.stages import (
    DocumentSavingStage,
    FormattingExecutionStage,
    LoadConfigStage,
    LoadDocxStage,
    ParagraphAlignmentStage,
    StyleDefinitionFixStage,
    SummaryGenerationStage,
    TreeBuildingStage,
    TreeNormalizationStage,
)
from wordformat.pipeline.stages_md import (
    DocumentCreationStage,
    LoadMarkdownStage,
    MarkdownParseStage,
)


def auto_format_thesis_document(
    jsonpath: str | list,
    docxpath: str,
    configpath: Optional[str] = None,
    savepath: str = "output/",
    check=True,
    presets: Optional[list] = None,
    sync: bool = False,
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
        presets (Optional[list]): 预设名列表（presets/ 目录下的 .py 插件），按顺序加载，
                                 后加载的预设覆盖先加载的。
        sync (bool): 仅 apply 模式生效。编辑后的 JSON 与原 docx 段落数可不一致
                                 （增/删/改/重排），建树前先结构对齐同步进原文档（见 sync_apply）。

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
        ...     check=True,
        ...     presets=["tsinghua"]
        ... )
    """

    ctx = FormatContext(
        docx_path=docxpath,
        json_path=jsonpath,
        config_path=configpath,
        save_dir=savepath,
        check=check,
        sync=sync and not check,
    )
    _prepare_hooks(ctx, presets)
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
        DocumentSavingStage(),
    ]
    if ctx.sync:
        # af --sync：建树前先把编辑后 JSON 的增删改同步进原 docx 结构，
        # 使段落数与节点数恢复 1:1，后续 zip/对齐检查不变。
        from wordformat.pipeline.sync_apply import StructureSyncStage

        pipeline.insert(2, StructureSyncStage())
    for stage in pipeline:
        ctx = stage.process(ctx)
    return ctx.output_path


def _prepare_hooks(ctx, presets) -> None:
    """注册内置机制 handlers，并装配默认文档领域（thesis），再加载预设。

    顺序：内置机制 → 文档领域（可覆盖/补充内置逻辑） → 预设（最后注册，
    可 unregister 禁用前述任何回调）。
    """
    from wordformat.domains import load_domain
    from wordformat.handlers import register_builtin_handlers

    register_builtin_handlers()
    load_domain("thesis")
    _load_presets_into_context(ctx, presets)


def _load_presets_into_context(ctx, presets) -> None:
    """加载预设列表：注册 hook 回调，manifest 存入 ctx 供配置合并。"""
    if not presets:
        return
    from wordformat.preset_loader import load_presets

    ctx.preset_manifests = load_presets(list(presets))
    logger.info(f"已加载预设: {', '.join(presets)}")


def md_to_docx(
    md_path: str,
    config_path: str | None = None,
    save_dir: str = "output/",
    presets: Optional[list] = None,
):
    """将 Markdown 文件转换为格式化后的 .docx 文档。

    流程：
        1. 加载 YAML 格式配置（可选）；
        2. 读取并解析 Markdown 文件为扁平段落列表；
        3. 构建 FormatNode 文档树；
        4. 创建新的 Word 文档，逐节点生成段落并填入文本；
        5. 修正样式定义；
        6. 应用格式规则；
        7. 后处理（标题编号、引用超链接）；
        8. 保存输出 .docx 文件。

    Args:
        md_path: Markdown 源文件路径。
        config_path: YAML 格式规范配置文件路径，为 None 时使用内置默认配置。
        save_dir: 输出目录。
        presets: 预设名列表，按顺序加载（后加载覆盖先加载）。

    Returns:
        生成的 .docx 文件路径。
    """
    from pathlib import Path

    from wordformat.utils import ensure_directory_exists, get_file_name

    ctx = FormatContext(
        md_path=md_path,
        config_path=config_path or "",
        save_dir=save_dir,
        check=False,
    )
    _prepare_hooks(ctx, presets)

    pipeline: list[PipelineStage] = [
        LoadConfigStage(),
        LoadMarkdownStage(),
        MarkdownParseStage(),
        TreeBuildingStage(),
        DocumentCreationStage(),
        StyleDefinitionFixStage(),
        FormattingExecutionStage(),
    ]

    for stage in pipeline:
        ctx = stage.process(ctx)

    # 保存前 hook：编号/超链接等机制与预设统一在此执行（复用 before_document_save）
    from wordformat.hooks import hooks

    hooks.emit(
        "before_document_save",
        document=ctx.document,
        config=ctx.config_model,
        ctx=ctx,
        check=False,
    )

    # 重命名输出文件
    ensure_directory_exists(save_dir)
    filename = get_file_name(md_path)
    out_path = Path(save_dir) / f"{filename}--生成版.docx"
    ctx.document.save(str(out_path))
    logger.info(f"保存文件到 {out_path}")
    return str(out_path)


def _validate_jd_categories(json_path) -> None:
    """jd 入口预检：未知 category 直接报错，避免静默丢段。

    TreeBuilding/create_node 对未知类别只记 warning 并返回 None，jd 是从 JSON
    重建文档，静默丢条会造成内容缺失且难以发现。
    """
    from wordformat.structure.document_builder import DocumentBuilder
    from wordformat.structure.settings import CATEGORY_TO_CLASS

    items = DocumentBuilder.load_paragraphs(json_path)
    if not isinstance(items, list):
        raise ValueError("jd JSON 顶层必须是条目列表")
    invalid: list[tuple[int, str]] = []
    for idx, item in enumerate(items, start=1):
        if not isinstance(item, dict):
            invalid.append((idx, "<非对象条目>"))
            continue
        category = item.get("category")
        if not category:
            invalid.append((idx, "<缺少 category>"))
        elif category not in CATEGORY_TO_CLASS:
            invalid.append((idx, str(category)))
    if invalid:
        detail = "，".join(f"第 {i} 条 {c!r}" for i, c in invalid[:10])
        suffix = " 等" if len(invalid) > 10 else ""
        raise ValueError(
            f"jd JSON 含未知 category：{detail}{suffix}。"
            f"可用类别：{', '.join(sorted(CATEGORY_TO_CLASS))}"
        )


def json_to_docx(
    json_path: str,
    config_path: str | None = None,
    save_dir: str = "output/",
    presets: Optional[list] = None,
):
    """将编辑后的识别 JSON 转换为格式化后的 .docx 文档（jd 命令，重建模式）。

    与 md_to_docx 同构，仅把「读 md + 解析 md」换成 TreeBuildingStage 原生
    的 json_path 加载：以 JSON 列表为内容源从空白文档逐节点建段，天然支持
    段落的增/删/改/重排，标题/题注编号按新树序重算。

    边界：内容以 JSON 为准——原 docx 的图片二进制/表格内容不在 gj JSON 里，
    往返会丢（需保留原文档对象时用 af --sync）；新增图片/表格用
    figure_image（paragraph=图片路径）/ table_object（table_rows）条目。

    Args:
        json_path: gj 产出（或 agent 编辑后）的扁平 JSON 文件路径。
        config_path: YAML 格式规范配置文件路径，为 None 时使用内置默认配置。
        save_dir: 输出目录。
        presets: 预设名列表，按顺序加载（后加载覆盖先加载）。

    Returns:
        生成的 .docx 文件路径。
    """
    from pathlib import Path

    from wordformat.utils import ensure_directory_exists, get_file_name

    ctx = FormatContext(
        json_path=json_path,
        config_path=config_path or "",
        save_dir=save_dir,
        check=False,
    )
    # 先加载领域/预设：它们可能注册自定义 category，再据此做 jd 类别预检
    _prepare_hooks(ctx, presets)
    _validate_jd_categories(json_path)

    pipeline: list[PipelineStage] = [
        LoadConfigStage(),
        TreeBuildingStage(),
        TreeNormalizationStage(),
        DocumentCreationStage(),
        StyleDefinitionFixStage(),
        FormattingExecutionStage(),
    ]

    for stage in pipeline:
        ctx = stage.process(ctx)

    # 保存前 hook：编号/超链接/页眉页脚/页边距与预设统一在此执行
    from wordformat.hooks import hooks

    hooks.emit(
        "before_document_save",
        document=ctx.document,
        config=ctx.config_model,
        ctx=ctx,
        check=False,
    )

    ensure_directory_exists(save_dir)
    filename = get_file_name(json_path)
    out_path = Path(save_dir) / f"{filename}--生成版.docx"
    ctx.document.save(str(out_path))
    logger.info(f"保存文件到 {out_path}")
    return str(out_path)
