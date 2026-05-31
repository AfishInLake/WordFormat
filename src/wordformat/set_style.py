#! /usr/bin/env python
# @Time    : 2026/1/11 19:51
# @Author  : afish
# @File    : set_style.py
from pathlib import Path
from typing import Optional

from docx import Document
from loguru import logger

from wordformat.config.config import get_config, init_config
from wordformat.config.datamodel import NodeConfigRoot

# ---- 从 orchestration/ 导入（原 set_style.py 中的函数已拆分到各子模块） ----
from wordformat.orchestration.binding import (  # noqa: E402, F401
    _sync_deletions,
    _sync_insertions,
)
from wordformat.orchestration.binding import (
    bind_and_sync as _bind_and_sync,
)
from wordformat.orchestration.style_fixer import (  # noqa: E402, F401
    collect_all_style_configs as _collect_all_style_configs,
)
from wordformat.orchestration.style_fixer import (
    fix_all_style_definitions as _fix_all_style_definitions,
)
from wordformat.orchestration.style_fixer import (
    fix_style_paragraph_properties as _fix_style_paragraph_properties,
)
from wordformat.orchestration.style_fixer import (
    fix_style_run_properties as _fix_style_run_properties,
)
from wordformat.orchestration.table_formatter import (  # noqa: E402
    format_table_content,
)
from wordformat.rules import (
    AbstractContentCN,
    AbstractContentEN,
    AbstractTitleCN,
    AbstractTitleEN,
    FormatNode,
    ReferenceEntry,
    References,
)
from wordformat.settings import VOIDNODELIST
from wordformat.utils import ensure_directory_exists
from wordformat.word_structure.document_builder import DocumentBuilder
from wordformat.word_structure.utils import (
    promote_bodytext_in_subtrees_of_type,
)


def apply_format_check_to_all_nodes(
    root_node: FormatNode, document, config, check=True
):
    """
    递归遍历文档树中的所有节点，
    对每个具有 check_format 方法的节点执行该方法。

    :param root_node: 树的根节点（FormatNode 或其子类实例）
    :param document: docx文档的实例
    :param config: 配置文件
    :param check: 用来控制是仅检查还是仅修改
    """
    from wordformat.rules.caption import CaptionFigure, CaptionTable
    from wordformat.utils import parse_caption_text

    chapter_index: int = 0
    figure_counter: dict[int, int] = {}
    table_counter: dict[int, int] = {}

    def traverse(node, parent_category="", current_chapter: int = 0):
        nonlocal chapter_index

        category = (
            node.value.get("category", "") if isinstance(node.value, dict) else ""
        )

        # 遇到一级标题时递增章节号
        if category == "heading_level_1":
            chapter_index += 1
            current_chapter = chapter_index

        if hasattr(node, "check_format"):
            try:
                # top 节点直接关联的 body_text 不参与格式化（如封面页、原创性声明等）
                # 但间接关联的 body_text（作为 heading 子节点）正常格式化
                is_top_direct_body_text = (
                    parent_category == "top" and category == "body_text"
                )
                if category not in VOIDNODELIST and not is_top_direct_body_text:
                    node.load_config(config)

                    # 对题注节点注入章节号和顺序号
                    if isinstance(node, (CaptionFigure, CaptionTable)):
                        # 检查是否为续表/续图：保留原标题注编号，不递增计数器
                        text = node.paragraph.text.strip() if node.paragraph else ""
                        parsed = parse_caption_text(text)
                        if (
                            parsed
                            and parsed.get("is_continued")
                            and parsed.get("chapter_num") is not None
                            and parsed.get("number_num") is not None
                        ):
                            chapter = parsed["chapter_num"]
                            seq = parsed["number_num"]
                        else:
                            chapter = current_chapter if current_chapter > 0 else 0
                            if isinstance(node, CaptionFigure):
                                counter = figure_counter
                            else:
                                counter = table_counter
                            counter[chapter] = counter.get(chapter, 0) + 1
                            seq = counter[chapter]
                        node.chapter_number = chapter
                        node.sequence_number = seq
                        if hasattr(config, "numbering"):
                            node._numbering_cfg = config.numbering.captions

                    if node.paragraph:
                        # 先执行内容替换（check/format 两种模式均执行）
                        node.apply_replace(document)
                        if check:
                            node.check_format(document)
                        else:
                            node.apply_format(document)
            except Exception as e:
                logger.warning(f"Node {node} not format, because: {str(e)}")
                raise e

        # 目录和附录的子节点跳过格式化（top 节点本身跳过格式化，但子节点需要遍历）
        SKIP_CHILDREN_CATEGORIES = {"heading_mulu", "heading_fulu"}
        if category not in SKIP_CHILDREN_CATEGORIES:
            for child in node.children:
                traverse(
                    child, parent_category=category, current_chapter=current_chapter
                )

    traverse(root_node)


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
        ...     "thesis_structure.json",
        ...     "draft.docx",
        ...     "formatted_with_comments.docx",
        ...     "format_rules.yaml"
        ... )
    """
    from wordformat.utils import get_file_name

    if configpath:
        init_config(configpath)
        try:
            config_model = get_config()
            logger.info("配置文件验证通过")
        except Exception as e:
            logger.error(f"配置加载失败: {str(e)}")
            raise
    else:
        config_model = NodeConfigRoot()
        logger.info("未提供配置文件，使用默认配置")

    ensure_directory_exists(savepath)

    filename_without_ext = get_file_name(docxpath)
    root_node = DocumentBuilder.build_from_json(jsonpath, config=config_model)
    # 注意：不再过滤 body_text 节点，body_text 也需要格式化（首行缩进、字体等）
    document = Document(docxpath)

    if not check:
        style_list = []
        for style in document.styles:
            style_list.append(style.name)
        logger.info(f"可用的样式有：{style_list}")

    _bind_and_sync(root_node, document, check)

    # 替换摘要节点
    subtress_dict = {
        AbstractTitleCN: AbstractContentCN,
        AbstractTitleEN: AbstractContentEN,
        References: ReferenceEntry,
    }
    for key, value in subtress_dict.items():
        promote_bodytext_in_subtrees_of_type(
            root_node, parent_type=key, target_type=value
        )
    # 执行格式化前，先统一修正样式定义（清除主题色、设置字体等）
    if not check:
        _fix_all_style_definitions(document, config_model)

    # 执行格式化
    apply_format_check_to_all_nodes(root_node, document, config_model, check)

    # 表格内容格式化
    format_table_content(document, config_model, check)

    # 处理标题自动编号（仅在格式化模式下执行，检查模式不修改编号）
    if (
        not check
        and hasattr(config_model, "numbering")
        and config_model.numbering.enabled
    ):
        from wordformat.numbering import process_heading_numbering

        process_heading_numbering(
            root_node, document, config_model.numbering, config_model.headings
        )

    # 创建引用超链接（仅在格式化模式下执行）
    if not check:
        from wordformat.hyperlinks import create_citation_hyperlinks

        create_citation_hyperlinks(root_node, document)

    savepath = Path(savepath)
    if check:
        docx_path = str(savepath / f"{filename_without_ext}--标注版.docx")
    else:
        docx_path = str(savepath / f"{filename_without_ext}--修改版.docx")
    logger.info(f"保存文件到 {docx_path}")
    document.save(docx_path)
    return docx_path
