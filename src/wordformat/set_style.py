#! /usr/bin/env python
# @Time    : 2026/1/11 19:51
# @Author  : afish
# @File    : set_style.py
from pathlib import Path

from docx import Document
from loguru import logger

from wordformat.config.config import get_config, init_config
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
from wordformat.utils import get_paragraph_xml_fingerprint, ensure_directory_exists
from wordformat.word_structure.document_builder import DocumentBuilder
from wordformat.word_structure.utils import (
    find_and_modify_first,
    promote_bodytext_in_subtrees_of_type,
)


def _fix_all_heading_style_definitions(document: Document, config_model):
    """在格式化开始前，统一修正所有 heading 样式定义。

    修正内容：
    1. 清除样式定义中的主题色引用（themeColor），替换为配置指定的颜色
    2. 确保样式定义存在（不存在则创建）

    这样做的原因：Word 的样式继承机制下，如果样式定义使用了主题色，
    段落内的 run 即使没有显式设置颜色，也会继承样式中的主题色。
    仅修改 run 级别的颜色不够，必须同时修正样式定义本身。
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from wordformat.style.style_enum import FontColor, _ensure_style_exists

    heading_config = getattr(config_model, "headings", None)
    if not heading_config:
        return

    # 收集所有 heading 级别的配置
    level_map = {
        "level_1": getattr(heading_config, "level_1", None),
        "level_2": getattr(heading_config, "level_2", None),
        "level_3": getattr(heading_config, "level_3", None),
    }

    for level_key, level_cfg in level_map.items():
        if level_cfg is None:
            continue

        style_name = getattr(level_cfg, "builtin_style_name", None)
        if not style_name:
            continue

        # 确保样式存在
        _ensure_style_exists(document, style_name)

        try:
            style = document.styles[style_name]
        except KeyError:
            logger.warning(f"样式 '{style_name}' 创建失败，跳过修正")
            continue

        style_element = style.element
        rPr = style_element.find(qn("w:rPr"))

        # 修正字体颜色：清除主题色，设置为配置指定的颜色
        if rPr is not None:
            color_elem = rPr.find(qn("w:color"))
            has_theme = False
            if color_elem is not None:
                has_theme = (
                    color_elem.get(qn("w:themeColor")) is not None
                    or color_elem.get(qn("w:themeTint")) is not None
                    or color_elem.get(qn("w:themeShade")) is not None
                )

            if has_theme:
                font_color_str = getattr(level_cfg, "font_color", "黑色") or "黑色"
                try:
                    fc = FontColor(font_color_str)
                    rgb_tuple = fc.rel_value
                    hex_color = f"{rgb_tuple[0]:02X}{rgb_tuple[1]:02X}{rgb_tuple[2]:02X}"

                    rPr.remove(color_elem)
                    new_color = OxmlElement("w:color")
                    new_color.set(qn("w:val"), hex_color)
                    rPr.append(new_color)

                    logger.info(f"已修正样式 '{style_name}' 主题色 → #{hex_color}")
                except Exception as e:
                    logger.warning(f"修正样式 '{style_name}' 颜色失败: {e}")


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

    def traverse(node):
        category = node.value.get("category", "") if isinstance(node.value, dict) else ""

        if hasattr(node, "check_format"):
            try:
                if category not in VOIDNODELIST:
                    node.load_config(config)
                    logger.debug(node)
                    if node.paragraph:
                        if check:
                            node.check_format(document)
                        else:
                            node.apply_format(document)
            except Exception as e:
                logger.warning(f"Node {node} not format, beacuse: {str(e)}")
                raise e

        # 目录和附录的子节点跳过格式化（top 节点本身跳过格式化，但子节点需要遍历）
        SKIP_CHILDREN_CATEGORIES = {"heading_mulu", "heading_fulu"}
        if category not in SKIP_CHILDREN_CATEGORIES:
            for child in node.children:
                traverse(child)

    traverse(root_node)


def xg(root_node, paragraph):
    """
    根据段落对象查找对应的节点
    :param root_node: 树的根节点
    :param paragraph: docx文档的段落对象
    :return: 找到的节点
    """

    def condition(node):
        if getattr(node, "fingerprint", False):
            return node.fingerprint == get_paragraph_xml_fingerprint(paragraph)
        return False

    return find_and_modify_first(root=root_node, condition=condition)


def auto_format_thesis_document(
        jsonpath: str | list,
        docxpath: str,
        configpath: str,
        savepath: str = "output/",
        check=True,
):
    """自动对学位论文文档进行格式校验与批注。

    该函数根据结构化 JSON 描述和 YAML 格式配置，对指定的 Word 文档进行格式合规性检查，
    并在不符合规范的位置插入批注（comments）。主要用于学术论文（如本科/硕士/博士论文）
    的自动化格式审查。

    流程说明：
        1. 从 JSON 文件加载文档逻辑结构树；
        2. 过滤掉类别为 'body_text' 的节点（保留标题、摘要、参考文献等结构节点）；
        3. 加载 Word 文档，并将每个非空段落匹配到对应的结构节点；
        4. 对特定子树（如中英文摘要、参考文献）执行节点提升操作，确保内容节点正确挂载；
        5. 遍历所有结构节点，依据配置文件中的格式规则进行校验，并在文档中添加批注；
        6. 保存带批注的文档到指定路径。

    Args:
        check (bool): 用来控制是仅检查还是仅修改
        jsonpath (str): 文档逻辑结构的 JSON 文件路径 或 json 数据，描述各章节/段落的语义类型。
        docxpath (str): 待处理的原始 Word (.docx) 文档路径。
        savepath (str): 处理完成后带批注的文档保存路径。
        configpath (str): 格式规范配置文件（YAML）路径，支持继承与合并。

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

    init_config(configpath)
    try:
        config_model = get_config()
        logger.info("配置文件验证通过")
    except Exception as e:
        logger.error(f"配置加载失败: {str(e)}")
        raise

    ensure_directory_exists(savepath)

    filename_without_ext = get_file_name(docxpath)
    root_node = DocumentBuilder.build_from_json(jsonpath, config=config_model)
    root_node.children = [
        node for node in root_node.children if node.value.get("category") != "body_text"
    ]
    document = Document(docxpath)

    if not check:
        style_list = []
        for style in document.styles:
            style_list.append(style.name)
        logger.info(f"可用的样式有：{style_list}")

    for paragraph in document.paragraphs:
        if not paragraph.text:
            continue
        node = xg(root_node, paragraph)
        if node:
            node.paragraph = paragraph

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
        _fix_all_heading_style_definitions(document, config_model)

    # 执行格式化
    apply_format_check_to_all_nodes(root_node, document, config_model, check)

    # 处理标题自动编号（仅在格式化模式下执行，检查模式不修改编号）
    if not check and hasattr(config_model, "numbering") and config_model.numbering.enabled:
        from wordformat.numbering import process_heading_numbering
        process_heading_numbering(root_node, document, config_model.numbering, config_model.headings)
    savepath = Path(savepath)
    if check:
        docx_path = str(savepath / f"{filename_without_ext}--标注版.docx")
    else:
        docx_path = str(savepath / f"{filename_without_ext}--修改版.docx")
    logger.info(f"保存文件到 {docx_path}")
    document.save(docx_path)
    return docx_path
