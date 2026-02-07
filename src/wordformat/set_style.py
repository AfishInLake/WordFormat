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
from wordformat.tree import print_tree
from wordformat.utils import get_paragraph_xml_fingerprint
from wordformat.word_structure.document_builder import DocumentBuilder
from wordformat.word_structure.utils import (
    find_and_modify_first,
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
    """

    def traverse(node):
        # 执行当前节点的格式检查（如果定义了）
        if hasattr(node, "check_format"):
            try:
                # HACK: 应该使用not in list
                if node.value.get("category") != "top":
                    node.load_config(config)
                    # TODO: 添加判断的参数
                    logger.debug(node)
                    if node.paragraph:
                        if check:
                            node.check_format(document)
                        else:
                            node.apply_format(document)
            except Exception as e:
                logger.warning(f"Node {node} not format, beacuse: {str(e)}")
                raise e

        # 递归处理所有子节点
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
        config_model = get_config()  # 首次调用：触发load()
        logger.info("配置文件验证通过")
    except Exception as e:
        logger.error(f"配置加载失败: {str(e)}")
        raise

    filename_without_ext = get_file_name(docxpath)
    root_node = DocumentBuilder.build_from_json(jsonpath, config=config_model)
    root_node.children = [
        node for node in root_node.children if node.value.get("category") != "body_text"
    ]
    document = Document(docxpath)
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
    print_tree(root_node)
    # FIXME: 临时解决
    """
    观察到不属于正文的内容被处理，需要剪枝
    word样式太多，需要考虑重置
    """
    apply_format_check_to_all_nodes(root_node, document, config_model, check)
    savepath = Path(savepath)
    savepath.mkdir(exist_ok=True)
    if check:
        docx_path = str(savepath / f"{filename_without_ext}--标注版.docx")
    else:
        docx_path = str(savepath / f"{filename_without_ext}--修改版.docx")
    logger.info(f"保存文件到 {docx_path}")
    document.save(docx_path)
    return docx_path
