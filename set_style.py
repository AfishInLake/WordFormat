#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 19:51
# @Author  : afish
# @File    : set_style.py
from docx import Document
from loguru import logger

from src.rules import *
from src.tree import print_tree
from src.utils import get_paragraph_xml_fingerprint
from src.word_structure.document_builder import DocumentBuilder
from src.word_structure.utils import find_and_modify_first, promote_bodytext_in_subtrees_of_type


def apply_format_check_to_all_nodes(root_node, document):
    """
    递归遍历文档树中的所有节点，
    对每个具有 check_format 方法的节点执行该方法。

    :param root_node: 树的根节点（FormatNode 或其子类实例）
    :param document: docx文档的实例
    """

    def traverse(node):
        # 执行当前节点的格式检查（如果定义了）
        if hasattr(node, 'check_format'):
            try:
                # TODO: 应该使用not in list
                if node.value.get('category') != 'top':
                    node.check_format(document)
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
        if getattr(node, 'fingerprint', False):
            return node.fingerprint == get_paragraph_xml_fingerprint(paragraph)
        return False

    return find_and_modify_first(
        root=root_node,
        condition=condition
    )


def main():
    root_node = DocumentBuilder.build_from_json('毕业设计说明书.json')
    root_node.children = [
        node for node in root_node.children
        if node.value.get('category') != 'body_text'
    ]
    document = Document('毕业设计说明书.docx')
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
        References: ReferenceEntry
    }
    for key, value in subtress_dict.items():
        promote_bodytext_in_subtrees_of_type(
            root_node,
            parent_type=key,
            target_type=value
        )
    print_tree(root_node)
    apply_format_check_to_all_nodes(root_node, document)
    document.save('毕业设计说明书-修改版.docx')


if __name__ == '__main__':
    main()
