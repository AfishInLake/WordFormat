#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 11:34
# @Author  : afish
# @File    : print_tree.py
import json

from src.rules.abstract import AbstractTitleCN, Keywords
from src.rules.body import BodyText
from src.rules.heading import Heading
from src.rules.node import FormatNode
from src.rules.references import References, Acknowledgements, ReferenceEntry
from src.tree import Tree, Stack, print_tree
from src.word_structure.document_builder import DocumentBuilder


# if __name__ == '__main__':
#     # 创建一个树
#     tree = Tree({'category': 'top', "paragraph": "/"})
#     with open('论文修改测试.json', 'r', encoding='utf-8') as f:
#         tmp_list = json.load(f)
#     stack = Stack()
#     root_node = FormatNode(
#         paragraph={'category': 'top', 'paragraph': '[ROOT]'},
#         expected_rule={},
#         level=0
#     )
#     # 替换 tree.root
#     tree.root = root_node
#     stack.push(tree.root)
#     for index, item in enumerate(tmp_list):
#         category = item['category']
#         # ========================
#         # 1. 创建新节点 + 设置 level
#         # ========================
#         node: FormatNode
#         if category == 'abstract_chinese_title':
#             node = AbstractTitleCN(item, {}, level=1)
#         elif category == 'abstract_english_title':
#             node = AbstractTitleCN(item, {}, level=1)
#         elif category == 'keywords_chinese':
#             node = Keywords(item, {}, level=3)
#         elif category == 'keywords_english':
#             node = Keywords(item, {}, level=3)
#         elif category == 'heading_level_1':
#             node = Heading(item, {}, level=1)
#         elif category == 'heading_level_2':
#             node = Heading(item, {}, level=2)
#         elif category == 'heading_level_3':
#             node = Heading(item, {}, level=3)
#         elif category == 'heading_fulu':
#             node = Heading(item, {}, level=1)  # 或 4，按需
#         elif category == 'references_title':
#             node = References(item, {}, level=1)
#         elif category == 'acknowledgements_title':
#             node = Acknowledgements(item, {}, level=1)
#         elif category == 'reference_entry':
#             node = ReferenceEntry(item, {}, level=2)
#         elif category in ('body_text', 'caption_figure', 'caption_table', 'other'):
#             node = BodyText(item, {}, level=999)  # 叶子节点，level 很大，不影响标题
#         else:
#             print("Unknown category:", item)
#             continue
#
#         # ========================
#         # 2. 如果是标题类节点（需要层级管理）
#         # ========================
#         if category in (
#                 'abstract_chinese_title', 'abstract_english_title',
#                 'keywords_chinese', 'keywords_english',
#                 'heading_level_1', 'heading_level_2', 'heading_level_3',
#                 'heading_fulu', 'references_title', 'acknowledgements_title'
#         ):
#             # 👇 关键：弹出所有 level >= 当前 node.level 的节点
#             while not stack.is_empty():
#                 top = stack.peek()
#                 if hasattr(top, 'level') and top.level >= node.level:
#                     stack.pop()
#                 else:
#                     break
#
#             # 挂到当前父节点
#             parent = stack.peek()
#             parent.add_child_node(node)
#             stack.push(node)
#
#         # ========================
#         # 3. 如果是非标题节点（正文、题注等）
#         # ========================
#         else:
#             # 直接挂到最近的标题（栈顶）
#             if not stack.is_empty():
#                 parent = stack.peek()
#                 parent.add_child_node(node)
#             else:
#                 tree.root.add_child_node(node)  # 安全兜底
#
#     print_tree(tree.root)
def main():
    root_node = DocumentBuilder.build_from_json('论文修改测试.json')
    print_tree(root_node)
if __name__ == '__main__':
    main()