#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 11:34
# @Author  : afish
# @File    : print_tree.py

from src.tree import print_tree
from src.word_structure.document_builder import DocumentBuilder

def main():
    root_node = DocumentBuilder.build_from_json('毕业设计说明书.json')
    print_tree(root_node)


if __name__ == '__main__':
    main()
