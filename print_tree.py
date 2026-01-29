#! /usr/bin/env python
# @Time    : 2026/1/11 11:34
# @Author  : afish
# @File    : print_tree.py

from src.tree import print_tree
from src.word_structure.document_builder import DocumentBuilder

JSON_PATH = "output/论文.json"


def main():
    root_node = DocumentBuilder.build_from_json(JSON_PATH)
    print_tree(root_node)


if __name__ == "__main__":
    main()
