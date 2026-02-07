#! /usr/bin/env python
# @Time    : 2026/1/11 11:34
# @Author  : afish
# @File    : print_tree.py
from wordformat.config.config import get_config, init_config
from wordformat.tree import print_tree
from wordformat.word_structure.document_builder import DocumentBuilder

JSON_PATH = "AAAA/1.json"
YAML_PATH = "example/undergrad_thesis.yaml"


def main():
    init_config(YAML_PATH)
    config_model = get_config()
    root_node = DocumentBuilder.build_from_json(JSON_PATH, config=config_model)
    print_tree(root_node)


if __name__ == "__main__":
    main()
