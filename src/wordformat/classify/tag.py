#! /usr/bin/env python
# @Time    : 2025/12/22 21:49
# @Author  : afish
# @File    : main.py

from wordformat.base import DocxBase


def set_tag_main(
    docx_path: str, configpath: str = None, domain: str = "thesis"
) -> list[dict]:
    """
    此入口用来生成段落文本标记，返回json数据

    :param docx_path: 传入的docx文件路径
    :param configpath: yaml配置文件路径（可选，当前未使用）
    :param domain: 文档领域（默认 thesis），决定分类后处理规则
    """
    dox = DocxBase(docx_path, configpath=configpath, domain=domain)
    a = dox.parse()
    return a
