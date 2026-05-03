#! /usr/bin/env python
# @Time    : 2025/12/22 21:49
# @Author  : afish
# @File    : main.py
import json
import os
from loguru import logger

from wordformat.base import DocxBase



def set_tag_main(docx_path: str, configpath: str = None) -> list[dict]:
    """
    此入口用来生成段落文本标记，返回json数据

    :param docx_path: 传入的docx文件路径
    :param configpath: yaml配置文件路径（可选，当前未使用）
    """
    dox = DocxBase(docx_path, configpath=configpath)
    a = dox.parse()
    return a