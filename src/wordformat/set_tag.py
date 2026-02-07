#! /usr/bin/env python
# @Time    : 2025/12/22 21:49
# @Author  : afish
# @File    : main.py
import json

from loguru import logger

from wordformat.base import DocxBase


def run(docx_path: str, json_save_path: str, configpath) -> list:
    dox = DocxBase(docx_path, configpath=configpath)
    a = dox.parse()

    with open(json_save_path, "w", encoding="utf-8") as f:
        json.dump(a, f, ensure_ascii=False, indent=4)
    logger.info(f"保存成功：{json_save_path}")
    return a


def set_tag_main(docx_path: str, json_save_path, configpath):
    return run(docx_path, json_save_path, configpath)
