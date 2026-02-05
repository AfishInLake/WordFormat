#! /usr/bin/env python
# @Time    : 2025/12/22 21:49
# @Author  : afish
# @File    : main.py
import json

from loguru import logger

from src.base import DocxBase

with open("src/system_prompt.txt", encoding="utf-8") as f:
    system_prompt = f.read()


def run(docx_path: str, json_save_path: str, configpath):
    dox = DocxBase(docx_path, system_prompt=system_prompt, configpath=configpath)
    a = dox.parse()

    with open(json_save_path, "w", encoding="utf-8") as f:
        json.dump(a, f, ensure_ascii=False, indent=4)
    logger.info(f"保存成功：{json_save_path}")


def set_tag_main(docx_path: str, json_save_path, configpath):
    run(docx_path, json_save_path, configpath)
