#! /usr/bin/env python
# @Time    : 2025/12/22 21:49
# @Author  : afish
# @File    : main.py
import asyncio
import json

from loguru import logger

from src.base import DocxBase

with open("src/system_prompt.txt", encoding="utf-8") as f:
    system_prompt = f.read()


async def run(docx_path: str, json_save_path: str):
    # TODO:需要尝试加载json文件，对比指纹，如果指纹一致，则不处理，否则处理以节省token
    dox = DocxBase(docx_path, system_prompt=system_prompt)
    a = await dox.parse()

    with open(json_save_path, "w", encoding="utf-8") as f:
        json.dump(a, f, ensure_ascii=False, indent=4)
    logger.info(f"保存成功：{json_save_path}")


def main(docx_path: str, json_save_path):
    asyncio.run(run(docx_path, json_save_path))
