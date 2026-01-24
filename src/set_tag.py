#! /usr/bin/env python
# @Time    : 2025/12/22 21:49
# @Author  : afish
# @File    : main.py
import asyncio
import json
from pathlib import Path

from src.base import DocxBase
from utils import get_file_name

with open("src/system_prompt.txt", encoding="utf-8") as f:
    system_prompt = f.read()


async def run(docx_path: str, json_save_path: str = "tmp/"):
    # TODO:需要尝试加载json文件，对比指纹，如果指纹一致，则不处理，否则处理以节省token
    savepath = Path(json_save_path)
    savepath.mkdir(exist_ok=True)
    file_name = get_file_name(docx_path)
    dox = DocxBase(docx_path, system_prompt=system_prompt)
    a = await dox.parse()

    with open(f"{savepath}/{file_name}.json", "w", encoding="utf-8") as f:
        json.dump(a, f, ensure_ascii=False, indent=4)


def main(docx_path: str, json_save_path: str = "tmp/"):
    asyncio.run(run(docx_path, json_save_path))
