#! /usr/bin/env python
# @Time    : 2025/12/22 21:49
# @Author  : afish
# @File    : main.py
import json

from src.base import DocxBase

with open("../system_prompt.txt", encoding="utf-8") as f:
    system_prompt = f.read()

path = "../tmp/毕业设计说明书.docx"


async def main():
    # TODO:需要尝试加载json文件，对比指纹，如果指纹一致，则不处理，否则处理以节省token
    dox = DocxBase(path, system_prompt=system_prompt)
    a = await dox.parse()
    with open("../tmp/毕业设计说明书.json", "w", encoding="utf-8") as f:
        json.dump(a, f, ensure_ascii=False, indent=4)


if __name__ == "__main__":
    import asyncio

    asyncio.run(main())
