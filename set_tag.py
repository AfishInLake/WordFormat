#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2025/12/22 21:49
# @Author  : afish
# @File    : main.py
import json

from src.base import DocxBase, Style

with open('system_prompt.txt', 'r', encoding='utf-8') as f:
    system_prompt = f.read()


async def main():
    dox = DocxBase(
        r"论文修改测试.docx",
        system_prompt=system_prompt
    )
    a = await dox.parse(Style(r'G:\desktop\RosAi\RosAi\WordParse\undergrad_thesis.yaml'))
    with open('论文修改测试.json', 'w', encoding='utf-8') as f:
        json.dump(a, f, ensure_ascii=False, indent=4)



if __name__ == '__main__':
    import asyncio

    asyncio.run(main())
