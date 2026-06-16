#! /usr/bin/env python
# @Time    : 2025/12/22 21:49
# @Author  : afish
# @File    : main.py

from wordformat.base import DocxBase


def set_tag_main(
    docx_path: str, configpath: str = None, image_dir: str | None = None
) -> list[dict]:
    """
    此入口用来生成段落文本标记，返回json数据

    :param docx_path: 传入的docx文件路径
    :param configpath: yaml配置文件路径（可选，当前未使用）
    :param image_dir:  图片导出目录（可选，设置后图片 blob 保存为文件并在 JSON 中记录路径）
    """
    dox = DocxBase(docx_path, configpath=configpath)
    a = dox.parse(image_dir=image_dir)
    return a
