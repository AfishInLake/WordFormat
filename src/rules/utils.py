#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/12 15:18
# @Author  : afish
# @File    : utils.py

import re
from typing import Optional, Tuple
from docx.text.run import Run


def split_run_by_regex(
        run: Run,
        pattern: str,
        flags: int = 0
) -> Optional[Tuple[Run, Optional[Run]]]:
    """
    在给定的 Run 中，用正则表达式匹配前缀，并将其拆分为两个 Run。

    参数:
        run (Run): 要拆分的原始 Run。
        pattern (str): 正则表达式，用于匹配“前缀”部分（如 "关键词："）。
                        必须使用捕获组 `( )` 来明确界定前缀。
        flags (int): re 模块的标志位，如 re.IGNORECASE。

    返回:
        Tuple[Run, Optional[Run]] | None:
            - 如果匹配成功：
                - 第一个元素：包含前缀的新 Run（原 run 被修改为此内容）
                - 第二个元素：包含剩余文本的新 Run（如果非空），否则为 None
            - 如果未匹配：返回 None

    示例:
        run.text = "关键词：离心泵"
        prefix_run, suffix_run = split_run_by_regex(run, r'(关[^a-zA-Z0-9\u4e00-\u9fff]*键[^a-zA-Z0-9\u4e00-\u9fff]*词[^a-zA-Z0-9\u4e00-\u9fff]*：?)')
        # prefix_run.text == "关键词："
        # suffix_run.text == "离心泵"
    """
    text = run.text
    match = re.search(pattern, text, flags)
    if not match:
        return None

    # 确保正则中有捕获组
    if len(match.groups()) == 0:
        raise ValueError("正则表达式必须包含捕获组 ( ) 来界定前缀")

    prefix = match.group(1)  # 使用第一个捕获组作为前缀
    prefix_end = match.end(1)
    suffix = text[prefix_end:]

    # 修改原 run 为前缀
    run.text = prefix

    # 创建后缀 run（仅当有内容时）
    suffix_run = None
    if suffix.strip():  # 非空白才创建
        suffix_run = run._parent.add_run(suffix)

    return run, suffix_run


def _is_chinese_font_name(name: str) -> bool:
    """判断字体名称是否包含中文字符"""
    return bool(re.search(r'[\u4e00-\u9fff]', name))
