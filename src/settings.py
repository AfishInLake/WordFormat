#! /usr/bin/env python
# @Time    : 2026/1/18 11:48
# @Author  : afish
# @File    : settings.py
import os

from dotenv import load_dotenv

load_dotenv()

# 从环境变量获取配置，如果不存在则使用默认值
API_KEY = os.getenv("WORDFORMAT_API_KEY", "")
MODEL = os.getenv("WORDFORMAT_MODEL", "")
MODEL_URL = os.getenv("WORDFORMAT_MODEL_URL", "")


CHARACTER_STYLE_CHECKS = {
    "bold": True,
    "italic": True,
    "underline": True,
    "font_size": True,
    "font_color": False,
    "font_name": False,
}
