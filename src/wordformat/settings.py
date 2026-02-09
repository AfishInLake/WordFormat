#! /usr/bin/env python
# @Time    : 2026/1/18 11:48
# @Author  : afish
# @File    : settings.py
import os
import sys
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

# 设置工作目录
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent.parent.parent

HOST = os.getenv("HOST", "127.0.0.1")
PORT = int(os.getenv("PORT", "8000"))
SERVER_HOST = f"http://{HOST}:{PORT}"

# 从环境变量获取配置，如果不存在则使用默认值
API_KEY = os.getenv("WORDFORMAT_API_KEY", "")
MODEL = os.getenv("WORDFORMAT_MODEL", "")
MODEL_URL = os.getenv("WORDFORMAT_MODEL_URL", "")

BATCH_SIZE = os.getenv("BATCH_SIZE", 64)
ONNX_VERSION = "20260204"
CHARACTER_STYLE_CHECKS = {
    "bold": True,
    "italic": True,
    "underline": True,
    "font_size": True,
    "font_color": False,
    "font_name": False,
}
