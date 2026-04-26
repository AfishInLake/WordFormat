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
    # 打包为可执行文件时，使用可执行文件所在目录
    BASE_DIR = Path(sys.executable).parent
elif os.getenv("WORDFORMAT_BASE_DIR"):
    # 支持通过环境变量自定义工作目录
    BASE_DIR = Path(os.getenv("WORDFORMAT_BASE_DIR")).resolve()
else:
    # 开发模式：从 settings.py 位置向上查找项目根目录
    # （包含 pyproject.toml 或 .git 的目录）
    _candidate = Path(__file__).resolve().parent.parent.parent
    if (_candidate / "pyproject.toml").exists() or (_candidate / ".git").exists():
        BASE_DIR = _candidate
    else:
        BASE_DIR = Path.cwd()

HOST = os.getenv("HOST", "127.0.0.1")
PORT = int(os.getenv("PORT", "8000"))
SERVER_HOST = f"http://{HOST}:{PORT}"

# 从环境变量获取配置，如果不存在则使用默认值
API_KEY = os.getenv("WORDFORMAT_API_KEY", "")
MODEL = os.getenv("WORDFORMAT_MODEL", "")
MODEL_URL = os.getenv("WORDFORMAT_MODEL_URL", "")

BATCH_SIZE = os.getenv("BATCH_SIZE", 64)
ONNX_VERSION = "20260204"

VOIDNODELIST = [
    "top"
]
