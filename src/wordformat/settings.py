#! /usr/bin/env python
# @Time    : 2026/1/18 11:48
# @Author  : afish
# @File    : settings.py
import os
import sys
from pathlib import Path

from dotenv import load_dotenv

from wordformat._version import __version__ as VERSION  # noqa: F401

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

# 单次最多处理的段落数。基准（int8，2/1 线程）：batch 32 处于吞吐最高梯队，
# 峰值内存增量≈0；64 多占 25MB、128 多占 131MB 且吞吐不升反降。
BATCH_SIZE = int(os.getenv("BATCH_SIZE", "32"))
ONNX_VERSION = "20260204"

# ONNX 推理线程数：2 核服务器上用满（intra=2），inter 保持 1。
# 机器与其他服务争抢 CPU 时，环境变量 ONNX_INTRA_OP_THREADS=1 可降档省一半 CPU。
ONNX_INTRA_OP_THREADS = int(os.getenv("ONNX_INTRA_OP_THREADS", "2"))
ONNX_INTER_OP_THREADS = int(os.getenv("ONNX_INTER_OP_THREADS", "1"))

VOIDNODELIST = [
    "top",
    "heading_mulu",
    "heading_fulu",
    "other",  # 封面、声明页等无需格式化的内容
    "document_title",  # 文档标题（后处理扩展标签，不参与格式化）
    "footer",  # 页脚/AI 生成声明（后处理扩展标签，不参与格式化）
]
