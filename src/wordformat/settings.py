#! /usr/bin/env python
# @Time    : 2026/1/18 11:48
# @Author  : afish
# @File    : settings.py
import os
import sys
from pathlib import Path

from dotenv import load_dotenv
from loguru import logger

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
    "toc_entry",  # 目录条目（规则识别，不参与格式化）
]

# 结构类别策略默认值（可由 YAML `structure` 段逐类别覆盖）：
# - void：不参与格式化的类别（按类别名单驱动，不是某个具体业务开关）
# - skip_children：子树不遍历的类别
DEFAULT_SKIP_CHILDREN_CATEGORIES = ["heading_mulu", "heading_fulu", "other"]


def resolve_structure_categories(config) -> dict[str, list[str]]:
    """从配置 `structure` 段解析结构类别名单（缺省 = 内置默认值）。

    YAML 示例::

        structure:
          void_categories: [top, heading_mulu, heading_fulu, other, ...]
          skip_children_categories: [heading_mulu, heading_fulu, other]

    想让某个类别参与格式化（如附录）时，把它从 void_categories 与
    skip_children_categories 中移除即可，不绑定任何业务语义。
    名单是权威的：显式写空列表（如 ``void_categories: []``）表示不排除
    任何类别，不会回退到内置默认值。唯一例外是 root 直属的 ``body_text``
    （封面/声明类无标题归属内容），它由遍历逻辑单独隐式排除。
    """
    structure = {}
    if isinstance(config, dict):
        structure = config.get("structure") or {}
    if not isinstance(structure, dict):
        structure = {}

    if "terminal_categories" in structure:
        logger.warning(
            "structure.terminal_categories 已移除，该配置将被忽略"
            "（目录/附录的子树隔离由 LEVEL_MAP 层级决定）"
        )

    def _pick(key: str, default: list[str]) -> list[str]:
        value = structure.get(key)
        if isinstance(value, (list, tuple)):
            return [str(v) for v in value]
        return list(default)

    return {
        "void": _pick("void_categories", VOIDNODELIST),
        "skip_children": _pick(
            "skip_children_categories", DEFAULT_SKIP_CHILDREN_CATEGORIES
        ),
    }
