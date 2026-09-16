#! /usr/bin/env python3
"""旧模型下载脚本（已退役）。

int8 小模型（thesis_paragraph_classifier_int8.onnx）已直接入库，随包分发，
不再需要从 GitHub Release 下载。此脚本仅保留版本检查逻辑：
模型存在即退出 0，避免任何人误运行而拉取 99MB 旧 bert 模型入库。
"""

import os
import sys
from pathlib import Path

# ================== 配置 ==================
# 旧模型下载地址已废弃（bert_paragraph_classifier.onnx 99MB）
OUTPUT_DIR = Path("src/wordformat/data/model")
INT8_MODEL = OUTPUT_DIR / "thesis_paragraph_classifier_int8.onnx"

if __name__ == "__main__":
    if INT8_MODEL.exists() and INT8_MODEL.stat().st_size > 0:
        print(
            f"int8 模型已随包分发：{INT8_MODEL.name} "
            f"({INT8_MODEL.stat().st_size / 1024 / 1024:.1f} MB)，无需下载。"
        )
        sys.exit(0)
    print(
        "未找到随包分发的 int8 模型，请确认仓库包含 "
        "src/wordformat/data/model/thesis_paragraph_classifier_int8.onnx",
        file=sys.stderr,
    )
    sys.exit(1)
