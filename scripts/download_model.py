#! /usr/bin/env python
# @Time    : 2026/2/14 17:05
# @Author  : afish
# @File    : download_model.py
# !/usr/bin/env python3
"""
Download ONNX model from GitHub Releases with retry and progress bar.
Cross-platform compatible (Windows/Linux/macOS).
"""

import os
import sys
from pathlib import Path

import requests
from tqdm import tqdm

# ================== 配置 ==================
MODEL_URL = "https://github.com/AfishInLake/WordFormat/releases/download/bert_paragraph_classifier.onnx/bert_paragraph_classifier.onnx"
OUTPUT_DIR = Path("src/wordformat/data/model")
OUTPUT_FILE = OUTPUT_DIR / "bert_paragraph_classifier.onnx"
MAX_RETRIES = 3
CHUNK_SIZE = 8192  # 8KB chunks
# 模型版本（从环境变量读取，CI 中设置为 "modelv1.0"）
MODEL_VERSION = os.getenv("MODEL_VERSION", "modelv1.0")
# 版本标记文件路径（与模型同目录）
VERSION_FILE = OUTPUT_DIR / ".model_version"


# =========================================
def is_model_up_to_date() -> bool:
    """Check if the existing model matches the expected version."""
    if not OUTPUT_FILE.exists():
        return False
    if not VERSION_FILE.exists():
        return False
    try:
        with open(VERSION_FILE) as f:
            current_version = f.read().strip()
        return current_version == MODEL_VERSION
    except Exception:
        return False


def download_with_progress(url: str, output_path: Path, max_retries: int = 3):
    """Download file with progress bar and retry logic."""
    output_path.parent.mkdir(parents=True, exist_ok=True)

    for attempt in range(max_retries + 1):
        try:
            print(f"Attempt {attempt + 1}/{max_retries + 1} to download from: {url}")
            response = requests.get(url, stream=True, timeout=30)
            response.raise_for_status()  # Raise HTTPError for bad responses

            total_size = int(response.headers.get("content-length", 0))
            with (
                open(output_path, "wb") as f,
                tqdm(
                    desc="Downloading model",
                    total=total_size,
                    unit="B",
                    unit_scale=True,
                    unit_divisor=1024,
                    ascii=True,
                    leave=True,
                ) as pbar,
            ):
                for chunk in response.iter_content(chunk_size=CHUNK_SIZE):
                    if chunk:
                        f.write(chunk)
                        pbar.update(len(chunk))

            # Verify file exists and has content
            if output_path.exists() and output_path.stat().st_size > 0:
                print(f"Successfully downloaded model to: {output_path}")
                print(f"File size: {output_path.stat().st_size} bytes")
                return True
            else:
                raise RuntimeError("Downloaded file is empty or missing")

        except Exception as e:
            print(f"Download failed (attempt {attempt + 1}): {e}", file=sys.stderr)
            if attempt < max_retries:
                print("Retrying in 5 seconds...")
                import time

                time.sleep(5)
            else:
                print("All retries exhausted. Exiting.", file=sys.stderr)
                return False

    return False


if __name__ == "__main__":
    # 检查是否需要下载
    if is_model_up_to_date():
        print(f"Model is up-to-date (version: {MODEL_VERSION}). Skipping download.")
        sys.exit(0)

    print(f"Model missing or outdated. Expected version: {MODEL_VERSION}")

    # 执行下载
    success = download_with_progress(MODEL_URL, OUTPUT_FILE, MAX_RETRIES)

    if success:
        # 下载成功后写入版本文件
        VERSION_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(VERSION_FILE, "w") as f:
            f.write(MODEL_VERSION)
        print(f"Version file written: {VERSION_FILE}")

    sys.exit(0 if success else 1)
