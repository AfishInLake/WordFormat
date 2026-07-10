"""文件系统工具。"""

import os

from loguru import logger


def get_file_name(file_name: str) -> str:
    basename = os.path.basename(file_name)
    filename_without_ext, _ = os.path.splitext(basename)
    return filename_without_ext


def ensure_is_directory(path):
    """检查 path 是否为已存在的文件夹，否则抛出 ValueError。"""
    if not os.path.exists(path):
        raise ValueError(f"路径不存在: '{path}'")
    if not os.path.isdir(path):
        raise ValueError(f"路径不是一个文件夹（它可能是一个文件）: '{path}'")


def ensure_directory_exists(path):
    """确保路径存在，不存在则递归创建，是文件则抛出 ValueError。"""
    if os.path.exists(path):
        if not os.path.isdir(path):
            raise ValueError(f"路径已存在但不是文件夹：'{path}'")
    else:
        os.makedirs(path, exist_ok=True)
        logger.info(f"已创建文件夹：'{path}'")
