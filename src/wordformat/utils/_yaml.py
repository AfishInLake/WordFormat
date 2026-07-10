"""YAML 配置加载。"""

from typing import Any

import yaml


def load_yaml_with_merge(file_path: str) -> dict[str, Any]:
    """加载 YAML 文件，正确处理 <<: *anchor 合并语法。"""
    with open(file_path, encoding="utf-8") as f:
        return yaml.load(f, Loader=yaml.FullLoader)
