#! /usr/bin/env python
"""配置加载器。"""

from __future__ import annotations

from wordformat.config.models import NodeConfigRoot
from wordformat.utils import load_yaml_with_merge

_config: NodeConfigRoot | None = None


def load_config(path: str) -> NodeConfigRoot:
    global _config
    raw = load_yaml_with_merge(path)
    _config = NodeConfigRoot(**raw)
    return _config


def get_config() -> NodeConfigRoot:
    if _config is None:
        raise RuntimeError("config not loaded, call load_config first")
    return _config


def init_config(path: str):
    """向后兼容别名。"""
    load_config(path)


def clear_config():
    """向后兼容别名。"""
    global _config
    _config = None


class ConfigNotLoadedError(RuntimeError):
    """向后兼容别名。"""

    pass


class LazyConfig:
    """向后兼容别名。"""

    def __init__(self, config_path: str | None = None):
        self._config_path = config_path

    def init(self, config_path: str):
        self._config_path = config_path

    def load(self):
        return load_config(self._config_path)

    def get(self):
        return get_config()

    def clear(self):
        clear_config()
