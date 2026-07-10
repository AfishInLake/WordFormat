#! /usr/bin/env python
"""配置加载器。

LazyConfig 是普通类（非单例），每个实例独立加载和缓存。
模块级函数 init_config/get_config/clear_config 操作共享默认实例，方便生产代码。
测试可直接创建 LazyConfig(path) 避免全局状态污染。
"""

from __future__ import annotations

from loguru import logger

from wordformat.utils import load_yaml_with_merge


class LazyConfig:
    """懒加载配置管理器，每个实例独立。"""

    def __init__(self, config_path: str | None = None):
        self._config: dict | None = None
        self._config_path: str | None = config_path
        self._loaded: bool = False

    def init(self, config_path: str) -> None:
        self._config_path = config_path
        self._loaded = False
        logger.info(f"配置路径已设置: {config_path}")

    def load(self) -> dict:
        if not self._config_path:
            raise ConfigNotLoadedError("请先调用 init(config_path) 或传入配置路径")
        from wordformat.config.models import NodeConfigRoot

        raw = load_yaml_with_merge(self._config_path)
        self._config = NodeConfigRoot(**raw)
        self._loaded = True
        logger.info("配置加载完成")
        return self._config

    def get(self) -> dict:
        if self._config_path and not self._loaded:
            self.load()
        if self._config is None:
            raise ConfigNotLoadedError("配置加载失败，无法获取")
        return self._config

    @property
    def config_path(self) -> str | None:
        return self._config_path

    def clear(self):
        self._config = None
        self._config_path = None
        self._loaded = False


class ConfigNotLoadedError(Exception):
    pass


# 共享默认实例 —— 生产代码通过 init_config/get_config 使用
_default_config = LazyConfig()


def init_config(config_path: str):
    _default_config.init(config_path)


def get_config() -> dict:
    return _default_config.get()


def clear_config():
    _default_config.clear()
