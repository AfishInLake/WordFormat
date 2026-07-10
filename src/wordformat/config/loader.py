#! /usr/bin/env python
# @Time    : 2026/1/26 10:25
# @Author  : afish
# @File    : config.py
from typing import Optional

from loguru import logger

from wordformat.utils import load_yaml_with_merge


class LazyConfig:
    """懒加载配置管理器（单例）"""

    _instance: Optional["LazyConfig"] = None
    _config: Optional[dict] = None
    _config_path: Optional[str] = None
    _loaded: bool = False

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def init(self, config_path: str) -> None:
        self._config_path = config_path
        self._loaded = False
        logger.info(f"懒加载配置已初始化路径: {config_path}")

    def load(self) -> dict:
        if not self._config_path:
            raise ConfigNotLoadedError("请先调用 init(config_path) 初始化配置路径")
        try:
            self._config = load_yaml_with_merge(self._config_path)
            self._loaded = True
            logger.info("配置加载完成")
            return self._config
        except Exception as e:
            logger.error(f"配置文件加载失败: {str(e)}")
            raise

    def get(self) -> dict:
        if not self._loaded:
            self.load()
        if self._config is None:
            raise ConfigNotLoadedError("配置加载失败，无法获取")
        return self._config

    @property
    def config_path(self) -> Optional[str]:
        return self._config_path

    def clear(self):
        self._config = None
        self._config_path = None
        self._loaded = False


class ConfigNotLoadedError(Exception):
    pass


lazy_config = LazyConfig()


def init_config(config_path: str):
    lazy_config.init(config_path)


def get_config() -> dict:
    return lazy_config.get()


def clear_config():
    lazy_config.clear()
