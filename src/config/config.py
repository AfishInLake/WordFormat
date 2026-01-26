#! /usr/bin/env python
# @Time    : 2026/1/26 10:25
# @Author  : afish
# @File    : config.py
# src/config.py
from typing import Optional

from loguru import logger
from pydantic import ValidationError

from src.utils import load_yaml_with_merge

from .datamodel import NodeConfigRoot

# 单例配置存储（私有变量，仅通过接口访问）
_global_config: Optional[NodeConfigRoot] = None


class LazyConfig:
    """懒加载配置管理器（单例）"""

    _instance: Optional["LazyConfig"] = None
    _config: Optional[NodeConfigRoot] = None
    _config_path: Optional[str] = None
    _loaded: bool = False

    def __new__(cls):
        """单例模式：确保全局只有一个实例"""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def init(self, config_path: str) -> None:
        """
        初始化配置路径（仅记录路径，不立即加载）
        调用时机：项目启动时（如auto_format_thesis_document开头）
        """
        self._config_path = config_path
        self._loaded = False  # 重置加载状态
        logger.info(f"懒加载配置已初始化路径: {config_path}")

    def load(self) -> NodeConfigRoot:
        """
        实际加载并验证配置（首次调用时执行）
        内部调用，外部无需手动调用
        """
        if not self._config_path:
            raise ConfigNotLoadedError("请先调用 init(config_path) 初始化配置路径")

        try:
            # 加载并合并YAML配置
            raw_config = load_yaml_with_merge(self._config_path)
            # Pydantic验证配置结构
            self._config = NodeConfigRoot(**raw_config)
            self._loaded = True
            logger.info("懒加载配置加载并验证通过")
            return self._config
        except ValidationError as e:
            logger.error(f"配置结构验证失败: {e}")
            raise
        except Exception as e:
            logger.error(f"配置文件加载失败: {str(e)}")
            raise

    def get(self) -> NodeConfigRoot:
        """
        获取配置（核心懒加载逻辑）
        - 首次调用：自动执行load()加载配置
        - 后续调用：直接返回已加载的配置
        """
        if not self._loaded:
            self.load()
        if self._config is None:
            raise ConfigNotLoadedError("配置加载失败，无法获取")
        return self._config

    @property
    def config_path(self) -> Optional[str]:
        """获取已初始化的配置路径"""
        return self._config_path

    def clear(self):
        """清空配置（仅用于测试/重置）"""
        self._config = None
        self._config_path = None
        self._loaded = False


class ConfigNotLoadedError(Exception):
    """配置未加载时的自定义异常"""

    pass


# 全局懒加载配置实例（导入时仅创建空实例，不加载配置）
lazy_config = LazyConfig()


# 便捷函数：对外暴露的极简接口
def init_config(config_path: str):
    """初始化配置路径（供外部调用）"""
    lazy_config.init(config_path)


def get_config() -> NodeConfigRoot:
    """获取配置（懒加载核心入口）"""
    return lazy_config.get()


def clear_config():
    """清空配置（测试用）"""
    lazy_config.clear()
