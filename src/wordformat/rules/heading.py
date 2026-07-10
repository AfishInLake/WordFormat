#! /usr/bin/env python
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : heading.py

from loguru import logger

from wordformat.config.models import HeadingLevelConfig, NodeConfigRoot
from wordformat.rules.node import FormatNode


class BaseHeadingNode(FormatNode[HeadingLevelConfig]):
    """标题节点基类（复用1/2/3级标题的通用逻辑）"""

    LEVEL: int = 0  # 标题层级（1/2/3）
    NODE_TYPE: str = ""
    CONFIG_MODEL = HeadingLevelConfig

    def _get_level_config(self, root_config: NodeConfigRoot) -> HeadingLevelConfig:
        """根据层级获取对应配置"""
        level_config_map = {
            1: root_config.headings.level_1,
            2: root_config.headings.level_2,
            3: root_config.headings.level_3,
        }
        target_config = level_config_map.get(self.LEVEL, root_config.headings.level_1)
        return target_config

    def load_config(self, root_config: dict | NodeConfigRoot):
        """重载加载配置方法，自动匹配对应层级的配置"""
        try:
            if isinstance(root_config, dict):
                # 修复：使用单下划线 _config（匹配基类@property的底层属性）
                level_config_dict = root_config.get("headings", {}).get(
                    f"level_{self.LEVEL}", {}
                )
                self._config = level_config_dict  # 正确赋值给单下划线私有属性
                logger.debug(f"{self.LEVEL}级标题字典配置：{self._config}")
                self._pydantic_config = self.CONFIG_MODEL(
                    **self._config
                )  # 读取赋值后的 _config

            elif isinstance(root_config, NodeConfigRoot):
                # 修复：先赋值 _pydantic_config，再同步到 _config
                self._pydantic_config = self._get_level_config(root_config)
                self._config = (
                    self._pydantic_config.model_dump()
                )  # 读取底层 _pydantic_config
            else:
                raise TypeError(
                    f"配置类型不支持：{type(root_config)}，仅支持dict或NodeConfigRoot"
                )

        except Exception as e:
            logger.error(f"{self.LEVEL}级标题配置加载失败：{str(e)}")
            raise  # 抛出异常，避免后续使用错误配置


# 各层级标题节点（无需重写check_format，直接复用基类逻辑）
class HeadingLevel1Node(BaseHeadingNode):
    """一级标题节点"""

    LEVEL = 1
    NODE_TYPE = "headings.level_1"
    NODE_LABEL = "一级标题"


class HeadingLevel2Node(BaseHeadingNode):
    """二级标题节点"""

    LEVEL = 2
    NODE_TYPE = "headings.level_2"
    NODE_LABEL = "二级标题"


class HeadingLevel3Node(BaseHeadingNode):
    """三级标题节点"""

    LEVEL = 3
    NODE_TYPE = "headings.level_3"
    NODE_LABEL = "三级标题"
