#! /usr/bin/env python
# @Time    : 2026/1/11 22:25
# @Author  : afish
# @File    : settings.py
"""结构层配置视图。

只暴露注册表的别名，**不 import 任何具体节点或领域模块**：
装配由 ``wordformat.assembly.register_nodes()`` 在包初始化时完成
（``wordformat/__init__.py``），因此这里读到的一定是完整注册表。
这样 structure 层不再反向依赖 rules/domains，可独立复用。
"""

from wordformat.structure.registry import _level_registry, _registry

CATEGORY_TO_CLASS = _registry
LEVEL_MAP = _level_registry
