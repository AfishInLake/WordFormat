#! /usr/bin/env python
"""FormatNode 子类的自动注册 + 配置导出机制。

每个 FormatNode 子类用 @register(category, level=...) 声明，
框架自动构建 CATEGORY_TO_CLASS 和 LEVEL_MAP。

注册表现已实例化（缝 B）：``default_registry`` 为默认领域（论文）实例，
``get_registry(domain)`` 创建/返回独立领域实例，多领域（公文、计划书……）
可注册各自的节点类而不互相覆盖。模块级 ``_registry`` / ``_level_registry``
与 ``register()`` / ``export_defaults()`` 为兼容垫片，指向默认实例，
现有 rules/ 模块与既有代码无需改动。
"""

import dataclasses
from typing import Dict, Optional, Type


class Registry:
    """领域级 FormatNode 注册表（每领域一份实例，互不覆盖）。"""

    def __init__(self) -> None:
        self._registry: Dict[str, Type] = {}
        self._level_registry: Dict[str, int] = {}

    def register(self, category: str, level: Optional[int] = None):
        """装饰器：将 FormatNode 子类注册到本实例。

        Usage:
            @registry.register("abstract_chinese_title", level=1)
            class AbstractTitleCN(FormatNode):
                ...
        """

        def decorator(cls):
            self._registry[category] = cls
            if level is not None:
                self._level_registry[category] = level
            return cls

        return decorator

    def get(self, category: str, default=None):
        """按类别取节点类（兼容 dict.get 语义）。"""
        return self._registry.get(category, default)

    def __contains__(self, category: str) -> bool:
        return category in self._registry

    def items(self):
        """类别 → 节点类 的迭代视图（遍历注册节点用）。"""
        return self._registry.items()

    def setdefault(self, category: str, cls: Type, level: Optional[int] = None) -> None:
        """为无 FormatNode 的类别登记占位节点类与层级（后处理扩展标签）。

        不覆盖已注册的真实节点类；level 仅在未登记时写入。
        """
        self._registry.setdefault(category, cls)
        if level is not None:
            self._level_registry.setdefault(category, level)

    def level_of(self, category: str) -> int:
        """类别逻辑层级；未登记返回 999（非标题）。"""
        return self._level_registry.get(category, 999)

    def reverse_category(self, cls: Type) -> Optional[str]:
        """节点类 → 注册类别名（树提升时同步 value 身份字段用）。"""
        for cat, c in self._registry.items():
            if c is cls:
                return cat
        return None

    def export_defaults(self) -> dict:
        """遍历本实例注册的 FormatNode 子类，根据 DEFAULTS + NODE_TYPE 生成完整配置。

        返回可直接写入 .yaml 的 dict，样式字段为平铺写法。
        """
        from wordformat.config.dotdict import flatten_style_config
        from wordformat.style.diff import WarningConfig

        result: dict = {}
        result["template_name"] = "未知模板"
        result["style_checks_warning"] = dataclasses.asdict(WarningConfig())
        result["numbering"] = {"enabled": False}
        # 页眉页脚配置段（issue#93）：非 FormatNode 的独立段，同样导出供模板配置
        import copy

        from wordformat.header_footer import DEFAULT_HF_CONFIG

        result["header_footer"] = copy.deepcopy(DEFAULT_HF_CONFIG)

        for cls in self._registry.values():
            defaults = getattr(cls, "DEFAULTS", {})
            node_type = getattr(cls, "NODE_TYPE", "")
            if defaults and node_type:
                _deep_set(result, node_type, flatten_style_config(defaults))

        return result


def _deep_set(d: dict, path: str, value: dict) -> None:
    """将 value 按点分路径写入嵌套 dict。"""
    parts = path.split(".")
    current = d
    for part in parts[:-1]:
        if part not in current:
            current[part] = {}
        current = current[part]
    # 最后一级：深度合并
    target = current.setdefault(parts[-1], {})
    target.update(value)


# ===== 默认领域实例（论文域，兼容既有行为）=====
default_registry = Registry()

# ===== 兼容垫片：旧引用方式指向默认实例 =====
_registry: Dict[str, Type] = default_registry._registry
_level_registry: Dict[str, int] = default_registry._level_registry

# 领域名 → 独立 Registry 实例（多领域共存）
_DOMAIN_REGISTRIES: Dict[str, Registry] = {}


def register(category: str, level: Optional[int] = None):
    """装饰器：注册到默认领域注册表（rules/*.py 无感知，行为不变）。"""
    return default_registry.register(category, level=level)


def export_defaults() -> dict:
    """导出默认领域（论文）的完整配置树。"""
    return default_registry.export_defaults()


def get_registry(domain: str) -> Registry:
    """返回指定领域的注册表实例；未知领域创建独立实例（不污染默认表）。"""
    if domain == "thesis":
        return default_registry
    return _DOMAIN_REGISTRIES.setdefault(domain, Registry())
