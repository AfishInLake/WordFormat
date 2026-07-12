#! /usr/bin/env python
"""FormatNode 子类的自动注册 + 配置导出机制。

每个 FormatNode 子类用 @register(category, level=...) 声明，
框架自动构建 CATEGORY_TO_CLASS 和 LEVEL_MAP。

export_defaults() 遍历所有注册类，按 NODE_TYPE 路径重建完整 YAML 配置树。
"""

import dataclasses

_registry: dict[str, type] = {}
_level_registry: dict[str, int] = {}


def register(category: str, level: int | None = None):
    """装饰器：将 FormatNode 子类注册到全局映射表。

    Usage:
        @register("abstract_chinese_title", level=1)
        class AbstractTitleCN(FormatNode):
            ...
    """

    def decorator(cls):
        _registry[category] = cls
        if level is not None:
            _level_registry[category] = level
        return cls

    return decorator


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


def export_defaults() -> dict:
    """遍历所有注册的 FormatNode 子类，根据 DEFAULTS + NODE_TYPE 生成完整配置。

    返回可直接写入 .yaml 的嵌套 dict。
    """
    from wordformat.style.diff import WarningConfig

    result: dict = {}
    result["template_name"] = "未知模板"
    result["style_checks_warning"] = dataclasses.asdict(WarningConfig())
    result["numbering"] = {"enabled": False}

    for cls in _registry.values():
        defaults = getattr(cls, "DEFAULTS", {})
        node_type = getattr(cls, "NODE_TYPE", "")
        if defaults and node_type:
            _deep_set(result, node_type, dict(defaults))

    return result
