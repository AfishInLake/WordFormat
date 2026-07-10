#! /usr/bin/env python
"""FormatNode 子类的自动注册机制。

每个 FormatNode 子类用 @register(category, level=...) 声明，
框架自动构建 CATEGORY_TO_CLASS 和 LEVEL_MAP，无需手动维护映射表。
"""

_registry: dict[str, type] = {}
_level_registry: dict[str, int] = {}


def register(category: str, level: int | None = None):
    """装饰器：将 FormatNode 子类注册到全局映射表。

    Usage:
        @register("abstract_chinese_title", level=1)
        class AbstractTitleCN(FormatNode[AbstractTitleConfig]):
            ...
    """

    def decorator(cls):
        _registry[category] = cls
        if level is not None:
            _level_registry[category] = level
        return cls

    return decorator
