#! /usr/bin/env python
# @Time    : 2026/9/28
# @Author  : afish
# @File    : __init__.py
"""文档领域（document domain）注册表。

框架（pipeline / hooks / handlers）不直接依赖任何具体文档类型；
每个文档类型（论文、公文、计划书……）是独立领域模块，通过
``@register_domain(name)`` 注册自己的规则节点装配与 hook handler，
由编排入口按 ``document_type`` 加载。目前内置领域：

- ``thesis``：学术论文（domains/thesis.py）

新增文档类型步骤：
1. 在 domains/ 下新建模块（如 gongwen.py），用 @register_domain 注册；
2. 在模块内 import 规则节点类（触发 @register 进入节点注册表）并实现
   register_handlers()，把该文档类型的处理逻辑（树提升、题注编号注入、
   摘要统计等）注册为 hook handler；
3. 编排入口按 document_type 装配（见 pipeline/orchestrate.py）。
"""

from __future__ import annotations

import importlib
import pkgutil
from typing import Callable, Dict, List

from loguru import logger

from wordformat.hooks import HookRegistry, hooks

# 领域名 → 注册函数（register_handlers(registry)）
_DOMAIN_REGISTRY: Dict[str, Callable] = {}

# 已注册回调的领域集合（幂等装配）
_LOADED_DOMAINS: set = set()


def register_domain(name: str):
    """装饰器：声明领域注册函数（签名 register_handlers(registry)）。

    Usage:
        @register_domain("thesis")
        def register_handlers(registry):
            ...
    """

    def decorator(fn: Callable):
        _DOMAIN_REGISTRY[name] = fn
        return fn

    return decorator


def list_domains() -> List[str]:
    """返回已注册的领域名列表。"""
    return sorted(_DOMAIN_REGISTRY)


def load_domain(name: str, registry: HookRegistry = hooks, force: bool = False) -> None:
    """装配指定领域：执行其 register_handlers(registry)。

    幂等：同一领域只装配一次（force 可强制，测试用）。
    未知领域抛 ValueError。
    """
    register_fn = _DOMAIN_REGISTRY.get(name)
    if register_fn is None:
        raise ValueError(
            f"未知文档领域 '{name}'（已注册: {', '.join(list_domains()) or '空'}）"
        )
    if not force and name in _LOADED_DOMAINS:
        return
    register_fn(registry)
    _LOADED_DOMAINS.add(name)
    logger.debug(f"领域 '{name}' 已装配")


def reset_domains() -> None:
    """清空领域装配记录（测试用，不会清空已注册回调本身）。"""
    _LOADED_DOMAINS.clear()


# 自动发现 domains/ 下所有领域模块（thesis.py、gongwen.py……），
# 导入后 @register_domain 装饰器会把注册函数写入 _DOMAIN_REGISTRY
def _discover_domain_modules() -> None:
    for module_info in pkgutil.iter_modules(__path__):
        importlib.import_module(f"{__name__}.{module_info.name}")


_discover_domain_modules()
