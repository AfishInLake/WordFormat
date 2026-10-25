#! /usr/bin/env python
"""文档领域分类后处理注册表（框架侧）。

分类（parse）链路的领域差异由各领域模块提供"整链函数"，经本注册表分派：
论文域的整链实现位于 ``wordformat.domains.thesis.postprocess``（随领域包
自动发现加载，导入即通过 ``register_postprocess`` 注册）。

与 domains/__init__.py 的事件回调注册表分工：事件注册表负责格式化链路的
领域差异，本注册表负责分类（parse）链路的领域差异。

协议是**整体替换**语义（整链函数），而非布尔开关——新领域需要的不是
"禁用某条规则"而是"改写某条规则"。
"""

from typing import Callable, Dict

# 领域名 → 整链函数（领域模块导入时注册，未装配的领域分派返回 None）
_POSTPROCESS_REGISTRY: Dict[str, Callable] = {}


def register_postprocess(name: str, fn: Callable) -> None:
    """注册文档领域的分类后处理实现（整体替换语义）。"""
    _POSTPROCESS_REGISTRY[name] = fn


def get_postprocess(domain: str) -> Callable | None:
    """返回指定领域的分类后处理函数；领域未装配（未注册）返回 None。"""
    return _POSTPROCESS_REGISTRY.get(domain)
