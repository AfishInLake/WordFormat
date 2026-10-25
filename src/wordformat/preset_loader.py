#! /usr/bin/env python
# @Time    : 2026/9/28
# @Author  : afish
# @File    : preset_loader.py
"""Preset 预设加载器。

预设是 presets/ 目录下的 .py 插件文件，约定：

1. 文件顶部可用 ``PRESET_MANIFEST`` 字典声明需要的 YAML 配置项
   （缺省格式配置），加载时自动深合并进用户配置（用户配置优先）；
2. 必须暴露 ``register(hooks)`` 函数，在其中调用 ``hooks.register(event, cb)``
   注册 Hook 回调。

多个预设叠加加载时，后加载的 manifest 覆盖先加载的（“预设优先级”）。
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from typing import Callable, Dict, List, Optional

from loguru import logger

from wordformat.hooks import HookRegistry, hooks
from wordformat.settings import BASE_DIR

# presets/ 目录默认位于项目根
PRESET_DIR = Path(BASE_DIR) / "presets"

# 模板文件名不视为可用预设
TEMPLATE_NAME = "_template.py"


def preset_dir() -> Path:
    """返回 presets 目录（不存在时创建）。"""
    directory = PRESET_DIR
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def list_presets() -> List[str]:
    """列出 presets/ 目录下所有可用预设名（按文件名，不含扩展名）。

    非 .py 文件、_template.py 均排除。目录不存在时返回空列表。
    """
    directory = PRESET_DIR
    if not directory.exists():
        return []
    names = []
    for f in sorted(directory.iterdir()):
        if not f.is_file() or f.suffix != ".py":
            continue
        if f.name in (TEMPLATE_NAME, "__init__.py"):
            continue
        names.append(f.stem)
    return names


def find_preset(name: str) -> Path:
    """按名称定位预设文件路径；未找到抛 FileNotFoundError。"""
    if not isinstance(name, str) or not name.strip():
        raise FileNotFoundError(f"预设名不能为空: {name!r}")
    if "\x00" in name or "/" in name or "\\" in name:
        raise FileNotFoundError(f"预设名不合法: {name!r}")
    path = PRESET_DIR / f"{name}.py"
    if not path.exists():
        available = ", ".join(list_presets()) or "（空）"
        raise FileNotFoundError(f"预设 '{name}' 不存在（presets/ 下可用: {available}）")
    return path


def _load_module(path: Path) -> Optional[object]:
    """动态加载单个 .py 模块并缓存到 sys.modules，避免重复加载。"""
    module_name = f"wordformat_preset_{path.stem}"
    if module_name in sys.modules:
        return sys.modules[module_name]
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        logger.error(f"无法解析预设文件: {path}")
        return None
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    except Exception as e:
        sys.modules.pop(module_name, None)
        logger.error(f"加载预设 {path.name} 失败: {e}")
        raise
    return module


def load_preset(
    name: str,
    registry: HookRegistry = hooks,
) -> Dict:
    """加载指定预设：执行其 register() 并返回 manifest 配置。

    Args:
        name: 预设名（presets/<name>.py）。
        registry: 回调注册目标，默认全局 hooks。

    Returns:
        PRESET_MANIFEST 字典（未声明时为空 dict）。

    Raises:
        FileNotFoundError: 预设不存在或未定义 register()。
    """
    path = find_preset(name)
    module = _load_module(path)
    if module is None:
        raise FileNotFoundError(f"预设加载失败: {path}")

    register_fn: Optional[Callable] = getattr(module, "register", None)
    if not callable(register_fn):
        raise FileNotFoundError(
            f"预设 '{name}' 未定义 register(hooks) 函数，请参考 presets/_template.py"
        )

    manifest = getattr(module, "PRESET_MANIFEST", None)
    manifest_dict = dict(manifest) if isinstance(manifest, dict) else {}
    if registry.is_preset_loaded(name):
        # 幂等：长驻服务/测试中重复加载同一预设不得重复注册回调
        logger.debug(f"预设 '{name}' 已加载，跳过重复注册")
        return manifest_dict

    register_fn(registry)
    registry.mark_preset_loaded(name)
    logger.info(f"预设 '{name}' 已加载（{path.name}）")
    return manifest_dict


def load_presets(
    names: Optional[List[str]] = None,
    registry: HookRegistry = hooks,
) -> List[Dict]:
    """加载多个预设并返回 manifest 列表（按加载顺序，后者覆盖前者）。

    Args:
        names: 预设名列表；None / 空列表表示不加载任何预设。
        registry: 回调注册目标，默认全局 hooks。

    Returns:
        manifest 列表，顺序与 names 一致（供调用方按深度合并）。
    """
    if not names:
        return []
    manifests: List[Dict] = []
    for name in names:
        manifests.append(load_preset(name, registry=registry))
    return manifests


def clear_presets(registry: HookRegistry = hooks) -> None:
    """清空预设注册的全部回调（测试用）。"""
    registry.clear()
