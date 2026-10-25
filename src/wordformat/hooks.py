#! /usr/bin/env python
# @Time    : 2026/9/28
# @Author  : afish
# @File    : hooks.py
"""Hook 事件调度器（Preset 插件系统核心）。

预设（preset）通过 register() 注册事件回调，框架在流水线关键位置 emit()
触发回调。回调签名统一为 ``callback(event_data: dict) -> dict | None``：

- 通过修改 event_data 中传入的可变对象（config、root_node、issues 等）实现定制；
- 返回 dict 时会被合并进事件数据，实现覆盖式定制（如返回 {"text": "..."} 改写批注文本）。

事件不区分注册顺序，同一事件按注册先后依次调用。
"""

from __future__ import annotations

from collections import defaultdict
from typing import Any, Callable, Dict, List

from loguru import logger

# ------------------------------------------------------------------
# 标准事件名（与流水线对应关系见 issue#94）
# ------------------------------------------------------------------
EVENT_ON_CONFIG_LOADED = "on_config_loaded"
EVENT_ON_TREE_BUILT = "on_tree_built"
EVENT_BEFORE_NODE_FORMAT = "before_node_format"
EVENT_AFTER_NODE_FORMAT = "after_node_format"
EVENT_ON_CAPTION_CHECK = "on_caption_check"
EVENT_ON_HEADING_NUMBERING = "on_heading_numbering"
EVENT_ON_TREE_NORMALIZE = "on_tree_normalize"
EVENT_ON_FORMAT_BEGIN = "on_format_begin"
EVENT_ON_SUMMARY_BUILD = "on_summary_build"
EVENT_BEFORE_DOCUMENT_SAVE = "before_document_save"
EVENT_ON_COMMENT = "on_comment"

ALL_EVENTS = (
    EVENT_ON_CONFIG_LOADED,
    EVENT_ON_TREE_BUILT,
    EVENT_BEFORE_NODE_FORMAT,
    EVENT_AFTER_NODE_FORMAT,
    EVENT_ON_CAPTION_CHECK,
    EVENT_ON_HEADING_NUMBERING,
    EVENT_ON_TREE_NORMALIZE,
    EVENT_ON_FORMAT_BEGIN,
    EVENT_ON_SUMMARY_BUILD,
    EVENT_BEFORE_DOCUMENT_SAVE,
    EVENT_ON_COMMENT,
)


class HookRegistry:
    """事件注册与触发的调度器。"""

    def __init__(self):
        self._callbacks: Dict[str, List[tuple]] = defaultdict(list)
        self._has_handlers = False
        # 已成功注册的预设名：同一 registry 重复 load_preset 时不重复挂回调
        self._loaded_presets: set[str] = set()

    # ------------------------------------------------------------------
    # 注册
    # ------------------------------------------------------------------
    def register(self, event: str, callback: Callable, name: str = None) -> None:
        """注册事件回调，重复注册同一回调会被追加调用。

        Args:
            event: 事件名（标准事件见 ALL_EVENTS，也允许自定义事件）。
            callback: ``callback(event_data: dict) -> dict | None``，
                返回 dict 时合并进事件数据。
            name: 可选回调名（内置机制用 "builtin.*" 前缀，供 unregister 精确定位）。
        """
        if event not in ALL_EVENTS:
            logger.warning(
                f"注册了未定义的事件 '{event}'，预设可能无法生效，"
                f"标准事件: {', '.join(ALL_EVENTS)}"
            )
        if not callable(callback):
            logger.warning(f"事件 '{event}' 的回调不是可调用对象，已忽略: {callback}")
            return
        cb_name = name or getattr(callback, "__name__", str(callback))
        self._callbacks[event].append((cb_name, callback))
        self._has_handlers = True
        logger.debug(f"已注册 hook: {event} <- {cb_name}")

    def unregister(self, event: str, name: str) -> bool:
        """按回调名取消注册（预设可用它禁用框架内置机制，如 builtin.tree_normalize）。

        Args:
            event: 事件名。
            name: 注册时传入的回调名；name 为 None 时清空该事件全部回调。

        Returns:
            True 表示至少移除一个回调。
        """
        if name is None:
            existed = bool(self._callbacks.get(event))
            self._callbacks.pop(event, None)
            self._refresh_has_handlers()
            return existed
        callbacks = self._callbacks.get(event)
        if not callbacks:
            return False
        remain = [item for item in callbacks if item[0] != name]
        removed = len(remain) != len(callbacks)
        if removed:
            if remain:
                self._callbacks[event] = remain
            else:
                self._callbacks.pop(event, None)
            self._refresh_has_handlers()
        return removed

    def _refresh_has_handlers(self) -> None:
        """unregister 后重算 has_handlers。"""
        self._has_handlers = any(bool(cbs) for cbs in self._callbacks.values())

    def is_preset_loaded(self, name: str) -> bool:
        """该 registry 是否已成功注册过指定预设。"""
        return name in self._loaded_presets

    def mark_preset_loaded(self, name: str) -> None:
        """标记预设已注册，供 load_preset 做幂等控制。"""
        self._loaded_presets.add(name)

    # ------------------------------------------------------------------
    # 触发
    # ------------------------------------------------------------------
    def emit(self, event: str, **event_data: Any) -> dict:
        """触发事件，返回回调合并后的数据（浅拷贝）。

        未注册任何回调时直接返回原数据，零开销（热路径 safe）。
        回调可通过三种方式影响流水线：

        1. 修改传入的可变对象（config 字典、issues 列表等）；
        2. 返回 dict，其键值会覆盖 event_data 中同名键；
        3. 返回 None，表示仅执行副作用。
        """
        callbacks = self._callbacks.get(event)
        if not callbacks:
            return event_data
        data: dict = dict(event_data)
        for _name, cb in callbacks:
            try:
                result = cb(data)
            except Exception as e:
                logger.error(f"Hook 事件 '{event}' 回调 {cb} 执行失败: {e}")
                continue
            if isinstance(result, dict):
                data.update(result)
        return data

    # ------------------------------------------------------------------
    # 工具
    # ------------------------------------------------------------------
    def is_registered(self, event: str) -> bool:
        """是否已注册该事件的回调。"""
        return bool(self._callbacks.get(event))

    def registered_events(self) -> List[str]:
        """返回已注册回调的事件名列表。"""
        return [e for e, cbs in self._callbacks.items() if cbs]

    def registered_callbacks(self, event: str = None) -> List[tuple]:
        """返回 (名称, 回调) 列表；event 为 None 时返回全部事件。"""
        if event is not None:
            return list(self._callbacks.get(event, []))
        return [item for cbs in self._callbacks.values() for item in cbs]

    def clear(self) -> None:
        """清空所有回调与预设加载标记（测试与热重载用）。"""
        self._callbacks.clear()
        self._has_handlers = False
        self._loaded_presets.clear()

    @property
    def has_handlers(self) -> bool:
        """全局是否有任何注册回调（emit 仍会做空转短路，此属性供调用方预判）。"""
        return self._has_handlers


# 全局唯一调度器：preset 与框架共用同一个实例
hooks = HookRegistry()
