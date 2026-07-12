#! /usr/bin/env python
# @Time    : 2026/1/24 19:47
# @Author  : afish
# @File    : datamodel.py
"""配置模型定义。

Pydantic 配置模型已被 DEFAULTS + DotDict 替代。
NodeConfigRoot 保留为 dict 子类，提供向后兼容的点号访问和 collect_style_configs()。
"""


class NodeConfigRoot(dict):
    """配置根节点 —— dict 子类，支持点号访问和 collect_style_configs()。

    可通过 **YAML_dict 构造，cfg.headings.level_1.font_size 等价于
    cfg["headings"]["level_1"]["font_size"]。
    """

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    def __getattr__(self, key: str):
        try:
            val = self[key]
        except KeyError:
            return None
        if isinstance(val, dict) and not isinstance(val, NodeConfigRoot):
            return NodeConfigRoot(**val)
        return val

    def __setattr__(self, key, value):
        self[key] = value

    def model_dump(self) -> dict:
        return dict(self)

    def collect_style_configs(self) -> dict[str, dict]:
        style_map: dict[str, dict] = {}
        _walk_config_for_styles(self, style_map)
        return style_map


def _resolve_builtin_style_name(cfg) -> str | None:
    from wordformat.style.defs import BuiltInStyle

    raw = (
        cfg.get("builtin_style_name")
        if isinstance(cfg, dict)
        else getattr(cfg, "builtin_style_name", None)
    )
    if not raw:
        return None
    try:
        return BuiltInStyle(raw).rel_value
    except Exception:
        return raw


class NumberingLevelConfig(dict):
    """编号级别配置 —— dict 子类，向后兼容编号模块。"""

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    def __getattr__(self, key):
        return self.get(key)


class NumberingConfig(dict):
    """编号总配置 —— dict 子类，向后兼容编号模块。"""

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    def __getattr__(self, key):
        val = self.get(key)
        if isinstance(val, dict) and not isinstance(val, NumberingConfig):
            return (
                NumberingConfig(**val)
                if key in ("captions",)
                else NumberingLevelConfig(**val)
            )
        return val


def _walk_config_for_styles(obj, style_map: dict[str, object]) -> None:
    if not isinstance(obj, dict):
        return
    eng_name = _resolve_builtin_style_name(obj)
    if eng_name and isinstance(eng_name, str):
        style_map[eng_name] = obj
    for _key, val in obj.items():
        if isinstance(val, dict):
            _walk_config_for_styles(val, style_map)
