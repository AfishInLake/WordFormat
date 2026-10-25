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
            # fallback：到 paragraph / font 子字典中查找
            for sub in ("paragraph", "font"):
                sub_dict = self.get(sub)
                if isinstance(sub_dict, dict) and key in sub_dict:
                    val = sub_dict[key]
                    break
            else:
                return None
        if isinstance(val, dict) and not isinstance(val, NodeConfigRoot):
            return NodeConfigRoot(**val)
        return val

    def __setattr__(self, key, value):
        self[key] = value

    def model_dump(self) -> dict:
        return dict(self)

    def collect_style_configs(self) -> dict[str, dict]:
        """收集每个内置样式在所有节点中一致的格式字段。

        同名样式常被多个节点复用（Normal 被正文、参考文献、致谢共享）。
        Word 的样式定义只有一份，若直接写入某个节点的完整配置，会把该
        节点的字体/缩进等泄漏给其他未显式覆盖该属性的段落（典型如表格
        单元格、页眉页脚）。因此这里只保留所有节点都显式给出且取值相同
        的字段；存在冲突的字段交给各节点的直接格式处理。

        锚点模板（global_format）和表格对象（tables.object）不对应实际
        段落，跳过以免污染同名样式。
        """
        occurrences: dict[str, list[dict]] = {}
        _walk_config_for_styles(self, occurrences)
        return {
            name: _intersect_configs(configs) for name, configs in occurrences.items()
        }


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


_STYLE_SUBDICT_KEYS = ("paragraph", "font")

# 这些配置项不是实际段落，不能参与样式定义收集
_STYLE_EXCLUDED_PATHS = {
    ("global_format",),
    ("tables", "object"),
}


def _flatten_style_config(obj: dict) -> dict:
    """把嵌套的 paragraph / font 字段摊平到节点配置。

    顶层字段优先于子字典字段，与 DotDict/NodeConfigRoot 的点号查找顺序
    保持一致。
    """
    merged: dict = {}
    for sub_key in _STYLE_SUBDICT_KEYS:
        sub_cfg = obj.get(sub_key)
        if isinstance(sub_cfg, dict):
            merged.update(sub_cfg)
    merged.update(
        {key: value for key, value in obj.items() if key not in _STYLE_SUBDICT_KEYS}
    )
    return merged


def _intersect_configs(configs: list[dict]) -> dict:
    """只保留所有配置都显式给出且取值一致的字段。"""
    common_keys = set(configs[0])
    for cfg in configs[1:]:
        common_keys &= set(cfg)
    return {
        key: configs[0][key]
        for key in common_keys
        if all(cfg[key] == configs[0][key] for cfg in configs[1:])
    }


def _walk_config_for_styles(obj, occurrences: dict[str, list[dict]], path=()) -> None:
    if not isinstance(obj, dict):
        return
    if path in _STYLE_EXCLUDED_PATHS:
        return
    style_cfg = _flatten_style_config(obj)
    eng_name = _resolve_builtin_style_name(style_cfg)
    if eng_name and isinstance(eng_name, str):
        occurrences.setdefault(eng_name, []).append(style_cfg)
    for key, val in obj.items():
        # paragraph / font 已合并进当前节点，不能作为独立节点重复收集
        if key in _STYLE_SUBDICT_KEYS or not isinstance(val, dict):
            continue
        _walk_config_for_styles(val, occurrences, (*path, key))
