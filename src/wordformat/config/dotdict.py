"""轻量级点号访问字典，替代 Pydantic 模型做配置容器。"""


class DotDict(dict):
    """支持点号访问的字典，递归转换嵌套 dict。

    cfg.alignment 等价于 cfg["alignment"]
    cfg.chinese_title.font_size 等价于 cfg["chinese_title"]["font_size"]
    """

    def __getattr__(self, key: str):
        try:
            val = self[key]
        except KeyError:
            return None
        if isinstance(val, dict):
            return DotDict(val)
        return val

    def __setattr__(self, key: str, value) -> None:
        self[key] = value

    def __delattr__(self, key: str) -> None:
        try:
            del self[key]
        except KeyError as e:
            raise AttributeError(key) from e


def deep_merge(base: dict, override: dict) -> dict:
    """深度合并 override 到 base，返回新 dict。"""
    result = dict(base)
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = deep_merge(result[k], v)
        else:
            result[k] = v
    return result
