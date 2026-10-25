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


# 全局格式默认值，所有节点以此为底（与 example/undergrad_thesis.yaml 的
# global_format 保持一致：0 行段前段后、1.5 倍行距、正文两端对齐）
BASE_FORMAT: dict[str, object] = {
    "paragraph": {
        "alignment": "两端对齐",
        "space_before": "0行",
        "space_after": "0行",
        "line_spacingrule": "1.5倍行距",
        "line_spacing": "1.5倍",
        "left_indent": "0字符",
        "right_indent": "0字符",
        "first_line_indent": "2字符",
        "builtin_style_name": "正文",
    },
    "font": {
        "chinese_font_name": "宋体",
        "english_font_name": "Times New Roman",
        "font_size": "小四",
        "font_color": "黑色",
        "bold": False,
        "italic": False,
        "underline": False,
    },
}


def deep_merge(base: dict, override: dict) -> dict:
    """深度合并 override 到 base，返回新 dict。"""
    result = dict(base)
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = deep_merge(result[k], v)
        else:
            result[k] = v
    return result


def flatten_style_config(obj: dict) -> dict:
    """把代码内 DEFAULTS 的 paragraph/font 分组摊平成与 YAML 一致的平铺结构。

    顶层字段优先于分组字段，供节点配置合并与 ``wordf config`` 导出使用。
    YAML 配置本身只接受平铺写法，因此该函数只处理 Python 常量。
    """
    if not isinstance(obj, dict):
        return obj
    result: dict = {}
    for group in ("paragraph", "font"):
        sub = obj.get(group)
        if isinstance(sub, dict):
            result.update(sub)
    for key, value in obj.items():
        if key in ("paragraph", "font"):
            continue
        result[key] = flatten_style_config(value) if isinstance(value, dict) else value
    return result
