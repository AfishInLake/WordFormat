"""批注文本格式化工具。

所有 add_comment 调用最终由此模块生成符合标准格式的批注文本：
    [位置]-[问题类型]：[现状]，规范：[标准]
"""

# diff_type → 中文问题类型
CHAR_DIFF_LABELS: dict[str, str] = {
    "font_size": "字号错误",
    "bold": "加粗错误",
    "italic": "斜体错误",
    "underline": "下划线错误",
    "font_color": "字体颜色错误",
    "font_name_cn": "中文字体错误",
    "font_name_en": "英文字体错误",
}

PARA_DIFF_LABELS: dict[str, str] = {
    "alignment": "对齐错误",
    "first_line_indent": "首行缩进错误",
    "line_spacing": "行距问题",
    "line_spacing_rule": "行距类型问题",
    "space_before": "段前间距错误",
    "space_after": "段后间距错误",
    "left_indent": "左缩进错误",
    "right_indent": "右缩进错误",
    "builtin_style_name": "内置样式错误",
}


# 问题类型 → 严重等级
SEVERITY_MAP: dict[str, str] = {
    "行距问题": "严重",
    "对齐错误": "严重",
    "首行缩进错误": "一般",
    "段前间距错误": "一般",
    "段后间距错误": "一般",
    "行距类型问题": "一般",
    "左缩进错误": "一般",
    "右缩进错误": "一般",
    "内置样式错误": "一般",
    "字号错误": "一般",
    "加粗错误": "一般",
    "数量过少": "一般",
    "数量过多": "一般",
    "编号错误": "一般",
    "章节号错误": "一般",
    "分隔符错误": "提醒",
    "标签错误": "提醒",
    "间距错误": "提醒",
    "格式错误": "提醒",
    "斜体错误": "提醒",
    "下划线错误": "提醒",
    "字体颜色错误": "提醒",
    "中文字体错误": "提醒",
    "英文字体错误": "提醒",
    "标点错误": "提醒",
}

_DEFAULT_SEVERITY = "一般"

# 严重等级排序权重（数值越小越严重）
SEVERITY_ORDER: dict[str, int] = {"严重": 0, "一般": 1, "提醒": 2}


def get_severity(comment_text: str) -> str:
    """从批注文本中提取严重等级。格式：位置-问题类型：..."""
    try:
        prop = comment_text.split("-", 1)[1].split("：")[0]
    except (IndexError, AttributeError):
        return _DEFAULT_SEVERITY
    return SEVERITY_MAP.get(prop, _DEFAULT_SEVERITY)


def format_comment(target: str, property_name: str, actual: str, standard: str) -> str:
    """生成标准批注文本：位置-问题类型：现状，规范：标准。"""
    return f"{target}-{property_name}：{actual}，规范：{standard}"
