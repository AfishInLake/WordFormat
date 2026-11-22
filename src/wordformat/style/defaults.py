"""样式兜底默认值与缩进方向枚举（全库单一数据源）。

这些值只在"配置完全缺省"时使用，属于产品默认排版约定而非环境配置。
集中定义避免同一个默认字体/字号在多处拷贝后分叉。
"""

from enum import Enum

DEFAULT_CN_FONT = "宋体"
DEFAULT_EN_FONT = "Times New Roman"
DEFAULT_FONT_SIZE = "小四"
DEFAULT_FONT_COLOR = "黑色"
DEFAULT_CODE_FONT = "Consolas"

#: 中文 1 字符 ≈ 240 twips（小四字号下的常用折算基准）
TWIPS_PER_CHAR = 240

#: 行距上限：超过该值的"行"配置按 10 行兜底（Word 实际可用范围）
MAX_LINE_SPACING = 10.0

#: 内置编号悬挂缩进：每级 210 twips
HANGING_TWIPS_PER_LEVEL = 210

#: 左右缩进比较容差（字符）：WPS 常用 0.25/0.26 字符微调，视觉等同 0
INDENT_TOLERANCE_CHARS = 0.5

#: 首行缩进比较容差（字符）：2 字符在五号/小四下折算为 1.8~2.2
FIRST_LINE_INDENT_TOLERANCE_CHARS = 0.3


class IndentSide(str, Enum):
    """段落左/右缩进方向（写入与读取共用同一词汇）。"""

    LEFT = "left"
    RIGHT = "right"

    def __str__(self) -> str:
        return str(self.value)


#: 历史写法 "R"/"X" 与标准写法 "left"/"right" 的兼容映射
_INDENT_SIDE_ALIASES = {
    "R": IndentSide.LEFT,
    "X": IndentSide.RIGHT,
    "left": IndentSide.LEFT,
    "right": IndentSide.RIGHT,
}


def resolve_indent_side(value) -> IndentSide:
    """把 "R"/"X"/"left"/"right" 统一为 IndentSide；非法值立即报错。"""
    if isinstance(value, IndentSide):
        return value
    try:
        return _INDENT_SIDE_ALIASES[value]
    except (KeyError, TypeError) as exc:
        raise ValueError(
            f"无效的缩进类型: {value!r}；indent_type 必须是 'left'/'right'"
            "（或历史写法 'R'/'X'）"
        ) from exc
