"""wordformat.math — LaTeX → OMML 公式渲染。

将 LaTeX 数学表达式转换为 Word 原生 OMML (Office Math Markup Language)，
支持分数、根式、上下标、大型运算符、重音符号、希腊字母等。

Usage:
    from wordformat.math import add_display_math, add_inline_math, latex_to_omath

    # 块级公式
    add_display_math(doc, r"\frac{1}{2\\pi}\\int_{-\\infty}^{\\infty}e^{-x^2}dx")

    # 内联公式
    add_inline_math(paragraph, "其中 $E=mc^2$ 是质能方程")
"""

from wordformat.math.omml import (
    add_display_math,
    add_inline_math,
    latex_to_omath,
    latex_to_omath_para,
    set_cell_math,
)

__all__ = [
    "latex_to_omath",
    "latex_to_omath_para",
    "add_display_math",
    "add_inline_math",
    "set_cell_math",
]
