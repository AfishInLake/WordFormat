"""图片段落和表格对象节点。"""

from wordformat.rules.node import FormatNode
from wordformat.structure.registry import register
from wordformat.style.comments import format_comment


@register("figure_image")
class FigureImage(FormatNode):
    """图片段落节点（包含内联图片的段落，非题注）。

    只检查对齐和首行缩进，不检查行距、字体等。
    """

    NODE_TYPE = "figure_image"
    NODE_LABEL = "图片段落"
    DEFAULTS = {"alignment": "居中对齐", "first_line_indent": "0字符"}
    DEFAULT_RULES = {}

    def _base(self, doc, p: bool, r: bool):
        """仅检查对齐和首行缩进。"""
        from wordformat.style.defs import Alignment, FirstLineIndent
        from wordformat.style.diff import _format_para_value

        cfg = self.pydantic_config
        expected_align = Alignment(str(cfg.alignment or "居中对齐"))
        actual_align = expected_align.get_from_paragraph(self.paragraph)
        if expected_align != actual_align:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=format_comment(
                    self.NODE_LABEL,
                    "对齐错误",
                    _format_para_value("alignment", actual_align),
                    _format_para_value("alignment", expected_align),
                ),
            )

        expected_indent = FirstLineIndent(str(cfg.first_line_indent or "0字符"))
        actual_indent = expected_indent.get_from_paragraph(self.paragraph)
        if expected_indent != actual_indent:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=format_comment(
                    self.NODE_LABEL,
                    "首行缩进错误",
                    _format_para_value("first_line_indent", actual_indent),
                    _format_para_value("first_line_indent", expected_indent),
                ),
            )


@register("table_object")
class TableObject(FormatNode):
    """表格对象节点（表格整体格式，非题注）。"""

    NODE_TYPE = "table_object"
    NODE_LABEL = "表格对象"
    DEFAULTS = {}
    DEFAULT_RULES = {}  # 表格对象格式由 Word 表格属性控制
