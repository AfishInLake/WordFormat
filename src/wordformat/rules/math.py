#! /usr/bin/env python
"""数学公式节点。公式本身由 OMML 渲染，Node 仅占位确保不被跳过。"""

from wordformat.config.dotdict import BASE_FORMAT, deep_merge
from wordformat.rules.node import FormatNode
from wordformat.structure.registry import register


@register("math_block", level=5)
class MathBlock(FormatNode):
    """块级公式"""

    NODE_TYPE = "math.block"
    NODE_LABEL = "块级公式"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            # 带编号的公式常用“居中公式 + 右对齐编号”的制表位写法，
            # 段落对齐因此可能是两端/左对齐，不以此为错
            "paragraph": {"alignment": ["居中对齐", "两端对齐", "左对齐"]},
            # 公式编号字号常比正文小一档（五号）
            "font": {"font_size": ["小四", "五号"]},
        },
    )

    def _base(self, doc, p: bool, r: bool):
        """公式段必须居中：段落未显式设置对齐时也要写入 / 报告。

        纯 OMML 公式段没有 w:r，段落样式检查会跳过；段落对齐的“未设置”
        按继承处理也不报告，导致公式不居中却无人发现。此处单独兜底：
        apply 时写入居中，check 时对带编号（有 run）的公式段报出批注。
        """
        from wordformat.style.diff import DIFFResult, ParagraphStyle
        from wordformat.style.reader import paragraph_get_alignment

        para = self.paragraph
        if para is None or paragraph_get_alignment(para) is not None:
            return
        ps = ParagraphStyle.from_config(self.pydantic_config)
        if ps.alignment.value != "居中对齐":
            return
        if not p:
            ps.alignment.format(docx_obj=para)
            return
        if not para.runs:
            return
        diff = DIFFResult(
            "alignment",
            ps.alignment,
            None,
            f"对齐方式期待{str(ps.alignment)};",
            0,
        )
        self.add_comment(
            doc=doc,
            runs=para.runs,
            text=ParagraphStyle.to_string([diff], target=self.NODE_LABEL),
        )
