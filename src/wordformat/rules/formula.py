"""FormulaNode — LaTeX 公式渲染为 OMML。render() 管结构，_base() 管段落格式。"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from wordformat.rules.node import FormatNode

if TYPE_CHECKING:
    from docx import Document
    from docx.text.paragraph import Paragraph


class FormulaNode(FormatNode):
    """公式节点 —— LaTeX → OMML，块级居中零缩进，行内随正文。"""

    NODE_TYPE = "formula"
    CONFIG_MODEL = None
    CONFIG_PATH = None

    def __init__(
        self,
        value: Any,
        level: int | float,
        paragraph=None,
        expected_rule: dict[str, Any] | None = None,
    ):
        super().__init__(
            value=value, level=level, paragraph=paragraph, expected_rule=expected_rule
        )
        self.type = "formula"
        self.content.setdefault("latex", "")
        self.content.setdefault("display", True)

    def load_config(self, full_config) -> None:
        """公式暂不绑定 Pydantic 配置，段落样式在 _base() 中处理。"""

    # ── 格式：与其他节点一致的 _base 模式 ──────────────────────────────

    def _base(self, doc: Document, p: bool, r: bool) -> None:  # noqa: FBT001
        """块级公式：居中 + 零缩进。行内公式：不设段落样式。"""
        if self.paragraph is None:
            return
        if not self.content.get("display", True):
            return  # 行内公式不强制段落样式

        from wordformat.style.check_format import ParagraphStyle

        ps = ParagraphStyle(
            alignment="居中对齐",
            first_line_indent="0字符",
            left_indent="0字符",
        )
        if p:
            issues = ps.diff_from_paragraph(self.paragraph)
        else:
            issues = ps.apply_to_paragraph(self.paragraph)
        if issues:
            self.add_comment(
                doc=doc, runs=self.paragraph.runs, text=ParagraphStyle.to_string(issues)
            )

    # ── 声明式接口 ──────────────────────────────────────────────────

    def extract(self, paragraph: Paragraph) -> dict:
        return {"text": ""}

    def render(self, document: Document) -> OxmlElement:
        """创建 OMML 段落 —— 只负责公式结构，样式由 _base() 处理。"""
        from wordformat.math.omml import latex_to_omath, latex_to_omath_para

        latex = self.content.get("latex", "")
        display = self.content.get("display", True)
        p_el = OxmlElement("w:p")

        if not latex.strip():
            return p_el

        try:
            om = latex_to_omath_para(latex) if display else latex_to_omath(latex)
            if om is not None:
                p_el.append(om)
                return p_el
        except Exception:
            pass

        # 回退：显示 LaTeX 原文
        r_el = OxmlElement("w:r")
        t_el = OxmlElement("w:t")
        t_el.text = latex
        t_el.set(qn("xml:space"), "preserve")
        r_el.append(t_el)
        p_el.append(r_el)
        return p_el

    def patch(self, paragraph: Paragraph, document: Document) -> list[str]:
        return []

    def get_alignment_text(self) -> str:
        latex = self.content.get("latex", "")
        return f"[FORMULA:{latex[:60]}]" if latex else "[FORMULA]"

    def to_value_dict(self) -> dict:
        d = dict(self.value) if isinstance(self.value, dict) else {}
        d["category"] = "formula"
        d["paragraph"] = ""
        meta = d.setdefault("meta", {})
        meta["latex"] = self.content.get("latex", "")
        meta["display"] = self.content.get("display", True)
        return d
