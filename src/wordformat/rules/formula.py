"""FormulaNode — LaTeX 公式渲染为 OMML。"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from wordformat.rules.node import FormatNode

if TYPE_CHECKING:
    from docx import Document
    from docx.text.paragraph import Paragraph

MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MATH_FONT = "Cambria Math"
_DEFAULT_SZ = "22"


def _w(tag: str) -> str:
    return f"{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{tag}"


def _m(tag: str) -> str:
    return f"{{{MATH_NS}}}{tag}"


class FormulaNode(FormatNode):
    """公式节点 —— LaTeX → OMML，支持行内和块级公式。"""

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
        self.content.setdefault("display", True)  # True=块级 $$, False=行内 $

    def load_config(self, full_config) -> None:
        """公式暂不绑定格式配置。"""

    def _base(self, doc: Document, p: bool, r: bool) -> None:  # noqa: FBT001
        """公式格式由 OMML 自行控制。"""

    def check_format(self, doc: Document):
        """跳过格式校验。"""

    def apply_format(self, doc: Document):
        """跳过段落级格式操作。"""

    def extract(self, paragraph: Paragraph) -> dict:
        return {"text": ""}

    def render(self, document: Document) -> OxmlElement:
        from wordformat.math.omml import latex_to_omath, latex_to_omath_para

        latex = self.content.get("latex", "")
        display = self.content.get("display", True)

        p_el = OxmlElement("w:p")

        # 块级公式：居中 + 零缩进
        if display:
            pPr = OxmlElement("w:pPr")
            jc = OxmlElement("w:jc")
            jc.set(qn("w:val"), "center")
            pPr.append(jc)
            ind = OxmlElement("w:ind")
            ind.set(qn("w:firstLine"), "0")
            ind.set(qn("w:left"), "0")
            pPr.append(ind)
            p_el.append(pPr)

        if not latex.strip():
            return p_el

        try:
            if display:
                om = latex_to_omath_para(latex)
            else:
                om = latex_to_omath(latex)
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
