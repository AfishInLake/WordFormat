"""TableNode — 表格段落节点。"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from wordformat.rules.node import FormatNode

if TYPE_CHECKING:
    from docx import Document
    from docx.text.paragraph import Paragraph


class TableNode(FormatNode):
    """表格节点 —— 从 content 中的 rows/cols/cells 渲染 w:tbl。"""

    NODE_TYPE = "table"
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
        self.type = "table"
        self.content.setdefault("rows", 1)
        self.content.setdefault("cols", 1)
        self.content.setdefault("cells", [])

    def load_config(self, full_config) -> None:
        """表格暂不绑定格式配置。"""

    def _base(self, doc: Document, p: bool, r: bool) -> None:  # noqa: FBT001
        """表格格式待实现。"""

    def check_format(self, doc: Document):
        """表格暂不校验格式。"""

    def apply_format(self, doc: Document):
        """表格暂不应用格式（跳过段落级 _clean_paragraph_edge_spaces）。"""

    def extract(self, paragraph: Paragraph) -> dict:
        return {"text": ""}

    def render(self, document: Document) -> OxmlElement:
        rows = self.content.get("rows", 1)
        cols = self.content.get("cols", 1)
        cells = self.content.get("cells", [])

        tbl = OxmlElement("w:tbl")
        # 表格属性
        tblPr = OxmlElement("w:tblPr")
        tblW = OxmlElement("w:tblW")
        tblW.set(qn("w:w"), "5000")
        tblW.set(qn("w:type"), "pct")
        tblPr.append(tblW)
        # tblGrid（python-docx 解析表格必需）
        tblGrid = OxmlElement("w:tblGrid")
        for _ in range(cols):
            gridCol = OxmlElement("w:gridCol")
            gridCol.set(qn("w:w"), str(5000 // max(cols, 1)))
            tblGrid.append(gridCol)
        tbl.append(tblGrid)
        # 边框
        tblBorders = OxmlElement("w:tblBorders")
        for border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
            border = OxmlElement(f"w:{border_name}")
            border.set(qn("w:val"), "single")
            border.set(qn("w:sz"), "4")
            border.set(qn("w:space"), "0")
            border.set(qn("w:color"), "000000")
            tblBorders.append(border)
        tblPr.append(tblBorders)
        tbl.append(tblPr)

        for r in range(rows):
            tr = OxmlElement("w:tr")
            for c in range(cols):
                tc = OxmlElement("w:tc")
                p = OxmlElement("w:p")
                r_elem = OxmlElement("w:r")
                t = OxmlElement("w:t")
                cell_text = ""
                if r < len(cells) and c < len(cells[r]):
                    cell_text = str(cells[r][c])
                t.text = cell_text
                t.set(qn("xml:space"), "preserve")
                r_elem.append(t)
                p.append(r_elem)
                tc.append(p)
                tr.append(tc)
            tbl.append(tr)
        return tbl

    def patch(self, paragraph: Paragraph, document: Document) -> list[str]:
        return []

    def get_alignment_text(self) -> str:
        cells = self.content.get("cells", [])
        summary = "|".join("|".join(str(c) for c in row) for row in cells)
        return f"[TABLE:{summary}]" if summary else "[TABLE]"

    def to_value_dict(self) -> dict:
        d = dict(self.value) if isinstance(self.value, dict) else {}
        d["category"] = "table"
        d["paragraph"] = ""
        meta = d.setdefault("meta", {})
        meta["rows"] = self.content.get("rows", 1)
        meta["cols"] = self.content.get("cols", 1)
        meta["cells"] = self.content.get("cells", [])
        return d
