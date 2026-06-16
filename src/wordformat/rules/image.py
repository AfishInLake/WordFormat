"""ImageNode — 图片段落节点，声明式管理图片的提取、渲染与修补。"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from wordformat.rules.node import FormatNode

if TYPE_CHECKING:
    from docx import Document
    from docx.text.paragraph import Paragraph


class ImageNode(FormatNode):
    """图片节点 —— render() 管结构，_base() 管段落格式。"""

    NODE_TYPE = "image"
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
        self.type = "image"
        self.content.setdefault("image_source", "")
        self.content.setdefault("image_blob", None)
        self.content.setdefault("image_sha256", "")
        self.content.setdefault("width_emu", 0)
        self.content.setdefault("height_emu", 0)
        self.content.setdefault("alignment", "center")

    def load_config(self, full_config) -> None:
        """图片节点暂不绑定 Pydantic 格式配置，段落样式在 _base() 中统一处理。"""

    # ── 格式：与其他节点一致的 _base 模式 ──────────────────────────────

    def _base(self, doc: Document, p: bool, r: bool) -> None:  # noqa: FBT001
        """图片段落格式：居中 + 零缩进。"""
        if self.paragraph is None:
            return
        from wordformat.style.check_format import ParagraphStyle

        ps = ParagraphStyle(
            alignment="居中对齐",
            first_line_indent="0字符",
            left_indent="0字符",
        )
        if p:
            paragraph_issues = ps.diff_from_paragraph(self.paragraph)
        else:
            paragraph_issues = ps.apply_to_paragraph(self.paragraph)
        if paragraph_issues:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=ParagraphStyle.to_string(paragraph_issues),
            )

    # ── 声明式接口 ──────────────────────────────────────────────────

    def extract(self, paragraph: Paragraph) -> dict:
        from wordformat.media import ImageRegistry

        info = ImageRegistry.extract_from_paragraph(paragraph)
        return {
            "text": "",
            "image_source": info.get("source", ""),
            "image_blob": info.get("blob"),
            "image_sha256": info.get("sha256", ""),
            "width_emu": info.get("width_emu", 0),
            "height_emu": info.get("height_emu", 0),
            "alignment": info.get("alignment", "center"),
        }

    def render(self, document: Document) -> OxmlElement:
        """创建 w:drawing 段落 —— 只负责 OOXML 结构，样式由 _base() 处理。"""
        from wordformat.media import ImageRegistry

        blob = self.content.get("image_blob")
        cx = self.content.get("width_emu", 0)
        cy = self.content.get("height_emu", 0)

        # 文件路径 → blob
        if blob is None:
            path = self.content.get("image_path", "")
            if path:
                import os as _os

                if _os.path.exists(path):
                    with open(path, "rb") as _f:
                        blob = _f.read()
                    import hashlib as _hl

                    self.content["image_sha256"] = _hl.sha256(blob).hexdigest()
                    self.content["image_blob"] = blob

        if blob is None:
            return self._build_placeholder(cx, cy)

        rId = ImageRegistry.register_blob(document, blob)
        if cx <= 0 or cy <= 0:
            try:
                img_part = document.part.related_parts.get(rId)
                if img_part is not None:
                    cx = cx or img_part._img.cx
                    cy = cy or img_part._img.cy
            except Exception:
                cx = cx or 914400
                cy = cy or 685800

        return _build_drawing_paragraph(rId, cx, cy)

    def patch(self, paragraph: Paragraph, document: Document) -> list[str]:
        """比较图片内容与尺寸，必要时重建。段落样式由 _base() 处理。"""
        current = self.extract(paragraph)
        changes = []

        current_sha = current.get("image_sha256", "")
        target_sha = self.content.get("image_sha256", "")
        if target_sha and current_sha != target_sha:
            new_elem = self.render(document)
            parent = paragraph._element.getparent()
            if parent is not None:
                parent.replace(paragraph._element, new_elem)
            changes.append(
                f"图片已替换: {current.get('image_source', '?')} → {self.content.get('image_source', '?')}"
            )
            return changes

        if current.get("width_emu") != self.content.get("width_emu") or current.get(
            "height_emu"
        ) != self.content.get("height_emu"):
            _update_extent(
                paragraph,
                self.content.get("width_emu", 0),
                self.content.get("height_emu", 0),
            )
            changes.append("图片尺寸已更新")

        return changes

    def get_alignment_text(self) -> str:
        sha = self.content.get("image_sha256", "")
        if sha:
            return f"[IMAGE:{sha}]"
        source = self.content.get("image_source", "")
        if source:
            return f"[IMAGE:{source}]"
        return ""

    def _build_placeholder(self, cx: int, cy: int) -> OxmlElement:
        """blob 不可用时创建占位符段落（不含 pPr，样式由 _base 处理）。"""
        if cx <= 0:
            cx = 914400
        if cy <= 0:
            cy = 685800
        p_el = OxmlElement("w:p")
        r_el = OxmlElement("w:r")
        drawing = OxmlElement("w:drawing")
        inline = OxmlElement("wp:inline")
        extent = OxmlElement("wp:extent")
        extent.set("cx", str(cx))
        extent.set("cy", str(cy))
        inline.append(extent)
        drawing.append(inline)
        r_el.append(drawing)
        p_el.append(r_el)
        return p_el

    def to_value_dict(self) -> dict:
        d = dict(self.value) if isinstance(self.value, dict) else {}
        d["category"] = "image"
        d["paragraph"] = ""
        meta = d.setdefault("meta", {}) if isinstance(d.get("meta"), dict) else {}
        if not isinstance(meta, dict):
            meta = {}
            d["meta"] = meta
        meta["image_source"] = self.content.get("image_source", "")
        meta["image_sha256"] = self.content.get("image_sha256", "")
        meta["width_emu"] = self.content.get("width_emu", 0)
        meta["height_emu"] = self.content.get("height_emu", 0)
        meta["alignment"] = self.content.get("alignment", "center")
        d.pop("drawings", None)
        return d


# ── 辅助函数 ────────────────────────────────────────────────────────


def _build_drawing_paragraph(rId: str, cx: int, cy: int) -> OxmlElement:
    """构建 w:drawing 结构（不含 pPr，段落样式由 _base() 统一处理）。"""
    PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

    p_el = OxmlElement("w:p")
    r_el = OxmlElement("w:r")
    drawing = OxmlElement("w:drawing")

    inline = OxmlElement("wp:inline")
    inline.set("distT", "0")
    inline.set("distB", "0")
    inline.set("distL", "0")
    inline.set("distR", "0")

    extent = OxmlElement("wp:extent")
    extent.set("cx", str(cx))
    extent.set("cy", str(cy))
    inline.append(extent)

    effectExtent = OxmlElement("wp:effectExtent")
    for attr in ("l", "t", "r", "b"):
        effectExtent.set(attr, "0")
    inline.append(effectExtent)

    docPr = OxmlElement("wp:docPr")
    docPr.set("id", "1")
    docPr.set("name", "Picture")
    inline.append(docPr)

    cNvGraphicFramePr = OxmlElement("wp:cNvGraphicFramePr")
    inline.append(cNvGraphicFramePr)

    graphic = OxmlElement("a:graphic")
    graphicData = OxmlElement("a:graphicData")
    graphicData.set("uri", PIC_NS)
    graphic.append(graphicData)

    pic = OxmlElement("pic:pic")
    nvPicPr = OxmlElement("pic:nvPicPr")
    cNvPr = OxmlElement("pic:cNvPr")
    cNvPr.set("id", "0")
    cNvPr.set("name", "Picture")
    nvPicPr.append(cNvPr)
    nvPicPr.append(OxmlElement("pic:cNvPicPr"))
    pic.append(nvPicPr)

    blipFill = OxmlElement("pic:blipFill")
    blip = OxmlElement("a:blip")
    blip.set(qn("r:embed"), rId)
    blipFill.append(blip)
    blipFill.append(OxmlElement("a:stretch"))
    pic.append(blipFill)

    spPr = OxmlElement("pic:spPr")
    xfrm = OxmlElement("a:xfrm")
    off = OxmlElement("a:off")
    off.set("x", "0")
    off.set("y", "0")
    xfrm.append(off)
    ext = OxmlElement("a:ext")
    ext.set("cx", str(cx))
    ext.set("cy", str(cy))
    xfrm.append(ext)
    spPr.append(xfrm)
    prstGeom = OxmlElement("a:prstGeom")
    prstGeom.set("prst", "rect")
    spPr.append(prstGeom)
    pic.append(spPr)

    graphicData.append(pic)
    inline.append(graphic)
    drawing.append(inline)
    r_el.append(drawing)
    p_el.append(r_el)
    return p_el


def _update_extent(paragraph: Paragraph, cx: int, cy: int) -> None:
    extent = paragraph._element.find(".//" + qn("wp:extent"))
    if extent is not None:
        extent.set("cx", str(cx))
        extent.set("cy", str(cy))
