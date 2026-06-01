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
    """图片节点 —— 声明式 extract / render / patch。"""

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
        # 声明式内容字段（初始化默认值）
        self.content.setdefault("image_source", "")
        self.content.setdefault("image_blob", None)
        self.content.setdefault("image_sha256", "")
        self.content.setdefault("width_emu", 0)
        self.content.setdefault("height_emu", 0)
        self.content.setdefault("alignment", "center")

    def load_config(self, full_config) -> None:
        """图片节点暂不绑定 Pydantic 格式配置。"""

    def _base(self, doc: Document, p: bool, r: bool) -> None:  # noqa: FBT001
        """图片格式由 render/patch 管理，_base 为空操作。"""

    # ── 声明式接口 ──────────────────────────────────────────────────

    def extract(self, paragraph: Paragraph) -> dict:
        """从 docx 图片段落提取 blob、尺寸、对齐方式。"""
        from wordformat.orchestration.image_registry import ImageRegistry

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
        """从 content 中的 blob / 文件路径 / 尺寸创建完整 w:drawing 段落。

        优先级：image_blob > image_path > 占位符
        """
        from wordformat.orchestration.image_registry import ImageRegistry

        blob = self.content.get("image_blob")
        cx = self.content.get("width_emu", 0)
        cy = self.content.get("height_emu", 0)
        alignment = self.content.get("alignment", "center")

        # 如果提供了文件路径，从文件读取 blob
        if blob is None:
            path = self.content.get("image_path", "")
            if path:
                import os as _os

                if _os.path.exists(path):
                    with open(path, "rb") as _f:
                        blob = _f.read()
                    # 更新 sha256
                    import hashlib as _hl

                    self.content["image_sha256"] = _hl.sha256(blob).hexdigest()
                    self.content["image_blob"] = blob

        if blob is None:
            return self._build_placeholder(cx, cy, alignment)

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

        return _build_drawing_paragraph(rId, cx, cy, alignment)

    def patch(self, paragraph: Paragraph, document: Document) -> list[str]:
        """比较虚拟内容与真实段落，必要时重建。"""
        current = self.extract(paragraph)
        changes = []

        # 比较图片内容（sha256）
        current_sha = current.get("image_sha256", "")
        target_sha = self.content.get("image_sha256", "")
        if target_sha and current_sha != target_sha:
            new_elem = self.render(document)
            parent = paragraph._element.getparent()
            if parent is not None:
                parent.replace(paragraph._element, new_elem)
            old_src = current.get("image_source", "?")
            new_src = self.content.get("image_source", "?")
            changes.append(f"图片已替换: {old_src} → {new_src}")
            return changes

        # 仅比较尺寸
        if current.get("width_emu") != self.content.get("width_emu") or current.get(
            "height_emu"
        ) != self.content.get("height_emu"):
            _update_extent(
                paragraph,
                self.content.get("width_emu", 0),
                self.content.get("height_emu", 0),
            )
            changes.append("图片尺寸已更新")

        # 比较对齐
        if current.get("alignment", "center") != self.content.get(
            "alignment", "center"
        ):
            _update_alignment(paragraph, self.content.get("alignment", "center"))
            changes.append("图片对齐已更新")

        return changes

    def get_alignment_text(self) -> str:
        """返回稳定的内容摘要（图片 blob 的 SHA256），不依赖 XML 指纹。"""
        sha = self.content.get("image_sha256", "")
        if sha:
            return f"[IMAGE:{sha}]"
        # 回退：用 image_source 作为弱标识
        source = self.content.get("image_source", "")
        if source:
            return f"[IMAGE:{source}]"
        return ""

    def _build_placeholder(self, cx: int, cy: int, alignment: str) -> OxmlElement:
        """blob 不可用时创建占位符段落。"""
        if cx <= 0:
            cx = 914400
        if cy <= 0:
            cy = 685800
        p_el = OxmlElement("w:p")
        pPr = OxmlElement("w:pPr")
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), _jc_val(alignment))
        pPr.append(jc)
        ind = OxmlElement("w:ind")
        ind.set(qn("w:firstLine"), "0")
        ind.set(qn("w:left"), "0")
        pPr.append(ind)
        p_el.append(pPr)
        # r + drawing 占位
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
        """序列化 —— 排除二进制 blob，元数据放入 meta。"""
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


def _jc_val(alignment: str) -> str:
    _map = {"left": "left", "center": "center", "right": "right", "justify": "both"}
    return _map.get(alignment, "center")


def _build_drawing_paragraph(rId: str, cx: int, cy: int, alignment: str) -> OxmlElement:
    """构建包含完整 w:drawing 结构的 w:p 元素。"""
    PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

    p_el = OxmlElement("w:p")

    # w:pPr → w:jc + w:ind（图片不缩进）
    pPr = OxmlElement("w:pPr")
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), _jc_val(alignment))
    pPr.append(jc)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:firstLine"), "0")
    ind.set(qn("w:left"), "0")
    pPr.append(ind)
    p_el.append(pPr)

    # w:r
    r_el = OxmlElement("w:r")
    drawing = OxmlElement("w:drawing")

    # wp:inline
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

    # a:graphic（nsmap 已含 a/pic/wp/r 前缀，lxml 自动处理声明）
    graphic = OxmlElement("a:graphic")
    graphicData = OxmlElement("a:graphicData")
    graphicData.set("uri", PIC_NS)
    graphic.append(graphicData)

    # pic:pic
    pic = OxmlElement("pic:pic")

    nvPicPr = OxmlElement("pic:nvPicPr")
    cNvPr = OxmlElement("pic:cNvPr")
    cNvPr.set("id", "0")
    cNvPr.set("name", "Picture")
    nvPicPr.append(cNvPr)
    cNvPicPr = OxmlElement("pic:cNvPicPr")
    nvPicPr.append(cNvPicPr)
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


def _update_alignment(paragraph: Paragraph, alignment: str) -> None:
    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        paragraph._element.insert(0, pPr)
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        pPr.append(jc)
    jc.set(qn("w:val"), _jc_val(alignment))
