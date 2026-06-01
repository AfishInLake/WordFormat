"""图片 blob 提取与注册 —— 在源文档和目标文档之间传输图片二进制数据。"""

import hashlib
import io

from docx import Document
from docx.text.paragraph import Paragraph


class ImageRegistry:
    """管理图片 blob 的提取、缓存和跨文档注册。"""

    @staticmethod
    def extract_from_paragraph(paragraph: Paragraph) -> dict:
        """从含 w:drawing 的段落提取图片元数据和二进制。"""
        from docx.oxml.ns import qn as _qn

        for r_elem in paragraph._element:
            if r_elem.tag.split("}")[-1] != "r":
                continue
            drawing = r_elem.find(_qn("w:drawing"))
            if drawing is None:
                continue
            blip = drawing.find(".//" + _qn("a:blip"))
            if blip is None:
                continue
            rId = blip.get(_qn("r:embed"))
            if rId is None:
                continue
            related_part = paragraph.part.related_parts.get(rId)
            if related_part is None:
                continue
            blob = related_part.blob
            extent = drawing.find(".//" + _qn("wp:extent"))
            cx = int(extent.get("cx", 0)) if extent is not None else 0
            cy = int(extent.get("cy", 0)) if extent is not None else 0
            # 对齐方式
            pPr = paragraph._element.find(_qn("w:pPr"))
            alignment = "center"
            if pPr is not None:
                jc = pPr.find(_qn("w:jc"))
                if jc is not None:
                    val = jc.get(_qn("w:val"), "center")
                    jc_map = {"left": "left", "right": "right", "both": "justify"}
                    alignment = jc_map.get(val, "center")
            return {
                "blob": blob,
                "rId": rId,
                "source": related_part.partname,
                "width_emu": cx,
                "height_emu": cy,
                "sha256": hashlib.sha256(blob).hexdigest(),
                "alignment": alignment,
            }
        return {
            "blob": None,
            "rId": None,
            "source": "",
            "width_emu": 0,
            "height_emu": 0,
            "sha256": "",
            "alignment": "center",
        }

    @staticmethod
    def register_blob(document: Document, blob: bytes) -> str:
        """将图片 blob 注册到目标文档，返回 rId。重复 blob 复用已有关系。"""
        stream = io.BytesIO(blob)
        rId, _ = document.part.get_or_add_image(stream)
        return rId

    @staticmethod
    def transfer_image(source_para: Paragraph, target_doc: Document) -> str | None:
        """将图片从源段落传输到目标文档，返回目标文档中的 rId。"""
        info = ImageRegistry.extract_from_paragraph(source_para)
        if info["blob"] is None:
            return None
        return ImageRegistry.register_blob(target_doc, info["blob"])
