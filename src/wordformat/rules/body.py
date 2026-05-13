#! /usr/bin/env python
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : body.py

import re
from copy import deepcopy

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from wordformat.config.datamodel import BodyTextConfig
from wordformat.rules.node import FormatNode
from wordformat.style.check_format import CharacterStyle, ParagraphStyle

# 匹配正文中的参考文献交叉引用标记
# 支持: [1] [1,2] [1-3] [1,2,3-5] [1，2] [1、2] [1, 2]
_CITATION_PATTERN = re.compile(r"\[[\d\-,—、，\s]+\]")


class BodyText(FormatNode[BodyTextConfig]):
    """正文节点"""

    NODE_TYPE = "body_text"
    CONFIG_MODEL = BodyTextConfig
    CONFIG_PATH = "body_text"

    def apply_replace(self, doc=None) -> bool:
        """文本替换后清除引用标记格式。

        正文段落中的引用标记（如 [1]）通常带有上标/下标格式，
        替换文本按字符数分配后，非引用位置的字符不应保留这些特殊格式。
        """
        replaced = super().apply_replace(doc)
        if not replaced or self.paragraph is None:
            return replaced

        for run in self.paragraph.runs:
            rPr = run._element.find(qn("w:rPr"))
            if rPr is not None:
                vertAlign = rPr.find(qn("w:vertAlign"))
                if vertAlign is not None:
                    rPr.remove(vertAlign)

        return replaced

    def _base(self, doc, p: bool, r: bool):
        """检查正文段落的字符与段落格式是否符合规范。"""
        cfg = self.pydantic_config
        ps = ParagraphStyle.from_config(cfg)
        if p:
            paragraph_issues = ps.diff_from_paragraph(self.paragraph)
        else:
            paragraph_issues = ps.apply_to_paragraph(self.paragraph)

        cstyle = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )

        for run in self.paragraph.runs:
            if r:
                diff_result = cstyle.diff_from_run(run)
            else:
                diff_result = cstyle.apply_to_run(run)
            if diff_result:
                self.add_comment(
                    doc=doc, runs=run, text=CharacterStyle.to_string(diff_result)
                )

        if paragraph_issues:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=ParagraphStyle.to_string(paragraph_issues),
            )

        return []

    def apply_format(self, doc):
        """格式化正文段落，之后将引用标记自动设为上标。"""
        super().apply_format(doc)
        self._apply_citation_superscript()

    # ------------------------------------------------------------------
    # 引用上标
    # ------------------------------------------------------------------

    def _apply_citation_superscript(self):
        """扫描正文段落，将 [1] [1,2] [1-3] 等引用标记设置为上标。

        在 apply_replace 已清除 vertAlign 的前提下，重新给引用位置的 run
        添加 w:vertAlign=superscript。若引用跨多个 run 或与正文混在同一 run，
        会先分割 run 再设置上标。
        """
        para = self.paragraph
        if para is None:
            return

        runs = list(para.runs)
        full_text = "".join(r.text for r in runs if r.text)
        citations = [(m.start(), m.end()) for m in _CITATION_PATTERN.finditer(full_text)]
        if not citations:
            return

        # ---- 1. 分割 run，使每个引用标记独占 run ----
        para_elem = para._element
        r_elems = para_elem.findall(qn("w:r"))
        cum = 0
        for r_elem in r_elems:
            t_elem = r_elem.find(qn("w:t"))
            if t_elem is None or t_elem.text is None:
                continue
            text = t_elem.text
            r_start = cum
            r_end = cum + len(text)

            # 找出此 run 内所有需要切分的位置（引用起止边界）
            split_points = set()
            for c_start, c_end in citations:
                if r_start < c_start < r_end:
                    split_points.add(c_start - r_start)
                if r_start < c_end < r_end:
                    split_points.add(c_end - r_start)

            if split_points:
                rPr = r_elem.find(qn("w:rPr"))
                segments = []
                last = 0
                for pos in sorted(split_points):
                    segments.append((last, pos))
                    last = pos
                segments.append((last, len(text)))

                # 更新第一个片段（原地修改）
                t_elem.text = text[segments[0][0]:segments[0][1]]
                if t_elem.text and t_elem.text != t_elem.text.strip():
                    t_elem.set(qn("xml:space"), "preserve")

                # 为后续片段创建新 <w:r>
                ref = r_elem
                for seg_start, seg_end in segments[1:]:
                    seg_text = text[seg_start:seg_end]
                    if not seg_text:
                        continue
                    new_r = OxmlElement("w:r")
                    if rPr is not None:
                        new_r.append(deepcopy(rPr))
                    new_t = OxmlElement("w:t")
                    new_t.text = seg_text
                    if seg_text != seg_text.strip():
                        new_t.set(qn("xml:space"), "preserve")
                    new_r.append(new_t)
                    ref.addnext(new_r)
                    ref = new_r

            cum += len(text)

        # ---- 2. 给完全位于引用范围内的 run 设置上标 ----
        r_elems = para_elem.findall(qn("w:r"))
        cum = 0
        for r_elem in r_elems:
            t_elem = r_elem.find(qn("w:t"))
            if t_elem is None or t_elem.text is None:
                continue
            text = t_elem.text
            r_start = cum
            r_end = cum + len(text)

            for c_start, c_end in citations:
                if r_start >= c_start and r_end <= c_end:
                    rPr = r_elem.find(qn("w:rPr"))
                    if rPr is None:
                        rPr = OxmlElement("w:rPr")
                        r_elem.insert(0, rPr)
                    if rPr.find(qn("w:vertAlign")) is None:
                        va = OxmlElement("w:vertAlign")
                        va.set(qn("w:val"), "superscript")
                        rPr.append(va)
                    break

            cum += len(text)
