#! /usr/bin/env python
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : body.py

import re
from copy import deepcopy

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from wordformat.config.dotdict import BASE_FORMAT, deep_merge
from wordformat.rules.node import FormatNode
from wordformat.structure.registry import register

# 匹配正文中的参考文献交叉引用标记
# 支持: [1] [1,2] [1-3] [1,2,3-5] [1，2] [1、2] [1, 2]
_CITATION_PATTERN = re.compile(r"\[[\d\-,—、，\s]+\]")

# 半角标点 → 对应全角标点
_HALF_TO_FULL = {
    ",": "，",
    ".": "。",
    ";": "；",
    ":": "：",
    "(": "（",
    ")": "）",
    "[": "【",
    "]": "】",
    '"': "“",
    "'": "‘",
    "!": "！",
    "?": "？",
}
# 匹配中文上下文中的半角标点（段首/中文后 + 标点 + 中文/段尾）
_PUNCT_PATTERN = re.compile(r"(^|[一-鿿])([,.;:()\[\]\"\'\!\?])([一-鿿]|$)")


def _split_run_at(paragraph, start: int, end: int):
    """拆分段落 run，返回只含 [start, end) 字符的独立 run。"""
    from docx.oxml import OxmlElement

    para_elem = paragraph._element
    cum = 0
    for r_elem in para_elem.findall(qn("w:r")):
        t_elem = r_elem.find(qn("w:t"))
        if t_elem is None or t_elem.text is None:
            continue
        rl = len(t_elem.text)
        r_start, r_end = cum, cum + rl
        if r_start <= start and end <= r_end:
            i0 = start - r_start
            i1 = end - r_start
            before_text = t_elem.text[:i0]
            punct_text = t_elem.text[i0:i1]
            after_text = t_elem.text[i1:]
            # 修改原 run 为标点前的文本
            t_elem.text = before_text
            if before_text and before_text != before_text.strip():
                t_elem.set(qn("xml:space"), "preserve")
            # 创建标点 run
            rPr = r_elem.find(qn("w:rPr"))
            punct_r = OxmlElement("w:r")
            if rPr is not None:
                punct_r.append(deepcopy(rPr))
            punct_t = OxmlElement("w:t")
            punct_t.text = punct_text
            punct_t.set(qn("xml:space"), "preserve")
            punct_r.append(punct_t)
            r_elem.addnext(punct_r)
            # 创建标点后的 run
            if after_text:
                after_r = OxmlElement("w:r")
                if rPr is not None:
                    after_r.append(deepcopy(rPr))
                after_t = OxmlElement("w:t")
                after_t.text = after_text
                after_t.set(qn("xml:space"), "preserve")
                after_r.append(after_t)
                punct_r.addnext(after_r)
            # 返回新 run 对象
            for r in paragraph.runs:
                if r._element is punct_r:
                    return r
            return None
        cum += rl
    return None


@register("body_text")
class BodyText(FormatNode):
    """正文节点"""

    NODE_TYPE = "body.text"
    NODE_LABEL = "正文段落"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            "paragraph": {"alignment": "两端对齐"},
            "font": {
                "chinese_font_name": "宋体",
                "english_font_name": "Times New Roman",
            },
            "rules": {"punctuation": {"enabled": True}},
        },
    )
    RULES = {"punctuation": "_check_punctuation"}

    def _check_punctuation(self, doc, rule_cfg, p: bool = False):
        """检测中文正文中的半角标点，锚在具体字符上（拆分 run）。"""
        if self.paragraph is None:
            return
        # 找出引用标记区间，跳过引用内的标点
        cite_ranges = [
            (m.start(), m.end())
            for m in _CITATION_PATTERN.finditer(self.paragraph.text)
        ]
        para_text = self.paragraph.text
        for m in _PUNCT_PATTERN.finditer(para_text):
            p0, p1 = m.start(2), m.end(2)
            if any(cs <= p0 < ce for cs, ce in cite_ranges):
                continue
            half = m.group(2)
            full = _HALF_TO_FULL.get(half, half)
            # 拆分 run，使标点独占一个 run
            target_run = _split_run_at(self.paragraph, p0, p1)
            if target_run is None:
                target_run = (
                    self.paragraph.runs[0]
                    if self.paragraph.runs
                    else self.paragraph.runs
                )
            msg = (
                f"{self.NODE_LABEL}-提醒-标点问题："
                f'使用了半角英文标点符号"{half}(英文)"，'
                f'规范：应使用全角中文标点符号"{full}(中文)"'
            )
            self.add_comment(doc=doc, runs=target_run, text=msg)

    def apply_replace(self, doc=None) -> bool:
        """文本替换后清除引用标记格式。

        正文段落中的引用标记（如 [1]）通常带有上标/下标格式，
        替换文本按字符数分配后，非引用位置的字符不应保留这些特殊格式。
        """
        replaced = super().apply_replace(doc)
        if not replaced or self.paragraph is None:
            return replaced

        for run in self.paragraph.runs:
            # superscript is not None ⇔ 该 run 存在 w:vertAlign（上标或下标）
            if run.font.superscript is not None:
                run.font.superscript = None  # 移除 w:vertAlign

        return replaced

    def apply_format(self, doc):
        """格式化正文段落，引用标记上标仅限第一章。"""
        super().apply_format(doc)
        chapter = (
            self.value.get("chapter_number", 0) if isinstance(self.value, dict) else 0
        )
        if chapter == 1:
            self._apply_citation_superscript()

    # ------------------------------------------------------------------
    # 引用上标（仅第一章）
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
        citations = [
            (m.start(), m.end()) for m in _CITATION_PATTERN.finditer(full_text)
        ]
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
                t_elem.text = text[segments[0][0] : segments[0][1]]
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
                    # CT_R.get_or_add_rPr() → CT_RPr.superscript 官方封装
                    r_elem.get_or_add_rPr().superscript = True
                    break

            cum += len(text)
