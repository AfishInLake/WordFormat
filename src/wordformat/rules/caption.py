#! /usr/bin/env python
# @Time    : 2026/1/11 19:43
# @Author  : afish
# @File    : caption.py


from wordformat.config.datamodel import (
    CaptionNumberingConfig,
    FiguresConfig,
    TablesConfig,
)
from wordformat.rules.node import FormatNode
from wordformat.style.check_format import CharacterStyle, ParagraphStyle
from wordformat.utils import parse_caption_text


def _replace_paragraph_text(paragraph, new_text: str) -> None:
    """替换段落全部文本，保留第一个 run 的格式，清除其余 run。"""
    runs = paragraph.runs
    if not runs:
        return
    runs[0].text = new_text
    for run in runs[1:]:
        run.text = ""


def _check_caption_numbering(
    node: FormatNode,
    document,
    expected_label: str,
    numbering_cfg: CaptionNumberingConfig,
) -> None:
    """检查题注编号格式，有问题则添加批注。"""
    paragraph = node.paragraph
    if not paragraph:
        return

    chapter = node.value.get("chapter_number", 0)
    seq = node.value.get("sequence_number", 0)
    separator = numbering_cfg.separator

    text = paragraph.text.strip()
    if not text:
        return

    parsed = parse_caption_text(text)
    if parsed is None:
        node.add_comment(document, paragraph.runs, f"题注格式无法识别：'{text[:50]}'")
        return

    issues = []
    if parsed["label"] != expected_label:
        issues.append(f"题注标签应为'{expected_label}'，当前为'{parsed['label']}'")
    # 检查标签与编号之间的空格（跳过续前缀再检查）
    check_text = text[1:].lstrip() if parsed.get("is_continued") else text
    label_with_space = check_text.startswith(f"{expected_label} ")
    if label_with_space != numbering_cfg.label_number_space:
        want = "有空格" if numbering_cfg.label_number_space else "无空格"
        issues.append(f"标签与编号间应为{want}")
    ch = parsed.get("chapter_num")
    if ch is not None and ch != chapter:
        issues.append(f"章节号应为{chapter}，当前为{ch}")
    if parsed["separator"] != separator:
        issues.append(f"分隔符应为'{separator}'，当前为'{parsed['separator']}'")
    num = parsed.get("number_num")
    if num is not None and num != seq:
        issues.append(f"题注编号应为{seq}，当前为{num}")

    if issues:
        node.add_comment(document, paragraph.runs, "；".join(issues))


def _apply_caption_numbering(
    node: FormatNode,
    expected_label: str,
    numbering_cfg: CaptionNumberingConfig,
) -> None:
    """修正题注编号：按正确格式重写题注文本。"""
    paragraph = node.paragraph
    if not paragraph:
        return

    chapter = node.value.get("chapter_number", 0)
    seq = node.value.get("sequence_number", 0)
    separator = numbering_cfg.separator

    text = paragraph.text.strip()
    if not text:
        return

    parsed = parse_caption_text(text)
    name = parsed["name"] if parsed else text
    is_continued = parsed.get("is_continued", False) if parsed else False
    label_text = f"续{expected_label}" if is_continued else expected_label
    label_part = f"{label_text} " if numbering_cfg.label_number_space else label_text
    new_text = f"{label_part}{chapter}{separator}{seq} {name}"
    _replace_paragraph_text(paragraph, new_text)


class CaptionFigure(FormatNode[FiguresConfig]):
    """题注-图片"""

    NODE_TYPE = "figures"
    CONFIG_MODEL = FiguresConfig
    CONFIG_PATH = "figures"

    def _base(self, doc, p: bool, r: bool):
        cfg = self.pydantic_config
        # 段落样式
        ps = ParagraphStyle.from_config(cfg)
        if p:
            paragraph_issues = ps.diff_from_paragraph(self.paragraph)
        else:
            paragraph_issues = ps.apply_to_paragraph(self.paragraph)

        # 字符样式
        cstyle = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )

        # 检查每个 run 的字符格式
        for run in self.paragraph.runs:
            if r:
                diff_result = cstyle.diff_from_run(run)
            else:
                diff_result = cstyle.apply_to_run(run)
            if diff_result:
                self.add_comment(
                    doc=doc, runs=run, text=CharacterStyle.to_string(diff_result)
                )

        # 检查段落格式差异
        if paragraph_issues:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=ParagraphStyle.to_string(paragraph_issues),
            )

        # 题注编号校验/修正（仅当 value 为 dict 且有 _numbering_cfg 时才执行）
        if isinstance(self.value, dict):
            numbering_cfg = self.value.get("_numbering_cfg")
            if numbering_cfg and numbering_cfg.enabled:
                prefix = cfg.caption_prefix or "图"
                if p:
                    _check_caption_numbering(self, doc, prefix, numbering_cfg)
                else:
                    _apply_caption_numbering(self, prefix, numbering_cfg)


class CaptionTable(FormatNode[TablesConfig]):
    """题注-表格"""

    NODE_TYPE = "tables"
    CONFIG_MODEL = TablesConfig
    CONFIG_PATH = "tables"

    def _base(self, doc, p: bool, r: bool):
        cfg = self.pydantic_config
        # 段落样式
        ps = ParagraphStyle.from_config(cfg)
        if p:
            paragraph_issues = ps.diff_from_paragraph(self.paragraph)
        else:
            paragraph_issues = ps.apply_to_paragraph(self.paragraph)

        # 字符样式
        cstyle = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )

        # 检查每个 run 的字符格式
        for run in self.paragraph.runs:
            if r:
                diff_result = cstyle.diff_from_run(run)
            else:
                diff_result = cstyle.apply_to_run(run)
            if diff_result:
                self.add_comment(
                    doc=doc, runs=run, text=CharacterStyle.to_string(diff_result)
                )

        # 检查段落格式差异
        if paragraph_issues:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=ParagraphStyle.to_string(paragraph_issues),
            )

        # 题注编号校验/修正（仅当 value 为 dict 且有 _numbering_cfg 时才执行）
        if isinstance(self.value, dict):
            numbering_cfg = self.value.get("_numbering_cfg")
            if numbering_cfg and numbering_cfg.enabled:
                prefix = cfg.caption_prefix or "表"
                if p:
                    _check_caption_numbering(self, doc, prefix, numbering_cfg)
                else:
                    _apply_caption_numbering(self, prefix, numbering_cfg)
