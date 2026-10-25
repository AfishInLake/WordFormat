#! /usr/bin/env python
# @Time    : 2026/1/11 19:43
# @Author  : afish
# @File    : caption.py

import re

from wordformat.config.dotdict import BASE_FORMAT, deep_merge
from wordformat.rules.node import FormatNode
from wordformat.structure.registry import register
from wordformat.utils import parse_caption_text


def _get_cfg(cfg, key, default=None):
    """兼容 dict 和对象访问。"""
    if isinstance(cfg, dict):
        return cfg.get(key, default)
    return getattr(cfg, key, default)


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
    numbering_cfg,
) -> None:
    """检查题注编号格式，有问题则添加批注。"""
    paragraph = node.paragraph
    if not paragraph:
        return

    value = node.value if isinstance(node.value, dict) else {}
    chapter = value.get("chapter_number", 0)
    seq = value.get("sequence_number", 0)
    separator = _get_cfg(numbering_cfg, "separator", ".")

    text = paragraph.text.strip()
    if not text:
        return
    # “续表：/续表”只是跨页续表的引导标签，本身不带编号，不参与编号校验
    if re.fullmatch(r"续?[图表]\s*[:：]?", text):
        return

    parsed = parse_caption_text(text)
    from wordformat.style.comments import format_comment

    target = getattr(node, "NODE_LABEL", "题注")

    if parsed is None:
        node.add_comment(
            document,
            paragraph.runs,
            format_comment(
                target, "格式错误", f"无法识别'{text[:50]}'", "正确题注格式"
            ),
        )
        return

    issues = []
    if parsed["label"] != expected_label:
        issues.append(
            format_comment(
                target,
                "标签错误",
                f"当前为'{parsed['label']}'",
                f"应为'{expected_label}'",
            )
        )
    # 检查标签与编号之间的空格（跳过续前缀再检查）
    check_text = text[1:].lstrip() if parsed.get("is_continued") else text
    label_with_space = check_text.startswith(f"{expected_label} ")
    label_space = _get_cfg(numbering_cfg, "label_number_space", False)
    if label_with_space != label_space:
        want = "有空格" if label_space else "无空格"
        issues.append(format_comment(target, "间距错误", "当前不符合", f"应为{want}"))
    ch = parsed.get("chapter_num")
    if ch is not None and ch != chapter:
        issues.append(
            format_comment(
                target,
                "章节号错误",
                f"当前为{ch}",
                f"应为{chapter}",
            )
        )
    if parsed["separator"] != separator:
        issues.append(
            format_comment(
                target,
                "分隔符错误",
                f"当前为'{parsed['separator']}'",
                f"应为'{separator}'",
            )
        )
    num = parsed.get("number_num")
    if num is not None and num != seq:
        issues.append(
            format_comment(
                target,
                "编号错误",
                f"当前为{num}",
                f"应为{seq}",
            )
        )

    if issues:
        # 题注检查 hook：预设可追加自定义题注规则（追加到 issues 或覆写）
        from wordformat.hooks import hooks

        result = hooks.emit("on_caption_check", node=node, parsed=parsed, issues=issues)
        issues = result.get("issues", issues)

    if issues:
        node.add_comment(document, paragraph.runs, "\n".join(issues))


def _apply_caption_numbering(
    node: FormatNode,
    expected_label: str,
    numbering_cfg,
) -> None:
    """修正题注编号：按正确格式重写题注文本。"""
    paragraph = node.paragraph
    if not paragraph:
        return

    value = node.value if isinstance(node.value, dict) else {}
    chapter = value.get("chapter_number", 0)
    seq = value.get("sequence_number", 0)
    separator = _get_cfg(numbering_cfg, "separator", ".")

    text = paragraph.text.strip()
    if not text:
        return

    parsed = parse_caption_text(text)
    name = parsed["name"] if parsed else text
    is_continued = parsed.get("is_continued", False) if parsed else False
    label_text = f"续{expected_label}" if is_continued else expected_label
    label_space = _get_cfg(numbering_cfg, "label_number_space", False)
    label_part = f"{label_text} " if label_space else label_text
    new_text = f"{label_part}{chapter}{separator}{seq} {name}"
    _replace_paragraph_text(paragraph, new_text)


@register("caption_figure")
class CaptionFigure(FormatNode):
    """题注-图片"""

    NODE_TYPE = "figures.caption"
    NODE_LABEL = "图注"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            "paragraph": {
                # 题注常借用正文样式，居中/两端/左对齐、0/2 字符缩进都常见
                "alignment": ["居中对齐", "两端对齐", "左对齐"],
                "first_line_indent": ["0字符", "2字符"],
                "space_before": ["0行", "0.5行"],
                "space_after": ["0行", "0.5行"],
                "line_spacingrule": [
                    "1.5倍行距",
                    "单倍行距",
                    "固定值",
                    "多倍行距",
                ],
                "builtin_style_name": "题注",
            },
            "font": {"font_size": ["五号", "小四"]},
            "caption_prefix": "图",
            "rules": {
                "caption_numbering": {
                    "enabled": True,
                    "separator": ".",
                    "label_number_space": False,
                }
            },
        },
    )
    RULES = {"caption_numbering": "_handle_caption_numbering"}

    def _para_checks_to_skip(self) -> set[str]:
        # 题注只有一行，行距数值没有视觉意义
        return {"line_spacing"}

    def _handle_caption_numbering(self, doc, rule_cfg, p: bool):
        """题注编号校验/修正"""
        prefix = self.pydantic_config.caption_prefix or "图"
        if p:
            _check_caption_numbering(self, doc, prefix, rule_cfg)
        else:
            _apply_caption_numbering(self, prefix, rule_cfg)


@register("caption_table")
class CaptionTable(FormatNode):
    """题注-表格"""

    NODE_TYPE = "tables.caption"
    NODE_LABEL = "表注"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            "paragraph": {
                "alignment": ["两端对齐", "居中对齐", "左对齐"],
                "first_line_indent": ["2字符", "0字符"],
                "space_before": ["0行", "0.5行"],
                "space_after": ["0行", "0.5行"],
                "line_spacingrule": [
                    "1.5倍行距",
                    "单倍行距",
                    "固定值",
                    "多倍行距",
                ],
                "builtin_style_name": "题注",
            },
            "font": {"font_size": ["小四", "五号"]},
            "caption_prefix": "表",
            "rules": {
                "caption_numbering": {
                    "enabled": True,
                    "separator": ".",
                    "label_number_space": False,
                }
            },
        },
    )
    RULES = {"caption_numbering": "_handle_caption_numbering"}

    def _para_checks_to_skip(self) -> set[str]:
        return {"line_spacing"}

    def _handle_caption_numbering(self, doc, rule_cfg, p: bool):
        """题注编号校验/修正"""
        prefix = self.pydantic_config.caption_prefix or "表"
        if p:
            _check_caption_numbering(self, doc, prefix, rule_cfg)
        else:
            _apply_caption_numbering(self, prefix, rule_cfg)
