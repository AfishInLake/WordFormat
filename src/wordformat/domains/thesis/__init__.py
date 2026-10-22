#! /usr/bin/env python
# @Time    : 2026/9/28
# @Author  : afish
# @File    : thesis.py
"""论文领域（thesis）装配。

论文特有的处理逻辑全部集中在此，框架（pipeline / hooks / handlers）不再直接
依赖论文规则类。流水线在各阶段 emit 标准事件，本模块注册下列领域回调：

- ``thesis.tree_normalize``（on_tree_normalize）：中英文摘要 / 参考文献 / 致谢子树提升；
- ``thesis.caption_numbering``（on_format_begin + before_node_format）：
  题注章节号/顺序号注入与正文 chapter_number 维护；
- ``thesis.summary_build``（on_summary_build）：检测报告摘要统计。

新增文档类型（公文、计划书……）时在 domains/ 下新建模块，
用 @register_domain 注册后由编排入口装配（见 pipeline/orchestrate.py）。
"""

from __future__ import annotations

import re

from wordformat.domains import register_domain
from wordformat.hooks import HookRegistry, hooks

# 导入即注册论文域分类后处理整链（框架侧分派: get_postprocess("thesis")）
from . import postprocess as _postprocess  # noqa: F401, E402

# 领域回调名（预设可用 hooks.unregister(事件, 名字) 禁用）
NAME_TREE_NORMALIZE = "thesis.tree_normalize"
NAME_CAPTION_NUMBERING = "thesis.caption_numbering"
NAME_SUMMARY_BUILD = "thesis.summary_build"


# ----------------------------------------------------------------------
# 1) on_tree_normalize —— 论文子树提升
# ----------------------------------------------------------------------
_MAPPINGS = None


def _thesis_mappings():
    """论文子树提升映射（惰性 import 论文规则类）。"""
    global _MAPPINGS
    if _MAPPINGS is None:
        from wordformat.domains.thesis.rules.abstract import (
            AbstractContentCN,
            AbstractContentEN,
            AbstractTitleCN,
            AbstractTitleEN,
        )
        from wordformat.domains.thesis.rules.acknowledgement import (
            Acknowledgements,
            AcknowledgementsCN,
        )
        from wordformat.domains.thesis.rules.references import (
            ReferenceEntry,
            References,
        )

        _MAPPINGS = [
            (AbstractTitleCN, AbstractContentCN),
            (AbstractTitleEN, AbstractContentEN),
            (References, ReferenceEntry),
            (Acknowledgements, AcknowledgementsCN),
        ]
    return _MAPPINGS


def _normalize_subtrees(data: dict):
    """Abstract/References/Acknowledgements 子树内的 body_text 提升为内容节点。

    事件数据：root_node、config、check。
    """
    root_node = data.get("root_node")
    if root_node is None:
        return
    from wordformat.structure.utils import promote_bodytext_in_subtrees_of_type

    for parent_cls, target_cls in _thesis_mappings():
        promote_bodytext_in_subtrees_of_type(
            root_node, parent_type=parent_cls, target_type=target_cls
        )


# ----------------------------------------------------------------------
# 2) on_format_begin + before_node_format —— 题注章节号注入
# ----------------------------------------------------------------------
# 遍历状态：章节号、图表顺序号计数器（before_node_format 按 DFS 顺序调用，
# 与流水线遍历顺序一致；on_format_begin 时重置）
_state = {"chapter": 0, "fig": {}, "tab": {}}


def _reset_caption_state(data: dict):
    """格式化开始：重置题注编号所需的章节/图序/表序状态。"""
    _state.update(chapter=0, fig={}, tab={})


# 目录条目形态：编号前缀 + 制表符页码（"1 绪论\t1"）或点线页码（"1.1 引言……7"）。
# 模型常把目录行判为 heading_level_1，章节计数必须排除，否则题注章节号整体偏移。
_TOC_ENTRY_RE = re.compile(
    r"^\s*(?:第[一二三四五六七八九十百\d]+[章节篇]|\d+(?:\.\d+)*|[一二三四五六七八九十]+、)"
    r".*?(?:\t\d+|[.．·]{2,}\s*\d+)\s*$"
)


def _inject_caption_numbering(data: dict):
    """题注节点注入章节号与顺序号，其余节点 setdefault 当前章节号。

    原 FormattingExecutionStage 硬编码逻辑：遇到一级标题递增章节号；
    题注为"续图/续表"时保留原文编号（不递增计数器），否则按章节计数；
    正文节点注入 chapter_number（BodyText 第一章引用上标需要）。
    """
    node = data.get("node")
    if node is None:
        return
    value = node.value if isinstance(node.value, dict) else {}
    category = value.get("category", "")

    # 遇到一级标题时递增章节号（目录条目不递增，见 _TOC_ENTRY_RE）
    if category == "heading_level_1":
        text = node.paragraph.text if node.paragraph else ""
        if not isinstance(text, str):
            text = ""
        if not _TOC_ENTRY_RE.match(text):
            _state["chapter"] += 1
    chapter = _state["chapter"]

    # 题注节点：注入章节号 + 顺序号（续表/续图保留原标题注编号）
    if category in ("caption_figure", "caption_table"):
        from wordformat.utils import parse_caption_text

        text = node.paragraph.text.strip() if node.paragraph else ""
        parsed = parse_caption_text(text)
        if (
            parsed
            and parsed.get("is_continued")
            and parsed.get("chapter_num") is not None
            and parsed.get("number_num") is not None
        ):
            chapter = parsed["chapter_num"]
            seq = parsed["number_num"]
        else:
            chapter = chapter if chapter > 0 else 0
            counter = _state["fig"] if category == "caption_figure" else _state["tab"]
            counter[chapter] = counter.get(chapter, 0) + 1
            seq = counter[chapter]
        value["chapter_number"] = chapter
        value["sequence_number"] = seq
    else:
        # 给所有节点注入章节号（BodyText 第一章引用上标需要）
        value.setdefault("chapter_number", chapter)


# ----------------------------------------------------------------------
# 3) on_summary_build —— 论文检测报告摘要
# ----------------------------------------------------------------------
def _build_summary(data: dict):
    """生成检测报告摘要文本（仅 check 模式）。

    事件数据：root_node、document、config、check。
    返回 {"summary": text} 供 SummaryGenerationStage 写入批注。
    """
    if not data.get("check", False):
        return None
    root_node = data.get("root_node")
    if root_node is None:
        return None
    summary = _compose_summary(root_node, data.get("document"), data.get("config"))
    return {"summary": summary} if summary else None


def _compose_summary(root_node, document, config_model) -> str:
    """遍历树和错误统计，生成论文检测报告摘要文本（字数/关键词/参考文献规范校验）。"""
    from wordformat.domains.thesis.rules.keywords import KeywordsCN, KeywordsEN
    from wordformat.rules.node import FormatNode
    from wordformat.utils import count_chinese_chars, has_chinese

    stats = FormatNode._error_stats
    total = stats["total"]

    def _collect_section(node, sections):
        cls_name = type(node).__name__
        para = node.paragraph
        if para and para.text.strip():
            text = para.text.strip()
            if cls_name == "AbstractContentCN":
                cn_chars = count_chinese_chars(text)
                if cn_chars:
                    sections["abstract_cn_chars"] = (
                        sections.get("abstract_cn_chars", 0) + cn_chars
                    )
            elif cls_name == "AbstractContentEN":
                sections["abstract_en_words"] = sections.get(
                    "abstract_en_words", 0
                ) + len(text.split())
            elif cls_name == "KeywordsCN":
                kws = KeywordsCN.extract_keywords(text)
                if kws:
                    sections["keyword_cn_count"] = len(kws)
            elif cls_name == "KeywordsEN":
                kws = KeywordsEN.extract_keywords(text)
                if kws:
                    sections["keyword_en_count"] = len(kws)
            elif cls_name == "ReferenceEntry":
                if has_chinese(text):
                    sections["ref_cn"] = sections.get("ref_cn", 0) + 1
                else:
                    sections["ref_en"] = sections.get("ref_en", 0) + 1
        # 处理混合节点：AbstractTitleContentCN/EN 的子 BodyText 是摘要正文
        if cls_name == "AbstractTitleContentCN":
            for child in node.children:
                cp = child.paragraph
                if cp and cp.text.strip():
                    cnt = count_chinese_chars(cp.text.strip())
                    sections["abstract_cn_chars"] = (
                        sections.get("abstract_cn_chars", 0) + cnt
                    )
        elif cls_name == "AbstractTitleContentEN":
            for child in node.children:
                cp = child.paragraph
                if cp and cp.text.strip():
                    cnt = len(cp.text.strip().split())
                    sections["abstract_en_words"] = (
                        sections.get("abstract_en_words", 0) + cnt
                    )
        for child in node.children:
            _collect_section(child, sections)

    sections: dict = {}
    _collect_section(root_node, sections)

    # 计算万字差错率
    total_chars = (
        sum(len(p.text) for p in document.paragraphs if p.text and p.text.strip())
        if document
        else 0
    )
    error_rate = (total / max(total_chars, 1)) * 10000 if total else 0

    # 模板名（从 config 读取）
    if config_model:
        template_name = getattr(config_model, "template_name", None) or "未知模板"
    else:
        template_name = "未知模板"

    lines = [
        "检测结果：",
        f"检测模板：《{template_name}》",
        f"检测错误数：{total}，万字差错率：{error_rate:.1f}",
        f"错误：{stats.get('错误', 0)}，提醒：{stats.get('提醒', 0)}",
    ]

    # 字数问题
    word_issues = []
    if sections.get("abstract_cn_chars"):
        word_issues.append(
            f"中文摘要：规范：300字左右，原文：{sections['abstract_cn_chars']}字"
        )
    if sections.get("abstract_en_words"):
        word_issues.append(
            f"英文摘要：规范：300字左右，原文：{sections['abstract_en_words']}词"
        )
    if sections.get("keyword_cn_count"):
        word_issues.append(
            f"中文关键词：规范：3-5个，原文：{sections['keyword_cn_count']}个"
        )
    if sections.get("keyword_en_count"):
        word_issues.append(
            f"英文关键词：规范：3-5个，原文：{sections['keyword_en_count']}个"
        )
    ref_cn = sections.get("ref_cn", 0)
    ref_en = sections.get("ref_en", 0)
    if ref_cn or ref_en:
        word_issues.append(
            f"参考文献：规范：不少于15条，原文：中文{ref_cn}条;外文{ref_en}条"
        )
    if word_issues:
        lines.append("字数问题：")
        lines.extend(word_issues)

    lines.append("说明：")
    lines.append(
        "1.请确保文档中正确使用换行符，硬回车（Enter）：指换行且生成新段落；软回车（Shift+Enter）：指换行但不生成新段落。"
    )
    lines.append("2.图片请使用“嵌入型”环绕方式，表格为无环绕方式。")
    lines.append("3.提醒不计算错误。")

    return "\n".join(lines)


# ----------------------------------------------------------------------
# 注册
# ----------------------------------------------------------------------
@register_domain("thesis")
def register_handlers(registry: HookRegistry = hooks) -> None:
    """注册论文领域全部回调（幂等：同名回调不重复注册）。"""
    existing = {name for name, _cb in registry.registered_callbacks()}

    if NAME_TREE_NORMALIZE not in existing:
        registry.register("on_tree_normalize", _normalize_subtrees, NAME_TREE_NORMALIZE)
    if NAME_CAPTION_NUMBERING not in existing:
        registry.register(
            "on_format_begin", _reset_caption_state, NAME_CAPTION_NUMBERING
        )
        registry.register(
            "before_node_format", _inject_caption_numbering, NAME_CAPTION_NUMBERING
        )
    if NAME_SUMMARY_BUILD not in existing:
        registry.register("on_summary_build", _build_summary, NAME_SUMMARY_BUILD)
