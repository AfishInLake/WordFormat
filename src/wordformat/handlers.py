#! /usr/bin/env python
# @Time    : 2026/9/28
# @Author  : afish
# @File    : handlers.py
"""框架内置机制 —— 统一注册为 Hook 回调。

原有硬编码的流水线逻辑（标题编号/超链接、页眉页脚）注册为 hook handler，
由 register_builtin_handlers() 一次性注册到全局 hooks 调度器。流水线 stage
只负责在正确时机 emit，机制本身可被预设禁用（hooks.unregister('builtin.*')）
或在其前后追加自定义逻辑。

论文特有的机制（子树提升、题注编号注入、检测摘要统计）已迁出至
domains/thesis.py，由文档领域装配，不再属于框架内置机制。

内置回调命名约定：``builtin.<机制名>``。
"""

from __future__ import annotations

from typing import List

from loguru import logger

from wordformat.hooks import HookRegistry, hooks

# 内置机制回调名（预设可用 hooks.unregister(event, name) 禁用）
BUILTIN_POST_PROCESS = "builtin.post_process"
BUILTIN_HEADER_FOOTER = "builtin.header_footer"


# ----------------------------------------------------------------------
# 1) before_document_save —— 标题编号 + 引用超链接（原 PostProcessingStage）
# ----------------------------------------------------------------------
def _post_processing(data: dict):
    """标题自动编号与引用超链接（仅 apply 模式）。

    事件数据：document、config、ctx（FormatContext）、check。
    """
    if data.get("check", False):
        return
    ctx = data.get("ctx")
    if ctx is None or ctx.root_node is None:
        return
    config_model = ctx.config_model
    numbering = getattr(config_model, "numbering", None)
    if numbering and getattr(numbering, "enabled", False):
        from wordformat.numbering import process_heading_numbering

        process_heading_numbering(
            ctx.root_node,
            ctx.document,
            numbering,
            getattr(config_model, "headings", None),
        )
    from wordformat.hyperlinks import create_citation_hyperlinks

    create_citation_hyperlinks(ctx.root_node, ctx.document)


# ----------------------------------------------------------------------
# 2) before_document_save —— 页眉页脚检查/修正（issue#93）
# ----------------------------------------------------------------------
def _header_footer(data: dict):
    """页眉页脚识别：check 检查并批注，apply 修正并报告残留。

    事件数据：document、config、ctx、check。
    """
    document = data.get("document")
    if document is None:
        return
    config_model = data.get("config")
    check = data.get("check", False)
    from wordformat.header_footer import apply_header_footer, check_header_footer

    if check:
        issues = check_header_footer(document, config_model)
        if issues:
            _anchor_comments(document, issues)
    else:
        remaining = apply_header_footer(document, config_model)
        if remaining:
            _anchor_comments(document, remaining)


def _anchor_comments(document, issues: List[str]) -> None:
    """把页眉页脚问题作为批注锚定到文档第一段，并计入错误统计。

    页眉/页脚 run 悬挂批注受 python-docx 限制（comments 仅存在于
    document part），因此统一锚定首段，与检测摘要批注同层。
    """
    from wordformat.rules.node import FormatNode
    from wordformat.style.comments import (
        SEVERITY_ORDER,
        add_styled_comment,
        get_severity,
        split_comment_line,
    )

    if not document.paragraphs:
        logger.warning("文档无正文段落，页眉页脚批注无法锚定，已跳过")
        return
    para = document.paragraphs[0]
    if not para.runs:
        para.add_run("")
    paragraphs = [split_comment_line(t) for t in issues]
    add_styled_comment(document, para.runs, paragraphs)
    for text in issues:
        sev = get_severity(text)
        if sev not in SEVERITY_ORDER:
            sev = "错误"
        FormatNode._error_stats["total"] += 1
        FormatNode._error_stats[sev] = FormatNode._error_stats.get(sev, 0) + 1


# ----------------------------------------------------------------------
# 注册
# ----------------------------------------------------------------------
_BUILTIN_NAMES = (
    BUILTIN_POST_PROCESS,
    BUILTIN_HEADER_FOOTER,
)


def register_builtin_handlers(
    registry: HookRegistry = hooks, force: bool = False
) -> None:
    """注册全部内置机制回调（幂等：已注册的同名回调不重复注册）。

    Args:
        registry: 目标 HookRegistry，默认全局 hooks。
        force: 强制重复注册（测试用）。
    """
    existing = {name for name, _cb in registry.registered_callbacks()}

    def _register(event, callback, name):
        registry.register(event, callback, name=name)

    if force or BUILTIN_POST_PROCESS not in existing:
        _register("before_document_save", _post_processing, BUILTIN_POST_PROCESS)
    if force or BUILTIN_HEADER_FOOTER not in existing:
        _register("before_document_save", _header_footer, BUILTIN_HEADER_FOOTER)
    logger.debug("内置机制已注册为 hook handler")
