#! /usr/bin/env python
# @Time    : 2026/10/22
# @Author  : afish
# @File    : sync_apply.py
"""af --sync：以编辑后的 JSON 为准，在原 docx 上做结构同步（增/删/改/重排段落）。

现有 af 依赖「树节点 ↔ document.paragraphs 1:1 zip」，JSON 与原文档段落数必须
相等，因此无法表达增删。本模块在建树之前先做一次**结构对齐**：把编辑后的 JSON
扁平列表 A 与文档段落列表 B 按文本归一化 key（图片段用哨兵 ``"<image>"``
统一，防空文本区漂移及图片段误删）喂给 difflib。

注意 difflib 的 opcode 以「A 变换为 B」为方向，与本功能（把 B 同步为 A）相反：
- ``insert``（B 有而 A 无的区段）→ JSON 删了这些条目 → 真删原段落；
- ``delete``（A 有而 B 无的区段）→ JSON 新增条目 → 在锚点 ``insert_paragraph_before``
  插段；尾部无锚点的按序 ``document.add_paragraph``；
- ``replace`` → 两区段配对逐段 ``_set_paragraph_text``（图片段跳过），多出
  条目插在区尾锚点之前（保持 JSON 顺序）、多出段落删除；
- ``equal`` 不动。

同步后段落数与条目数恢复 1:1，后续建树/zip/格式化流程不变。

边界（见 SKILL.md）：
- 表格内容不在 paragraphs 流中：原表格无法经 JSON 删改；JSON 新增的
  ``table_object`` 条目会被标记 ``sync_skip`` 并忽略（建表请用 jd 重建）；
- difflib 把「移动」拆成删+插，含图片/域代码的段落不可移动；
- replace 段仅清 run 重写文本，段内超链接（如目录条目）文本不参与同步。
"""

from __future__ import annotations

import difflib
from typing import List

from wordformat.log_config import logger
from wordformat.utils import para_contains_image

# 图片段哨兵：JSON 的 figure_image 条目与 docx 含图段落共用，保证对齐稳定
IMAGE_SENTINEL = "<image>"


def _normalize(text) -> str:
    return (text or "").strip()


def _item_text(item: dict) -> str:
    """JSON 条目的目标文本；figure_image/table_object 不含可写入文档的文本。"""
    if item.get("category") in ("figure_image", "table_object"):
        return ""
    from wordformat.utils import resolve_item_text

    return _normalize(resolve_item_text(item))


def _item_key(item: dict) -> str:
    if item.get("category") == "figure_image":
        return IMAGE_SENTINEL
    return _item_text(item)


def _para_key(para) -> str:
    if para_contains_image(para):
        return IMAGE_SENTINEL
    return _normalize(para.text)


def _set_paragraph_text(para, text: str) -> None:
    """清 run 后单 run 写入文本（保留段落 pPr；超链接/域代码不在 runs 中，不处理）。"""
    for run in list(para.runs):
        run._element.getparent().remove(run._element)
    if text:
        para.add_run(text)


def _remove_paragraphs(paras: List) -> None:
    for para in paras:
        para._element.getparent().remove(para._element)


def _insert_item(document, paras: List, anchor_idx: int, item: dict) -> None:
    """在 B 侧 anchor_idx 段落前插入条目；anchor 越界（文档尾部）则追加。

    同一锚点连续插入多条时须重复传同一 anchor_idx：每次都插在锚段之前，
    依次得到 item0, item1, ..., anchor 的正确顺序。
    """
    text = _item_text(item)
    if anchor_idx < len(paras):
        paras[anchor_idx].insert_paragraph_before(text)
    else:
        document.add_paragraph(text)


def _apply_opcode(
    document,
    items: list[dict],
    paras: List,
    tag: str,
    i1: int,
    i2: int,
    j1: int,
    j2: int,
) -> tuple[int, int, int]:
    """处理单条 opcode，同步 B 侧区段到 A 侧形态，返回 (删, 插, 改) 计数。

    difflib 方向为「A 变换为 B」，与同步目标（B→A）相反：
    insert 表示 B 多出的段落（JSON 删了）；delete 表示 A 多出的条目（JSON 新增）。
    """
    if tag == "insert":
        _remove_paragraphs(paras[j1:j2])
        return j2 - j1, 0, 0
    if tag == "delete":
        # 同一锚点依次插在 B[j1] 之前（越界则追加），得到正确条目顺序
        for item in items[i1:i2]:
            _insert_item(document, paras, j1, item)
        return 0, i2 - i1, 0
    if tag == "replace":
        a_slice, b_slice = items[i1:i2], paras[j1:j2]
        n_edit = 0
        for item, para in zip(a_slice, b_slice, strict=False):
            if para_contains_image(para):
                continue  # 哨兵区不覆写图片段文本，保留原图
            _set_paragraph_text(para, _item_text(item))
            n_edit += 1
        # 配对的区段已逐段覆写文本；多出的条目按 JSON 顺序插在区尾 B[j2]
        # 之前（越界则追加），多出的段落删除。
        for item in a_slice[len(b_slice) :]:
            _insert_item(document, paras, j2, item)
        if len(b_slice) > len(a_slice):
            _remove_paragraphs(b_slice[len(a_slice) :])
        return (
            max(0, len(b_slice) - len(a_slice)),
            max(0, len(a_slice) - len(b_slice)),
            n_edit,
        )
    return 0, 0, 0


def sync_document_structure(document, items: list[dict]) -> None:
    """以 JSON 条目列表为准，原地对 document 段落做增/删/改，使两者一一对应。

    Args:
        document: python-docx Document（会被就地修改）
        items: 编辑后的 JSON 扁平条目列表（category/paragraph/replace 字段）
    """
    paras = list(document.paragraphs)
    # 表格条目不对应段落（原表格由 exists 标记保留，sync 不支持新增表格），
    # 不参与对齐，避免把相邻段落误配为表格而清空文本。
    text_items = [it for it in items if it.get("category") != "table_object"]
    keys_a = [_item_key(it) for it in text_items]
    keys_b = [_para_key(p) for p in paras]
    matcher = difflib.SequenceMatcher(None, keys_a, keys_b, autojunk=False)
    n_del = n_ins = n_edit = 0

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        d, ins, e = _apply_opcode(document, text_items, paras, tag, i1, i2, j1, j2)
        n_del += d
        n_ins += ins
        n_edit += e

    if not document.paragraphs:
        # Word 要求正文至少保留一个段落
        document.add_paragraph("")

    logger.info(f"结构同步完成：改 {n_edit} 段，插 {n_ins} 段，删 {n_del} 段")


class StructureSyncStage:
    """af --sync 流水线阶段：建树前先把编辑后的 JSON 同步进原文档结构。

    同步后将条目列表放入 ctx.paragraphs，供 TreeBuildingStage 直接建树
    （免去二次读文件），段落数与 JSON 条目数一致，后续 zip 天然对齐。
    """

    def process(self, ctx):
        from wordformat.log_config import logger
        from wordformat.structure.document_builder import DocumentBuilder

        items = ctx.paragraphs or DocumentBuilder.load_paragraphs(ctx.json_path)
        # af --sync 不支持新增表格（建表走 jd）；标 sync_skip 让 zip/建表阶段跳过，
        # 原文档已有表格（exists）同样不重复插入。
        for item in items:
            if (
                isinstance(item, dict)
                and item.get("category") == "table_object"
                and not item.get("exists")
            ):
                item["sync_skip"] = True
                logger.warning(
                    "af --sync 不支持新增表格条目，已忽略；如需增删表格请用 jd 重建"
                )
        sync_document_structure(ctx.document, items)
        if (
            not any(
                isinstance(it, dict) and it.get("category") != "table_object"
                for it in items
            )
            and ctx.document.paragraphs
        ):
            # 内容被删空时 sync 会补一个空段落（Word 要求），这里补占位条目，
            # 避免后续「节点数 ≠ 段落数」对齐检查失败。
            items.append({"category": "body_text", "paragraph": "", "score": 1.0})
        ctx.paragraphs = items
        return ctx
