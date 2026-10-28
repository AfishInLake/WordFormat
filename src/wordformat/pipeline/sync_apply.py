#! /usr/bin/env python
# @Time    : 2026-10-22
# @Author  : afish
# @File    : sync_apply.py
"""af --sync：以编辑后的 JSON 为准，在原 docx 上做结构同步（增/删/改/重排段落）。

现有 af 依赖「树节点 ↔ document.paragraphs 1:1 zip」，JSON 与原文档段落数必须
相等，因此无法表达增删。本模块在建树之前先做一次**结构对齐**，有两条路径：

**按身份对齐（``para_id`` 存在时）** —— gj 已把每段的 ``w14:paraId`` 写进文档和
JSON，两侧按该 id 精确配对：

- JSON 条目带 id 且文档中存在 → 同一段落，文本有变化就按 run 级字符改写；
- JSON 条目不带 id（agent/人新增的条目）→ 新插入一段，落点由它在 JSON 里紧邻的
  条目决定（整体按 JSON 顺序重排 body），不再依赖文本相似度；
- 文档段落未被任何条目引用 → JSON 删掉了它 → 真删该段；
- 表格条目按 ``table_id``（表格首单元格段落的 ``w14:paraId``）认领原有 ``w:tbl``：
  删条目＝删那个表格，不带 ``table_id`` 的新条目＝在该位置新建表格，条目顺序变了＝
  移动表格；``exists=True`` 却缺 ``table_id``（表格身份上线前的旧 JSON）时退回按
  ``exists`` 顺序归位且不删表。

**按文本对齐（旧 JSON 或文档无身份时）** —— 把 JSON 扁平列表 A 与文档段落列表 B 按
文本归一化 key（图片段用哨兵 ``"<image>"`` 统一，防空文本区漂移及图片段误删）喂给
difflib。注意 difflib 的 opcode 以「A 变换为 B」为方向，与本功能（把 B 同步为 A）
相反：

- ``insert``（B 有而 A 无的区段）→ JSON 删了这些条目 → 真删原段落；
- ``delete``（A 有而 B 无的区段）→ JSON 新增条目 → 在锚点 ``insert_paragraph_before``
  插段；尾部无锚点的按序 ``document.add_paragraph``；
- ``replace`` → 两区段配对逐段改写文本，多出条目插在区尾锚点之前（保持 JSON 顺序）、
  多出段落删除；
- ``equal`` 不动。

文本对齐靠文本找锚点，重复文本（同一篇里出现 4 次的「本章小结」）无法区分指代哪一段；
身份对齐没有这个限制。两条路径同步后都恢复「段落数 == 条目数」，后续建树/zip/格式化
流程不变。

边界（见 SKILL.md）：
- 表格内容不在 paragraphs 流中：表格整体可按 ``table_id`` 删除、移动、新增，但单元格
  文本仍不能经 JSON 改写（``table_rows`` 只读）；
- 文本对齐把「移动」拆成删+插，含图片/域代码的段落不可移动；
- 段落文本改写只动文字 run：加粗/斜体/超链接按字符对齐保留，含图片、域代码、公式的
  run 不改写（``Run.text`` 赋值会清空 run 内容，写进去等于删图、毁页码域）。
"""

from __future__ import annotations

import difflib
from typing import List

from docx.oxml.ns import qn
from loguru import logger

from wordformat.para_ids import ensure_para_ids, get_para_id, get_table_id
from wordformat.text_edit import set_paragraph_text
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


def _set_paragraph_text(para, text: str) -> bool:
    """按字符对齐改写段落文本，保留 run 边界上的既有格式；返回是否改写。"""
    return set_paragraph_text(para, text)


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


def _align_by_text(document, items: list[dict]) -> tuple[int, int, int, int, int]:
    """旧路径：按归一化文本做 difflib 对齐（无段落身份时）。表格一律不动。"""
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
    return n_edit, n_ins, n_del, 0, 0


# ---------------------------------------------------------------------------
# 按段落身份对齐（gj 写入 w14:paraId 后可用）
# ---------------------------------------------------------------------------


def _identity_ready(document, items: list[dict]) -> bool:
    """两侧都能给出唯一段落身份时，才允许按身份对齐。

    只要文档里有一个 body 段落缺 id 或 id 重复，就退回文本对齐：身份不完整时按身份
    删段等于凭不存在的键误删，宁可退回旧行为。
    """
    if not any(isinstance(it, dict) and it.get("para_id") for it in items):
        return False
    body_ids = [get_para_id(p._p) for p in document.paragraphs]
    return all(body_ids) and len(set(body_ids)) == len(body_ids)


def _take_existing(para_by_id: dict, item: dict):
    """取出条目指向的既有段落；无 id 或 id 在文档中已不存在时返回 None（按新增处理）。"""
    para_id = item.get("para_id")
    if not para_id:
        return None
    return para_by_id.pop(para_id, None)


def _reorder_body(document, ordered: List) -> None:
    """按 JSON 顺序重排 body：lxml 插入即移动，增删改与「移动段落」一次解决。"""
    body = document.element.body
    structural = (qn("w:p"), qn("w:tbl"))
    tail = next((el for el in body.iterchildren() if el.tag not in structural), None)
    for el in ordered:
        if tail is None:
            body.append(el)
        else:
            tail.addprevious(el)


def _body_tables(body) -> List:
    """body 直接子级的 ``w:tbl`` 元素，顺序即文档序（与 gj 的表格导出范围一致）。"""
    return [el for el in body.iterchildren() if el.tag == qn("w:tbl")]


def _table_identity_ready(items: list[dict], table_elements: List) -> bool:
    """表格能否按身份归位/删除/新增。

    删表格是破坏性动作，身份证据不足时宁可维持旧的「原地不动 + 告警」。判据是
    ``exists``：gj 给每个原有表格条目写 ``exists=True``，所以「带 exists 却没有
    ``table_id``」只可能是表格身份上线之前导出的 JSON——它指向的是既有表格，此时
    按身份删除等于凭不存在的键误删。不带 ``exists`` 的条目本身就是新建意图，
    没有 ``table_id`` 是正常的，不妨碍身份模式生效。

    JSON 完全没有表格条目时，只有段落条目普遍带 ``para_id``（说明这份 JSON 出自
    同一文档的 gj，见过全部表格）才认定「表格被删了」——手写 JSON 忽略表格
    不等于要删表格。
    """
    if not all(get_table_id(el) for el in table_elements):
        return False
    table_items = [
        it
        for it in items
        if isinstance(it, dict) and it.get("category") == "table_object"
    ]
    if any(it.get("exists") and not it.get("table_id") for it in table_items):
        return False
    if not table_items and table_elements:
        return any(isinstance(it, dict) and it.get("para_id") for it in items)
    return True


def _claim_table_by_id(item: dict, tbl_by_id: dict, document, ordered: List) -> bool:
    """按 ``table_id`` 认领既有表格元素，认领不到则放占位段（返回 True 表示新增）。

    新建表格不在这里做：先放一个空段落占住 JSON 里的位置，真正的建表由
    ``TableObject._try_insert_table`` 在该段落后插入表格并删掉占位段，
    这样建表逻辑只有一份，jd/md 与 af --sync 共用。
    """
    table_id = item.get("table_id")
    el = tbl_by_id.pop(table_id, None) if table_id else None
    if el is None:
        ordered.append(document.add_paragraph("")._p)
        return True
    ordered.append(el)
    return False


def _place_table_item(
    item: dict, tbl_by_id, table_elements: List, cursor: int, document, ordered
) -> tuple[bool, int]:
    """把一个表格条目落到 ``ordered`` 上，返回 ``(是否新增, 归位后的游标)``。"""
    if tbl_by_id is None:
        # 旧 JSON（无表格身份）：按 exists 顺序归位既有表格，既不新增也不删除
        if item.get("exists") and cursor < len(table_elements):
            ordered.append(table_elements[cursor])
            return False, cursor + 1
        return False, cursor
    return _claim_table_by_id(item, tbl_by_id, document, ordered), cursor


def _align_by_identity(document, items: list[dict]) -> tuple[int, int, int, int, int]:
    """按 ``para_id`` 对齐：JSON 为准，段落可增、删、改、移位；表格按 ``table_id`` 同理。"""
    para_by_id = {get_para_id(p._p): p for p in document.paragraphs}
    table_elements = _body_tables(document.element.body)
    if _table_identity_ready(items, table_elements):
        tbl_by_id: dict | None = {get_table_id(el): el for el in table_elements}
    else:
        tbl_by_id = None
    ordered: List = []
    n_edit = n_ins = n_del = n_tbl_ins = n_tbl_del = 0
    table_cursor = 0

    for item in items:
        if item.get("category") == "table_object":
            added, table_cursor = _place_table_item(
                item, tbl_by_id, table_elements, table_cursor, document, ordered
            )
            n_tbl_ins += added
            continue
        para = _take_existing(para_by_id, item)
        if para is None:
            para = document.add_paragraph(_item_text(item))
            n_ins += 1
        elif _set_paragraph_text(para, _item_text(item)):
            n_edit += 1
        ordered.append(para._p)

    n_del = len(para_by_id)
    _remove_paragraphs(list(para_by_id.values()))
    n_tbl_del = _drop_unclaimed_tables(tbl_by_id, table_elements, table_cursor)
    _reorder_body(document, ordered)
    ensure_para_ids(document)  # 新插入的段落也要有身份，否则下一轮同步断链
    return n_edit, n_ins, n_del, n_tbl_ins, n_tbl_del


def _drop_unclaimed_tables(tbl_by_id, table_elements: List, table_cursor: int) -> int:
    """处理 JSON 没认领的表格，返回删除数。

    身份模式没被认领的就是被删掉的条目（``_table_identity_ready`` 已确认身份证据
    充分）；无身份时（旧 JSON）一个都不删，只告警未归位的表格原地不动。
    """
    if tbl_by_id is None:
        if table_cursor != len(table_elements):
            logger.warning(
                f"文档中有 {len(table_elements)} 个表格，JSON 只描述到 {table_cursor} 个，"
                "未被引用的表格保持原位不动"
            )
        return 0
    for el in tbl_by_id.values():
        el.getparent().remove(el)
    return len(tbl_by_id)


def _mark_unsupported_new_tables(document, items: list[dict]) -> None:
    """身份证据不足时，JSON 新增的表格条目仍按老行为忽略（建表走 jd）。

    标记必须发生在 ``_align_by_identity`` 之前且与它的表格分支保持一致：一旦身份
    可用，新表格条目就要落一个占位段落，此时若仍标 ``sync_skip``，占位段就没有任何
    节点绑定，「段落数 == 条目数」被打破，建树后的 zip 对齐检查直接报错。
    """
    table_elements = _body_tables(document.element.body)
    if _identity_ready(document, items) and _table_identity_ready(
        items, table_elements
    ):
        return
    for item in items:
        if (
            isinstance(item, dict)
            and item.get("category") == "table_object"
            and not item.get("exists")
        ):
            item["sync_skip"] = True
            logger.warning(
                "本次同步无法按身份新增表格（缺 para_id/table_id 或在按文本对齐），"
                "该条目已忽略；如需增删表格请用 jd 重建或补齐身份后重试"
            )


def sync_document_structure(document, items: list[dict]) -> None:
    """以 JSON 条目列表为准，原地对 document 的段落/表格做增/删/改/重排，使两者一一对应。

    Args:
        document: python-docx Document（会被就地修改）
        items: 编辑后的 JSON 扁平条目列表（category/paragraph/para_id/table_id/replace）
    """
    _mark_unsupported_new_tables(document, items)
    if _identity_ready(document, items):
        n_edit, n_ins, n_del, n_tbl_ins, n_tbl_del = _align_by_identity(document, items)
    else:
        if any(isinstance(it, dict) and it.get("para_id") for it in items):
            logger.warning(
                "JSON 带段落身份但文档侧不完整（未重新 gj 或在 Word 里被重建），"
                "已退回按文本对齐"
            )
        n_edit, n_ins, n_del, n_tbl_ins, n_tbl_del = _align_by_text(document, items)

    if not document.paragraphs:
        # Word 要求正文至少保留一个段落
        document.add_paragraph("")

    tables = f"，表格 +{n_tbl_ins} −{n_tbl_del}" if (n_tbl_ins or n_tbl_del) else ""
    logger.info(f"结构同步完成：改 {n_edit} 段，插 {n_ins} 段，删 {n_del} 段{tables}")


class StructureSyncStage:
    """af --sync 流水线阶段：建树前先把编辑后的 JSON 同步进原文档结构。

    同步后将条目列表放入 ctx.paragraphs，供 TreeBuildingStage 直接建树
    （免去二次读文件），段落数与 JSON 条目数一致，后续 zip 天然对齐。
    """

    def process(self, ctx):
        from wordformat.structure.document_builder import DocumentBuilder

        items = ctx.paragraphs or DocumentBuilder.load_paragraphs(ctx.json_path)
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
