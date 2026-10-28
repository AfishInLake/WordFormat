"""段落/表格身份：`w14:paraId` 的读取与补齐。

编辑 JSON 要能可靠地折射成文档结构编辑，前提是每个段落有一个**跨进程稳定**的身份：

- python-docx 的 ``Paragraph`` 包装对象每次访问 ``document.paragraphs`` 都重新构造，
  ``p1 is p2`` 恒为 False，不能当身份；
- lxml 元素代理只在同一进程内稳定，而 gj 与 af 是两次独立进程；
- 数组下标会随增删整体漂移，用下标当身份等于回到「节点-段落错位」的老问题。

OOXML 已经自带了这个身份：Word 给每个 ``w:p`` 写 ``w14:paraId``（文档内唯一的 8 位
十六进制），python-docx 保存时原样保留，用户在 Word 里编辑后 Word 也会维护它。因此
这里只负责给缺失的段落（python-docx 生成的文件、本工具新插入的段）补上，不另造私有
属性，避免 Word 兼容性与校验风险。
"""

from __future__ import annotations

import secrets

from docx.oxml.ns import nsmap, qn
from loguru import logger

W14 = nsmap["w14"]
W14_PARA_ID = f"{{{W14}}}paraId"

# paraId 为 32 位非零值；避开高位保留区间，取值域 [1, 0x7FFFFFFE]
_ID_MIN = 0x00000001
_ID_RANGE = 0x7FFFFFFE - _ID_MIN + 1


def get_para_id(p_el) -> str | None:
    """段落元素 ``w:p`` 上的 ``w14:paraId``，缺失时返回 None。"""
    return p_el.get(W14_PARA_ID)


def paragraph_para_id(paragraph) -> str | None:
    """``Paragraph`` 对象的段落身份（转发到其 XML 元素）。"""
    return get_para_id(paragraph._p)


def _existing_ids(body) -> set[str]:
    """文档中已占用的 paraId，含表格单元格段落，避免与 Word 已有的值撞号。"""
    return {el.get(W14_PARA_ID) for el in body.iter(qn("w:p")) if el.get(W14_PARA_ID)}


def _mint_id(used: set[str]) -> str:
    while True:
        token = f"{secrets.randbelow(_ID_RANGE) + _ID_MIN:08X}"
        if token not in used:
            used.add(token)
            return token


def ensure_para_ids(document) -> int:
    """给 body 直接子级段落补齐缺失的 ``w14:paraId``，返回新写入的数量。

    只处理 body 层段落（JSON 可寻址的那批）；表格单元格段落交给 Word 自己维护，
    唯一性检查时仍会避开它们的取值。
    """
    body = document.element.body
    used = _existing_ids(body)
    minted = 0
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        if el.get(W14_PARA_ID):
            continue
        el.set(W14_PARA_ID, _mint_id(used))
        minted += 1
    if minted:
        logger.debug(f"补齐段落身份：新写入 {minted} 个 w14:paraId")
    return minted


def _table_anchor_p(tbl_el):
    """表格身份锚点：``w:tbl`` 首行首单元格的首个 ``w:p``（取不到时返回 None）。"""
    anchors = tbl_el.xpath("./w:tr[1]/w:tc[1]/w:p[1]")
    return anchors[0] if anchors else None


def get_table_id(tbl_el) -> str | None:
    """表格身份 = 首单元格段落的 ``w14:paraId``。

    OOXML 没给 ``w:tbl`` 定义 id，但单元格段落照样带 ``w14:paraId``；借它当身份可以
    复用 Word 的维护（用户在 Word 里编辑表格不会丢 id），不必另造私有属性。
    """
    anchor = _table_anchor_p(tbl_el)
    return get_para_id(anchor) if anchor is not None else None


def ensure_table_ids(document) -> int:
    """给 body 层每个表格的首单元格段落补齐身份，返回新写入的数量。

    只处理 body 直接子级的 ``w:tbl``，与 gj 导出表格条目的范围一致（嵌套表格不单独
    寻址，它会与其宿主表格共用取值域）。
    """
    body = document.element.body
    used = _existing_ids(body)
    minted = 0
    for tbl in body.iterchildren():
        if tbl.tag != qn("w:tbl"):
            continue
        anchor = _table_anchor_p(tbl)
        if anchor is None or get_para_id(anchor):
            continue
        anchor.set(W14_PARA_ID, _mint_id(used))
        minted += 1
    if minted:
        logger.debug(f"补齐表格身份：新写入 {minted} 个 w14:paraId")
    return minted
