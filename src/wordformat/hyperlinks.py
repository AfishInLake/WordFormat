#! /usr/bin/env python
"""
正文引用 → 参考文献条目 超链接模块

功能：
  1. 为每个 ReferenceEntry 段落创建 Word 书签
  2. 将正文中的 [1] [1,2] [1-3] 引用标记包裹为超链接，
     点击可跳转到对应的参考文献条目
"""

import re
from collections import deque

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from loguru import logger

from wordformat.rules.body import _CITATION_PATTERN


def create_citation_hyperlinks(root_node, document):
    """为参考文献创建书签，并将正文中的引用标记替换为可点击的超链接。

    调用时机：格式化 + 编号完成后。

    Args:
        root_node: 文档树根节点
        document: docx Document 对象
    """
    from wordformat.rules.references import ReferenceEntry

    # 1. 按文档顺序收集所有 ReferenceEntry 节点
    ref_entries = []
    _collect_nodes_of_type(root_node, ReferenceEntry, ref_entries)
    if not ref_entries:
        return

    # 2. 为每个参考文献条目创建书签
    bookmark_names = []
    for i, entry in enumerate(ref_entries, start=1):
        para = getattr(entry, "paragraph", None)
        if para is None:
            bookmark_names.append(None)
            continue
        bm_name = f"_Ref{i}"
        _insert_bookmark(para, bm_name, i)
        bookmark_names.append(bm_name)

    # 3. 遍历正文段落，将引用标记包裹为超链接
    body_nodes = []
    _collect_nodes_of_type(root_node, None, body_nodes, collect_body=True)
    for node in body_nodes:
        para = getattr(node, "paragraph", None)
        if para is None:
            continue
        _wrap_citations_in_hyperlinks(para, bookmark_names)

    logger.info(f"引用超链接创建完成：{len(ref_entries)} 条参考文献，"
                f"{len(body_nodes)} 个正文段落已处理")


def _collect_nodes_of_type(root_node, node_type, result, collect_body=False):
    """遍历树，收集指定类型的节点（按文档顺序）。

    Args:
        root_node: 树根节点
        node_type: 要收集的节点类型（如 ReferenceEntry），None 时按 collect_body 收集
        result: 收集结果列表（原地修改）
        collect_body: 当 node_type 为 None 时，收集 BodyText 节点
    """
    from wordformat.rules.body import BodyText

    queue = deque([root_node])
    while queue:
        node = queue.popleft()
        if node_type is not None:
            if isinstance(node, node_type):
                result.append(node)
        elif collect_body and isinstance(node, BodyText):
            result.append(node)
        queue.extend(node.children)


# ---------------------------------------------------------------------------
# 书签
# ---------------------------------------------------------------------------

# 全局书签 ID 计数器（每个文档内唯一）
_bookmark_id_counter = 1000


def _next_bookmark_id():
    global _bookmark_id_counter
    _bookmark_id_counter += 1
    return _bookmark_id_counter


def _insert_bookmark(paragraph, bookmark_name: str, bookmark_id: int):
    """在段落开头（pPr 之后、第一个 run 之前）插入 Word 书签。"""
    para_elem = paragraph._element
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), bookmark_name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))

    # 找到 pPr 之后的插入位置（或开头）
    insert_at = 0
    pPr = para_elem.find(qn("w:pPr"))
    if pPr is not None:
        # 获取 pPr 在父元素中的索引
        children = list(para_elem)
        try:
            insert_at = children.index(pPr) + 1
        except ValueError:
            insert_at = 0

    para_elem.insert(insert_at, start)
    para_elem.insert(insert_at + 1, end)


# ---------------------------------------------------------------------------
# 超链接
# ---------------------------------------------------------------------------


def _wrap_citations_in_hyperlinks(paragraph, bookmark_names: list):
    """将段落中匹配引用标记的 run 包裹为超链接。

    Args:
        paragraph: docx Paragraph 对象
        bookmark_names: 与引用编号对应的书签名列表（1-indexed）
    """
    para_elem = paragraph._element
    # 使用列表快照，因为我们会修改 paragraph.runs
    r_elems = list(para_elem.findall(qn("w:r")))

    for r_elem in r_elems:
        t_elem = r_elem.find(qn("w:t"))
        if t_elem is None or t_elem.text is None:
            continue

        text = t_elem.text
        m = _CITATION_PATTERN.fullmatch(text)
        if not m:
            continue

        # 提取引用编号
        ref_nums = _parse_ref_numbers(text)
        if not ref_nums or ref_nums[0] > len(bookmark_names):
            continue

        # 链接到第一个引用编号对应的参考文献
        anchor = bookmark_names[ref_nums[0] - 1]
        if anchor is None:
            continue

        # 在 rPr 中添加 Hyperlink 字符样式（保留已有格式如上标）
        rPr = r_elem.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            r_elem.insert(0, rPr)
        if rPr.find(qn("w:rStyle")) is None:
            rStyle = OxmlElement("w:rStyle")
            rStyle.set(qn("w:val"), "Hyperlink")
            rPr.insert(0, rStyle)

        # 创建 <w:hyperlink> 包裹 run
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("w:anchor"), anchor)
        hyperlink.set(qn("w:history"), "1")
        # 将 hyperlink 插入到 r_elem 之前，再将 r_elem 移入 hyperlink
        r_elem.addprevious(hyperlink)
        hyperlink.append(r_elem)


def _parse_ref_numbers(text: str) -> list[int]:
    """从引用标记文本中提取参考编号列表。

    >>> _parse_ref_numbers("[1]") → [1]
    >>> _parse_ref_numbers("[1,2,3]") → [1, 2, 3]
    >>> _parse_ref_numbers("[1-3]") → [1, 2, 3]
    >>> _parse_ref_numbers("[1,3-5]") → [1, 3, 4, 5]
    """
    inner = text.strip("[]")
    numbers = []
    for part in re.split(r"[,，、]\s*", inner):
        part = part.strip()
        m = re.match(r"(\d+)\s*-\s*(\d+)", part)
        if m:
            start, end = int(m.group(1)), int(m.group(2))
            numbers.extend(range(start, end + 1))
        else:
            try:
                numbers.append(int(part))
            except ValueError:
                continue
    return numbers
