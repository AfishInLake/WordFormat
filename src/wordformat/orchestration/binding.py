#! /usr/bin/env python
"""段落绑定同步 —— 将虚拟节点树与 docx 文档段落对齐并同步。"""

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from loguru import logger

from wordformat.rules.node import FormatNode
from wordformat.utils import _get_para_alignment_text, align_paragraphs
from wordformat.word_structure.utils import collect_nodes_in_order


def bind_and_sync(root_node: FormatNode, document: Document, check: bool):
    """将虚拟节点树与 docx 文档同步。

    1. 收集 docx 段落和树节点（文档顺序）
    2. 内容摘要序列对齐 → matches / insertions / deletions
    3. 匹配节点绑定 paragraph → extract → patch
    4. 插入节点创建新段落、删除节点移除多余段落
    """
    docx_paras = [p for p in document.paragraphs if _get_para_alignment_text(p)]
    tree_nodes = collect_nodes_in_order(root_node)
    json_entries = [n.value for n in tree_nodes]

    matches, insertions, deletions = align_paragraphs(json_entries, docx_paras)

    # Phase 1: 绑定 + Extract + Patch（仅 matched 节点）
    for json_idx, para in matches.items():
        node = tree_nodes[json_idx]
        node.paragraph = para
        if not check and node.type in ("paragraph", "", None):
            extracted = node.extract(para)
            node.content.update(extracted)
            changes = node.patch(para, document)
            if changes:
                logger.debug(f"节点 #{json_idx} patch: {', '.join(changes)}")

    # Phase 2: 插入 / 删除
    if insertions:
        if check:
            logger.warning(
                f"{len(insertions)} 个 JSON 条目在文档中找不到对应段落（需插入）"
            )
            for idx in sorted(insertions):
                node = tree_nodes[idx]
                text_preview = node.value.get("paragraph", "")[:50]
                logger.warning(f"  插入 #{idx}: {text_preview}...")
        else:
            _sync_insertions(tree_nodes, matches, insertions, document)

    if deletions:
        if check:
            logger.warning(
                f"{len(deletions)} 个文档段落在 JSON 中找不到对应条目（需删除）"
            )
            for idx in sorted(deletions):
                text_preview = docx_paras[idx].text[:50]
                logger.warning(f"  删除 #{idx}: {text_preview}...")
        else:
            _sync_deletions(docx_paras, deletions)


def _sync_insertions(
    tree_nodes: list[FormatNode],
    matches: dict[int, "Paragraph"],
    insertions: set[int],
    document: "Document",
):
    """为 insertions 中的每个节点在 docx 中创建段落。

    优先使用节点的 render() 方法；若节点未实现则回退到 _build_paragraph_element。
    """
    matched_indices = sorted(matches.keys())
    # 无匹配锚点时按升序插入（追加到 body 末尾），有锚点时按降序（避免索引偏移）
    reverse_order = bool(matched_indices)

    for json_idx in sorted(insertions, reverse=reverse_order):
        next_match = None
        for mi in matched_indices:
            if mi > json_idx:
                next_match = mi
                break

        prev_match = None
        for mi in reversed(matched_indices):
            if mi < json_idx:
                prev_match = mi
                break

        anchor_para = None
        use_addnext = False
        if prev_match is not None:
            anchor_para = matches[prev_match]
            use_addnext = True
        elif next_match is not None:
            anchor_para = matches[next_match]
            use_addnext = False

        node = tree_nodes[json_idx]
        new_elem = node.render(document)

        if anchor_para is not None:
            if use_addnext:
                anchor_para._element.addnext(new_elem)
            else:
                anchor_para._element.addprevious(new_elem)
        else:
            body_elem = None
            try:
                body_elem = document.element.body
            except Exception:
                pass
            if body_elem is None:
                logger.warning("无法确定文档 body 位置，跳过插入")
                continue
            body_elem.append(new_elem)

        # 仅文本类型节点创建 Paragraph 包装（表格/图片不需要）
        if node.type in ("paragraph", "") or not node.type:
            node.paragraph = Paragraph(new_elem, document._body)
        node._is_insertion = True

        logger.info(
            f"已插入{node.type} #{json_idx}: {node.value.get('paragraph', '')[:50]}..."
        )


def _sync_deletions(
    docx_paras: list["Paragraph"],
    deletions: set[int],
):
    """移除 deletions 对应的 docx 段落（降序处理避免索引偏移）。"""
    for docx_idx in sorted(deletions, reverse=True):
        elem = docx_paras[docx_idx]._element
        parent = elem.getparent()
        if parent is not None:
            text_preview = docx_paras[docx_idx].text[:50]
            parent.remove(elem)
            logger.info(f"已删除段落 #{docx_idx}: {text_preview}...")


# ── 遗留构建函数（非声明式节点类型的回退） ──────────────────────────


def _build_paragraph_element(node: FormatNode) -> "OxmlElement":
    """构建 w:p 段落元素（遗留，新节点应使用 render()）。"""
    new_p = OxmlElement("w:p")
    text = node.value.get("paragraph", "")
    if text.strip():
        new_r = OxmlElement("w:r")
        new_t = OxmlElement("w:t")
        new_t.text = text
        new_t.set(qn("xml:space"), "preserve")
        new_r.append(new_t)
        new_p.append(new_r)
    return new_p


def _build_table_element(node: FormatNode) -> "OxmlElement":
    """从 node.content 构建 w:tbl 表格元素（遗留）。"""
    tbl = OxmlElement("w:tbl")
    rows = node.content.get("rows", 1)
    cols = node.content.get("cols", 1)
    cells = node.content.get("cells", [])

    for r in range(rows):
        tr = OxmlElement("w:tr")
        for c in range(cols):
            tc = OxmlElement("w:tc")
            p = OxmlElement("w:p")
            r_elem = OxmlElement("w:r")
            t = OxmlElement("w:t")
            cell_text = ""
            if r < len(cells) and c < len(cells[r]):
                cell_text = str(cells[r][c])
            t.text = cell_text
            t.set(qn("xml:space"), "preserve")
            r_elem.append(t)
            p.append(r_elem)
            tc.append(p)
            tr.append(tc)
        tbl.append(tr)
    return tbl
