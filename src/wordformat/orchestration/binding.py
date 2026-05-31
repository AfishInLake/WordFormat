#! /usr/bin/env python
"""段落绑定同步 —— 将虚拟节点树与 docx 文档段落对齐并同步。"""

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from loguru import logger

from wordformat.rules.node import FormatNode
from wordformat.utils import align_paragraphs
from wordformat.word_structure.utils import collect_nodes_in_order


def bind_and_sync(root_node: FormatNode, document: Document, check: bool):
    """将虚拟节点树与 docx 文档同步。

    1. 收集非空 docx 段落和树节点（文档顺序）
    2. 文本序列对齐 → matches / insertions / deletions
    3. 匹配节点直接绑定 paragraph
    4. apply 模式下：插入节点创建新段落、删除节点移除多余段落
    5. check 模式下：仅记录差异日志，不修改文档
    """
    docx_paras = [p for p in document.paragraphs if p.text.strip()]
    tree_nodes = collect_nodes_in_order(root_node)
    json_entries = [n.value for n in tree_nodes]

    if len(json_entries) != len(tree_nodes):
        logger.warning(
            f"节点数量 ({len(tree_nodes)}) 与 JSON 条目数量 ({len(json_entries)}) 不一致"
        )

    matches, insertions, deletions = align_paragraphs(json_entries, docx_paras)

    # 绑定匹配的段落
    for json_idx, para in matches.items():
        tree_nodes[json_idx].paragraph = para

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


def _sync_insertions(  # noqa: C901
    tree_nodes: list[FormatNode],
    matches: dict[int, "Paragraph"],
    insertions: set[int],
    document: "Document",
):
    """为 insertions 中的每个节点在 docx 中创建段落。

    按 json_index 降序处理（从后往前），每个新段落插入到最近的
    已匹配邻接段落之后，确保插入顺序正确。
    """
    matched_indices = sorted(matches.keys())

    for json_idx in sorted(insertions, reverse=True):
        # 找最近的后续已匹配节点
        next_match = None
        for mi in matched_indices:
            if mi > json_idx:
                next_match = mi
                break

        # 找最近的前置已匹配节点
        prev_match = None
        for mi in reversed(matched_indices):
            if mi < json_idx:
                prev_match = mi
                break

        # 确定锚点段落
        anchor_para = None
        use_addnext = False
        if prev_match is not None:
            anchor_para = matches[prev_match]
            use_addnext = True
        elif next_match is not None:
            anchor_para = matches[next_match]
            use_addnext = False

        # 按节点类型构建对应的 XML 元素
        new_elem = _build_element(tree_nodes[json_idx])

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

        # 包装并绑定（paragraph 类型用 Paragraph，其他类型暂时只存 element 引用）
        parent = new_elem.getparent()
        if tree_nodes[json_idx].type == "paragraph":
            tree_nodes[json_idx].paragraph = Paragraph(new_elem, parent)
        else:
            tree_nodes[json_idx].paragraph = Paragraph(new_elem, parent)  # fallback
        tree_nodes[json_idx]._is_insertion = True

        logger.info(
            f"已插入{tree_nodes[json_idx].type} #{json_idx}: "
            f"{tree_nodes[json_idx].value.get('paragraph', '')[:50]}..."
        )


def _build_element(node) -> "OxmlElement":
    """按节点类型构建对应的 OOXML 元素。"""
    if node.type == "paragraph" or not node.type:
        return _build_paragraph_element(node)
    elif node.type == "table":
        return _build_table_element(node)
    elif node.type == "image":
        return _build_image_element(node)
    else:
        logger.warning(f"未知节点类型 '{node.type}'，回退为 paragraph")
        return _build_paragraph_element(node)


def _build_paragraph_element(node) -> "OxmlElement":
    """构建 w:p 段落元素。"""
    new_p = OxmlElement("w:p")
    new_r = OxmlElement("w:r")
    new_t = OxmlElement("w:t")
    new_t.text = node.value.get("paragraph", "")
    new_t.set(qn("xml:space"), "preserve")
    new_r.append(new_t)
    new_p.append(new_r)
    return new_p


def _build_table_element(node) -> "OxmlElement":
    """从 node.content 构建 w:tbl 表格元素。"""
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


def _build_image_element(node) -> "OxmlElement":
    """从 node.content 构建 w:drawing 图片元素（内嵌占位图）。"""
    path = node.content.get("path", "")
    width = node.content.get("width", 300)
    height = node.content.get("height", 200)

    drawing = OxmlElement("w:drawing")
    inline = OxmlElement("wp:inline")
    extent = OxmlElement("wp:extent")
    extent.set("cx", str(width * 9525))
    extent.set("cy", str(height * 9525))
    inline.append(extent)
    drawing.append(inline)

    if path:
        logger.info(f"图片节点引用路径: {path}")
    return drawing


def _sync_deletions(
    docx_paras: list["Paragraph"],
    deletions: set[int],
):
    """移除 deletions 对应的 docx 段落。

    按 docx_index 降序处理，避免索引偏移影响后续删除。
    """
    for docx_idx in sorted(deletions, reverse=True):
        elem = docx_paras[docx_idx]._element
        parent = elem.getparent()
        if parent is not None:
            text_preview = docx_paras[docx_idx].text[:50]
            parent.remove(elem)
            logger.info(f"已删除段落 #{docx_idx}: {text_preview}...")
