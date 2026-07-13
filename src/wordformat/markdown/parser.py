#! /usr/bin/env python
# @Time    : 2026/7/12
# @Author  : afish
# @File    : parser.py
"""Markdown 解析器，基于 mistune AST 将 Markdown 文本转为扁平段落列表。"""

from __future__ import annotations

from typing import Any

import mistune


def parse_markdown(md_text: str) -> list[dict[str, Any]]:
    """解析 Markdown 文本，返回与现有 JSON 格式兼容的扁平段落列表。

    category 映射：
      h1 → heading_level_1, h2 → heading_level_2, h3+ → heading_level_3
      普通段落 → body_text
      仅含图片的段落 → figure_image

    每个段落附带 inline_marks，记录行内格式片段（bold/italic/code 等），
    供 DocumentCreationStage 创建多 run 段落时使用。
    """
    md = mistune.create_markdown(
        renderer=None, plugins=["math", "table", "strikethrough", "url"]
    )
    ast = md(md_text)
    result: list[dict[str, Any]] = []
    _walk_blocks(ast, result)
    return result


# ---------------------------------------------------------------------------
# 内部辅助
# ---------------------------------------------------------------------------


def _walk_blocks(blocks: list[dict], result: list[dict]) -> None:
    for block in blocks:
        btype = block["type"]

        if btype == "blank_line" or btype == "thematic_break":
            continue

        if btype == "block_math":
            result.append(_make_block_math(block))
        elif btype == "heading":
            result.append(_make_heading(block))
        elif btype in ("paragraph", "block_text"):
            item = _make_paragraph(block)
            if isinstance(item, list):
                result.extend(item)
            elif item:
                result.append(item)
        elif btype == "table":
            result.append(_make_table(block))
        elif btype == "block_code":
            result.extend(_make_code_block(block))
        elif btype == "list":
            _walk_list(block, result)
        elif btype == "block_quote":
            _walk_blocks(block.get("children", []), result)


def _make_block_math(block: dict) -> dict:
    return {
        "category": "math_block",
        "paragraph": block.get("raw", ""),
        "score": 1.0,
        "inline_marks": [
            {"text": block.get("raw", ""), "math": True, "math_display": True}
        ],
    }


def _make_heading(block: dict) -> dict:
    level = block["attrs"]["level"]
    category = f"heading_level_{min(level, 3)}"
    text = _extract_text(block.get("children", []))
    segments = _extract_segments(block.get("children", []))
    return {
        "category": category,
        "paragraph": text,
        "score": 1.0,
        "inline_marks": segments,
    }


def _make_paragraph(block: dict) -> dict | list[dict] | None:
    children = block.get("children", [])

    # 仅含图片或块级公式的段落
    non_empty = [c for c in children if c["type"] != "softbreak"]
    if len(non_empty) == 1 and non_empty[0]["type"] == "block_math":
        return _make_block_math(non_empty[0])
    if len(non_empty) == 1 and non_empty[0]["type"] == "image":
        img_node = non_empty[0]
        url = img_node["attrs"].get("url", "")
        alt_text = _extract_text(img_node.get("children", [])).strip()
        # 图片在前，题注在后（图题在下）
        result: list[dict] = [
            {
                "category": "figure_image",
                "paragraph": url,
                "score": 1.0,
                "inline_marks": [],
            }
        ]
        if alt_text:
            result.append(
                {
                    "category": "caption_figure",
                    "paragraph": alt_text,
                    "score": 1.0,
                    "inline_marks": [],
                }
            )
        return result

    text = _extract_text(children)
    if not text.strip():
        return None
    segments = _extract_segments(children)
    return {
        "category": "body_text",
        "paragraph": text,
        "score": 1.0,
        "inline_marks": segments,
    }


def _make_table(block: dict) -> dict:
    rows: list[list[str]] = []
    for child in block.get("children", []):
        if child["type"] == "table_head":
            # table_head: cells are direct children (no row wrapper)
            cells = [
                _extract_text(cell.get("children", []))
                for cell in child.get("children", [])
            ]
            rows.append(cells)
        elif child["type"] == "table_body":
            # table_body: children are table_row elements
            for row in child.get("children", []):
                cells = [
                    _extract_text(cell.get("children", []))
                    for cell in row.get("children", [])
                ]
                rows.append(cells)
    return {
        "category": "table_object",
        "paragraph": "",
        "score": 1.0,
        "inline_marks": [],
        "table_rows": rows,
    }


def _make_code_block(block: dict) -> list[dict]:
    """代码块按行拆成多个段落，每行独立为一个 body_text 节点。"""
    text = block.get("raw", "")
    lines = text.split("\n")
    result: list[dict] = []
    for line in lines:
        line = line.strip("\r")
        if not line.strip():
            continue
        result.append(
            {
                "category": "body_text",
                "paragraph": line,
                "score": 1.0,
                "inline_marks": [{"text": line, "code": True}],
            }
        )
    return result


def _walk_list(block: dict, result: list[dict]) -> None:
    for item in block.get("children", []):
        if item["type"] == "list_item":
            for child in item.get("children", []):
                _walk_blocks([child], result)


def _extract_text(children: list[dict]) -> str:
    """从 inline children 提取纯文本。"""
    parts: list[str] = []

    def walk(nodes):
        for node in nodes:
            if node["type"] == "text":
                parts.append(node["raw"].replace("\r\n", "\n").replace("\r", "\n"))
            elif node["type"] == "codespan":
                parts.append(node.get("raw", ""))
            elif node["type"] == "inline_math":
                parts.append(f"${node.get('raw', '')}$")
            elif node["type"] == "block_math":
                parts.append(f"$${node.get('raw', '')}$$")
            elif node["type"] == "softbreak":
                parts.append(" ")
            elif node["type"] == "linebreak":
                parts.append("\n")
            elif "children" in node:
                walk(node["children"])

    walk(children)
    return "".join(parts)


def _extract_segments(children: list[dict]) -> list[dict]:  # noqa: C901
    """从 inline children 提取带格式标记的文本片段。"""
    segments: list[dict] = []

    def walk(nodes, attrs):  # noqa: C901
        for node in nodes:
            ntype = node["type"]
            if ntype == "text":
                if node["raw"]:
                    text = node["raw"].replace("\r\n", "\n").replace("\r", "\n")
                    segments.append({"text": text, **attrs})
            elif ntype in ("strong", "emphasis", "strikethrough"):
                extra = {
                    "strong": "bold",
                    "emphasis": "italic",
                    "strikethrough": "strikethrough",
                }[ntype]
                walk(node["children"], {**attrs, extra: True})
            elif ntype == "link":
                walk(node["children"], {**attrs, "link_url": node["attrs"]["url"]})
            elif ntype == "image":
                segments.append(
                    {"text": "", **attrs, "image_url": node["attrs"]["url"]}
                )
            elif ntype == "codespan":
                segments.append({"text": node.get("raw", ""), **attrs, "code": True})
            elif ntype == "inline_math":
                segments.append({"text": node.get("raw", ""), **attrs, "math": True})
            elif ntype == "block_math":
                segments.append(
                    {
                        "text": node.get("raw", ""),
                        **attrs,
                        "math": True,
                        "math_display": True,
                    }
                )
            elif ntype == "softbreak":
                segments.append({"text": " ", **attrs})
            elif ntype == "linebreak":
                segments.append({"text": "\n", **attrs})

    walk(children, {})
    return segments
