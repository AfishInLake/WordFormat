#! /usr/bin/env python
# @Time    : 2026/7/12
# @Author  : afish
# @File    : parser.py
"""Markdown 解析器，基于 mistune AST 将 Markdown 文本转为扁平段落列表。"""

from __future__ import annotations

import re
from typing import Any

import mistune
from loguru import logger

from wordformat.categories import Category

# Markdown 没有题注语法，作者通常手写「表1-1 名称」「图1-1 名称」紧贴对象。
# 该正则要求「标签 + 编号 + 分隔符（空格/冒号）+ 文字」，避免把
# 「表1所示的数据……」这类正文句误判成题注。
_CAPTION_RE = re.compile(
    r"^\s*续?\s*(?P<label>[图表])\s*\d+(?:\s*[-－.．–—]\s*\d+)*\s*(?:[：:]|\s)\s*\S"
)

# 题注是单行短标签；过长或带句末标点的段落按正文处理，不提升
_CAPTION_MAX_CHARS = 80
_CAPTION_FORBIDDEN_CHARS = "。；;！？!?"

# 题注与对象的相邻方向：表题在表上方，图题在图下方
_CAPTION_NEIGHBOR_ORDER = {
    Category.TABLE_OBJECT: ((Category.CAPTION_TABLE, "表"), (-1, 1)),
    Category.FIGURE_IMAGE: ((Category.CAPTION_FIGURE, "图"), (1, -1)),
}


def parse_markdown(md_text: str) -> list[dict[str, Any]]:
    """解析 Markdown 文本，返回与现有 JSON 格式兼容的扁平段落列表。

    category 映射：
      h1 → heading_level_1, h2 → heading_level_2, h3+ → heading_level_3
      普通段落 → body_text
      仅含图片的段落 → figure_image（alt 文本生成 caption_figure）
      紧贴表格/图片的「表X-Y 名称」「图X-Y 名称」段落 → caption_table / caption_figure

    每个段落附带 inline_marks，记录行内格式片段（bold/italic/code 等），
    供 DocumentCreationStage 创建多 run 段落时使用。
    """
    md = mistune.create_markdown(
        renderer=None, plugins=["math", "table", "strikethrough", "url"]
    )
    ast = md(md_text)
    result: list[dict[str, Any]] = []
    _walk_blocks(ast, result)
    _promote_adjacent_captions(result)
    _warn_uncaptioned_images(result)
    return result


# ---------------------------------------------------------------------------
# 内部辅助
# ---------------------------------------------------------------------------


def _promote_adjacent_captions(items: list[dict[str, Any]]) -> None:
    """把紧贴表格/图片、形如「表X-Y 名称」「图X-Y 名称」的正文段提升为题注。

    作者手写的题注行会被当成 body_text，落到 docx 里就是普通正文，既没有
    题注样式也没有自动编号。题注与对象相邻（中间的空行已在解析阶段跳过）
    且文本形态匹配时，就地改 category，内容与位置保持不变。

    图片已由 alt 文本生成题注时不再认领邻居，避免同一张图出现两条题注。
    """
    for idx, item in enumerate(items):
        category = item.get("category")
        mapping = _CAPTION_NEIGHBOR_ORDER.get(category)
        if mapping is None:
            continue
        (caption_category, label), neighbors = mapping
        if caption_category == Category.CAPTION_FIGURE and _image_has_caption(
            items, idx
        ):
            continue
        for offset in neighbors:
            pos = idx + offset
            if not 0 <= pos < len(items):
                continue
            candidate = items[pos]
            if candidate.get("category") != Category.BODY_TEXT:
                continue
            if (
                category == Category.TABLE_OBJECT
                and offset == 1
                and pos + 1 < len(items)
                and items[pos + 1].get("category") == Category.TABLE_OBJECT
            ):
                # 两表之间的「表X-Y」行属于下一张表（表题在上），不被上一张认领
                continue
            if _looks_like_caption(str(candidate.get("paragraph", "")), label):
                candidate["category"] = caption_category
                candidate["score"] = 1.0
                break


def _image_has_caption(items: list[dict[str, Any]], idx: int) -> bool:
    """图片是否已有题注：alt 生成的题注紧随其后，上一张图的题注不算本图的。"""
    nxt = idx + 1
    if nxt < len(items) and items[nxt].get("category") == Category.CAPTION_FIGURE:
        return True
    prev = idx - 1
    if (
        0 <= prev < len(items)
        and items[prev].get("category") == Category.CAPTION_FIGURE
    ):
        before_caption = prev - 1
        owns_previous_image = (
            0 <= before_caption < len(items)
            and items[before_caption].get("category") == Category.FIGURE_IMAGE
        )
        return not owns_previous_image
    return False


def _looks_like_caption(text: str, label: str) -> bool:
    text = text.strip()
    if not text or len(text) > _CAPTION_MAX_CHARS:
        return False
    if any(ch in text for ch in _CAPTION_FORBIDDEN_CHARS):
        return False
    match = _CAPTION_RE.match(text)
    return bool(match) and match.group("label") == label


def _warn_uncaptioned_images(items: list[dict[str, Any]]) -> None:
    """图片既无 alt 题注、也无相邻手写题注时告警（论文图无题注是硬伤）。"""
    for idx, item in enumerate(items):
        if item.get("category") != Category.FIGURE_IMAGE:
            continue
        if _image_has_caption(items, idx):
            continue
        url = str(item.get("paragraph", "")) or "<空路径>"
        logger.warning(
            "图片缺少题注，未生成图题注："
            f"{url}；"
            "请在 Markdown 中写成 ![图1 题注文字](路径)，"
            "或在图片旁一行写「图1-1 题注文字」"
        )


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
        elif btype == "block_html":
            # 块级原始 HTML：剔标签后按正文保留，避免内容静默消失
            raw = re.sub(r"<[^>]+>", " ", block.get("raw", "")).strip()
            if raw:
                logger.warning(f"Markdown 含块级 HTML，已降级为纯文本：{raw[:50]}")
                result.append(
                    {"category": Category.BODY_TEXT, "paragraph": raw, "score": 1.0}
                )
        else:
            logger.warning(f"未处理的 Markdown 块类型，已跳过：{btype}")


def _make_block_math(block: dict) -> dict:
    return {
        "category": Category.MATH_BLOCK,
        "paragraph": block.get("raw", ""),
        "score": 1.0,
        "inline_marks": [
            {"text": block.get("raw", ""), "math": True, "math_display": True}
        ],
    }


def _make_heading(block: dict) -> dict:
    level = block["attrs"]["level"]
    # 未知级别立即 ValueError，避免拼出永不命中的类别名
    category = Category(f"heading_level_{min(level, 3)}")
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
                "category": Category.FIGURE_IMAGE,
                "paragraph": url,
                "score": 1.0,
                "inline_marks": [],
            }
        ]
        if alt_text:
            result.append(
                {
                    "category": Category.CAPTION_FIGURE,
                    "paragraph": alt_text,
                    "score": 1.0,
                    "inline_marks": [],
                }
            )
        # 无 alt 时不生成题注；由 _warn_uncaptioned_images 在相邻题注识别后统一告警
        return result

    text = _extract_text(children)
    if not text.strip():
        return None
    segments = _extract_segments(children)
    return {
        "category": Category.BODY_TEXT,
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
        "category": Category.TABLE_OBJECT,
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
                "category": Category.BODY_TEXT,
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
