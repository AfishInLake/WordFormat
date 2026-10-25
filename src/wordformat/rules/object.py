"""图片段落和表格对象节点。"""

from pathlib import Path

from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches

from wordformat.rules.node import FormatNode
from wordformat.structure.registry import register
from wordformat.style.comments import format_comment
from wordformat.style.units import extract_unit_from_string

_EMU_PER_CM = 360_000
_WIDTH_TOLERANCE_CM = 0.01

# 环绕方式：中文标签与 OOXML 档位双向接受
_WRAP_LABELS = {
    "嵌入型": "嵌入型",
    "inline": "嵌入型",
    "四周型": "四周型",
    "square": "四周型",
    "紧密型": "紧密型",
    "tight": "紧密型",
    "穿越型": "穿越型",
    "through": "穿越型",
    "上下型": "上下型",
    "topandbottom": "上下型",
    "衬于文字下方": "衬于文字下方",
    "behind": "衬于文字下方",
    "浮于文字上方": "浮于文字上方",
    "infront": "浮于文字上方",
}


def _resolve_wrap(value) -> str | None:
    """把配置值归一为中文环绕标签；未配置返回 None。"""
    if value is None or str(value).strip() == "":
        return None
    key = str(value).strip()
    label = _WRAP_LABELS.get(key) or _WRAP_LABELS.get(key.lower())
    if label is None:
        raise ValueError(
            f"图片环绕方式非法: {value!r}（可用: {sorted(set(_WRAP_LABELS.values()))}）"
        )
    return label


def _detect_wrap(drawing) -> str | None:
    """读取 wp:inline / wp:anchor 判断实际环绕方式。"""
    if drawing.find(qn("wp:inline")) is not None:
        return "嵌入型"
    anchor = drawing.find(qn("wp:anchor"))
    if anchor is None:
        return None
    for tag, label in (
        ("wp:wrapSquare", "四周型"),
        ("wp:wrapTight", "紧密型"),
        ("wp:wrapThrough", "穿越型"),
        ("wp:wrapTopAndBottom", "上下型"),
    ):
        if anchor.find(qn(tag)) is not None:
            return label
    if anchor.find(qn("wp:wrapNone")) is not None:
        behind = anchor.get("behindDoc")
        return "衬于文字下方" if behind in ("1", "true") else "浮于文字上方"
    return None


def _drawing_extent_emu(drawing) -> int | None:
    """图片显示宽度（wp:extent/@cx，EMU）。"""
    container = drawing.find(qn("wp:inline"))
    if container is None:
        container = drawing.find(qn("wp:anchor"))
    if container is None:
        return None
    extent = container.find(qn("wp:extent"))
    if extent is None:
        return None
    try:
        return int(extent.get("cx"))
    except (TypeError, ValueError):
        return None


def _body_width_emu(paragraph, doc) -> int | None:
    """段落所属节的版心宽度：页宽 − 左右页边距。"""
    body = doc.element.body
    target = paragraph._p
    while target is not None and target.getparent() is not None:
        if target.getparent() is body:
            break
        target = target.getparent()
    sections = doc.sections
    if not sections:
        return None
    index = 0
    if target is not None and target.getparent() is body:
        children = list(body)
        try:
            pos = children.index(target)
        except ValueError:
            pos = 0
        index = sum(
            1
            for child in children[:pos]
            if child.tag == qn("w:p") and child.xpath("./w:pPr/w:sectPr")
        )
    section = sections[min(index, len(sections) - 1)]
    if section.page_width is None:
        return None
    width = int(section.page_width) - int(section.left_margin or 0)
    width -= int(section.right_margin or 0)
    return width if width > 0 else None


def _resolve_max_width_emu(value) -> int | None:
    """宽度上限（EMU）；None 表示按版心。"""
    if value is None or str(value).strip() in ("", "版心", "body"):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return int(Cm(float(value)))
    unit = extract_unit_from_string(str(value))
    if not unit.is_valid:
        raise ValueError(f"图片宽度上限需带单位（如 15cm / 150mm）: {value!r}")
    emu = unit.convert_to_emu()
    if emu is None:
        raise ValueError(f"图片宽度上限单位不支持: {value!r}")
    return emu


@register("figure_image")
class FigureImage(FormatNode):
    """图片段落节点（包含内联图片的段落，非题注）。

    只检查对齐、首行缩进和图片版式（环绕方式、是否超出版心），
    不检查行距、字体等。md→docx 时会尝试从 value 中的路径插入实际图片。
    """

    NODE_TYPE = "figures.image"
    NODE_LABEL = "图片段落"
    DEFAULTS = {
        # 图片段落居中/两端/左对齐都常见
        "alignment": ["居中对齐", "两端对齐", "左对齐"],
        "rules": {
            "figure_layout": {
                "enabled": True,
                # wrap 不写则只查宽度；max_width 缺省按版心（页宽 − 左右边距）
                "wrap": None,
                "max_width": "版心",
            }
        },
    }
    DEFAULT_RULES = {}
    RULES = {"figure_layout": "_check_figure_layout"}

    def _base(self, doc, p: bool, r: bool):
        """仅检查对齐；图片段落的缩进由图片版式决定，不做要求。"""
        from wordformat.style.defs import Alignment
        from wordformat.style.diff import _format_para_value

        self._try_insert_image()

        cfg = self.pydantic_config
        align_cfg = cfg.alignment or "居中对齐"
        align_values = align_cfg if isinstance(align_cfg, list) else [align_cfg]
        expected_align = Alignment(str(align_values[0]))
        actual_align = expected_align.get_from_paragraph(self.paragraph)
        if not any(Alignment(str(v)) == actual_align for v in align_values):
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=format_comment(
                    self.NODE_LABEL,
                    "对齐错误",
                    _format_para_value("alignment", actual_align),
                    _format_para_value("alignment", expected_align),
                ),
            )

    def _try_insert_image(self) -> None:
        """如果段落为空且 value 中有图片路径，尝试插入图片。"""
        if self.paragraph is None:
            return
        # 已有非空 run 时跳过（图片已存在或来自 docx 而非 md）
        for run in self.paragraph.runs:
            if run.text and run.text.strip():
                return
        value = self.value
        if not isinstance(value, dict):
            return
        path = value.get("paragraph", "").strip()
        if not path:
            return
        img_path = Path(path)
        if not img_path.exists():
            return
        try:
            # 清空占位 run，插入图片
            for run in self.paragraph.runs:
                run.text = ""
            if self.paragraph.runs:
                self.paragraph.runs[0].add_picture(str(img_path), width=Inches(5.5))
            else:
                run = self.paragraph.add_run()
                run.add_picture(str(img_path), width=Inches(5.5))
        except Exception:
            pass  # 图片格式不支持等，保留占位文本

    def _check_figure_layout(self, doc, rule_cfg, p: bool = False):
        """环绕方式与宽度检查：只报批注，不改版式（check/apply 都报）。"""
        if self.paragraph is None:
            return
        drawings = self.paragraph._p.xpath(".//w:drawing")
        if not drawings:
            return
        expected_wrap = _resolve_wrap(rule_cfg.get("wrap"))
        limit_emu = _resolve_max_width_emu(rule_cfg.get("max_width"))
        if limit_emu is None:
            limit_emu = _body_width_emu(self.paragraph, doc)
        limit_cm = limit_emu / _EMU_PER_CM if limit_emu else None

        for drawing in drawings:
            if expected_wrap is not None:
                actual_wrap = _detect_wrap(drawing)
                if actual_wrap is not None and actual_wrap != expected_wrap:
                    self.add_comment(
                        doc=doc,
                        runs=self.paragraph.runs,
                        text=format_comment(
                            self.NODE_LABEL,
                            "环绕方式错误",
                            f"当前{actual_wrap}",
                            expected_wrap,
                        ),
                    )
            if limit_cm is None:
                continue
            width_emu = _drawing_extent_emu(drawing)
            if width_emu is None:
                continue
            width_cm = width_emu / _EMU_PER_CM
            if width_cm > limit_cm + _WIDTH_TOLERANCE_CM:
                self.add_comment(
                    doc=doc,
                    runs=self.paragraph.runs,
                    text=format_comment(
                        self.NODE_LABEL,
                        "图片超宽",
                        f"当前宽 {width_cm:.2f}cm",
                        f"不超过 {limit_cm:.2f}cm",
                    ),
                )


@register("table_object")
class TableObject(FormatNode):
    """表格对象节点（表格整体格式，非题注）。

    md→docx 时会从 value.table_rows 创建 Word 表格。
    """

    NODE_TYPE = "tables.object"
    NODE_LABEL = "表格对象"
    DEFAULTS = {}
    DEFAULT_RULES = {}

    def _base(self, doc, p: bool, r: bool):
        value = self.value
        if isinstance(value, dict) and (value.get("exists") or value.get("sync_skip")):
            # 原文档已有表格（exists）或 sync 场景忽略的表格：不重复建表、不改内容
            return
        self._try_insert_table(doc)

    def _table_profile(self) -> dict | None:
        """解析本条表格使用的样式档案。

        优先级：条目 `table_style` > YAML `tables.default`。未配置 tables 段
        或未指定档案名时返回 None（保持 Table Grid 现状）。
        """
        value = self.value if isinstance(self.value, dict) else {}
        tables = getattr(self, "_full_config", {}).get("tables") or {}
        if not isinstance(tables, dict):
            return None
        styles = tables.get("styles") or {}
        name = value.get("table_style") or tables.get("default")
        if not name:
            return None
        if not isinstance(styles, dict) or name not in styles:
            raise ValueError(
                f"表格样式档案不存在: {name!r}"
                f"（tables.styles 可用: {sorted(styles) if isinstance(styles, dict) else []}）"
            )
        profile = styles[name]
        if not isinstance(profile, dict):
            raise ValueError(f"tables.styles.{name} 必须是映射")
        return profile

    def _try_insert_table(self, doc) -> None:
        """如果 value 中有 table_rows，在段落位置后插入 Word 表格。"""
        value = self.value
        if not isinstance(value, dict):
            return
        rows = value.get("table_rows", [])
        if not rows:
            return
        profile = self._table_profile()
        num_cols = max(len(r) for r in rows) if rows else 1
        style_name = (profile or {}).get("style", "Table Grid")
        table = doc.add_table(rows=len(rows), cols=num_cols, style=style_name)
        for i, row_data in enumerate(rows):
            for j, cell_text in enumerate(row_data):
                table.cell(i, j).text = cell_text
        if profile:
            _apply_table_profile(table, profile)
        # 将表格从文档末尾移到当前段落后
        if self.paragraph is not None:
            anchor = self.paragraph._element
            anchor.addnext(table._element)
            # jd 生成的空占位段只用于定位；删除以免表格前多出一个空段
            if not self.paragraph.text.strip() and not self.paragraph._p.xpath(
                ".//w:drawing"
            ):
                anchor.getparent().remove(anchor)


# ---------------------------------------------------------------------------
# 通用表格样式（配置镜像 OOXML tblPr/trPr/tcPr，任意组合；三线表只是其中一种档案）
# ---------------------------------------------------------------------------

_BORDER_EDGES = {
    "top": "top",
    "left": "left",
    "bottom": "bottom",
    "right": "right",
    "inside_h": "insideH",
    "inside_v": "insideV",
}

# OOXML 要求 tblPr 子元素按固定顺序排列
_TBL_PR_ORDER = (
    "w:tblStyle",
    "w:tblpPr",
    "w:tblOverlap",
    "w:bidiVisual",
    "w:tblStyleRowBandSize",
    "w:tblStyleColBandSize",
    "w:tblW",
    "w:jc",
    "w:tblCellSpacing",
    "w:tblInd",
    "w:tblBorders",
    "w:shd",
    "w:tblLayout",
    "w:tblCellMar",
    "w:tblLook",
    "w:tblCaption",
    "w:tblDescription",
)

_TR_PR_ORDER = (
    "w:cnfStyle",
    "w:divId",
    "w:gridBefore",
    "w:gridAfter",
    "w:wBefore",
    "w:wAfter",
    "w:cantSplit",
    "w:trHeight",
    "w:tblHeader",
    "w:tblCellSpacing",
    "w:jc",
    "w:hidden",
)

_TC_PR_ORDER = (
    "w:cnfStyle",
    "w:tcW",
    "w:gridSpan",
    "w:hMerge",
    "w:vMerge",
    "w:tcBorders",
    "w:shd",
    "w:noWrap",
    "w:tcMar",
    "w:textDirection",
    "w:tcFitText",
    "w:vAlign",
    "w:hideMark",
)

_ALIGN_MAP = {
    "left": WD_TABLE_ALIGNMENT.LEFT,
    "center": WD_TABLE_ALIGNMENT.CENTER,
    "right": WD_TABLE_ALIGNMENT.RIGHT,
}


def _cm_to_dxa(value) -> int:
    return int(round(float(value) * 567))


def _ordered_insert(parent, tag: str, elem, order: tuple[str, ...]) -> None:
    """删除同名旧元素后按 OOXML 顺序插入。"""
    old = parent.find(qn(tag))
    if old is not None:
        parent.remove(old)
    successors = order[order.index(tag) + 1 :]
    parent.insert_element_before(elem, *successors)


def _build_borders(tag: str, borders: dict):
    """构建 w:tblBorders / w:tcBorders。style 取 OOXML val；none 显式清边。"""
    el = OxmlElement(tag)
    for edge, attrs in borders.items():
        edge_tag = _BORDER_EDGES.get(edge)
        if edge_tag is None:
            raise ValueError(
                f"表格边框边名非法: {edge!r}（可用: {sorted(_BORDER_EDGES)}）"
            )
        if isinstance(attrs, str):
            # 简写：top: double 等价于 top: {style: double}
            attrs = {"style": attrs}
        elif attrs is None:
            attrs = {}
        elif not isinstance(attrs, dict):
            raise ValueError(f"边框 {edge!r} 配置必须是映射或样式字符串: {attrs!r}")
        border = OxmlElement(f"w:{edge_tag}")
        style = attrs.get("style", "single")
        border.set(qn("w:val"), str(style))
        if style != "none":
            if "size_pt" in attrs:
                border.set(qn("w:sz"), str(int(round(float(attrs["size_pt"]) * 2))))
            if "color" in attrs:
                border.set(qn("w:color"), str(attrs["color"]))
            if "space_pt" in attrs:
                border.set(qn("w:space"), str(int(attrs["space_pt"])))
        el.append(border)
    return el


def _apply_table_props(table, tbl_pr, profile: dict) -> None:
    if "layout" in profile:
        table.autofit = profile["layout"] != "fixed"
    width = profile.get("width")
    if isinstance(width, dict) and width.get("type"):
        typ = str(width["type"])
        value = width.get("value", 0)
        if typ == "auto":
            w_type, w_val = "auto", 0
        elif typ == "pct":
            # OOXML pct 以 1/50% 为单位：100% -> 5000
            w_type, w_val = "pct", int(round(float(value) * 50))
        elif typ == "cm":
            w_type, w_val = "dxa", _cm_to_dxa(value)
        else:  # dxa
            w_type, w_val = "dxa", int(value)
        el = OxmlElement("w:tblW")
        el.set(qn("w:w"), str(w_val))
        el.set(qn("w:type"), w_type)
        _ordered_insert(tbl_pr, "w:tblW", el, _TBL_PR_ORDER)
    alignment = profile.get("alignment")
    if alignment in _ALIGN_MAP:
        table.alignment = _ALIGN_MAP[alignment]
    elif alignment:
        raise ValueError(f"表格对齐非法: {alignment!r}（可用: {sorted(_ALIGN_MAP)}）")
    indent = profile.get("indent_cm")
    if indent is not None:
        el = OxmlElement("w:tblInd")
        el.set(qn("w:w"), str(_cm_to_dxa(indent)))
        el.set(qn("w:type"), "dxa")
        _ordered_insert(tbl_pr, "w:tblInd", el, _TBL_PR_ORDER)


def _apply_borders(tbl_pr, borders) -> None:
    if not borders:
        return
    el = _build_borders("w:tblBorders", borders)
    _ordered_insert(tbl_pr, "w:tblBorders", el, _TBL_PR_ORDER)


def _apply_col_widths(table, spec) -> None:
    """列宽：custom 写逐列宽度；equal 保持默认等宽；auto 不干预。

    python-docx 的 Column.width 只改 w:tblGrid/w:gridCol，单元格自身的
    w:tcW 仍是建表时的等宽值；fixed 布局下不同 Word 版本对两者的优先级
    处理不一致，故 custom 模式同时写 gridCol 与每个单元格的 tcW。
    """
    if not spec:
        return
    mode = spec.get("mode", "equal")
    if mode != "custom":
        return
    widths = spec.get("custom") or []
    for idx, width in enumerate(widths):
        if idx >= len(table.columns):
            break
        cm = Cm(float(width))
        table.columns[idx].width = cm
        for row in table.rows:
            if idx < len(row.cells):
                row.cells[idx].width = cm


def _apply_cell_margins(tbl_pr, margins) -> None:
    if not margins:
        return
    el = OxmlElement("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        if side not in margins:
            continue
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(_cm_to_dxa(margins[side])))
        node.set(qn("w:type"), "dxa")
        el.append(node)
    _ordered_insert(tbl_pr, "w:tblCellMar", el, _TBL_PR_ORDER)


def _set_cell_fill(cell, fill) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), str(fill))
    _ordered_insert(tc_pr, "w:shd", shd, _TC_PR_ORDER)


def _set_cell_borders(cell, borders: dict) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    el = _build_borders("w:tcBorders", borders)
    _ordered_insert(tc_pr, "w:tcBorders", el, _TC_PR_ORDER)


def _add_tr_prop(row, tag: str) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn(tag)) is not None:
        return
    el = OxmlElement(tag)
    _ordered_insert(tr_pr, tag, el, _TR_PR_ORDER)


def _apply_header_row(table, header_row) -> None:
    if header_row.get("repeat") and len(table.rows):
        _add_tr_prop(table.rows[0], "w:tblHeader")
    if header_row.get("cant_split"):
        for row in table.rows:
            _add_tr_prop(row, "w:cantSplit")


def _apply_band_cell_props(row, band: dict) -> None:
    """单行内的底纹/加粗/边框覆盖。"""
    for cell in row.cells:
        if band.get("fill"):
            _set_cell_fill(cell, band["fill"])
        if band.get("borders"):
            _set_cell_borders(cell, band["borders"])
        if band.get("bold"):
            for para in cell.paragraphs:
                for run in para.runs:
                    run.bold = True


def _apply_row_band(table, band: dict) -> None:
    """行级覆盖：底纹/加粗/单元格边框。"""
    rows = band.get("rows", "all")
    if rows == "all":
        targets = list(range(len(table.rows)))
    elif isinstance(rows, int) and not isinstance(rows, bool):
        targets = [rows]
    elif isinstance(rows, (list, tuple)):
        try:
            targets = [int(i) for i in rows]
        except (TypeError, ValueError) as exc:
            raise ValueError(
                f"row_bands.rows 必须是 all、整数或整数列表: {rows!r}"
            ) from exc
    else:
        raise ValueError(f"row_bands.rows 必须是 all、整数或整数列表: {rows!r}")
    for idx in targets:
        # 负数会命中 Python 负索引，必须显式排除
        if not 0 <= idx < len(table.rows):
            continue
        _apply_band_cell_props(table.rows[idx], band)


def _apply_row_props(table, header_row, row_bands) -> None:
    if header_row:
        _apply_header_row(table, header_row)
    for band in row_bands or []:
        _apply_row_band(table, band)


def _apply_cell_run(table, cell_run) -> None:
    """单元格文本字符格式，仅应用显式给出的字段。

    CharacterStyle 的 bold/italic/underline 默认 False；若把未配置的字段也
    写回 run，会覆盖 row_bands/header_row 已设置的加粗。
    """
    if not cell_run:
        return
    from wordformat.style.diff import CharacterStyle

    field_map = {
        "chinese_font_name": "font_name_cn",
        "english_font_name": "font_name_en",
        "font_size": "font_size",
        "font_color": "font_color",
        "bold": "bold",
        "italic": "italic",
        "underline": "underline",
    }
    only = {diff_type for key, diff_type in field_map.items() if key in cell_run}
    if not only:
        return
    cstyle = CharacterStyle(
        font_name_cn=cell_run.get("chinese_font_name", "宋体"),
        font_name_en=cell_run.get("english_font_name", "Times New Roman"),
        font_size=cell_run.get("font_size", "小四"),
        font_color=cell_run.get("font_color", "BLACK"),
        bold=bool(cell_run.get("bold", False)),
        italic=bool(cell_run.get("italic", False)),
        underline=bool(cell_run.get("underline", False)),
    )
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    if run.text.strip():
                        cstyle.apply_to_run(run, only=only, force=True)


def _apply_table_profile(table, profile: dict) -> None:
    """应用一个表格样式档案（分发到各属性辅助函数）。"""
    tbl_pr = table._tbl.tblPr
    _apply_table_props(table, tbl_pr, profile)
    _apply_borders(tbl_pr, profile.get("borders"))
    _apply_col_widths(table, profile.get("col_widths"))
    _apply_cell_margins(tbl_pr, profile.get("cell_margins_cm"))
    # 先写统一的字符格式，再由行带/表头覆盖，避免默认值把行级加粗重置
    _apply_cell_run(table, profile.get("cell_run"))
    _apply_row_props(table, profile.get("header_row"), profile.get("row_bands"))
