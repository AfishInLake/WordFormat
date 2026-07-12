#! /usr/bin/env python
"""OOXML 样式继承链解析。

把「有效格式属性」的解析统一到一个入口：按 OOXML 继承链顺序遍历若干「源」
（rPr / pPr XML 元素），对每个源用「提取器」取某属性，返回链上第一个已设置的值。

继承链顺序
    run（字符）:  直接 rPr → 字符样式(rStyle)basedOn链 → 段落样式basedOn链
                 → 默认段落样式 → docDefaults/rPrDefault
    段落:        直接 pPr → 段落样式basedOn链 → 默认段落样式 → docDefaults/pPrDefault
    字体名额外解析主题字体（asciiTheme/eastAsiaTheme → theme1.xml 的 fontScheme）。

扩展方式：新增一个属性 = 写一个 `extractor(elem) -> value | _MISS` 函数即可，
无需再关心继承链遍历——遍历由 StyleResolver 统一负责，杜绝「查样式漏继承」。
"""

from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_UNDERLINE
from docx.oxml.ns import qn
from loguru import logger

# 链上「本源未设置该属性」的哨兵，区别于合法的 None（如主题色/主题字体不确定）
_MISS = object()


class ThemeRef:
    """字体主题引用（如 minorHAnsi / minorEastAsia），由 ThemeFontTable 兑现为具体字体名。"""

    __slots__ = ("token",)

    def __init__(self, token: str):
        """token 如 'minorHAnsi'、'majorEastAsia'。"""
        self.token = token


# ── 主题字体表 ────────────────────────────────────────────────────
class ThemeFontTable:
    """解析 theme part 的 fontScheme，把主题字体 token 映射为具体字体名。"""

    def __init__(self, document=None):
        """从 document 解析主题字体表；document 为 None 或解析失败时退化为空表。"""
        self._map: dict[str, str] = {}
        if document is not None:
            try:
                self._load(document)
            except Exception as e:  # theme 缺失/异常时退化为空表
                logger.debug(f"主题字体解析失败：{e}")

    def _load(self, document) -> None:
        from docx.oxml import parse_xml

        theme_part = next(
            (
                p
                for p in document.part.package.iter_parts()
                if "theme" in str(p.partname) and str(p.partname).endswith(".xml")
            ),
            None,
        )
        if theme_part is None:
            return
        root = parse_xml(theme_part.blob)
        for major_minor in ("major", "minor"):
            font = root.find(f".//{qn('a:' + major_minor + 'Font')}")
            if font is None:
                continue
            cap = major_minor.capitalize()  # Major / Minor
            latin = font.find(qn("a:latin"))
            ea = font.find(qn("a:ea"))
            cs = font.find(qn("a:cs"))
            if latin is not None and latin.get("typeface"):
                self._map[f"{major_minor}HAnsi"] = latin.get("typeface")
                self._map[f"{major_minor}Ascii"] = latin.get("typeface")
            if ea is not None and ea.get("typeface"):
                self._map[f"{major_minor}EastAsia"] = ea.get("typeface")
            if cs is not None and cs.get("typeface"):
                self._map[f"{major_minor}Bidi"] = cs.get("typeface")
            # 供按 "major"/"minor" 直接兜底
            self._map.setdefault(
                cap, latin.get("typeface") if latin is not None else ""
            )

    def resolve(self, token: str) -> str | None:
        """token 如 'minorHAnsi'；返回字体名，未知或空 → None。"""
        return self._map.get(token) or None


# ── 提取器：单个 rPr/pPr 元素 → 归一化值 | _MISS ──────────────────
# 字符级 rPr 均为 python-docx 的 CT_RPr，直接用其类型化子元素
# （.b/.i/.u/.sz/.color），复用官方 on/off、半点、颜色转换。
def x_bold(rPr):
    """从 rPr 提取加粗状态，未设置返回 _MISS。"""
    return _MISS if rPr.b is None else bool(rPr.b.val)


def x_italic(rPr):
    """从 rPr 提取斜体状态，未设置返回 _MISS。"""
    return _MISS if rPr.i is None else bool(rPr.i.val)


def x_underline(rPr):
    """从 rPr 提取下划线状态，未设置返回 _MISS。"""
    u = rPr.u
    if u is None:
        return _MISS
    return u.val not in (None, WD_UNDERLINE.NONE)


def x_size_pt(rPr):
    """从 rPr 提取字号（pt），未设置返回 _MISS。"""
    sz = rPr.sz
    if sz is None:
        return _MISS
    try:
        return sz.val.pt
    except (ValueError, TypeError):
        return _MISS


def x_color_rgb(rPr):
    """从 rPr 提取字体颜色 (r,g,b)；主题色返回 None，未设置返回 _MISS。"""
    c = rPr.color
    if c is None:
        return _MISS
    if c.themeColor is not None:
        return None  # 主题色：rgb 只是猜测，返回 None 表示不确定
    try:
        val = c.val
    except Exception:
        return (0, 0, 0)
    if val is None or val == "auto":
        return (0, 0, 0)
    return (val[0], val[1], val[2])  # RGBColor 支持索引


def _font_attr(rPr, literal: str, theme: str):
    """从 rFonts 提取字体名；直接字体名返回 str，主题 token 返回 ThemeRef，否则 _MISS。"""
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        return _MISS
    name = rFonts.get(qn(literal))
    if name:
        return name
    token = rFonts.get(qn(theme))
    if token:
        return ThemeRef(token)
    return _MISS


def x_font_ea(rPr):
    """提取东亚字体（eastAsia）或主题引用。"""
    return _font_attr(rPr, "w:eastAsia", "w:eastAsiaTheme")


def x_font_ascii(rPr):
    """提取西文字体（ascii）或主题引用。"""
    return _font_attr(rPr, "w:ascii", "w:asciiTheme")


_JC_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "start": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "end": WD_ALIGN_PARAGRAPH.RIGHT,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "distribute": WD_ALIGN_PARAGRAPH.DISTRIBUTE,
}


def x_alignment(pPr):
    """从 pPr 提取对齐方式（WD_ALIGN_PARAGRAPH），未设置返回 _MISS。"""
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        return _MISS
    return _JC_MAP.get(jc.get(qn("w:val")), _MISS)


def _spacing_lines(pPr, which: str):
    """从 pPr/spacing 提取段前/段后间距（行数）；autospacing 返回 None，未设置返回 _MISS。"""
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        return _MISS
    if spacing.get(qn(f"w:{which}Autospacing")) in ("1", "true"):
        return None
    val = spacing.get(qn(f"w:{which}Lines"))
    if val is None:
        return _MISS
    try:
        return round(int(val) / 100.0, 1)
    except (TypeError, ValueError):
        return _MISS


def x_space_before(pPr):
    """提取段前间距（行数）。"""
    return _spacing_lines(pPr, "before")


def x_space_after(pPr):
    """提取段后间距（行数）。"""
    return _spacing_lines(pPr, "after")


def x_line_spacing(pPr):
    """返回 (rule, value)：rule ∈ WD_LINE_SPACING，value 为倍数 float 或 None。"""
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        return _MISS
    rule = spacing.get(qn("w:lineRule"))
    if rule == "atLeast":
        return (WD_LINE_SPACING.AT_LEAST, None)
    if rule == "exact":
        return (WD_LINE_SPACING.EXACTLY, None)
    # auto 或未标注 lineRule：需 line 值算倍数
    line = spacing.get(qn("w:line"))
    if line is None:
        return _MISS
    try:
        return (WD_LINE_SPACING.MULTIPLE, int(line) / 240.0)
    except (TypeError, ValueError):
        return _MISS


def _ind_chars(pPr, attr: str, negate: bool = False):
    """从 pPr/ind 提取缩进（字符数）；negate=True 时取反（悬挂缩进用）。"""
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return _MISS
    val = ind.get(qn(attr))
    if val is None:
        return _MISS
    try:
        num = int(val) / 100.0
    except (TypeError, ValueError):
        return _MISS
    return -num if negate else num


def x_first_line_indent(pPr):
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return _MISS
    v = ind.get(qn("w:firstLineChars"))
    if v is not None:
        try:
            return int(v) / 100.0
        except (TypeError, ValueError):
            return _MISS
    v = ind.get(qn("w:hangingChars"))
    if v is not None:
        try:
            return -int(v) / 100.0
        except (TypeError, ValueError):
            return _MISS
    return _MISS


def x_left_indent(pPr):
    """提取左缩进（字符数）。"""
    return _ind_chars(pPr, "w:leftChars")


def x_right_indent(pPr):
    """提取右缩进（字符数）。"""
    return _ind_chars(pPr, "w:rightChars")


# ── 继承链解析器 ──────────────────────────────────────────────────
class StyleResolver:
    """按继承链解析段落/run 的有效格式属性；每文档构建一次并缓存。"""

    def __init__(self, document=None):
        self._by_id: dict[str, object] = {}
        self._rpr_default = None
        self._ppr_default = None
        self._default_para_style_id: str | None = None
        self.theme = ThemeFontTable(document)
        if document is not None:
            try:
                self._index(document)
            except Exception as e:
                logger.debug(f"样式索引构建失败，降级为仅直接格式：{e}")

    # -- 构建 --
    def _index(self, document) -> None:
        """遍历 styles.xml，构建 styleId -> CT_Style 映射和 docDefaults 引用。"""
        styles_el = document.styles.element
        for s in styles_el.findall(qn("w:style")):
            sid = s.get(qn("w:styleId"))
            if sid:
                self._by_id[sid] = s
            if (
                s.get(qn("w:type")) == "paragraph"
                and s.get(qn("w:default")) in ("1", "true")
                and self._default_para_style_id is None
            ):
                self._default_para_style_id = sid
        dd = styles_el.find(qn("w:docDefaults"))
        if dd is not None:
            rprd = dd.find(qn("w:rPrDefault"))
            if rprd is not None:
                self._rpr_default = rprd.find(qn("w:rPr"))
            pprd = dd.find(qn("w:pPrDefault"))
            if pprd is not None:
                self._ppr_default = pprd.find(qn("w:pPr"))

    @classmethod
    def _get_cached(cls, obj) -> StyleResolver:
        try:
            document = obj.part.document
        except Exception:
            return cls(None)  # 拿不到文档（如 Mock）→ 仅直接格式
        cached = getattr(document, "_wf_style_resolver", None)
        if not isinstance(cached, StyleResolver):
            cached = cls(document)
            try:
                document._wf_style_resolver = cached
            except Exception:
                pass
        return cached

    @classmethod
    def for_run(cls, run) -> StyleResolver:
        return cls._get_cached(run)

    @classmethod
    def for_paragraph(cls, paragraph) -> StyleResolver:
        return cls._get_cached(paragraph)

    # -- 样式链 --
    def _style_chain(self, style_id: str | None):
        """沿 w:basedOn 链向上遍历，yield 每个 CT_Style 元素。"""
        seen: set[str] = set()
        while style_id and style_id not in seen and style_id in self._by_id:
            seen.add(style_id)
            style = self._by_id[style_id]
            yield style
            based = style.find(qn("w:basedOn"))
            style_id = based.get(qn("w:val")) if based is not None else None

    @staticmethod
    def _child(elem, tag: str):
        return elem.find(qn(tag)) if elem is not None else None

    def _para_style_id(self, pPr) -> str | None:
        pStyle = self._child(pPr, "w:pStyle")
        return pStyle.get(qn("w:val")) if pStyle is not None else None

    # -- 源链 --
    def run_rpr_sources(self, run):
        """生成 run 字符属性继承链源：直接 rPr -> rStyle 链 -> 段落样式链 -> docDefaults。"""
        rPr = run._element.find(qn("w:rPr"))
        if rPr is not None:
            yield rPr
            if rPr.style is not None:  # CT_RPr.style → rStyle/@val
                for st in self._style_chain(rPr.style):
                    yield self._child(st, "w:rPr")
        # 段落样式链
        pPr = getattr(run._parent, "_p", None)
        pPr = pPr.find(qn("w:pPr")) if pPr is not None else None
        pid = self._para_style_id(pPr) if pPr is not None else None
        if pid is None:
            pid = self._default_para_style_id
        for st in self._style_chain(pid):
            yield self._child(st, "w:rPr")
        yield self._rpr_default

    def para_ppr_sources(self, paragraph):
        """生成段落属性继承链源：直接 pPr -> 段落样式链 -> docDefaults。"""
        pPr = paragraph._element.find(qn("w:pPr"))
        yield pPr
        pid = self._para_style_id(pPr) if pPr is not None else None
        if pid is None:
            pid = self._default_para_style_id
        for st in self._style_chain(pid):
            yield self._child(st, "w:pPr")
        yield self._ppr_default

    # -- 解析 --
    def _resolve(self, sources, extractor):
        """遍历源链，用 extractor 提取第一个非 _MISS 的值。"""
        for src in sources:
            if src is None:
                continue
            val = extractor(src)
            if val is not _MISS:
                return val
        return _MISS

    def resolve_run(self, run, extractor, default=None):
        """解析 run 有效字符属性，全链未设置返回 default。"""
        val = self._resolve(self.run_rpr_sources(run), extractor)
        return default if val is _MISS else val

    def resolve_para(self, paragraph, extractor, default=None):
        """解析段落有效段落属性，全链未设置返回 default。"""
        val = self._resolve(self.para_ppr_sources(paragraph), extractor)
        return default if val is _MISS else val

    def resolve_font(self, run, extractor) -> str | None:
        """字体名解析：把 ThemeRef 兑现为具体字体名；未设置 → None。"""
        val = self._resolve(self.run_rpr_sources(run), extractor)
        if val is _MISS:
            return None
        if isinstance(val, ThemeRef):
            return self.theme.resolve(val.token)
        return val
