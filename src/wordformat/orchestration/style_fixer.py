#! /usr/bin/env python
"""样式定义修正 —— 在格式化前统一修正 Word 样式定义中的字符和段落属性。"""

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from loguru import logger

from wordformat.config.datamodel import BaseModel, GlobalFormatConfig
from wordformat.style.style_enum import (
    Alignment,
    BuiltInStyle,
    FirstLineIndent,
    FontColor,
    FontSize,
    LeftIndent,
    LineSpacing,
    LineSpacingRule,
    RightIndent,
    SpaceAfter,
    SpaceBefore,
    _ensure_style_exists,
)


def collect_all_style_configs(config_model) -> dict:  # noqa: C901
    """遍历 NodeConfigRoot，收集所有唯一的 (英文样式名 → 配置段) 映射。

    对于被多个段引用的同一样式（如 body_text 和 references.content 都用 "Normal"），
    保留先遇到的配置段。
    """

    style_map: dict[str, object] = {}

    def _resolve_style_name(cfg) -> str | None:
        raw = getattr(cfg, "builtin_style_name", None)
        if not raw:
            return None
        try:
            return BuiltInStyle(raw).rel_value  # "正文" → "Normal", "题注" → "Caption"
        except Exception:
            return raw  # 自定义样式名直接使用

    def _walk(obj, path: str = ""):
        if isinstance(obj, GlobalFormatConfig):
            eng_name = _resolve_style_name(obj)
            if eng_name and eng_name not in style_map:
                style_map[eng_name] = obj
        if isinstance(obj, BaseModel):
            for f_name in type(obj).model_fields:
                val = getattr(obj, f_name)
                if isinstance(val, BaseModel):
                    _walk(val, f"{path}.{f_name}")
                elif isinstance(val, dict):
                    for _k, v in val.items():
                        if isinstance(v, BaseModel):
                            _walk(v, f"{path}.{_k}")

    _walk(config_model)
    return style_map


def fix_style_run_properties(style, cfg, style_name: str):  # noqa: C901
    """修正样式定义中的字符格式属性（w:rPr）。"""
    style_element = style.element

    # 确保 w:rPr 存在
    rPr = style_element.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        style_element.insert(0, rPr)

    # --- 字体名称 ---
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    cn_name = getattr(cfg, "chinese_font_name", None)
    if cn_name:
        rFonts.set(qn("w:eastAsia"), str(cn_name))
    en_name = getattr(cfg, "english_font_name", None)
    if en_name:
        rFonts.set(qn("w:ascii"), str(en_name))
        rFonts.set(qn("w:hAnsi"), str(en_name))

    # --- 字号 ---
    font_size = getattr(cfg, "font_size", None)
    if font_size is not None:
        try:
            fs = FontSize(font_size)
            pt_val = fs.rel_value
            half_pt = str(int(round(pt_val * 2)))
            sz = rPr.find(qn("w:sz"))
            if sz is None:
                sz = OxmlElement("w:sz")
                rPr.append(sz)
            sz.set(qn("w:val"), half_pt)
            szCs = rPr.find(qn("w:szCs"))
            if szCs is None:
                szCs = OxmlElement("w:szCs")
                rPr.append(szCs)
            szCs.set(qn("w:val"), half_pt)
        except Exception as e:
            logger.warning(f"设置样式 '{style_name}' 字号失败: {e}")

    # --- 字体颜色（清除主题色） ---
    font_color = getattr(cfg, "font_color", None)
    if font_color is not None:
        try:
            fc = FontColor(font_color)
            rgb = fc.rel_value
            hex_color = f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
            for old_color in rPr.findall(qn("w:color")):
                rPr.remove(old_color)
            new_color = OxmlElement("w:color")
            new_color.set(qn("w:val"), hex_color)
            rPr.append(new_color)
        except Exception as e:
            logger.warning(f"设置样式 '{style_name}' 颜色失败: {e}")

    # --- 加粗 ---
    bold = getattr(cfg, "bold", None)
    if bold is not None:
        b = rPr.find(qn("w:b"))
        if bold:
            if b is None:
                rPr.append(OxmlElement("w:b"))
        else:
            if b is not None:
                rPr.remove(b)

    # --- 斜体 ---
    italic = getattr(cfg, "italic", None)
    if italic is not None:
        i = rPr.find(qn("w:i"))
        if italic:
            if i is None:
                rPr.append(OxmlElement("w:i"))
        else:
            if i is not None:
                rPr.remove(i)

    # --- 下划线 ---
    underline = getattr(cfg, "underline", None)
    if underline is not None:
        u = rPr.find(qn("w:u"))
        if underline:
            if u is None:
                u = OxmlElement("w:u")
                u.set(qn("w:val"), "single")
                rPr.append(u)
        else:
            if u is not None:
                rPr.remove(u)


def fix_style_paragraph_properties(style, cfg, style_name: str):  # noqa: C901
    """修正样式定义中的段落格式属性（w:pPr）。

    设置：对齐方式、段前/段后间距、行距、首行缩进、左右缩进。
    直接在 w:pPr 元素上操作 XML，确保样式定义级别生效。
    """
    from wordformat.style.set_some import (
        _SetFirstLineIndent,
        _SetIndent,
        _SetLineSpacing,
        _SetSpacing,
    )

    style_element = style.element

    pPr = style_element.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        style_element.append(pPr)

    # --- 对齐方式 ---
    alignment = getattr(cfg, "alignment", None)
    if alignment is not None:
        try:
            al = Alignment(alignment)
            jc = pPr.find(qn("w:jc"))
            if jc is None:
                jc = OxmlElement("w:jc")
                pPr.append(jc)
            xml_val_map = {
                0: "left",
                1: "center",
                2: "right",
                3: "both",
                4: "distribute",
            }
            jc.set(qn("w:val"), xml_val_map.get(al.rel_value, "left"))
        except Exception as e:
            logger.warning(f"设置样式 '{style_name}' 对齐方式失败: {e}")

    # --- 段前/段后间距 ---
    for attr_name, cls, spacing_type in [
        ("space_before", SpaceBefore, "before"),
        ("space_after", SpaceAfter, "after"),
    ]:
        val = getattr(cfg, attr_name, None)
        if val is None:
            continue
        try:
            inst = cls(val)
            if inst.rel_unit == "hang":
                _SetSpacing._set_hang_on_pPr(pPr, spacing_type, inst.rel_value)
            else:
                logger.warning(
                    f"样式 '{style_name}' {attr_name} 使用了 '{inst.rel_unit}' 单位，"
                    f"样式定义仅支持'行'单位，已跳过"
                )
        except Exception as e:
            logger.warning(f"设置样式 '{style_name}' {attr_name} 失败: {e}")

    # --- 行距 ---
    line_spacingrule = getattr(cfg, "line_spacingrule", None)
    if line_spacingrule is not None:
        try:
            lsr = LineSpacingRule(line_spacingrule)
            rule_map = {
                0: "auto",
                1: "auto",
                2: "auto",
                3: "atLeast",
                4: "exact",
                5: "auto",
            }
            line_rule = rule_map.get(lsr.rel_value, "auto")

            line_spacing = getattr(cfg, "line_spacing", None)
            if line_spacing is not None:
                ls = LineSpacing(line_spacing)
                ls_val = ls.rel_value
                if ls.rel_unit == "倍":
                    line_val = ls_val * 240
                elif ls.rel_unit in ("pt",):
                    line_val = ls_val * 20
                else:
                    line_val = ls_val * 240
                _SetLineSpacing._set_on_pPr(pPr, line_rule, line_val)
            else:
                logger.warning(
                    f"样式 '{style_name}' 设置了 line_spacingrule 但未设置 line_spacing，已跳过行距"
                )
        except Exception as e:
            logger.warning(f"设置样式 '{style_name}' 行距失败: {e}")

    # --- 缩进 ---
    first_line_indent = getattr(cfg, "first_line_indent", None)
    if first_line_indent is not None:
        try:
            inst = FirstLineIndent(first_line_indent)
            if inst.rel_unit == "char":
                _SetFirstLineIndent._clear_ind_on_pPr(pPr)
                _SetFirstLineIndent._set_char_on_pPr(pPr, inst.rel_value)
            else:
                logger.warning(
                    f"样式 '{style_name}' first_line_indent 使用了 '{inst.rel_unit}' 单位，"
                    f"样式定义仅支持'字符'单位，已跳过"
                )
        except Exception as e:
            logger.warning(f"设置样式 '{style_name}' first_line_indent 失败: {e}")

    for attr_name, cls, indent_type in [
        ("left_indent", LeftIndent, "R"),
        ("right_indent", RightIndent, "X"),
    ]:
        val = getattr(cfg, attr_name, None)
        if val is None:
            continue
        try:
            inst = cls(val)
            if inst.rel_unit == "char":
                _SetIndent._set_char_on_pPr(pPr, indent_type, inst.rel_value)
            else:
                logger.warning(
                    f"样式 '{style_name}' {attr_name} 使用了 '{inst.rel_unit}' 单位，"
                    f"样式定义仅支持'字符'单位，已跳过"
                )
        except Exception as e:
            logger.warning(f"设置样式 '{style_name}' {attr_name} 失败: {e}")


def fix_all_style_definitions(document: Document, config_model):
    """在格式化开始前，统一修正文档中所有使用的样式定义。

    遍历配置中所有段（body_text、headings、abstract、figures、tables、
    references、acknowledgements 等），收集唯一的 builtin_style_name，
    然后修正每个样式定义的字符格式和段落格式，使其与配置一致。
    """
    style_configs = collect_all_style_configs(config_model)

    for eng_style_name, cfg in style_configs.items():
        _ensure_style_exists(document, eng_style_name)

        try:
            style = document.styles[eng_style_name]
        except KeyError:
            logger.warning(f"样式 '{eng_style_name}' 创建失败，跳过修正")
            continue

        fix_style_run_properties(style, cfg, eng_style_name)
        fix_style_paragraph_properties(style, cfg, eng_style_name)

        logger.debug(f"已修正样式定义: {eng_style_name}")
