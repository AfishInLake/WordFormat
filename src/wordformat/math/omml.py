"""
OML Renderer — LaTeX math → OMML (Office Math Markup Language) for .docx files.

Uses the pure-Python pipeline:
    LaTeX → latex2mathml → MathML → mathml2omml → OMML

Usage:
    from omml_renderer import latex_to_omath, add_display_math, add_inline_math

    # Parse a LaTeX formula into an OMML element tree
    om = latex_to_omath(r"w_j = \frac{q_j}{\\sum_{k=1}^{n} q_k}")

    # Add a display-math paragraph to a python-docx Document
    add_display_math(doc, r"w_j = \frac{1}{3} \\widehat{w}_j")

    # Render inline $...$ math inside a paragraph
    add_inline_math(paragraph, "The result is $C_{10}$ and $N_2$.")

Dependencies: lxml, python-docx, latex2mathml, mathml2omml
"""

from __future__ import annotations

import re

import latex2mathml.converter
import lxml.etree as ET
import mathml2omml

# ---------------------------------------------------------------------------
# Namespace constants
# ---------------------------------------------------------------------------

MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

MATH_FONT = "Cambria Math"
_DEFAULT_SZ = "22"


# ---------------------------------------------------------------------------
# LaTeX → OMML pipeline
# ---------------------------------------------------------------------------


def _latex_to_omml_xml(latex: str) -> str | None:
    """Convert a LaTeX math string to OMML XML string via MathML."""
    if not latex.strip():
        return None
    try:
        mathml = latex2mathml.converter.convert(latex)
    except Exception:
        return None
    if not mathml:
        return None
    try:
        return mathml2omml.convert(mathml)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Low-level helpers (used for paragraph-level rPr)
# ---------------------------------------------------------------------------


def _omml(tag: str) -> str:
    return f"{{{MATH_NS}}}{tag}"


def _make_wrpr() -> ET.Element:
    """Build a w:rPr element with math font and size."""
    wrpr = ET.Element(f"{{{WORD_NS}}}rPr")
    rf = ET.SubElement(wrpr, f"{{{WORD_NS}}}rFonts")
    rf.set(f"{{{WORD_NS}}}ascii", MATH_FONT)
    rf.set(f"{{{WORD_NS}}}hAnsi", MATH_FONT)
    sz = ET.SubElement(wrpr, f"{{{WORD_NS}}}sz")
    sz.set(f"{{{WORD_NS}}}val", _DEFAULT_SZ)
    return wrpr


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def _make_ctrl_pr() -> ET.Element:
    """Build an m:ctrlPr containing w:rPr."""
    cp = ET.Element(_omml("ctrlPr"))
    cp.append(_make_wrpr())
    return cp


def _post_process(omath: ET.Element) -> None:
    """Fix up mathml2omml output to match Word/WPS expectations."""
    # Add w:rPr to every m:r that lacks it (WPS needs font/size on each run)
    for mr in omath.iter(_omml("r")):
        if mr.find(f"{{{WORD_NS}}}rPr") is None:
            wrpr = _make_wrpr()
            # Insert w:rPr after m:rPr (or at position 0 if no m:rPr)
            mrpr = mr.find(_omml("rPr"))
            if mrpr is not None:
                idx = list(mr).index(mrpr)
                mr.insert(idx + 1, wrpr)
            else:
                mr.insert(0, wrpr)

    # Fix rad elements: add radPr, degHide, deg, ctrlPr as needed
    for rad in omath.iter(_omml("rad")):
        radpr = rad.find(_omml("radPr"))
        if radpr is None:
            radpr = ET.Element(_omml("radPr"))
            rad.insert(0, radpr)
        if radpr.find(_omml("degHide")) is None and rad.find(_omml("deg")) is None:
            ET.SubElement(radpr, _omml("degHide")).set(_omml("val"), "1")
        if radpr.find(_omml("ctrlPr")) is None:
            radpr.append(_make_ctrl_pr())
        deg = rad.find(_omml("deg"))
        if deg is None:
            deg = ET.Element(_omml("deg"))
            deg.append(_make_ctrl_pr())
            rp_idx = list(rad).index(radpr)
            rad.insert(rp_idx + 1, deg)
        elif deg.find(_omml("ctrlPr")) is None:
            deg.append(_make_ctrl_pr())


def latex_to_omath(latex: str) -> ET.Element | None:
    """Parse a LaTeX math expression into an ``m:oMath`` OMML element.

    Returns ``None`` if the string is empty or parsing fails.
    """
    omml_xml = _latex_to_omml_xml(latex)
    if not omml_xml:
        return None
    try:
        # mathml2omml output uses "m:" prefix without namespace declaration
        wrapped = f'<root xmlns:m="{MATH_NS}">{omml_xml}</root>'
        tree = ET.fromstring(wrapped.encode("utf-8"))
        omath = tree.find(f"{{{MATH_NS}}}oMath")
        if omath is not None:
            _post_process(omath)
        return omath
    except Exception:
        return None


def latex_to_omath_para(latex: str) -> ET.Element | None:
    """Parse LaTeX into an ``m:oMathPara`` element (for display / block math)."""
    om = latex_to_omath(latex)
    if om is None:
        return None
    omathpara = ET.Element(_omml("oMathPara"))
    omathpara.append(om)
    return omathpara


def add_display_math(doc, latex: str):
    """Add a centered display-math paragraph to a ``python-docx`` Document.

    The *latex* string should **not** contain ``$$`` delimiters.
    Returns the new paragraph.
    """
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ppr = p._p.find(f"{{{WORD_NS}}}pPr")
    if ppr is not None:
        ppr.append(_make_wrpr())
    om = latex_to_omath(latex)
    if om is not None:
        omathpara = ET.Element(_omml("oMathPara"))
        omathpara.append(om)
        p._p.append(omathpara)
    else:
        from docx.shared import Pt

        run = p.add_run(latex)
        run.font.size = Pt(11)
    return p


def add_inline_math(paragraph, text: str):
    """Render text containing ``$...$`` inline-math segments into *paragraph*.

    Math segments are rendered as proper OMML ``m:oMath`` elements.
    """
    parts = re.split(r"(\$[^$\n]+\$)", text)
    for part in parts:
        if part.startswith("$") and part.endswith("$"):
            math = part[1:-1].strip()
            om = latex_to_omath(math)
            if om is not None:
                paragraph._p.append(om)
            else:
                paragraph.add_run(math)
        elif part:
            paragraph.add_run(part)


def set_cell_math(cell, text: str):
    """Render text with ``$...$`` / ``$$...$$`` math into a table cell."""
    p = cell.paragraphs[0]
    p.clear()
    stripped = text.strip()
    if stripped.startswith("$$") and stripped.endswith("$$"):
        math = stripped[2:-2].strip()
        om = latex_to_omath(math)
        if om is not None:
            omathpara = ET.Element(_omml("oMathPara"))
            omathpara.append(om)
            p._p.append(omathpara)
        else:
            p.add_run(math)
    else:
        add_inline_math(p, text)
