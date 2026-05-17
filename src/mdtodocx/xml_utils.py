"""
XML utility functions for OOXML document manipulation.

Uses lxml for proper XML parsing and namespace-aware serialisation.
All functions operate on Element objects, not raw strings.

OOXML schema requires specific child ordering inside w:pPr, w:rPr, w:r, etc.
Any element placed out of order is technically invalid and triggers Word's
"unable to read content" recovery prompt. The helpers here keep elements
in schema order automatically.
"""

from __future__ import annotations
from lxml import etree as ET
from .config import W_NS, FontSize


def w(tag: str) -> str:
    """Create a namespaced tag name. w('rPr') → '{http://...}rPr'."""
    return f"{{{W_NS}}}{tag}"


def _local(tag: str) -> str:
    """Strip namespace prefix from a tag (Clark notation → local name)."""
    return tag.split("}")[-1] if "}" in tag else tag


# ============================================================
# OOXML SCHEMA CHILD-ORDER (subset, in document-spec order)
# ============================================================

# CT_PPr child order — paragraph properties
PPR_ORDER = (
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd",
    "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
    "sectPr", "pPrChange",
)

# CT_RPr child order — run properties
RPR_ORDER = (
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
    "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
    "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
    "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
    "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
    "eastAsianLayout", "specVanish", "oMath", "rPrChange",
)

# CT_TblPr child order — table properties (subset)
TBL_PR_ORDER = (
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
    "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
    "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption",
    "tblDescription", "tblPrChange",
)


# ============================================================
# CORE FIND / SET HELPERS
# ============================================================

def find(element: ET.Element, tag: str) -> ET.Element | None:
    """Find direct child by namespaced tag."""
    return element.find(w(tag))


def find_all(element: ET.Element, tag: str) -> list[ET.Element]:
    """Find all direct children by namespaced tag."""
    return element.findall(w(tag))


def find_desc(element: ET.Element, tag: str) -> ET.Element | None:
    """Find first descendant by namespaced tag (any depth)."""
    return element.find(f".//{w(tag)}")


def find_all_desc(element: ET.Element, tag: str) -> list[ET.Element]:
    """Find all descendants by namespaced tag (any depth)."""
    return element.findall(f".//{w(tag)}")


def set_attr(element: ET.Element, attr: str, value: str) -> None:
    """Set a namespaced attribute on element."""
    element.set(w(attr), value)


def get_attr(element: ET.Element, attr: str) -> str | None:
    """Get a namespaced attribute from element."""
    return element.get(w(attr))


# ============================================================
# SCHEMA-ORDERED INSERTION
# ============================================================

def _insert_ordered(parent: ET.Element, child: ET.Element, order: tuple[str, ...]) -> None:
    """
    Insert child into parent at the position required by `order`.

    Children whose tag is unknown to the schema list go at the end.
    Children with known tag are inserted before the first existing child
    that ranks higher in the schema.
    """
    child_name = _local(child.tag)
    if child_name not in order:
        parent.append(child)
        return
    target_idx = order.index(child_name)
    for i, existing in enumerate(parent):
        existing_name = _local(existing.tag)
        if existing_name not in order:
            continue
        if order.index(existing_name) > target_idx:
            parent.insert(i, child)
            return
    parent.append(child)


def ensure_child_ordered(
    parent: ET.Element, tag: str, order: tuple[str, ...]
) -> ET.Element:
    """Find or create a direct child element, inserting it in schema order."""
    child = find(parent, tag)
    if child is not None:
        return child
    child = ET.Element(w(tag))
    _insert_ordered(parent, child, order)
    return child


def ensure_child(parent: ET.Element, tag: str) -> ET.Element:
    """
    Find or create a direct child element (appended at end if created).

    Prefer `ensure_child_ordered` when the parent has a defined schema order.
    """
    child = find(parent, tag)
    if child is None:
        child = ET.SubElement(parent, w(tag))
    return child


# ============================================================
# RUN / RPR HELPERS
# ============================================================

def ensure_run_rpr(run: ET.Element) -> ET.Element:
    """
    Ensure run has rPr as its FIRST child, returning it.

    Per OOXML schema (CT_R), rPr must be the first child of w:r — before
    any w:t / w:br / etc. Using ET.SubElement would append rPr to the end,
    producing invalid XML that Word rejects with the recovery prompt.
    """
    rpr = find(run, "rPr")
    if rpr is None:
        rpr = ET.Element(w("rPr"))
        run.insert(0, rpr)
    return rpr


def get_ppr_rpr(para: ET.Element) -> ET.Element | None:
    """Get rPr inside pPr (paragraph run properties)."""
    ppr = find(para, "pPr")
    if ppr is None:
        return None
    return find(ppr, "rPr")


def ensure_ppr_rpr(para: ET.Element) -> ET.Element:
    """Ensure pPr exists and has rPr (in schema order). Returns the rPr."""
    ppr = ensure_child_ordered(para, "pPr", PPR_ORDER)
    return ensure_child_ordered(ppr, "rPr", PPR_ORDER)


def set_font_size(rpr: ET.Element, size_half_pts: int = FontSize.PT_14) -> None:
    """Set sz and szCs on an rPr element, keeping schema order."""
    for tag in ("sz", "szCs"):
        el = ensure_child_ordered(rpr, tag, RPR_ORDER)
        set_attr(el, "val", str(size_half_pts))


def ensure_bold(rpr: ET.Element) -> None:
    """Add bold (b + bCs) to rPr element if not present, in schema order."""
    for tag in ("b", "bCs"):
        ensure_child_ordered(rpr, tag, RPR_ORDER)


def make_rpr(size_half_pts: int, bold: bool = False) -> ET.Element:
    """Create a minimal rPr element with font size and optional bold."""
    rpr = ET.Element(w("rPr"))
    if bold:
        ensure_bold(rpr)
    set_font_size(rpr, size_half_pts)
    return rpr


def set_letter_spacing(rpr: ET.Element, value: int) -> None:
    """Set letter-spacing (w:spacing w:val) on rPr, in schema order."""
    spacing = ensure_child_ordered(rpr, "spacing", RPR_ORDER)
    set_attr(spacing, "val", str(value))


# ============================================================
# PARAGRAPH (PPR) HELPERS
# ============================================================

def set_spacing(para: ET.Element, line: int, line_rule: str = "auto") -> None:
    """Set line spacing (w:spacing line/lineRule) on paragraph's pPr."""
    ppr = ensure_child_ordered(para, "pPr", PPR_ORDER)
    spacing = ensure_child_ordered(ppr, "spacing", PPR_ORDER)
    set_attr(spacing, "line", str(line))
    set_attr(spacing, "lineRule", line_rule)


def set_alignment(para: ET.Element, value: str = "both") -> None:
    """Set justification alignment on paragraph's pPr."""
    ppr = ensure_child_ordered(para, "pPr", PPR_ORDER)
    jc = ensure_child_ordered(ppr, "jc", PPR_ORDER)
    set_attr(jc, "val", value)


def set_indent(para: ET.Element, left: int, first_line: int) -> None:
    """Set indentation on paragraph's pPr."""
    ppr = ensure_child_ordered(para, "pPr", PPR_ORDER)
    ind = ensure_child_ordered(ppr, "ind", PPR_ORDER)
    set_attr(ind, "left", str(left))
    set_attr(ind, "firstLine", str(first_line))


# ============================================================
# TABLE HELPERS
# ============================================================

def set_table_width(tbl_pr: ET.Element, width: int, width_type: str = "pct") -> None:
    """Set table width on tblPr."""
    tbl_w = ensure_child_ordered(tbl_pr, "tblW", TBL_PR_ORDER)
    set_attr(tbl_w, "w", str(width))
    set_attr(tbl_w, "type", width_type)


def set_table_borders(
    tbl_pr: ET.Element,
    sides: tuple[str, ...] = ("top", "left", "bottom", "right", "insideH", "insideV"),
    val: str = "single",
    sz: int = 4,
    color: str = "auto",
) -> None:
    """
    Set table borders on tblPr.

    Default: 1/2pt single black borders on all six sides (outer + inner grid).
    `sz` is in eighths of a point (4 = 1/2 pt, 8 = 1 pt).
    """
    tbl_borders = ensure_child_ordered(tbl_pr, "tblBorders", TBL_PR_ORDER)
    # Clear any pre-existing border children to avoid duplicates
    for child in list(tbl_borders):
        tbl_borders.remove(child)
    border_order = ("top", "left", "bottom", "right", "insideH", "insideV")
    for side in border_order:
        if side not in sides:
            continue
        b = ET.SubElement(tbl_borders, w(side))
        set_attr(b, "val", val)
        set_attr(b, "sz", str(sz))
        set_attr(b, "space", "0")
        set_attr(b, "color", color)


# ============================================================
# PRETTY-PRINT
# ============================================================

def indent_xml(element: ET.Element, level: int = 0) -> None:
    """Add pretty-print indentation to an XML tree (in-place)."""
    indent = "\n" + "  " * level
    if len(element):
        if not element.text or not element.text.strip():
            element.text = indent + "  "
        last_child = None
        for child in element:
            indent_xml(child, level + 1)
            last_child = child
        if last_child is not None and (not last_child.tail or not last_child.tail.strip()):
            last_child.tail = indent
    if not element.tail or not element.tail.strip():
        element.tail = indent if level else "\n"
