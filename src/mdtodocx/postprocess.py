"""
DOCX post-processing: fix tables, lists, and numbering via XML manipulation.

Pipeline: pandoc produces DOCX → this module fixes formatting → final DOCX.
Uses xml.etree.ElementTree for proper XML parsing (no regex on raw XML).
"""

import os
import tempfile
import zipfile
from lxml import etree as ET

from .config import FontSize, Spacing, Indent, Table, W_NS
from .xml_utils import (
    w, find, find_all, find_desc, find_all_desc,
    ensure_child, set_attr, get_attr,
    set_font_size, ensure_bold, make_rpr,
    get_ppr_rpr, ensure_ppr_rpr, ensure_run_rpr,
    set_spacing, set_alignment, set_indent,
    set_table_width, set_table_borders,
    set_letter_spacing,
)


def _register_namespaces() -> None:
    """Register all OOXML namespaces so ET.write() produces clean output."""
    namespaces = {
        "w": W_NS,
        "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
        "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    }
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)


def _load_xml(path: str):
    """Load XML file."""
    return ET.parse(path)


def _save_xml(tree, path: str) -> None:
    """
    Save XML with declaration and UTF-8 encoding.

    Office expects double-quoted, standalone="yes" XML declarations. Stdlib
    ElementTree emits single quotes and omits standalone, which some strict
    OOXML readers (including some Word versions) flag as malformed package
    parts. Post-process the file to normalise the declaration.
    """
    tree.write(path, xml_declaration=True, encoding="UTF-8")
    _normalise_xml_declaration(path)


def _normalise_xml_declaration(path: str) -> None:
    """Rewrite the XML declaration to Office's preferred form."""
    with open(path, "rb") as f:
        data = f.read()
    target_decl = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    if data.startswith(b"<?xml"):
        end = data.find(b"?>") + 2
        # Replace whatever ET wrote with the canonical declaration
        data = target_decl + data[end:]
    else:
        data = target_decl + b"\n" + data
    with open(path, "wb") as f:
        f.write(data)


# ============================================================
# TABLE FIXES
# ============================================================

def _fix_table_caption(para: ET.Element) -> None:
    """Fix a table caption paragraph: 14pt font + letter-spacing 20."""
    rpr = ensure_ppr_rpr(para)
    set_font_size(rpr, FontSize.PT_14)
    set_letter_spacing(rpr, Spacing.LETTER_TRACK)

    # Also fix each run's rPr — must be FIRST child of run (OOXML schema).
    for run in find_all(para, "r"):
        run_rpr = ensure_run_rpr(run)
        set_font_size(run_rpr, FontSize.PT_14)
        set_letter_spacing(run_rpr, Spacing.LETTER_TRACK)


def _fix_image_caption(para: ET.Element) -> None:
    """Fix an image caption paragraph: 12pt font, centered, 1.5 spacing."""
    rpr = ensure_ppr_rpr(para)
    set_font_size(rpr, FontSize.PT_12)
    set_spacing(para, Spacing.LINE_150)

    # Also fix each run's rPr — must be FIRST child of run (OOXML schema).
    for run in find_all(para, "r"):
        run_rpr = ensure_run_rpr(run)
        set_font_size(run_rpr, FontSize.PT_12)


def _is_table_caption(para: ET.Element) -> bool:
    """
    Check if paragraph is a table caption.

    Identification by paragraph style (`table_heading` set by Lua filter,
    or pandoc's default `TableCaption`). Avoids fuzzy text matching that
    would falsely flag body paragraphs containing "Таблица 1 - ..." in
    quoted text.
    """
    ppr = find(para, "pPr")
    if ppr is None:
        return False
    pstyle = find(ppr, "pStyle")
    if pstyle is None:
        return False
    return get_attr(pstyle, "val") in ("table_heading", "TableCaption")


def _fix_table_block(tbl: ET.Element) -> None:
    """Fix a single table: width, borders, header row, body rows."""
    # Fix table width to 100% and add full borders (outer + inner grid)
    tbl_pr = find(tbl, "tblPr")
    if tbl_pr is not None:
        set_table_width(tbl_pr, Table.WIDTH_PCT, "pct")
        set_table_borders(tbl_pr, sz=Table.BORDER_SZ)

    # Process rows
    rows = find_all(tbl, "tr")
    for i, row in enumerate(rows):
        is_header = (i == 0)

        # Also check for tblHeader element in trPr
        tr_pr = find(row, "trPr")
        if tr_pr is not None and find(tr_pr, "tblHeader") is not None:
            is_header = True

        size = FontSize.PT_12
        bold = is_header

        # Fix each cell's paragraphs
        for cell in find_all(row, "tc"):
            for para in find_all(cell, "p"):
                _fix_cell_paragraph(para, size, bold)


def _fix_cell_paragraph(para: ET.Element, size: int, bold: bool) -> None:
    """Fix font size and bold in a table cell paragraph."""
    # Fix pPr/rPr
    ppr_rpr = get_ppr_rpr(para)
    if ppr_rpr is not None:
        set_font_size(ppr_rpr, size)
        if bold:
            ensure_bold(ppr_rpr)

    # Fix each run — rPr must be FIRST child of run (OOXML schema).
    for run in find_all(para, "r"):
        rpr = ensure_run_rpr(run)
        set_font_size(rpr, size)
        if bold:
            ensure_bold(rpr)


def fix_tables(root: ET.Element) -> None:
    """Fix all tables in document.xml root element."""
    # Fix table captions
    for para in find_all_desc(root, "p"):
        if _is_table_caption(para):
            _fix_table_caption(para)

    # Fix table blocks
    for tbl in find_all_desc(root, "tbl"):
        _fix_table_block(tbl)


# ============================================================
# LIST FIXES
# ============================================================

def _has_num_pr(para: ET.Element) -> bool:
    """Check if paragraph has numbering properties (is a list item)."""
    ppr = find(para, "pPr")
    if ppr is None:
        return False
    return find(ppr, "numPr") is not None


def _get_list_style(para: ET.Element, num_fmt_map: dict) -> str:
    """Determine list style (marker_list or num_list) based on numbering format."""
    from .config import Styles
    ppr = find(para, "pPr")
    if ppr is None:
        return Styles.MARKER_LIST
    num_pr = find(ppr, "numPr")
    if num_pr is None:
        return Styles.MARKER_LIST
    num_id_el = find(num_pr, "numId")
    if num_id_el is None:
        return Styles.MARKER_LIST
    num_id = get_attr(num_id_el, "val")
    ilvl_el = find(num_pr, "ilvl")
    ilvl = get_attr(ilvl_el, "val") if ilvl_el is not None else "0"
    key = f"{num_id}:{ilvl}"
    fmt = num_fmt_map.get(key, "bullet")
    return Styles.NUM_LIST if fmt in ("decimal", "upperLetter", "lowerLetter") else Styles.MARKER_LIST


def _build_num_fmt_map(docx_path: str) -> dict:
    """Build map of numId:ilvl → numFmt from numbering.xml."""
    import tempfile
    result = {}
    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(docx_path, "r") as z:
            z.extractall(tmpdir)
        numbering_path = os.path.join(tmpdir, "word", "numbering.xml")
        if not os.path.exists(numbering_path):
            return result
        tree = ET.parse(numbering_path)
        root = tree.getroot()
        if root is None:
            return result
        # Build abstractNum map
        abstract_map = {}
        for abs_num in find_all_desc(root, "abstractNum"):
            abs_id = get_attr(abs_num, "abstractNumId")
            for lvl in find_all(abs_num, "lvl"):
                ilvl = get_attr(lvl, "ilvl") or "0"
                num_fmt = find(lvl, "numFmt")
                fmt = get_attr(num_fmt, "val") if num_fmt is not None else "bullet"
                abstract_map[f"{abs_id}:{ilvl}"] = fmt
        # Map num → abstractNum
        for num in find_all_desc(root, "num"):
            num_id = get_attr(num, "numId")
            abs_ref = find(num, "abstractNumId")
            if abs_ref is not None:
                abs_id = get_attr(abs_ref, "val")
                for key, fmt in abstract_map.items():
                    if key.startswith(f"{abs_id}:"):
                        ilvl = key.split(":")[1]
                        result[f"{num_id}:{ilvl}"] = fmt
    return result


def _set_paragraph_style(para: ET.Element, style_name: str) -> None:
    """Set paragraph style (pStyle) to given style name."""
    from .config import W_NS
    ppr = ensure_child(para, "pPr")
    pstyle = find(ppr, "pStyle")
    if pstyle is None:
        pstyle = ET.SubElement(ppr, w("pStyle"))
    set_attr(pstyle, "val", style_name)


def fix_list_styles(root: ET.Element, num_fmt_map: dict) -> None:
    """Fix list paragraph styles: aa/Compact → marker_list/num_list."""
    for para in find_all_desc(root, "p"):
        if _has_num_pr(para):
            style = _get_list_style(para, num_fmt_map)
            _set_paragraph_style(para, style)


def fix_caption_styles(root: ET.Element) -> None:
    """Fix table and image caption styles.

    Handles:
    - TableCaption → table_heading (Pandoc built-in → our style)
    - ImageCaption → table_heading + centering (Pandoc built-in for figures)
    - CaptionedFigure → main_text + centering (Pandoc built-in for image paragraphs)
    - image_heading → table_heading + centering (from Lua filter for image captions)
    """
    from .config import Styles
    for para in find_all_desc(root, "p"):
        ppr = find(para, "pPr")
        if ppr is None:
            continue
        pstyle = find(ppr, "pStyle")
        if pstyle is None:
            continue
        val = get_attr(pstyle, "val")

        if val == "TableCaption":
            set_attr(pstyle, "val", Styles.TABLE_HEADING)

        elif val == "ImageCaption":
            # Pandoc's built-in style for figure captions → our style + centering
            set_attr(pstyle, "val", Styles.TABLE_HEADING)
            set_alignment(para, "center")
            _fix_table_caption(para)

        elif val == "CaptionedFigure":
            # Pandoc's built-in style for image paragraphs → our style + centering
            set_attr(pstyle, "val", Styles.MAIN_TEXT)
            set_alignment(para, "center")

        elif val == "image_heading":
            # From Lua filter: image caption → image_heading + centering
            # Style stays as image_heading (mapped to imageheading by fix_pstyle_aliases)
            set_alignment(para, "center")
            _fix_image_caption(para)


def fix_image_paragraphs(root: ET.Element) -> None:
    """Center paragraphs that contain images (drawings).

    After the Lua filter converts Figure → main_text Div, Pandoc emits
    paragraphs with main_text style containing <w:drawing>. This function
    adds w:jc=center to those paragraphs so images are centered.
    """
    for para in find_all_desc(root, "p"):
        # Only process paragraphs that contain a drawing (image)
        if find_desc(para, "drawing") is None:
            continue
        set_alignment(para, "center")


# Pandoc emits pStyle values matching the style NAME from the reference docx
# when the lua filter uses {custom-style="…"}. The reference docx, however,
# uses condensed styleIds (no underscore). Word looks up styles by styleId,
# not by name, so unmapped values fall back to Normal — losing 14pt, list
# numbering, and triggering the "unable to read content" recovery prompt.
#
# This map matches the {styleId, name} pairs in docx/ref/word/styles.xml.
PSTYLE_NAME_TO_ID = {
    "marker_list": "markerlist",
    "num_list": "numlist",
    "table_heading": "tableheading",
    "main_text": "maintext",
    "center_heading": "centerheading",
    "image_heading": "imageheading",
}


def fix_pstyle_aliases(root: ET.Element) -> None:
    """Rewrite pStyle values from style names to the actual styleIds."""
    for ps in find_all_desc(root, "pStyle"):
        val = get_attr(ps, "val")
        if val in PSTYLE_NAME_TO_ID:
            set_attr(ps, "val", PSTYLE_NAME_TO_ID[val])


def _collect_style_ids(styles_root: ET.Element) -> set[str]:
    """Return the set of all styleIds defined in styles.xml."""
    ids = set()
    for s in find_all(styles_root, "style"):
        sid = get_attr(s, "styleId")
        if sid:
            ids.add(sid)
    return ids


def drop_dangling_style_refs(
    doc_root: ET.Element, valid_style_ids: set[str]
) -> None:
    """
    Remove pStyle / rStyle / tblStyle references that don't resolve to a
    style in styles.xml.

    Pandoc emits a small set of internal style names (Compact for table cells,
    FirstParagraph for the first body paragraph, Table for table style) that
    are not present in our reference docx. A dangling reference makes Word
    flag the document with the "unable to read content" recovery prompt.
    Stripping the reference makes Word fall back to its default ("Normal" /
    "Table Normal"), which is the desired behaviour for these paragraphs.
    """
    for tag in ("pStyle", "rStyle", "tblStyle"):
        # Find all references and remove those pointing to unknown styleIds
        for ref in list(find_all_desc(doc_root, tag)):
            val = get_attr(ref, "val")
            if val and val not in valid_style_ids:
                # Locate the parent and remove this child element
                # ElementTree doesn't expose parent pointers, so traverse.
                for parent in doc_root.iter():
                    if ref in list(parent):
                        parent.remove(ref)
                        break


def _fix_list_paragraph(para: ET.Element) -> None:
    """Fix a single list paragraph: 14pt, indent, spacing, alignment."""
    size = FontSize.PT_14

    # Fix paragraph properties
    set_spacing(para, Spacing.LINE_150)
    set_alignment(para, "both")
    set_indent(para, Indent.LEFT_ZERO, Indent.FIRST_LINE_125CM)

    # Fix pPr/rPr
    ppr_rpr = ensure_ppr_rpr(para)
    set_font_size(ppr_rpr, size)

    # Fix each run — rPr must be FIRST child of run (OOXML schema).
    for run in find_all(para, "r"):
        rpr = ensure_run_rpr(run)
        # Remove misplaced letter-spacing carried over from elsewhere
        spacing = find(rpr, "spacing")
        if spacing is not None:
            rpr.remove(spacing)
        set_font_size(rpr, size)


def fix_lists(root: ET.Element) -> None:
    """Fix all list paragraphs in document.xml root element."""
    for para in find_all_desc(root, "p"):
        if _has_num_pr(para):
            _fix_list_paragraph(para)


# ============================================================
# NUMBERING FIXES
# ============================================================

def _fix_bullet_level(lvl: ET.Element) -> None:
    """Fix a single bullet level: hyphen marker, remove Symbol font, 14pt."""
    # Change bullet character to hyphen
    lvl_text = find(lvl, "lvlText")
    if lvl_text is not None:
        set_attr(lvl_text, "val", "-")

    # Remove Symbol font override
    rpr = find(lvl, "rPr")
    if rpr is not None:
        rfonts = find(rpr, "rFonts")
        if rfonts is not None:
            rpr.remove(rfonts)
        set_font_size(rpr, FontSize.PT_14)
    else:
        # Add rPr with font size
        rpr = ET.SubElement(lvl, w("rPr"))
        set_font_size(rpr, FontSize.PT_14)


def _fix_misconfigured_decimal_level(lvl: ET.Element) -> None:
    """Fix levels with numFmt=decimal but lvlText='-': convert to bullet."""
    # Change format to bullet
    num_fmt = find(lvl, "numFmt")
    if num_fmt is not None:
        set_attr(num_fmt, "val", "bullet")

    # Ensure lvlText is hyphen
    lvl_text = find(lvl, "lvlText")
    if lvl_text is not None:
        set_attr(lvl_text, "val", "-")

    # Ensure rPr with font size 14pt
    rpr = find(lvl, "rPr")
    if rpr is not None:
        rfonts = find(rpr, "rFonts")
        if rfonts is not None:
            rpr.remove(rfonts)
        set_font_size(rpr, FontSize.PT_14)
    else:
        rpr = ET.SubElement(lvl, w("rPr"))
        set_font_size(rpr, FontSize.PT_14)


def _dedupe_abstract_num_nsids(root: ET.Element) -> None:
    """
    Ensure every <w:abstractNum> has a unique <w:nsid w:val="..."/> value.

    Pandoc, when augmenting the reference docx's numbering.xml with new
    abstractNums for ordered/bullet lists in the source, copies nsid values
    from the original abstractNums it cloned. The resulting duplicate nsids
    confuse Word: abstractNums sharing an nsid are treated as the same list
    definition, and when one of them carries a <w:pStyle> link to a paragraph
    style (e.g. markerlist or numlist), Word picks the linked one for *every*
    paragraph that references either abstractNum — even via an explicit
    paragraph-level numId. The visible effect is that bullet lists render as
    numbered and numbered lists render as bullets.

    Fix: walk all abstractNums, keep the first occurrence of each nsid, and
    assign unique nsid values to subsequent duplicates.
    """
    seen = set()
    next_unique = 0xF0000000  # high range, won't collide with existing values
    for an in find_all_desc(root, "abstractNum"):
        nsid = find(an, "nsid")
        if nsid is None:
            continue
        val = get_attr(nsid, "val")
        if val is None or val in seen:
            # Generate a fresh 8-hex-digit nsid value
            new_val = f"{next_unique:08X}"
            while new_val in seen:
                next_unique += 1
                new_val = f"{next_unique:08X}"
            set_attr(nsid, "val", new_val)
            seen.add(new_val)
            next_unique += 1
        else:
            seen.add(val)


def fix_numbering(docx_path: str) -> None:
    """
    Fix numbering.xml:
    - Bullet levels: hyphen marker, remove Symbol font, 14pt
    - Misconfigured decimal (numFmt=decimal + lvlText='-'): convert to bullet + 14pt
    - Deduplicate <w:nsid> values so Word does not collapse distinct lists
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(docx_path, "r") as z:
            z.extractall(tmpdir)

        numbering_path = os.path.join(tmpdir, "word", "numbering.xml")
        if not os.path.exists(numbering_path):
            return

        tree = _load_xml(numbering_path)
        root = tree.getroot()
        assert root is not None, "numbering.xml has no root element"

        for lvl in find_all_desc(root, "lvl"):
            num_fmt = find(lvl, "numFmt")
            if num_fmt is None:
                continue
            fmt = get_attr(num_fmt, "val")
            lvl_text = find(lvl, "lvlText")
            text_val = get_attr(lvl_text, "val") if lvl_text is not None else ""

            if fmt == "bullet":
                _fix_bullet_level(lvl)
            elif fmt == "decimal" and text_val == "-":
                # Misconfigured: decimal format with hyphen text → should be bullet
                _fix_misconfigured_decimal_level(lvl)

        _dedupe_abstract_num_nsids(root)

        # Pretty-print disabled: adds whitespace text nodes that Word rejects
        _save_xml(tree, numbering_path)

        # Repack docx
        _repack_docx(docx_path, tmpdir)


# ============================================================
# DOCX ZIP OPERATIONS
# ============================================================

def _repack_docx(docx_path: str, tmpdir: str) -> None:
    """Repack a directory into a DOCX ZIP file."""
    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as z:
        for dirpath, dirnames, filenames in os.walk(tmpdir):
            for fn in filenames:
                full_path = os.path.join(dirpath, fn)
                arcname = os.path.relpath(full_path, tmpdir)
                z.write(full_path, arcname)


def _extract_docx(docx_path: str, tmpdir: str) -> None:
    """Extract a DOCX ZIP file to a directory."""
    with zipfile.ZipFile(docx_path, "r") as z:
        z.extractall(tmpdir)


# ============================================================
# PACKAGE-LEVEL CLEANUP
# ============================================================

CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _read_xml(path: str):
    """Parse an XML file and return its root element."""
    return ET.parse(path).getroot()


def _write_xml(root: ET.Element, path: str) -> None:
    """Write an XML element tree to disk with normalised declaration."""
    _save_xml(ET.ElementTree(root), path)


def _write_xml_default_ns(root: ET.Element, path: str, ns_uri: str) -> None:
    """
    Write XML with a default (unprefixed) namespace.

    lxml does not support ``register_namespace("", uri)`` so we serialise
    to bytes and fix the generated ``ns0:`` prefix by replacing it with
    the bare tag name, then inserting the proper ``xmlns="..."`` declaration
    on the root element.
    """
    raw = ET.tostring(root, xml_declaration=True, encoding="UTF-8",
                       standalone=True)
    # lxml may output single-quoted declaration; normalise later.
    # Replace ns0: prefix with nothing and add default xmlns.
    ns_prefix = b"ns0:"
    if ns_prefix in raw:
        raw = raw.replace(ns_prefix, b"")
        # Insert xmlns="..." right after the root tag name
        root_tag = root.tag.split("}")[-1].encode()
        old = b"<" + root_tag
        new = b"<" + root_tag + b' xmlns="' + ns_uri.encode() + b'"'
        raw = raw.replace(old, new, 1)
    with open(path, "wb") as f:
        f.write(raw)
    _normalise_xml_declaration(path)


def _dedupe_content_types(ct_path: str) -> None:
    """
    Remove duplicate <Override> entries from [Content_Types].xml.

    Pandoc can emit duplicate content-type entries (e.g. when the reference
    docx already contains a placeholder image and pandoc adds the same image
    again). Word treats duplicate PartName overrides as a corruption signal
    and shows the "unable to read content" recovery prompt.
    """
    if not os.path.exists(ct_path):
        return
    ct_root = _read_xml(ct_path)
    seen: set[tuple[str, str]] = set()
    for child in list(ct_root):
        tag = child.tag
        if tag == f"{{{CT_NS}}}Override":
            key = (child.get("PartName") or "", child.get("ContentType") or "")
            if key in seen:
                ct_root.remove(child)
            else:
                seen.add(key)
    _write_xml_default_ns(ct_root, ct_path, CT_NS)


def cleanup_package(tmpdir: str) -> None:
    """
    Strip parts that pandoc creates but never references with content.

    Pandoc populates the package with an `odttf` font extension default,
    an empty `word/comments.xml`, an empty `docProps/custom.xml`, and an
    empty `word/_rels/footnotes.xml.rels`. The corresponding overrides in
    `[Content_Types].xml` and relationships in `_rels/.rels` /
    `word/_rels/document.xml.rels` reference these inert parts.

    Word treats this combination of "declared but unused / empty" parts as
    a corruption signal and shows the "unable to read content" recovery
    prompt. Removing the empty parts together with their declarations
    silences the prompt.

    Also deduplicates <Override> entries in [Content_Types].xml — pandoc
    can emit duplicate content-type entries when the reference docx already
    contains a placeholder image.

    Also normalises every XML file's declaration to the form Office writes
    (`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`). Some Word
    versions are stricter than the spec about this prologue.
    """
    word_dir = os.path.join(tmpdir, "word")
    docprops_dir = os.path.join(tmpdir, "docProps")
    rels_dir = os.path.join(tmpdir, "_rels")
    word_rels_dir = os.path.join(word_dir, "_rels")

    # Helper: read a small file's content as bytes
    def _read(path: str) -> bytes:
        with open(path, "rb") as f:
            return f.read()

    # Identify candidates for removal: empty pandoc-stub parts.
    candidates: list[tuple[str, str | None]] = []
    # (path, sentinel substring that confirms the file is empty/unused)
    candidates += [
        (os.path.join(word_dir, "comments.xml"), "<w:comments"),
        (os.path.join(docprops_dir, "custom.xml"), "<Properties"),
        (os.path.join(word_rels_dir, "footnotes.xml.rels"), "<Relationships"),
    ]

    removed_parts: set[str] = set()
    for path, sentinel in candidates:
        if not os.path.exists(path):
            continue
        content = _read(path)
        # "Empty" = self-closing root with no element children
        # Heuristic: the file's content has no non-namespace child element
        try:
            root = ET.fromstring(content)
        except ET.XMLSyntaxError:
            continue
        if len(root) == 0:
            os.remove(path)
            # Track the in-package path (relative, using forward slashes)
            rel = os.path.relpath(path, tmpdir).replace(os.sep, "/")
            removed_parts.add(rel)

    # 1) Always deduplicate [Content_Types].xml — pandoc can emit duplicate
    #    <Override> entries for images that already exist in the reference docx.
    ct_path = os.path.join(tmpdir, "[Content_Types].xml")
    if os.path.exists(ct_path):
        ct_root = _read_xml(ct_path)
        # First: deduplicate all Override entries
        _dedupe_content_types(ct_path)
        # Then: strip entries for removed parts and obsolete defaults
        # (re-read after dedup since _dedupe_content_types rewrote the file)
        ct_root = _read_xml(ct_path)
        if removed_parts:
            for child in list(ct_root):
                tag = child.tag
                if tag == f"{{{CT_NS}}}Override":
                    pn = (child.get("PartName") or "").lstrip("/")
                    if pn in removed_parts:
                        ct_root.remove(child)
                elif tag == f"{{{CT_NS}}}Default":
                    # Drop the obfuscated-font default if no .odttf files remain
                    if child.get("Extension") == "odttf":
                        ct_root.remove(child)
            _write_xml_default_ns(ct_root, ct_path, CT_NS)

    # 2) Strip Relationship entries in *.rels that target removed parts.
    if removed_parts:
        def _strip_rels(rels_path: str, target_filter) -> None:
            if not os.path.exists(rels_path):
                return
            rels_root = _read_xml(rels_path)
            for child in list(rels_root):
                target = child.get("Target") or ""
                if target_filter(target):
                    rels_root.remove(child)
            _write_xml_default_ns(rels_root, rels_path, PKG_REL_NS)

        pkg_rels = os.path.join(rels_dir, ".rels")
        _strip_rels(
            pkg_rels,
            # _rels/.rels targets are relative to package root
            lambda t: t.lstrip("/") in removed_parts,
        )

        doc_rels = os.path.join(word_rels_dir, "document.xml.rels")
        _strip_rels(
            doc_rels,
            # word/_rels/document.xml.rels targets are relative to word/
            lambda t: f"word/{t}" in removed_parts,
        )


def normalise_all_xml_declarations(tmpdir: str) -> None:
    """
    Rewrite every .xml / .rels file in the package so the prologue matches
    Office's canonical form: <?xml version="1.0" encoding="UTF-8" standalone="yes"?>.
    """
    for dirpath, _, filenames in os.walk(tmpdir):
        for fn in filenames:
            if fn.endswith(".xml") or fn.endswith(".rels"):
                _normalise_xml_declaration(os.path.join(dirpath, fn))


# ============================================================
# MAIN POST-PROCESSING ENTRY POINT
# ============================================================

def postprocess_docx(docx_path: str) -> None:
    """Apply all post-processing fixes to a DOCX file."""
    _register_namespaces()

    # Build numbering format map (bullet vs decimal)
    num_fmt_map = _build_num_fmt_map(docx_path)

    with tempfile.TemporaryDirectory() as tmpdir:
        _extract_docx(docx_path, tmpdir)

        doc_path = os.path.join(tmpdir, "word", "document.xml")
        tree = _load_xml(doc_path)
        root = tree.getroot()
        assert root is not None, "document.xml has no root element"

        # Load valid styleIds from styles.xml so we can drop dangling refs
        styles_path = os.path.join(tmpdir, "word", "styles.xml")
        valid_style_ids: set[str] = set()
        if os.path.exists(styles_path):
            styles_tree = _load_xml(styles_path)
            styles_root = styles_tree.getroot()
            if styles_root is not None:
                valid_style_ids = _collect_style_ids(styles_root)

        # Fix styles first (before formatting)
        fix_list_styles(root, num_fmt_map)
        fix_caption_styles(root)

        # Fix formatting
        fix_tables(root)
        fix_lists(root)
        fix_image_paragraphs(root)  # Center image paragraphs

        # Final step: rewrite pStyle values from style names to actual styleIds
        # so Word can resolve them (otherwise lists / captions fall back to Normal
        # and Word emits an "unable to read content" recovery prompt).
        fix_pstyle_aliases(root)

        # Drop pandoc-emitted style refs (Compact / FirstParagraph / Table)
        # that do not correspond to any style in styles.xml — Word treats
        # dangling refs as a corruption signal.
        if valid_style_ids:
            drop_dangling_style_refs(root, valid_style_ids)

        # Pretty-print disabled: adds whitespace text nodes that Word rejects
        _save_xml(tree, doc_path)

        # Drop empty pandoc-stub parts (comments.xml, custom.xml, etc.)
        # together with their content-type overrides and relationships.
        cleanup_package(tmpdir)

        # Final pass: align every XML prologue with Office's canonical form
        # (double-quoted, standalone="yes"). Some Word versions reject the
        # bare ET-style declaration with single quotes / no standalone.
        normalise_all_xml_declarations(tmpdir)

        _repack_docx(docx_path, tmpdir)

    # Fix numbering in a separate pass (different XML file)
    fix_numbering(docx_path)
