"""
Configuration constants for mdToDocx converter.

All magic numbers, paths, and style names in one place.
"""

import os
import sys

# ============================================================
# PATHS
# ============================================================

# When running from a PyInstaller bundle, data files are in sys._MEIPASS.
# Otherwise, project root is 2 levels up from this file.
if getattr(sys, "frozen", False):
    _DATA_ROOT = sys._MEIPASS
else:
    _DATA_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

PROJECT_ROOT = _DATA_ROOT
DEFAULT_REF = os.path.join(_DATA_ROOT, "ref.docx")
LUA_FILTER = os.path.join(_DATA_ROOT, "filters", "academic.lua")

# ============================================================
# STYLE NAMES — must match docx/ref/word/styles.xml
# ============================================================

class Styles:
    """Style IDs used in the reference DOCX template."""
    CENTER_HEADING = "center_heading"
    HEADING_1 = "heading_1"
    HEADING_2 = "heading_2"
    HEADING_3 = "heading_3"
    MAIN_TEXT = "main_text"
    MARKER_LIST = "marker_list"
    NUM_LIST = "num_list"
    TABLE_HEADING = "table_heading"
    IMAGE_HEADING = "image_heading"

# ============================================================
# FORMATTING CONSTANTS (half-points / twips / OOXML values)
# ============================================================

class FontSize:
    """Font sizes in half-points (OOXML w:sz units)."""
    PT_14 = 28  # 14pt × 2
    PT_12 = 24  # 12pt × 2

class Spacing:
    """Line spacing values (OOXML w:spacing units)."""
    LINE_150 = 360   # 1.5 line spacing (240 × 1.5)
    LETTER_TRACK = 20  # letter-spacing for table captions

class Indent:
    """Indentation values in twips (1 twip = 1/1440 inch)."""
    FIRST_LINE_125CM = 709  # 1.25cm ≈ 709 twips
    LEFT_ZERO = 0

class Table:
    """Table formatting constants."""
    WIDTH_PCT = 5000  # 100% width in OOXML pct units
    BORDER_SZ = 4     # border size in 1/8 pt (4 = 1/2 pt single line)

# ============================================================
# NAMESPACES
# ============================================================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W_NS}
