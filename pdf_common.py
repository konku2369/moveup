"""
Reusable PDF table export library.

Standalone module — no Tk, no app-specific imports.
Drop into any project that needs clean table PDFs with optional kawaii theming.

Usage::

    from pdf_common import build_section_pdf, PALETTE_KAWAII

    sections = [
        ("Inventory", ["SKU", "Product", "Qty"], [["A1", "Widget", 5], ...]),
    ]
    build_section_pdf("report.pdf", "My Report", "March 2026", sections,
                      palette=PALETTE_KAWAII)
"""

import os
from datetime import datetime
from typing import Any, Dict, List, Optional, Sequence, Tuple

from reportlab.lib import colors
from reportlab.lib.colors import Color
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


# ---------------------------------------------------------------------------
# Color palettes
# ---------------------------------------------------------------------------
# Each palette is a dict with keys: header, row_a, row_b, grid, text

PALETTE_KAWAII = {
    "header": Color(0.96, 0.90, 0.95),
    "row_a":  Color(0.995, 0.965, 0.985),
    "row_b":  Color(0.98, 0.92, 0.96),
    "grid":   Color(0.72, 0.68, 0.74),
    "text":   colors.black,
}

PALETTE_BW = {
    "header": Color(0.92, 0.92, 0.92),
    "row_a":  Color(1.0, 1.0, 1.0),
    "row_b":  Color(0.95, 0.95, 0.95),
    "grid":   Color(0.75, 0.75, 0.75),
    "text":   colors.black,
}

PALETTE_PLAIN = {
    "header": colors.lightgrey,
    "row_a":  colors.whitesmoke,
    "row_b":  colors.white,
    "grid":   colors.grey,
    "text":   colors.black,
}


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

def truncate_text(val: Any, max_len: int) -> str:
    """
    Convert *val* to a string and hard-truncate it to *max_len* characters.

    If truncation occurs, the last three characters of the result are replaced
    with ``"..."`` so the total length is still *max_len* (not *max_len* + 3).
    ``None`` is treated as an empty string.

    Parameters
    ----------
    val : Any
        Value to convert and truncate.  ``str(val)`` is called unconditionally.
    max_len : int
        Maximum character length of the returned string.

    Returns
    -------
    str
        The (possibly truncated) string.
    """
    s = str(val) if val is not None else ""
    return s if len(s) <= max_len else s[:max_len - 3] + "..."


# ---------------------------------------------------------------------------
# Column width heuristics
# ---------------------------------------------------------------------------

# Default multiplier rules applied by column name keywords.
# Product names need the most space (3x), barcodes are short numeric strings (0.7x).
# These multipliers are relative — they're normalized to sum to total_width.
_WIDTH_RULES: List[Tuple[Sequence[str], float]] = [
    (("product", "name"),       3.0),   # wide — long product names
    (("brand",),                1.5),   # medium — brand names
    (("barcode", "metrc", "sku"), 0.7), # narrow — short ID strings
    (("room", "location"),      1.0),   # standard
]


def compute_column_widths(
    columns: List[str],
    total_width: float = 742,
    overrides: Optional[Dict[str, float]] = None,
    min_width: float = 50,
) -> List[float]:
    """
    Auto-compute column widths that sum to exactly *total_width*.

    Algorithm:
    1. Compute a *base* width = ``total_width / len(columns)`` (clamped to ≥ 60).
    2. For each column, look up a multiplier from ``_WIDTH_RULES`` by matching
       keywords in the lower-cased column name.  Caller *overrides* take
       precedence over the rule table.
    3. Scale all raw widths proportionally so they sum to exactly *total_width*.
    4. Clamp each result to *min_width* to prevent zero-width columns.

    Default multipliers (from ``_WIDTH_RULES``):

    ============================================  ===========
    Column name contains                          Multiplier
    ============================================  ===========
    ``"product"`` or ``"name"``                   3.0 × base
    ``"brand"``                                   1.5 × base
    ``"barcode"``, ``"metrc"``, or ``"sku"``      0.7 × base
    ``"room"`` or ``"location"``                  1.0 × base
    anything else                                 1.0 × base
    ============================================  ===========

    Parameters
    ----------
    columns : list[str]
        Column header names.
    total_width : float
        Total available horizontal space in points (default 742 = landscape
        letter minus standard margins).
    overrides : dict[str, float] | None
        Map of column name → multiplier to apply instead of the rule table.
        Keys must match the column name exactly (case-sensitive).
    min_width : float
        Minimum width in points for any single column.

    Returns
    -------
    list[float]
        Per-column widths, one entry per column, summing to approximately
        *total_width* (may be slightly under due to *min_width* clamping).
    """
    overrides = overrides or {}
    base = max(60, int(total_width / max(1, len(columns))))
    raw: List[float] = []

    for col in columns:
        low = str(col).lower()

        # Check caller overrides first
        if col in overrides:
            raw.append(base * overrides[col])
            continue

        # Apply keyword rules
        mult = 1.0
        for keywords, m in _WIDTH_RULES:
            if any(k in low for k in keywords):
                mult = m
                break
        raw.append(base * mult)

    # Scale to fit total_width
    s = sum(raw)
    if s > 0:
        scale = total_width / s
        return [max(min_width, int(w * scale)) for w in raw]
    return [max(min_width, int(total_width / max(1, len(columns))))] * len(columns)


# ---------------------------------------------------------------------------
# Table style builder
# ---------------------------------------------------------------------------

def build_table_style(
    palette: Dict[str, Any],
    num_rows: int,
    header_font_size: int = 9,
    body_font_size: int = 8,
    padding: int = 4,
    extra_commands: Optional[List] = None,
) -> TableStyle:
    """
    Build a complete ReportLab ``TableStyle`` from a color palette dict.

    Sets header row styling (bold, larger font, colored background), body row
    alternating backgrounds (``row_a`` / ``row_b``), a thin grid, MIDDLE
    vertical alignment, LEFT horizontal alignment, and uniform cell padding.

    Parameters
    ----------
    palette : dict
        A palette dict with keys ``header``, ``row_a``, ``row_b``, ``grid``,
        and optionally ``text`` (defaults to black).  Use one of
        ``PALETTE_KAWAII``, ``PALETTE_BW``, or ``PALETTE_PLAIN``.
    num_rows : int
        Total number of rows including the header.  Not currently used in
        the style commands but kept for API consistency with callers that
        may need it for future per-row logic.
    header_font_size : int
        Point size for the header row.
    body_font_size : int
        Point size for data rows.
    padding : int
        Top and bottom cell padding in points.
    extra_commands : list | None
        Additional ReportLab TableStyle command tuples to append.  Processed
        after the base style so they can override any default setting.

    Returns
    -------
    TableStyle
        A ready-to-use ReportLab ``TableStyle`` instance.
    """
    commands = [
        # Header row
        ("FONTNAME",       (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",       (0, 0), (-1, 0), header_font_size),
        ("BACKGROUND",     (0, 0), (-1, 0), palette["header"]),
        ("TEXTCOLOR",      (0, 0), (-1, 0), palette.get("text", colors.black)),

        # Body rows
        ("FONTSIZE",       (0, 1), (-1, -1), body_font_size),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [palette["row_a"], palette["row_b"]]),

        # Grid & alignment
        ("GRID",           (0, 0), (-1, -1), 0.25, palette["grid"]),
        ("VALIGN",         (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",          (0, 0), (-1, -1), "LEFT"),

        # Padding
        ("TOPPADDING",     (0, 0), (-1, -1), padding),
        ("BOTTOMPADDING",  (0, 0), (-1, -1), padding),
    ]

    if extra_commands:
        commands.extend(extra_commands)

    return TableStyle(commands)


# ---------------------------------------------------------------------------
# Smart column alignment
# ---------------------------------------------------------------------------

_CENTER_KEYWORDS = ("qty", "count", "quantity", "changes", "age", "score",
                    "delta", "sells", "change", "unchanged")
_RIGHT_KEYWORDS  = ("cost", "price", "value", "retail", "total")


def auto_align_commands(columns: List[str]) -> List[tuple]:
    """
    Infer column alignment from column names and return TableStyle commands.

    Columns whose lower-cased names contain any ``_CENTER_KEYWORDS`` keyword
    (e.g. ``"qty"``, ``"count"``, ``"score"``) get ``ALIGN CENTER`` commands.
    Columns matching ``_RIGHT_KEYWORDS`` (e.g. ``"cost"``, ``"price"``) get
    ``ALIGN RIGHT``.  All other columns are left at the default LEFT alignment
    set by ``build_table_style()``.

    Parameters
    ----------
    columns : list[str]
        Column header names, in the same order as the table data.

    Returns
    -------
    list[tuple]
        ReportLab TableStyle command tuples (may be empty if no columns match).
        Pass as the *extra_commands* argument to ``build_table_style()`` or
        directly to ``TableStyle()``.
    """
    cmds = []
    for i, col in enumerate(columns):
        low = str(col).lower()
        if any(k in low for k in _CENTER_KEYWORDS):
            cmds.append(("ALIGN", (i, 0), (i, -1), "CENTER"))
        elif any(k in low for k in _RIGHT_KEYWORDS):
            cmds.append(("ALIGN", (i, 0), (i, -1), "RIGHT"))
    return cmds


# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------

_DATE_FORMAT = "%m/%d/%Y %I:%M %p"


def draw_footer(canvas, doc, timestamp_str: Optional[str] = None):
    """
    ReportLab page callback that draws a footer on every page.

    Renders three elements at y=20 pt from the bottom of the page:
    - **Left**: generation timestamp (``timestamp_str`` or ``datetime.now()``).
    - **Right**: ``"Page N"`` right-aligned to the right margin.
    - **Above both**: a thin 0.25pt light-grey horizontal rule at y=30.

    Designed to be passed as both ``onFirstPage`` and ``onLaterPages`` to
    ``SimpleDocTemplate.build()``.  The ``canvas.saveState()`` / ``restoreState()``
    calls ensure that footer drawing does not affect the main page content.

    Parameters
    ----------
    canvas : reportlab.pdfgen.canvas.Canvas
        The current page canvas (injected by ReportLab).
    doc : SimpleDocTemplate
        The document object (used for ``doc.pagesize``).
    timestamp_str : str | None
        Pre-formatted timestamp string to show in the footer.  If ``None``,
        the current local time is formatted as ``"MM/DD/YYYY HH:MM AM/PM"``.
    """
    canvas.saveState()
    w, _h = doc.pagesize
    y = 20

    canvas.setFont("Helvetica", 8)
    ts = timestamp_str or datetime.now().strftime(_DATE_FORMAT)
    canvas.drawString(40, y, ts)

    page_text = f"Page {canvas.getPageNumber()}"
    tw = canvas.stringWidth(page_text, "Helvetica", 8)
    canvas.drawString(w - 40 - tw, y, page_text)

    canvas.setLineWidth(0.25)
    canvas.setStrokeColor(colors.lightgrey)
    canvas.line(40, y + 10, w - 40, y + 10)
    canvas.restoreState()


# ---------------------------------------------------------------------------
# High-level PDF builder
# ---------------------------------------------------------------------------

Section = Tuple[Optional[str], List[str], List[List[Any]]]
# (section_title | None,  column_headers,  data_rows)


def build_section_pdf(
    path: str,
    title: str,
    subtitle: str,
    sections: List[Section],
    palette: Optional[Dict[str, Any]] = None,
    orientation: str = "landscape",
    margins: int = 24,
    footer: bool = True,
    width_overrides: Optional[Dict[str, float]] = None,
    extra_style_fn=None,
) -> str:
    """
    Build a multi-section table PDF.

    Parameters
    ----------
    path : str
        Output file path.
    title, subtitle : str
        Document header text.
    sections : list of (section_title, column_headers, data_rows)
        Each section becomes an optional heading + a styled table.
        *section_title* may be None to omit the heading.
    palette : dict, optional
        Color palette (default: PALETTE_PLAIN).
    orientation : "landscape" | "portrait"
        Page orientation.
    margins : int
        Page margins in points (all four sides).
    footer : bool
        Whether to draw timestamp + page-number footer.
    width_overrides : dict, optional
        Column-name → width multiplier overrides passed to compute_column_widths().
    extra_style_fn : callable, optional
        ``fn(columns, table_data) → list[tuple]`` returning additional
        TableStyle commands for app-specific styling (e.g., velocity colors).

    Returns
    -------
    str
        The output file path (same as *path*).
    """
    pal = palette or PALETTE_PLAIN
    pagesize = landscape(letter) if orientation == "landscape" else letter

    doc = SimpleDocTemplate(
        path,
        pagesize=pagesize,
        leftMargin=margins,
        rightMargin=margins,
        topMargin=margins,
        bottomMargin=margins,
        title=title,
    )

    styles = getSampleStyleSheet()
    story: List = []

    story.append(Paragraph(title, styles["Title"]))
    story.append(Paragraph(subtitle, styles["Normal"]))
    story.append(Spacer(1, 12))

    # Compute total available width from page dimensions
    pw, _ = pagesize
    total_width = pw - 2 * margins

    for section_title, columns, data_rows in sections:
        if section_title:
            story.append(Spacer(1, 14))
            story.append(Paragraph(section_title, styles["Heading2"]))
            story.append(Spacer(1, 6))

        if not data_rows:
            story.append(Paragraph("No data.", styles["Normal"]))
            continue

        table_data = [list(columns)] + [list(row) for row in data_rows]

        col_widths = compute_column_widths(
            columns, total_width=total_width, overrides=width_overrides,
        )

        # Auto-alignment + caller extras
        extra = auto_align_commands(columns)
        if extra_style_fn:
            extra.extend(extra_style_fn(columns, table_data))

        style = build_table_style(pal, len(table_data), extra_commands=extra)

        tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(style)
        story.append(tbl)

    # Build with optional footer
    if footer:
        doc.build(
            story,
            onFirstPage=draw_footer,
            onLaterPages=draw_footer,
        )
    else:
        doc.build(story)

    return path
