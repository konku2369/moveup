# pdf_export.py
import os
from datetime import datetime
from typing import Optional, Tuple, List

import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.colors import Color
from kawaii_settings import load_settings, compute_effective_profile
import random
import subprocess


def _auto_open_file(path: str):
    """Open a file with the system default application, cross-platform."""
    try:
        if os.name == "nt":
            os.startfile(path)
        elif hasattr(os, "uname") and os.uname().sysname == "Darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except Exception:
        pass


from data_core import (
    COLUMNS_TO_USE,
    TYPE_TRUNC_LEN,
    ellipses,
    sort_with_backstock_priority,
    sanitize_prefix,
)

# ------------------------------
# PDF profile (single source of truth)
# ------------------------------
PDF_PROFILE = {
    "font_size": 9,
    "header_font_size": 9,
    "cell_padding_top": 2,
    "cell_padding_bottom": 2,

    "type_trunc": TYPE_TRUNC_LEN,
    "product_trunc": 75,
    "room_trunc": 12,
    "barcode_tail_moveup": 6,
    "metrc_tail_audit": 8,

    "moveup_widths": [50, 345, 60, 60, 30],
    "audit_widths":  [50, 345, 60, 55, 35],
}

BASE_PDF_FILENAME = "Print_me_Filtered_Move_Up"
DATE_FORMAT = "%B %d, %Y — %I:%M %p"


# ------------------------------
# Common helpers
# ------------------------------
def _fmt_type(val: str) -> str:
    return ellipses(str(val or ""), int(PDF_PROFILE["type_trunc"]))


def _fmt_product(val: str) -> str:
    return ellipses(str(val or ""), int(PDF_PROFILE["product_trunc"]))


def _fmt_room(val: str) -> str:
    return ellipses(str(val or ""), int(PDF_PROFILE["room_trunc"]))


def _fmt_barcode_tail(val: str, n: int) -> str:
    s = "" if val is None else str(val).strip()
    return s[-n:] if len(s) > n else s


def _set_alpha(canvas, a: float):
    try:
        canvas.setFillAlpha(a)
        canvas.setStrokeAlpha(a)
    except Exception:
        pass


def _draw_footer(canvas, doc):
    canvas.saveState()
    w, _h = letter
    y = 20
    canvas.setFont("Helvetica", 8)
    canvas.drawString(40, y, datetime.now().strftime(DATE_FORMAT))
    page_text = f"Page {canvas.getPageNumber()}"
    text_width = canvas.stringWidth(page_text, "Helvetica", 8)
    canvas.drawString(w - 40 - text_width, y, page_text)
    canvas.setLineWidth(0.25)
    canvas.setStrokeColor(colors.lightgrey)
    canvas.line(40, y + 10, w - 40, y + 10)
    canvas.restoreState()


def _draw_star(canvas, x: float, y: float, r: float):
    canvas.saveState()
    canvas.line(x - r, y, x + r, y)
    canvas.line(x, y - r, x, y + r)
    canvas.line(x - r * 0.7, y - r * 0.7, x + r * 0.7, y + r * 0.7)
    canvas.line(x - r * 0.7, y + r * 0.7, x + r * 0.7, y - r * 0.7)
    canvas.restoreState()




def _draw_random_stars(canvas, w, h, stroke_col, sparkle_alpha, count, seed):
    rng = random.Random(seed)
    canvas.saveState()
    _set_alpha(canvas, sparkle_alpha)
    canvas.setStrokeColor(stroke_col)
    canvas.setLineWidth(1.0)

    # Scatter only in the four narrow margin bands — keeps stars off the table content
    for _ in range(int(count)):
        band = rng.randint(0, 3)
        if band == 0:       # bottom margin strip
            x = rng.uniform(44, w - 44)
            y = rng.uniform(28, 70)
        elif band == 1:     # top margin strip
            x = rng.uniform(44, w - 44)
            y = rng.uniform(h - 70, h - 28)
        elif band == 2:     # left margin strip
            x = rng.uniform(28, 50)
            y = rng.uniform(70, h - 70)
        else:               # right margin strip
            x = rng.uniform(w - 50, w - 28)
            y = rng.uniform(70, h - 70)
        r = rng.uniform(3.5, 8.5)
        _draw_star(canvas, x, y, r)

    canvas.restoreState()





def _draw_daisy(canvas, x: float, y: float, scale: float, stroke_col: Color, alpha: float):
    canvas.saveState()
    _set_alpha(canvas, alpha)
    canvas.setStrokeColor(stroke_col)
    canvas.setLineWidth(max(0.6, 1.2 * scale))

    petals = 10
    petal_r = 10 * scale
    petal_dist = 16 * scale
    for i in range(petals):
        canvas.saveState()
        canvas.translate(x, y)
        canvas.rotate((360 / petals) * i)
        canvas.translate(petal_dist, 0)
        canvas.scale(1.6, 1.0)
        canvas.circle(0, 0, petal_r, stroke=1, fill=0)
        canvas.restoreState()

    canvas.setLineWidth(max(0.6, 1.0 * scale))
    canvas.circle(x, y, 7.5 * scale, stroke=1, fill=0)
    canvas.restoreState()


def _draw_paw(canvas, x: float, y: float, scale: float, stroke_col: Color, alpha: float):
    canvas.saveState()
    _set_alpha(canvas, alpha)
    canvas.setStrokeColor(stroke_col)
    canvas.setLineWidth(max(0.6, 1.1 * scale))

    canvas.circle(x, y, 8.5 * scale, stroke=1, fill=0)
    toes = [
        (x - 9 * scale, y + 10 * scale),
        (x - 3 * scale, y + 13 * scale),
        (x + 3 * scale, y + 13 * scale),
        (x + 9 * scale, y + 10 * scale),
    ]
    for tx, ty in toes:
        canvas.circle(tx, ty, 3.3 * scale, stroke=1, fill=0)

    canvas.restoreState()


def _kawaii_profile_from_settings(printer_bw: bool) -> dict:
    """Load settings once and compute the effective profile."""
    s = load_settings()
    s.printer_bw = bool(printer_bw)
    return compute_effective_profile(s)


def _draw_kawaii_background(canvas, doc, prof: dict):
    """Draw kawaii decorations using a pre-loaded profile dict."""
    w, h = letter
    canvas.saveState()

    is_bw = bool(prof.get("printer_bw", False))

    tint_a = float(prof.get("tint_alpha", 0.0))
    stroke_a = float(prof.get("stroke_alpha", 0.08))
    sparkle_a = float(prof.get("sparkle_alpha", 0.06))
    border_a = float(prof.get("border_alpha", 0.09))
    stars_count = int(prof.get("stars_count", 10))
    daisy_count = int(prof.get("daisy_count", 6))
    paw_count = int(prof.get("paw_count", 4))

    sr, sg, sb = prof.get("stroke_rgb", (0.55, 0.40, 0.50))
    stroke_col = Color(float(sr), float(sg), float(sb))

    # --- Tint wash (skip in B/W — keep background pure white) ---
    if tint_a > 0 and not is_bw:
        tr, tg, tb = prof.get("tint_rgb", (1.0, 0.84, 0.92))
        _set_alpha(canvas, tint_a)
        canvas.setFillColor(Color(float(tr), float(tg), float(tb)))
        canvas.rect(0, 0, w, h, stroke=0, fill=1)

    # --- Border ---
    _set_alpha(canvas, border_a)
    canvas.setStrokeColor(stroke_col)
    canvas.setLineWidth(1.0)
    margin = 26
    canvas.rect(margin, margin + 18, w - 2 * margin, h - (2 * margin + 34), stroke=1, fill=0)

    # --- Big corner daisy (bottom-right margin, away from table content) ---
    cx, cy = w * 0.84, h * 0.11
    canvas.saveState()
    _set_alpha(canvas, stroke_a)
    canvas.setStrokeColor(stroke_col)
    canvas.setLineWidth(1.4)

    petals = 12
    petal_r = 64
    petal_dist = 112
    canvas.translate(cx, cy)
    canvas.rotate(18)

    for i in range(petals):
        canvas.saveState()
        canvas.rotate((360 / petals) * i)
        canvas.translate(petal_dist, 0)
        canvas.scale(1.6, 1.0)
        canvas.circle(0, 0, petal_r, stroke=1, fill=0)
        canvas.restoreState()

    canvas.setLineWidth(1.2)
    canvas.circle(0, 0, 66, stroke=1, fill=0)
    canvas.restoreState()

    # --- Sparkles ---
    _set_alpha(canvas, sparkle_a)
    canvas.setStrokeColor(stroke_col)
    canvas.setLineWidth(1.0)

    # Fixed sparkle points — all in outer margin bands, never over table content
    sparkle_points = [
        (w * 0.14, h * 0.07, 9),   # bottom-left margin
        (w * 0.86, h * 0.07, 8),   # bottom-right margin
        (w * 0.08, h * 0.50, 7),   # mid-left margin
        (w * 0.92, h * 0.50, 7),   # mid-right margin
        (w * 0.50, h * 0.05, 6),   # bottom-center margin
    ]
    for x, y, r in sparkle_points:
        _draw_star(canvas, x, y, r)

    # --- Random stars (count driven by element intensity setting) ---
    if stars_count > 0:
        _draw_random_stars(canvas, w, h, stroke_col, sparkle_a, stars_count, seed=42)

    # --- Daisies & paws (pool ordered so first entries are drawn at lower intensity) ---
    # All positions placed in outer margin bands — never over the table content area
    daisy_positions = [
        (w * 0.05, h * 0.12, 0.95),   # 1 – bottom-left corner
        (w * 0.95, h * 0.12, 0.95),   # 2 – bottom-right corner
        (w * 0.04, h * 0.88, 0.85),   # 3 – top-left corner
        (w * 0.96, h * 0.88, 0.85),   # 4 – top-right corner
        (w * 0.50, h * 0.06, 0.75),   # 5 – bottom-center
        (w * 0.50, h * 0.94, 0.75),   # 6 – top-center
        (w * 0.04, h * 0.50, 0.75),   # 7 – mid-left edge
        (w * 0.96, h * 0.50, 0.75),   # 8 – mid-right edge
        (w * 0.04, h * 0.33, 0.60),   # 9 – lower-mid-left edge
        (w * 0.96, h * 0.33, 0.60),   # 10 – lower-mid-right edge
    ]
    for x, y, s in daisy_positions[:daisy_count]:
        _draw_daisy(canvas, x, y, s, stroke_col, stroke_a)

    # All positions placed in outer margin bands — never over the table content area
    paw_positions = [
        (w * 0.08, h * 0.10, 0.80),   # 1 – bottom-left corner
        (w * 0.92, h * 0.10, 0.80),   # 2 – bottom-right corner
        (w * 0.07, h * 0.88, 0.75),   # 3 – top-left corner
        (w * 0.93, h * 0.88, 0.75),   # 4 – top-right corner
        (w * 0.50, h * 0.04, 0.70),   # 5 – bottom-center
        (w * 0.50, h * 0.96, 0.70),   # 6 – top-center
        (w * 0.04, h * 0.65, 0.65),   # 7 – left side mid
    ]
    for x, y, s in paw_positions[:paw_count]:
        _draw_paw(canvas, x, y, s, stroke_col, stroke_a)

    canvas.restoreState()


def _draw_page(canvas, doc, kawaii_pdf: bool, prof: Optional[dict] = None):
    if kawaii_pdf and prof is not None:
        _draw_kawaii_background(canvas, doc, prof)
    _draw_footer(canvas, doc)


def _table_style_kawaii(printer_bw: bool):
    if printer_bw:
        header_bg = Color(0.92, 0.92, 0.92)
        row_a = Color(1.0, 1.0, 1.0)
        row_b = Color(0.95, 0.95, 0.95)
        grid = Color(0.75, 0.75, 0.75)
    else:
        header_bg = Color(0.96, 0.90, 0.95)
        row_a = Color(0.995, 0.965, 0.985)
        row_b = Color(0.98, 0.92, 0.96)
        grid = Color(0.72, 0.68, 0.74)
    return header_bg, row_a, row_b, grid


# ------------------------------
# Move-Up PDF helpers
# ------------------------------
def _prep_moveup_table_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Returns a df with columns:
      Type, Product Name, Package Barcode, Room, Qty On Hand
    with formatting applied for PDF printing.
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=["Type", "Product Name", "Package Barcode", "Room", "Qty On Hand"])

    pdf_df = df.loc[:, COLUMNS_TO_USE].copy()
    # Data is already sorted by the caller; no need to re-sort here.

    pdf_df["Room"] = pdf_df["Room"].fillna("").astype(str).map(_fmt_room)
    pdf_df["Type"] = pdf_df["Type"].fillna("").astype(str).map(_fmt_type)
    pdf_df["Product Name"] = pdf_df["Product Name"].fillna("").astype(str).map(_fmt_product)
    pdf_df["Package Barcode"] = pdf_df["Package Barcode"].map(
        lambda x: _fmt_barcode_tail(x, int(PDF_PROFILE["barcode_tail_moveup"]))
    )
    return pdf_df[["Type", "Product Name", "Package Barcode", "Room", "Qty On Hand"]]


def _build_moveup_page_elements(
    df_chunk: pd.DataFrame,
    title: str,
    kawaii_pdf: bool,
    printer_bw: bool
):
    styles = getSampleStyleSheet()
    elements: List = []

    widths = PDF_PROFILE["moveup_widths"]
    fs = int(PDF_PROFILE["font_size"])

    elements.append(Paragraph(f"<b>{title}</b>", styles["Heading2"]))
    elements.append(Paragraph(datetime.now().strftime(DATE_FORMAT), styles["Normal"]))
    elements.append(Paragraph(" ", styles["Normal"]))

    headers = ["Type", "Product", "Barcode", "Location", "Qty"]
    table_data = [headers] + df_chunk.values.tolist()
    table = Table(table_data, colWidths=widths)

    if kawaii_pdf:
        header_bg, row_a, row_b, grid = _table_style_kawaii(printer_bw)
    else:
        header_bg = colors.lightgrey
        row_a = colors.gainsboro
        row_b = None
        grid = colors.grey

    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_bg),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), fs),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("GRID", (0, 0), (-1, -1), 0.5, grid),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), int(PDF_PROFILE["cell_padding_bottom"])),
        ("TOPPADDING", (0, 0), (-1, -1), int(PDF_PROFILE["cell_padding_top"])),
    ]))

    for i in range(1, len(table_data), 2):
        table.setStyle([("BACKGROUND", (0, i), (-1, i), row_a)])
        if kawaii_pdf and (i + 1) < len(table_data) and row_b is not None:
            table.setStyle([("BACKGROUND", (0, i + 1), (-1, i + 1), row_b)])

    elements.append(table)
    return elements


# ------------------------------
# Audit PDF helpers
# ------------------------------
def _build_audit_page_elements(
    df_chunk: pd.DataFrame,
    title: str,
    mode: str,
    kawaii_pdf: bool,
    printer_bw: bool
):
    """
    df_chunk must have columns:
      Type, Product, METRC, Room / Notes, QtyOrCount
    """
    styles = getSampleStyleSheet()
    elements: List = []

    widths = PDF_PROFILE["audit_widths"]
    fs = int(PDF_PROFILE["font_size"])

    elements.append(Paragraph(f"<b>{title}</b>", styles["Heading2"]))
    elements.append(Paragraph(datetime.now().strftime(DATE_FORMAT), styles["Normal"]))
    elements.append(Paragraph(" ", styles["Normal"]))

    headers = (
        ["Type", "Product", "METRC", " ", "Qty"]
        if mode == "master"
        else ["Type", "Product", "METRC", " ", "Count"]
    )

    table_data = [headers] + df_chunk.values.tolist()
    table = Table(table_data, colWidths=widths)

    if kawaii_pdf:
        header_bg, row_a, row_b, grid = _table_style_kawaii(printer_bw)
    else:
        header_bg = colors.lightgrey
        row_a = colors.gainsboro
        row_b = None
        grid = colors.grey

    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_bg),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), fs),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("ALIGN", (-1, 0), (-1, -1), "RIGHT"),
        ("GRID", (0, 0), (-1, -1), 0.5, grid),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), int(PDF_PROFILE["cell_padding_top"])),
        ("BOTTOMPADDING", (0, 0), (-1, -1), int(PDF_PROFILE["cell_padding_bottom"])),
    ]))

    for i in range(1, len(table_data), 2):
        table.setStyle([("BACKGROUND", (0, i), (-1, i), row_a)])
        if kawaii_pdf and row_b and (i + 1) < len(table_data):
            table.setStyle([("BACKGROUND", (0, i + 1), (-1, i + 1), row_b)])

    elements.append(table)
    return elements


# ------------------------------
# Public export functions
# ------------------------------
def export_moveup_pdf_paginated(
    move_up_df: pd.DataFrame,
    priority_df: Optional[pd.DataFrame],
    base_dir: str,
    timestamp: bool,
    prefix: Optional[str],
    auto_open: bool,
    items_per_page: int = 30,
    kawaii_pdf: bool = False,
    printer_bw: bool = False,
) -> str:
    parts = [BASE_PDF_FILENAME]
    if timestamp:
        parts.append(datetime.now().strftime("%Y-%m-%d_%H-%M"))
    pdf_filename = "_".join(parts) + ".pdf"
    if prefix:
        prefix = sanitize_prefix(prefix)
        pdf_filename = f"{prefix}_{pdf_filename}"
    output_path = os.path.join(base_dir, pdf_filename)

    doc = SimpleDocTemplate(output_path, pagesize=letter)
    elements: List = []

    if move_up_df is None:
        move_up_df = pd.DataFrame(columns=COLUMNS_TO_USE)

    prio = priority_df.copy() if priority_df is not None else pd.DataFrame(columns=COLUMNS_TO_USE)
    rest = move_up_df.loc[:, COLUMNS_TO_USE].copy() if not move_up_df.empty else pd.DataFrame(columns=COLUMNS_TO_USE)

    # Remove priority items from regular list to avoid duplicates
    if not prio.empty and not rest.empty:
        prio_bcs = set(prio["Package Barcode"].astype(str).str.strip().tolist())
        rest = rest[~rest["Package Barcode"].astype(str).str.strip().isin(prio_bcs)].copy()

    # Mark priority items with ⭐ prefix on Product Name
    if not prio.empty:
        prio["Product Name"] = "\u2b50 " + prio["Product Name"].astype(str)

    # Split regular items by room: backstock first, then other rooms
    room_lower = rest["Room"].astype(str).str.strip().str.lower() if not rest.empty else pd.Series([], dtype=str)
    back_raw = rest[room_lower == "backstock"].copy() if not rest.empty else pd.DataFrame(columns=COLUMNS_TO_USE)
    other_raw = rest[room_lower != "backstock"].copy() if not rest.empty else pd.DataFrame(columns=COLUMNS_TO_USE)

    # Combine into one list: priority at top → backstock → other rooms
    combined = pd.concat([prio, back_raw, other_raw], ignore_index=True)

    title = "Move-Up Inventory List"

    if not combined.empty:
        combined_pdf_df = _prep_moveup_table_df(combined)
        for start in range(0, len(combined_pdf_df), items_per_page):
            chunk = combined_pdf_df.iloc[start:start + items_per_page]
            elements += _build_moveup_page_elements(chunk, title, kawaii_pdf, printer_bw)
            if start + items_per_page < len(combined_pdf_df):
                elements.append(PageBreak())

    prof = _kawaii_profile_from_settings(printer_bw) if kawaii_pdf else None

    doc.build(
        elements,
        onFirstPage=lambda c, d: _draw_page(c, d, kawaii_pdf, prof),
        onLaterPages=lambda c, d: _draw_page(c, d, kawaii_pdf, prof),
    )

    if auto_open:
        _auto_open_file(output_path)
    return output_path


def export_audit_pdfs(
    df: pd.DataFrame,
    base_dir: str,
    title_text: str,
    sort_mode: str,
    kawaii_pdf: bool,
    printer_bw: bool,
    auto_open: bool,
    default_store: str = "Store",
    default_room: str = "Sales Floor",
    type_trunc_len: int = TYPE_TRUNC_LEN,
) -> Tuple[str, str]:
    """
    Generates two PDFs:
      - Audit_Master_*.pdf (Qty column filled)
      - Audit_Blank_*.pdf  (Count column blank)
    """

    if df is None or df.empty:
        raise ValueError("Nothing to export.")

    work = df.copy()

    # Ensure required display columns exist (even if blank)
    for col in ["Type", "Brand", "Product Name", "Package Barcode", "Room", "Qty On Hand"]:
        if col not in work.columns:
            work[col] = ""

    # Optional audit meta
    if "Distributor" not in work.columns:
        work["Distributor"] = ""
    if "Store" not in work.columns:
        work["Store"] = ""
    if "Size" not in work.columns:
        work["Size"] = ""

    # Normalize
    work["Distributor"] = work["Distributor"].fillna("").astype(str).str.strip()
    work.loc[work["Distributor"] == "", "Distributor"] = "Unknown Distributor"

    work["Store"] = work["Store"].fillna("").astype(str).str.strip()
    work.loc[work["Store"] == "", "Store"] = (default_store.strip() or "Store")

    work["Room"] = work["Room"].fillna("").astype(str).str.strip()
    work.loc[work["Room"] == "", "Room"] = (default_room.strip() or "Sales Floor")

    work["SizeFull"] = work["Size"].fillna("").astype(str).str.strip()

    # Sorting/grouping
    if sort_mode == "distributor_type_size_product":
        group_col = "Distributor"
        group_label = "Distributor"
        sort_keys = ["Distributor", "Type", "SizeFull", "Product Name"]
    elif sort_mode == "brand_type_product":
        group_col = "Brand"
        group_label = "Brand"
        sort_keys = ["Brand", "Type", "SizeFull", "Product Name"]
    else:
        group_col = "Type"
        group_label = "Type"
        sort_keys = ["Type", "Brand", "SizeFull", "Product Name"]

    # Display transforms
    def _type_disp(v: str) -> str:
        s = str(v or "").strip()
        return (s[:type_trunc_len] + "…") if len(s) > int(type_trunc_len) else s

    work["TypeDisp"] = work["Type"].fillna("").astype(str).map(_type_disp)
    work["ProductDisp"] = work["Product Name"].fillna("").astype(str).map(lambda v: _fmt_product(v))
    work["MetrcLast8"] = work["Package Barcode"].fillna("").astype(str).map(
        lambda v: _fmt_barcode_tail(v, int(PDF_PROFILE["metrc_tail_audit"]))
    )

    work["Qty On Hand"] = pd.to_numeric(work["Qty On Hand"], errors="coerce").fillna(0).astype(int)

    sort_keys = [k for k in sort_keys if k in work.columns]
    if sort_keys:
        work = work.sort_values(sort_keys, kind="stable").reset_index(drop=True)

    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    master_path = os.path.join(base_dir, f"Audit_Master_{stamp}.pdf")
    blank_path = os.path.join(base_dir, f"Audit_Blank_{stamp}.pdf")
    styles = getSampleStyleSheet()

    def _build(path: str, mode: str):
        doc = SimpleDocTemplate(path, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
        elements: List = []

        elements.append(Paragraph(f"<b>{title_text}</b>", styles["Heading1"]))
        elements.append(Paragraph(datetime.now().strftime(DATE_FORMAT), styles["Normal"]))
        elements.append(Paragraph(" ", styles["Normal"]))

        groups = [g for g in work[group_col].dropna().astype(str).unique().tolist() if str(g).strip()]
        if not groups:
            groups = ["(Ungrouped)"]

        for gi, gval in enumerate(groups):
            gdf = work[work[group_col].astype(str) == str(gval)].copy()

            qty_or_count = (
                gdf["Qty On Hand"].astype(int).astype(str).tolist()
                if mode == "master"
                else [""] * len(gdf)
            )

            audit_table_df = pd.DataFrame({
                "Type": gdf["TypeDisp"].fillna("").astype(str),
                "Product": gdf["ProductDisp"].fillna("").astype(str),
                "METRC": gdf["MetrcLast8"].fillna("").astype(str),
                "Room / Notes": [""] * len(gdf),
                "QtyOrCount": qty_or_count,
            })

            elements += _build_audit_page_elements(
                df_chunk=audit_table_df,
                title=f"{group_label}: {gval}",
                mode=mode,
                kawaii_pdf=kawaii_pdf,
                printer_bw=printer_bw,
            )

            if gi < len(groups) - 1:
                elements.append(PageBreak())

        doc.build(
            elements,
            onFirstPage=lambda c, d: _draw_page(c, d, kawaii_pdf=kawaii_pdf, prof=audit_prof),
            onLaterPages=lambda c, d: _draw_page(c, d, kawaii_pdf=kawaii_pdf, prof=audit_prof),
        )

    audit_prof = _kawaii_profile_from_settings(printer_bw) if kawaii_pdf else None

    _build(master_path, "master")
    _build(blank_path, "blank")

    if auto_open:
        _auto_open_file(master_path)
        _auto_open_file(blank_path)

    return master_path, blank_path