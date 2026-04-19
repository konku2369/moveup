"""
PDF export for move-up sticker sheets and audit reports.

Generates paginated portrait PDFs (move-up) and landscape PDFs (audit)
with optional kawaii decorations (stars, daisies, paws, cat faces).
Uses kawaii_settings.py for decoration profiles.

TWO TYPES OF PDF EXPORTS:
=========================
1. MOVE-UP STICKER PDF (export_moveup_pdf_paginated)
   - Portrait letter-size pages
   - Priority items (⭐) appear first, then backstock, then other rooms
   - Each page holds ~30-35 items (configurable)
   - Staff prints these, cuts them into strips, and sticks them on products
     to indicate "bring this to the sales floor"

2. AUDIT PDFs (export_audit_pdfs)
   - Generates TWO PDFs: Master (with qty) and Blank (empty count column)
   - Master PDF: reference copy showing what SHOULD be on the shelf
   - Blank PDF: staff walks the floor and writes actual counts
   - Grouped by distributor, brand, or type (user chooses sort mode)
   - Supports Accessory Audit mode (filters to accessories only)

KAWAII DECORATIONS:
  When kawaii mode is on, pages get decorative overlays:
  - Background tint (pink/lavender or greyscale)
  - Decorative border
  - Corner daisies, scattered stars, paw prints, cat faces
  - Element counts scale with elem_intensity slider
  - All decorations stay in page margins — never overlap the table content
  - Controlled by kawaii_pdf_settings.json via kawaii_settings.py

NOTE ON pdf_common.py:
  This file does NOT use pdf_common.py yet. The table styling here is
  more specialized (kawaii color blending, per-page decorations). The
  generic pdf_common.py is for NEW projects like EarthMed. This file
  stays as-is because it already works perfectly for moveup's needs.
"""

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
    except Exception as e:
        print(f"[moveup] _auto_open_file failed: {e}")


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
# Tuning constants for PDF table layout. Widths are in points (72pt = 1 inch)
# and are tuned for letter-size paper (8.5" × 11").
PDF_PROFILE = {
    "font_size": 9,
    "header_font_size": 9,
    "cell_padding_top": 2,
    "cell_padding_bottom": 2,

    "type_trunc": TYPE_TRUNC_LEN,    # max chars for Type column
    "product_trunc": 75,              # max chars for Product Name
    "room_trunc": 12,                 # max chars for Room
    "barcode_tail_moveup": 6,         # show last N digits of barcode in move-up PDF
    "metrc_tail_audit": 8,            # show last N digits in audit PDF

    # Column widths in points: [Type, Product Name, Barcode, Room, Qty]
    "moveup_widths": [50, 345, 60, 60, 30],
    "audit_widths":  [50, 345, 60, 55, 35],
}

BASE_PDF_FILENAME = "Print_me_Filtered_Move_Up"
DATE_FORMAT = "%B %d, %Y — %I:%M %p"


# ------------------------------
# Common helpers
# ------------------------------
def _fmt_field(val, key: str) -> str:
    """Truncate *val* to the character limit stored in PDF_PROFILE[key], appending '…' if cut."""
    return ellipses(str(val or ""), int(PDF_PROFILE[key]))

_fmt_type    = lambda v: _fmt_field(v, "type_trunc")
_fmt_product = lambda v: _fmt_field(v, "product_trunc")
_fmt_room    = lambda v: _fmt_field(v, "room_trunc")


def _fmt_barcode_tail(val: str, n: int) -> str:
    """Return the last *n* characters of the barcode string, or the full string if shorter than *n*."""
    s = "" if val is None else str(val).strip()
    return s[-n:] if len(s) > n else s


def _set_alpha(canvas, a: float):
    """Set fill and stroke alpha on the ReportLab canvas.

    Silently ignores AttributeError for ReportLab builds that don't support
    per-object transparency (very old versions), so kawaii decorations degrade
    gracefully to opaque rather than crashing the export.
    """
    try:
        canvas.setFillAlpha(a)
        canvas.setStrokeAlpha(a)
    except AttributeError:
        pass


def _draw_footer(canvas, doc):
    """Draw a page footer with timestamp (left) and page number (right), separated by a hairline rule.

    Called by _draw_page as the onFirstPage/onLaterPages callback so it runs once per PDF page.
    """
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
    """Draw a 4-line sparkle/star at (x, y) with radius r."""
    canvas.saveState()
    canvas.line(x - r, y, x + r, y)           # horizontal
    canvas.line(x, y - r, x, y + r)           # vertical
    canvas.line(x - r * 0.7, y - r * 0.7, x + r * 0.7, y + r * 0.7)  # diagonal ↘
    canvas.line(x - r * 0.7, y + r * 0.7, x + r * 0.7, y - r * 0.7)  # diagonal ↗
    canvas.restoreState()




def _draw_random_stars(canvas, w, h, stroke_col, sparkle_alpha, count, seed):
    """Scatter sparkle stars in page margin bands (avoiding table content area)."""
    rng = random.Random(seed)  # deterministic seed so same PDF = same star layout
    canvas.saveState()
    _set_alpha(canvas, sparkle_alpha)
    canvas.setStrokeColor(stroke_col)
    canvas.setLineWidth(1.0)

    # Four margin bands: top, bottom, left, right. Stars never overlap the table.
    for _ in range(int(count)):
        band = rng.randint(0, 3)
        if band == 0:       # bottom margin strip
            x = rng.uniform(44, w - 44)
            y = rng.uniform(28, 72)
        elif band == 1:     # top margin strip
            x = rng.uniform(44, w - 44)
            y = rng.uniform(h - 72, h - 28)
        elif band == 2:     # left margin strip
            x = rng.uniform(28, 52)
            y = rng.uniform(72, h - 72)
        else:               # right margin strip
            x = rng.uniform(w - 52, w - 28)
            y = rng.uniform(72, h - 72)
        r = rng.uniform(4.0, 13.0)
        _draw_star(canvas, x, y, r)

    canvas.restoreState()


def _draw_random_cats(canvas, w, h, stroke_col, stroke_alpha, count, seed):
    """Scatter *count* cat faces randomly across the four page-margin bands.

    Uses a deterministic RNG (seeded per-export) so the same PDF always gets the same
    cat positions. Scale varies from 0.28–0.58 so faces range from tiny to medium.
    """
    rng = random.Random(seed)
    for _ in range(int(count)):
        band = rng.randint(0, 3)
        if band == 0:       # bottom margin
            x = rng.uniform(50, w - 50)
            y = rng.uniform(30, 78)
        elif band == 1:     # top margin
            x = rng.uniform(50, w - 50)
            y = rng.uniform(h - 78, h - 30)
        elif band == 2:     # left margin
            x = rng.uniform(30, 58)
            y = rng.uniform(78, h - 78)
        else:               # right margin
            x = rng.uniform(w - 58, w - 30)
            y = rng.uniform(78, h - 78)
        scale = rng.uniform(0.28, 0.58)
        _draw_cat_face(canvas, x, y, scale, stroke_col, stroke_alpha)



def _draw_daisy(canvas, x: float, y: float, scale: float, stroke_col: Color, alpha: float):
    """Draw a 10-petal daisy at (x, y), scaled by *scale*, with the given stroke color and alpha.

    Petals are drawn as slightly squashed ellipses (1.6:1 aspect ratio) arranged evenly
    around a central circle. scale=1.0 produces petals at radius ~16pt and center circle ~7.5pt.
    """
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
    """Draw a cat paw print at (x, y): one large heel pad plus four small toe circles above it.

    At scale=1.0 the heel pad has radius ~8.5pt and each toe circle ~3.3pt.
    Toes are arranged in a gentle arc (−9, −3, +3, +9 × scale horizontally) ~10–13pt above center.
    """
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


def _draw_cat_face(canvas, x: float, y: float, scale: float, stroke_col: Color, alpha: float):
    """Draw a minimal line-art cat face at (x, y).

    Components:
    - Circular head (radius 9×scale)
    - Two pointed ear triangles projecting from upper sides of head
    - Two circular eyes, one circular nose
    - Six whiskers (3 per side), fanning slightly up, flat, and down

    scale=0.3–0.6 is typical for margin decoration; scale=1.0 is ~9pt radius (tiny).
    """
    canvas.saveState()
    _set_alpha(canvas, alpha)
    canvas.setStrokeColor(stroke_col)

    r = 9 * scale

    # Head
    canvas.setLineWidth(max(0.6, 1.1 * scale))
    canvas.circle(x, y, r, stroke=1, fill=0)

    # Ears — two pointed triangles on top
    canvas.setLineWidth(max(0.5, 0.9 * scale))
    for side in (-1, 1):
        ex = x + side * 6.5 * scale
        ey_base = y + r * 0.6
        p = canvas.beginPath()
        p.moveTo(ex - 4 * scale, ey_base)
        p.lineTo(ex + 4 * scale, ey_base)
        p.lineTo(ex + side * 2.5 * scale, ey_base + 8 * scale)
        p.close()
        canvas.drawPath(p, stroke=1, fill=0)

    # Eyes
    canvas.setLineWidth(max(0.5, 0.8 * scale))
    canvas.circle(x - 3 * scale, y + 1.5 * scale, 1.8 * scale, stroke=1, fill=0)
    canvas.circle(x + 3 * scale, y + 1.5 * scale, 1.8 * scale, stroke=1, fill=0)

    # Nose
    canvas.circle(x, y - 1.5 * scale, 1.0 * scale, stroke=1, fill=0)

    # Whiskers — 3 per side, fanning slightly up/flat/down
    canvas.setLineWidth(max(0.4, 0.6 * scale))
    wlen = 7 * scale
    wy = y - 1.5 * scale
    for side in (-1, 1):
        x0 = x + side * 1.8 * scale
        for dy in (-wlen * 0.35, 0.0, wlen * 0.35):
            canvas.line(x0, wy, x0 + side * wlen, wy + dy)

    canvas.restoreState()


def _kawaii_profile_from_settings(printer_bw: bool) -> dict:
    """Load settings once and compute the effective profile."""
    s = load_settings()
    s.printer_bw = bool(printer_bw)
    return compute_effective_profile(s)


def _draw_kawaii_tint_and_border(canvas, w, h, prof, stroke_col, is_bw, rng):
    """Tint wash + decorative border."""
    tint_a = float(prof.get("tint_alpha", 0.0))
    border_a = float(prof.get("border_alpha", 0.09))

    if tint_a > 0 and not is_bw:
        tr, tg, tb = prof.get("tint_rgb", (1.0, 0.84, 0.92))
        _set_alpha(canvas, tint_a)
        canvas.setFillColor(Color(float(tr), float(tg), float(tb)))
        canvas.rect(0, 0, w, h, stroke=0, fill=1)

    _set_alpha(canvas, border_a)
    canvas.setStrokeColor(stroke_col)
    canvas.setLineWidth(1.0)
    margin = 26
    canvas.rect(margin, margin + 18, w - 2 * margin, h - (2 * margin + 34), stroke=1, fill=0)


def _draw_kawaii_corner_daisy(canvas, w, h, stroke_col, stroke_a, rng, _jx, _jy):
    """Big corner daisy in bottom-right margin."""
    cx, cy = w * 0.84 + _jx() * 0.5, h * 0.11 + _jy() * 0.5
    canvas.saveState()
    _set_alpha(canvas, stroke_a)
    canvas.setStrokeColor(stroke_col)
    canvas.setLineWidth(1.4)

    petals = 12
    petal_r = 64
    petal_dist = 112
    canvas.translate(cx, cy)
    canvas.rotate(rng.uniform(0, 30))

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


def _draw_kawaii_sparkles(canvas, w, h, stroke_col, sparkle_a, rng, _jx, _jy):
    """Fixed sparkle points with small jitter in outer margin bands."""
    _set_alpha(canvas, sparkle_a)
    canvas.setStrokeColor(stroke_col)
    canvas.setLineWidth(1.0)

    sparkle_points = [
        (w * 0.14, h * 0.07, 9),
        (w * 0.86, h * 0.07, 8),
        (w * 0.08, h * 0.50, 7),
        (w * 0.92, h * 0.50, 7),
        (w * 0.50, h * 0.05, 6),
        (w * 0.25, h * 0.04, 7),
        (w * 0.75, h * 0.04, 7),
    ]
    for x, y, r in sparkle_points:
        _draw_star(canvas, x + _jx(), y + _jy(), r * rng.uniform(0.8, 1.2))


def _draw_kawaii_element_pools(canvas, w, h, stroke_col, stroke_a, sparkle_a,
                                prof, jitter_seed, _jx, _jy):
    """Stars, daisies, paws, and cat faces — counts driven by profile."""
    stars_count = int(prof.get("stars_count", 18))
    daisy_count = int(prof.get("daisy_count", 9))
    paw_count = int(prof.get("paw_count", 6))
    cat_count = int(prof.get("cat_count", 4))

    if stars_count > 0:
        _draw_random_stars(canvas, w, h, stroke_col, sparkle_a, stars_count, seed=jitter_seed + 1)

    # Kawaii element "pools" — positions ordered by visual priority.
    # At low elem_intensity, only the first few (corners) are drawn.
    # At max intensity, all positions are filled. The third value is scale (0-1).
    # Daisies — pool of 15 fixed positions
    daisy_positions = [
        (w * 0.05, h * 0.12, 0.95),   #  1 – bottom-left corner
        (w * 0.95, h * 0.12, 0.95),   #  2 – bottom-right corner
        (w * 0.04, h * 0.88, 0.85),   #  3 – top-left corner
        (w * 0.96, h * 0.88, 0.85),   #  4 – top-right corner
        (w * 0.50, h * 0.06, 0.75),   #  5 – bottom-center
        (w * 0.50, h * 0.94, 0.75),   #  6 – top-center
        (w * 0.04, h * 0.50, 0.75),   #  7 – mid-left edge
        (w * 0.96, h * 0.50, 0.75),   #  8 – mid-right edge
        (w * 0.04, h * 0.33, 0.65),   #  9 – lower-mid-left edge
        (w * 0.96, h * 0.33, 0.65),   # 10 – lower-mid-right edge
        (w * 0.04, h * 0.25, 0.60),   # 11 – left edge, lower quarter
        (w * 0.96, h * 0.25, 0.60),   # 12 – right edge, lower quarter
        (w * 0.04, h * 0.67, 0.60),   # 13 – left edge, upper quarter
        (w * 0.96, h * 0.67, 0.60),   # 14 – right edge, upper quarter
        (w * 0.28, h * 0.05, 0.55),   # 15 – bottom, quarter-left
    ]
    for x, y, s in daisy_positions[:daisy_count]:
        _draw_daisy(canvas, x + _jx(), y + _jy(), s, stroke_col, stroke_a)

    # Paws — pool of 10
    paw_positions = [
        (w * 0.08, h * 0.10, 0.80),   #  1 – bottom-left corner
        (w * 0.92, h * 0.10, 0.80),   #  2 – bottom-right corner
        (w * 0.07, h * 0.88, 0.75),   #  3 – top-left corner
        (w * 0.93, h * 0.88, 0.75),   #  4 – top-right corner
        (w * 0.50, h * 0.04, 0.70),   #  5 – bottom-center
        (w * 0.50, h * 0.96, 0.70),   #  6 – top-center
        (w * 0.04, h * 0.65, 0.65),   #  7 – left side mid
        (w * 0.96, h * 0.65, 0.65),   #  8 – right side mid
        (w * 0.04, h * 0.42, 0.60),   #  9 – left side low-mid
        (w * 0.96, h * 0.42, 0.60),   # 10 – right side low-mid
    ]
    for x, y, s in paw_positions[:paw_count]:
        _draw_paw(canvas, x + _jx(), y + _jy(), s, stroke_col, stroke_a)

    # Cat faces — scattered randomly
    if cat_count > 0:
        _draw_random_cats(canvas, w, h, stroke_col, stroke_a, cat_count, seed=jitter_seed + 2)


def _draw_kawaii_background(canvas, doc, prof: dict):
    """Draw all kawaii decorations for one page using a pre-computed profile dict.

    Rendering order (back to front):
      1. Background tint wash — color fill over the full page
      2. Decorative border rectangle — inside page margins
      3. Large watermark corner daisy — bottom-right margin
      4. Fixed sparkle points — outer margin bands with slight jitter
      5. Element pools — stars, daisy set, paw set, and cat faces (counts from profile)

    Parameters
    ----------
    prof : dict
        Output of compute_effective_profile(). Expected keys: printer_bw, stroke_alpha,
        sparkle_alpha, stroke_rgb, tint_alpha, tint_rgb, border_alpha, jitter_seed,
        stars_count, daisy_count, paw_count, cat_count.
    """
    w, h = letter
    canvas.saveState()

    is_bw = bool(prof.get("printer_bw", False))
    stroke_a = float(prof.get("stroke_alpha", 0.08))
    sparkle_a = float(prof.get("sparkle_alpha", 0.06))

    sr, sg, sb = prof.get("stroke_rgb", (0.55, 0.40, 0.50))
    stroke_col = Color(float(sr), float(sg), float(sb))

    jitter_seed = int(prof.get("jitter_seed", 42))
    rng = random.Random(jitter_seed)

    def _jx(): return rng.uniform(-w * 0.018, w * 0.018)
    def _jy(): return rng.uniform(-h * 0.018, h * 0.018)

    _draw_kawaii_tint_and_border(canvas, w, h, prof, stroke_col, is_bw, rng)
    _draw_kawaii_corner_daisy(canvas, w, h, stroke_col, stroke_a, rng, _jx, _jy)
    _draw_kawaii_sparkles(canvas, w, h, stroke_col, sparkle_a, rng, _jx, _jy)
    _draw_kawaii_element_pools(canvas, w, h, stroke_col, stroke_a, sparkle_a,
                               prof, jitter_seed, _jx, _jy)

    canvas.restoreState()


def _draw_page(canvas, doc, kawaii_pdf: bool, prof: Optional[dict] = None):
    """ReportLab page callback — draws kawaii background (if enabled) then the footer.

    Registered as both onFirstPage and onLaterPages so every page gets the same treatment.
    *prof* is computed once before doc.build() and passed via closure so all pages share
    the same jitter_seed (identical layout across pages in a single export run).
    """
    if kawaii_pdf and prof is not None:
        _draw_kawaii_background(canvas, doc, prof)
    _draw_footer(canvas, doc)


def _table_style_kawaii(printer_bw: bool, prof: Optional[dict] = None):
    """Compute four table colors (header_bg, row_a, row_b, grid) from the kawaii profile.

    When *prof* is provided the colors are blended from the profile's tint/stroke RGB and
    alpha values so table color intensity tracks the elem_intensity slider.  When *prof* is
    None, hardcoded defaults are used as a fallback.

    The internal _blend_w helper linearly interpolates a given RGB color toward white:
      alpha=0 → pure white, alpha=1 → full color.

    Multipliers (e.g. 12.0, 10.0 for color mode) are calibrated so that the default
    'Cute' preset (tint_alpha≈0.055, border_alpha≈0.10) reproduces the original hardcoded
    pink palette values exactly.  At 0% elem_intensity all table cells blend to white.

    Returns
    -------
    tuple[Color, Color, Color, Color]
        (header_bg, row_a, row_b, grid) as ReportLab Color objects.
    """

    def _bw(v):
        return max(0.0, min(1.0, v))

    # Blend an RGB color toward white by alpha amount.
    # alpha=0 → pure white, alpha=1 → full color. Used to scale table colors with intensity.
    def _blend_w(rgb, alpha):
        a = max(0.0, min(1.0, alpha))
        return Color(_bw(a * rgb[0] + (1 - a)), _bw(a * rgb[1] + (1 - a)), _bw(a * rgb[2] + (1 - a)))

    if printer_bw:
        if prof:
            ta = float(prof.get("tint_alpha", 0.007))
            ba = float(prof.get("border_alpha", 0.04))
            grey = (0.5, 0.5, 0.5)
            # Multipliers (23.0, 14.0, 12.5) are calibrated so that at the default
            # "Cute" preset (ta=0.007, ba=0.04), the blended output matches the
            # original hardcoded grey values (0.92, 0.95, 0.75). At 0% intensity
            # (ta→0), everything blends to white (invisible tint).
            header_bg = _blend_w(grey, min(ta * 23.0, 0.40))   # → ~0.92 grey
            row_a     = Color(1.0, 1.0, 1.0)                    # always white in B/W
            row_b     = _blend_w(grey, min(ta * 14.0, 0.30))   # → ~0.95 grey
            grid      = _blend_w(grey, min(ba * 12.5, 0.55))   # → ~0.75 grey
        else:
            header_bg = Color(0.92, 0.92, 0.92)
            row_a = Color(1.0, 1.0, 1.0)
            row_b = Color(0.95, 0.95, 0.95)
            grid = Color(0.75, 0.75, 0.75)
    else:
        # Color mode: blend pink/purple tint toward white based on intensity.
        # tint_rgb controls the base hue (pink↔purple via bg_hue_pct slider).
        if prof:
            ta = float(prof.get("tint_alpha", 0.055))
            ba = float(prof.get("border_alpha", 0.10))
            tint = prof.get("tint_rgb", (1.00, 0.86, 0.92))    # pink base
            stroke = prof.get("stroke_rgb", (0.55, 0.40, 0.50)) # darker accent for grid
            # Same calibration approach as B/W: multipliers tuned so default "Cute"
            # preset (ta=0.055) reproduces the original hardcoded pink palette.
            header_bg = _blend_w(tint,   min(ta * 12.0, 0.70))  # → ~(0.96, 0.90, 0.95)
            row_a     = _blend_w(tint,   min(ta * 4.5,  0.30))  # → ~(0.995, 0.965, 0.985)
            row_b     = _blend_w(tint,   min(ta * 10.0, 0.60))  # → ~(0.98, 0.92, 0.96)
            grid      = _blend_w(stroke, min(ba * 5.0,  0.65))  # → ~(0.72, 0.68, 0.74)
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

    missing = [c for c in COLUMNS_TO_USE if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns for PDF export: {missing}")

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
    printer_bw: bool,
    prof: Optional[dict] = None,
):
    """Build ReportLab flowable elements (title + table) for one page of move-up items."""
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
        header_bg, row_a, row_b, grid = _table_style_kawaii(printer_bw, prof=prof)
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
    printer_bw: bool,
    barcode_header: str = "METRC",
    prof: Optional[dict] = None,
):
    """Build ReportLab flowable elements (title + table) for one audit PDF page group.

    Parameters
    ----------
    df_chunk : pd.DataFrame
        Must have columns: Type, Product, METRC, Room / Notes, QtyOrCount.
    title : str
        Group header shown above the table, e.g. ``"Distributor: Green Thumb"``.
    mode : str
        ``"master"`` shows the Qty column filled in; ``"blank"`` shows a Count column
        with empty cells for staff to fill by hand during a floor audit.
    kawaii_pdf : bool
        When True uses profile-blended colors from *prof*.
    printer_bw : bool
        When True forces greyscale color palette.
    barcode_header : str
        Column header label for the barcode column (``"METRC"`` or ``"Barcode"``).
    prof : dict or None
        Pre-computed kawaii profile dict; None falls back to hardcoded defaults.
    """
    styles = getSampleStyleSheet()
    elements: List = []

    widths = PDF_PROFILE["audit_widths"]
    fs = int(PDF_PROFILE["font_size"])

    elements.append(Paragraph(f"<b>{title}</b>", styles["Heading2"]))
    elements.append(Paragraph(datetime.now().strftime(DATE_FORMAT), styles["Normal"]))
    elements.append(Paragraph(" ", styles["Normal"]))

    headers = (
        ["Type", "Product", barcode_header, " ", "Qty"]
        if mode == "master"
        else ["Type", "Product", barcode_header, " ", "Count"]
    )

    table_data = [headers] + df_chunk.values.tolist()
    table = Table(table_data, colWidths=widths)

    if kawaii_pdf:
        header_bg, row_a, row_b, grid = _table_style_kawaii(printer_bw, prof=prof)
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
    """Export a paginated move-up sticker PDF (portrait letter-size).

    Page layout order: priority items (prefixed with ⭐) → backstock items → all other rooms.
    Priority items in *priority_df* are deduplicated against *move_up_df* so no item appears twice.

    Parameters
    ----------
    move_up_df : pd.DataFrame
        Full filtered move-up DataFrame (must contain COLUMNS_TO_USE columns).
    priority_df : pd.DataFrame or None
        Subset of items flagged as priority; shown first with a ⭐ prefix. None means no priority section.
    base_dir : str
        Directory where the PDF file will be written.
    timestamp : bool
        When True appends a ``YYYY-MM-DD_HH-MM`` timestamp to the filename.
    prefix : str or None
        Optional store-name prefix prepended to the filename. Sanitized via sanitize_prefix().
    auto_open : bool
        When True calls _auto_open_file() after writing so the OS opens the PDF immediately.
    items_per_page : int
        Maximum rows per page; default 30. Larger values pack more items but smaller font appears.
    kawaii_pdf : bool
        When True renders kawaii background decorations behind the table on each page.
    printer_bw : bool
        When True forces greyscale decoration palette (for B/W printers).

    Returns
    -------
    str
        Absolute path to the written PDF file.
    """
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
    prof = _kawaii_profile_from_settings(printer_bw) if kawaii_pdf else None
    if prof is not None:
        prof["jitter_seed"] = random.randint(0, 999999)

    if not combined.empty:
        combined_pdf_df = _prep_moveup_table_df(combined)
        for start in range(0, len(combined_pdf_df), items_per_page):
            chunk = combined_pdf_df.iloc[start:start + items_per_page]
            elements += _build_moveup_page_elements(chunk, title, kawaii_pdf, printer_bw, prof=prof)
            if start + items_per_page < len(combined_pdf_df):
                elements.append(PageBreak())

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
    barcode_col: str | None = None,
) -> Tuple[str, str]:
    """Generate a matched pair of audit PDFs from the current inventory DataFrame.

    Two files are written to *base_dir*:
    - ``Audit_Master_<stamp>.pdf`` — Qty column filled in; reference copy for managers.
    - ``Audit_Blank_<stamp>.pdf``  — Count column blank; staff writes actual counts during floor walk.

    Both PDFs share the same sort order and grouping, so pages align 1:1 for cross-referencing.

    Parameters
    ----------
    df : pd.DataFrame
        Source inventory data. Missing display columns (Type, Brand, Room, etc.) are added as blank.
    base_dir : str
        Output directory for both PDF files.
    title_text : str
        Report title shown on the first page of each PDF (e.g. ``"Sales Floor Audit — Main Store"``).
    sort_mode : str
        Grouping/sort key. One of:
        - ``"distributor_type_size_product"`` — group by Distributor, sort by Type/Size/Product
        - ``"brand_type_product"`` — group by Brand, sort by Type/Product
        - anything else → group by Type, sort by Brand/Product
    kawaii_pdf : bool
        When True applies kawaii decoration profile to every page.
    printer_bw : bool
        When True uses greyscale decoration palette.
    auto_open : bool
        When True opens both PDFs with the system default viewer after writing.
    default_store : str
        Fallback store label when the Store column is missing or blank.
    default_room : str
        Fallback room label when the Room column is missing or blank.
    type_trunc_len : int
        Max characters for the Type column before truncation with '…'.
    barcode_col : str or None
        If provided and present in *df*, use this column for the barcode display instead of
        "Package Barcode". The column header label changes from "METRC" to "Barcode".

    Returns
    -------
    tuple[str, str]
        ``(master_path, blank_path)`` — absolute paths to the two written PDFs.

    Raises
    ------
    ValueError
        If *df* is None or empty (nothing to export).
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
    if barcode_col and barcode_col in work.columns:
        work["MetrcLast8"] = work[barcode_col].fillna("").astype(str).map(
            lambda v: _fmt_barcode_tail(v, int(PDF_PROFILE["metrc_tail_audit"]))
        )
        barcode_header = "Barcode"
    else:
        work["MetrcLast8"] = work["Package Barcode"].fillna("").astype(str).map(
            lambda v: _fmt_barcode_tail(v, int(PDF_PROFILE["metrc_tail_audit"]))
        )
        barcode_header = "METRC"

    work["Qty On Hand"] = pd.to_numeric(work["Qty On Hand"], errors="coerce").fillna(0).astype(int)

    sort_keys = [k for k in sort_keys if k in work.columns]
    if sort_keys:
        work = work.sort_values(sort_keys, kind="stable").reset_index(drop=True)

    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    master_path = os.path.join(base_dir, f"Audit_Master_{stamp}.pdf")
    blank_path = os.path.join(base_dir, f"Audit_Blank_{stamp}.pdf")
    styles = getSampleStyleSheet()

    audit_prof = _kawaii_profile_from_settings(printer_bw) if kawaii_pdf else None
    if audit_prof is not None:
        audit_prof["jitter_seed"] = random.randint(0, 999999)

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
                barcode_header=barcode_header,
                prof=audit_prof,
            )

            if gi < len(groups) - 1:
                elements.append(PageBreak())

        doc.build(
            elements,
            onFirstPage=lambda c, d: _draw_page(c, d, kawaii_pdf=kawaii_pdf, prof=audit_prof),
            onLaterPages=lambda c, d: _draw_page(c, d, kawaii_pdf=kawaii_pdf, prof=audit_prof),
        )

    _build(master_path, "master")
    _build(blank_path, "blank")

    if auto_open:
        _auto_open_file(master_path)
        _auto_open_file(blank_path)

    return master_path, blank_path