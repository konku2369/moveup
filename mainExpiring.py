"""
Expiring items analysis window.

Toplevel window that loads a METRC export independently, detects expiration
date columns via pattern matching, and displays items grouped by expiration
urgency. Supports PDF export and configurable day-range thresholds.
"""
# Sweed "Soon To Expire" Viewer (v7)
# - Subtle/tasteful kawaii UI theme (always on)
# - Toggle: Include missing expiration dates (default ON)
# - PDF export: Normal by default, optional Kawaii PDF toggle
# - Single-click bucket selection
# - Dynamic buckets based on Days window
# - Exclusions (accessories etc.)
#
# Updates:
# - Truncate "METRC-6" (derived from METRC last 6 digits) to 18 chars (display + PDF)
# - Truncate Product Name to 60 chars (display + PDF)
# - Auto-open PDF after export
# - Exclude Brand from PDF output only
# - Replace Subcategory with METRC last-6 digits (and label it "METRC-6")
# - Center Available Qty column in PDF

import os
import re
import sys
import subprocess
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

# PDF
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

from themes import MOVEUP_THEME, apply_theme
from data_core import truncate_text


APP_TITLE = "Sweed - Soon To Expire Viewer"

# Columns that must exist for header-row auto-detection in Excel/CSV files.
HEADER_DETECT_REQUIRED_COLS = ["Product Name", "Expiration Date"]

# Regex patterns for auto-detecting which column serves which role.
# All patterns are case-insensitive, substring-matching (no anchors unless explicit).
# first_matching_col() tries each pattern in order and returns the first match.
COL_PATTERNS = {
    "expiration": [r"expiration\s*date", r"\bexpir", r"\bexp\b", r"expires?", r"best\s*by", r"use\s*by"],
    "product": [r"product\s*name", r"\bproduct\b"],
    "category": [r"\bcategory\b"],
    "subcategory": [r"sub\s*category|subcategory"],
    "brand": [r"\bbrand\b"],
    "location": [r"\blocation\b"],
    "available_qty": [r"available\s*qty", r"\bavailable\b"],
    "qty": [r"^\s*qty\s*$", r"quantity", r"on\s*hand", r"count"],
    "metrc": [r"\bmetrc\b", r"\btag\b", r"rfid", r"package\s*tag", r"package\s*id", r"metrc\s*tag"],
}

# Product types excluded from the expiring report (not worth tracking expiry for)
DEFAULT_EXCLUSIONS = ["accessory", "accessories"]

# ----------------------------
# Truncation rules
# ----------------------------
TRUNCATE_SUBCATEGORY_TO = 18
TRUNCATE_PRODUCT_TO = 60


def open_file_with_default_app(path: str):
    """
    Cross-platform "open in default app".
    - Windows: os.startfile
    - macOS: open
    - Linux: xdg-open
    """
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except Exception as e:
        print(f"[expiring] Could not auto-open file: {e}")


def metrc_last6(val) -> str:
    """
    Return last 6 digits of METRC/tag value.
    If no digits exist, fall back to last 6 characters.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    digits = re.sub(r"\D+", "", s)
    if digits:
        return digits[-6:]
    return s[-6:] if len(s) >= 6 else s


def first_matching_col(columns, patterns):
    """Return the first column name matching any regex in *patterns*, or None."""
    cols = list(columns)
    for pat in patterns:
        rx = re.compile(pat, re.IGNORECASE)
        for c in cols:
            if rx.search(str(c)):
                return c
    return None


def detect_header_row_excel(path, max_rows=20):
    """Scan the first *max_rows* of an Excel file for the header row containing required columns."""
    preview = pd.read_excel(path, header=None, nrows=max_rows)

    def norm(x):
        if pd.isna(x):
            return ""
        return str(x).strip()

    for i in range(len(preview)):
        row_vals = [norm(x) for x in preview.iloc[i].tolist()]
        if all(req in row_vals for req in HEADER_DETECT_REQUIRED_COLS):
            return i
    return None


def detect_header_row_csv(path, max_rows=20):
    """Scan the first *max_rows* of a CSV file for the header row containing required columns."""
    preview = pd.read_csv(path, header=None, nrows=max_rows, dtype=str, encoding_errors="ignore")

    def norm(x):
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return ""
        return str(x).strip()

    for i in range(len(preview)):
        row_vals = [norm(x) for x in preview.iloc[i].tolist()]
        if all(req in row_vals for req in HEADER_DETECT_REQUIRED_COLS):
            return i
    return None


def parse_date_series(series: pd.Series) -> pd.Series:
    """Parse a Series to datetime, handling both string dates and Excel serial numbers."""
    s = series.copy()

    if pd.api.types.is_numeric_dtype(s):
        try:
            return pd.to_datetime(s, unit="D", origin="1899-12-30", errors="coerce")
        except Exception:
            pass

    return pd.to_datetime(s, errors="coerce")


def _clean_keyword_list(raw: str) -> list[str]:
    if not raw:
        return []
    parts = [p.strip().lower() for p in raw.split(",")]
    return [p for p in parts if p]


def _to_number(series: pd.Series) -> pd.Series:
    s = series.copy()
    if s.dtype == "O":
        s = s.astype(str).str.replace(",", "", regex=False).str.strip()
        s = s.replace({"": None, "nan": None, "None": None})
    return pd.to_numeric(s, errors="coerce")


def build_dynamic_buckets(max_days: int) -> list[tuple[str, int, int]]:
    """
    Buckets expand based on Days window.
    - 0–7, 8–14, 15–30
    - then 30-day blocks up to max_days: 31–60, 61–90, ...
    """
    max_days = int(max_days)
    if max_days < 0:
        max_days = 0

    buckets = []
    base = [("0–7 days", 0, 7), ("8–14 days", 8, 14), ("15–30 days", 15, 30)]
    for label, lo, hi in base:
        if lo > max_days:
            break
        buckets.append((label, lo, min(hi, max_days)))

    if max_days <= 30:
        return buckets

    start = 31
    while start <= max_days:
        end = min(start + 29, max_days)
        label = f"{start}–{end} days" if end != start else f"{start} days"
        buckets.append((label, start, end))
        start = end + 1

    return buckets


# ----------------------------
# PDF Export (Normal default, optional Kawaii)
# ----------------------------

def df_to_pdf(
    path: str,
    title: str,
    subtitle: str,
    df: pd.DataFrame,
    columns: list[str],
    kawaii_pdf: bool = False,
    product_col: str | None = None,
    metrc6_col: str | None = None,
):
    """Export a single DataFrame to a landscape PDF table with optional kawaii styling."""

    def _find_available_qty_col_index(cols: list[str]) -> int | None:
        for i, c in enumerate(cols):
            name = str(c).strip().lower()
            if ("available" in name and "qty" in name) or name in {"available", "qty", "quantity", "on hand", "count"}:
                return i
        return None

    doc = SimpleDocTemplate(
        path,
        pagesize=landscape(letter),
        leftMargin=24,
        rightMargin=24,
        topMargin=24,
        bottomMargin=24,
        title=title
    )
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(title, styles["Title"]))
    story.append(Paragraph(subtitle, styles["Normal"]))
    story.append(Spacer(1, 12))

    out = df[columns].copy()

    # Truncation for PDF output
    if product_col and product_col in out.columns:
        out[product_col] = out[product_col].map(lambda v: truncate_text(v, TRUNCATE_PRODUCT_TO))
    if metrc6_col and metrc6_col in out.columns:
        out[metrc6_col] = out[metrc6_col].map(lambda v: truncate_text(v, TRUNCATE_SUBCATEGORY_TO))

    for c in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[c]):
            out[c] = out[c].dt.strftime("%Y-%m-%d")
    out = out.fillna("")
    table_data = [columns] + out.values.tolist()

    total_width = 742
    base = max(60, int(total_width / max(1, len(columns))))
    col_widths = []
    for c in columns:
        name = str(c).lower()

        if "product" in name:
            col_widths.append(base * 3)

        elif "location" in name:
            col_widths.append(int(base * 1))

        # NEW: make METRC-6 skinnier
        elif "metrc" in name:
            col_widths.append(int(base * 0.7))  # tweak: 0.6–0.9

        elif "expiration" in name or "days to expire" in name or "days until" in name:
            col_widths.append(int(base * 1.0))

        else:
            col_widths.append(base)

    s = sum(col_widths)
    if s > 0:
        scale = total_width / s
        col_widths = [max(50, int(w * scale)) for w in col_widths]

    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)

    if kawaii_pdf:
        header_bg = colors.Color(0.96, 0.90, 0.95)
        row_a = colors.Color(0.995, 0.965, 0.985)
        row_b = colors.Color(0.98, 0.92, 0.96)
        grid = colors.Color(0.72, 0.68, 0.74)
    else:
        header_bg = colors.lightgrey
        row_a = colors.whitesmoke
        row_b = colors.white
        grid = colors.grey

    base_style = [
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("BACKGROUND", (0, 0), (-1, 0), header_bg),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),

        ("FONTSIZE", (0, 1), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.25, grid),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [row_a, row_b]),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]

    # NEW: center Available Qty column in PDF (header + body)
    qty_idx = _find_available_qty_col_index(columns)
    if qty_idx is not None:
        base_style.append(("ALIGN", (qty_idx, 0), (qty_idx, -1), "CENTER"))

    tbl.setStyle(TableStyle(base_style))

    story.append(tbl)
    doc.build(story)


# ----------------------------
# Data Model
# ----------------------------

class DataModel:
    """
    Data model for expiring items: loads a METRC file, detects columns, filters by days-to-expire.

    Standalone from main.py's data pipeline because the Expiring window runs as an
    independent Toplevel with its own file import. Does not share state with MoveUpGUI.
    """

    def __init__(self):
        """
        Initialise a blank DataModel with no data loaded.

        All ``col_*`` attributes start as ``None``; they are set by
        ``load_file()`` to the detected column names from the loaded file.
        ``df_raw`` holds the original loaded DataFrame; ``df_filtered`` holds
        the last result of ``filter_expiring()``.
        """
        self.df_raw: pd.DataFrame | None = None
        self.df_filtered: pd.DataFrame | None = None

        self.col_exp = None
        self.col_product = None
        self.col_available = None
        self.col_location = None
        self.col_category = None
        self.col_subcategory = None   # original detected subcategory (if any)
        self.col_brand = None

        # NEW
        self.col_metrc = None
        self.col_metrc6 = None  # the derived display/sort/export column ("METRC-6")

    def load_file(self, path: str):
        """
        Load an inventory file and detect all relevant columns.

        Supports ``.xlsx`` and ``.csv``.  Uses ``detect_header_row_excel()`` /
        ``detect_header_row_csv()`` to skip any preceding summary rows before
        the real column header.  After reading, performs column detection with
        ``first_matching_col()`` using the patterns in ``COL_PATTERNS``:

        - ``col_exp``: expiration date column (required)
        - ``col_product``: product name (required)
        - ``col_available``: available / qty column
        - ``col_location``: location/room column
        - ``col_category``: product type/category
        - ``col_subcategory``: sub-category (if present)
        - ``col_brand``: brand
        - ``col_metrc``: full METRC tag column
        - ``col_metrc6``: derived ``"METRC-6"`` column (last 6 chars of METRC tag)

        Expiration dates are parsed and coerced to ``datetime64`` via
        ``parse_date_series()``.  Raises ``ValueError`` if either the
        expiration or product column cannot be found.
        """
        ext = os.path.splitext(path)[1].lower()
        if ext == ".xlsx":
            header_row = detect_header_row_excel(path)
            if header_row is None:
                raise ValueError("Could not detect the header row in this XLSX export.")
            df = pd.read_excel(path, header=header_row)
        elif ext == ".csv":
            header_row = detect_header_row_csv(path)
            if header_row is None:
                raise ValueError("Could not detect the header row in this CSV export.")
            df = pd.read_csv(path, header=header_row)
        else:
            raise ValueError("Unsupported file type. Use .xlsx or .csv")

        df = df.dropna(axis=1, how="all")

        self.col_exp = first_matching_col(df.columns, COL_PATTERNS["expiration"])
        self.col_product = first_matching_col(df.columns, COL_PATTERNS["product"])
        self.col_available = first_matching_col(df.columns, COL_PATTERNS["available_qty"]) \
            or first_matching_col(df.columns, COL_PATTERNS["qty"])
        self.col_location = first_matching_col(df.columns, COL_PATTERNS["location"])
        self.col_category = first_matching_col(df.columns, COL_PATTERNS["category"])
        self.col_subcategory = first_matching_col(df.columns, COL_PATTERNS["subcategory"])
        self.col_brand = first_matching_col(df.columns, COL_PATTERNS["brand"])

        # NEW: find METRC/tag column
        self.col_metrc = first_matching_col(df.columns, COL_PATTERNS["metrc"])

        if not self.col_exp or not self.col_product:
            raise ValueError(
                f"Missing required columns. Found expiration={self.col_exp}, product={self.col_product}"
            )

        # Last 6 digits only: full METRC tags are 24+ chars but the last 6 are what staff use to identify a package on the floor.
        df["METRC-6"] = df[self.col_metrc].map(metrc_last6) if (self.col_metrc and self.col_metrc in df.columns) else ""
        self.col_metrc6 = "METRC-6"

        df[self.col_exp] = parse_date_series(df[self.col_exp])

        self.df_raw = df
        self.df_filtered = None

    def filter_expiring(
        self,
        days: int,
        include_expired: bool,
        include_missing: bool,
        search: str,
        exclude_keywords: list[str],
        location_filter: list[str] | None = None,
    ):
        """
        Filter the loaded DataFrame to expiring items and store in ``df_filtered``.

        Filtering pipeline:
        1. Compute a cutoff date = today + *days*.
        2. Select rows where expiration date is before the cutoff (``mask_soon``).
        3. If *include_expired* is ``False``, further restrict to dates ≥ today.
        4. Optionally include rows with missing/null expiration (``mask_missing``).
        5. Apply keyword exclusions: any row where Category, METRC-6, or Product
           Name contains any exclusion keyword (case-insensitive) is dropped.
        6. Apply *location_filter* if provided (exact match against room column).
        7. Compute ``"Days To Expire"`` column for sorting.
        8. Apply free-text *search* across all string columns.
        9. Sort by ``Days To Expire`` ascending (most urgent first).

        Sets ``self.df_filtered`` to the result and also returns it.

        Parameters
        ----------
        days : int
            Days-ahead window; items expiring within this range are included.
        include_expired : bool
            If ``True``, also include items that have already expired.
        include_missing : bool
            If ``True``, also include items with no expiration date.
        search : str
            Free-text search applied across all string columns.
        exclude_keywords : list[str]
            Product type keywords to exclude (e.g. ``["accessory"]``).
        location_filter : list[str] | None
            Restrict to specific locations.  ``None`` = no location filter.

        Returns
        -------
        pd.DataFrame | None
            Filtered DataFrame, or ``None`` if no data is loaded.
        """
        if self.df_raw is None:
            return None

        df = self.df_raw.copy()
        today = pd.Timestamp(datetime.now().date())
        cutoff = today + pd.Timedelta(days=int(days))

        exp = df[self.col_exp]
        mask_soon = exp.notna() & (exp <= cutoff)
        if not include_expired:
            mask_soon = mask_soon & (exp >= today)

        mask_missing = exp.isna()
        mask = (mask_soon | mask_missing) if include_missing else mask_soon
        df = df.loc[mask].copy()

        ex = [k.strip().lower() for k in exclude_keywords if k and k.strip()]
        if ex:
            cols_for_ex = [c for c in [self.col_category, self.col_metrc6, self.col_product] if c and c in df.columns]
            if cols_for_ex:
                blob = df[cols_for_ex].fillna("").astype(str).agg(" | ".join, axis=1).str.lower()
                ex_mask = pd.Series(False, index=df.index)
                for k in ex:
                    ex_mask = ex_mask | blob.str.contains(re.escape(k), na=False)
                df = df.loc[~ex_mask].copy()

        if location_filter and self.col_location and self.col_location in df.columns:
            df = df[df[self.col_location].isin(location_filter)].copy()

        df["Days To Expire"] = (df[self.col_exp].dt.normalize() - today).dt.days

        q = (search or "").strip().lower()
        if q:
            text_cols = [c for c in df.columns if df[c].dtype == "object"]
            if text_cols:
                big = df[text_cols].fillna("").astype(str).agg(" | ".join, axis=1).str.lower()
                df = df.loc[big.str.contains(q, na=False)]

        df["_sortkey"] = df["Days To Expire"].fillna(10**9)
        df = df.sort_values(by=["_sortkey"], ascending=True).drop(columns=["_sortkey"])

        self.df_filtered = df
        return df

    def build_action_list(self) -> pd.DataFrame | None:
        """
        Build the structured, export-ready action list from ``df_filtered``.

        Selects and standardises columns: Location, Category, METRC-6, Brand,
        Product Name, Expiration Date, Days To Expire, Available Qty.  Missing
        columns are created as empty strings rather than causing KeyErrors.
        Sorted by: Location → Category → METRC-6 → Expiration Date → Product Name.

        Returns
        -------
        pd.DataFrame | None
            Sorted action list, or ``None`` if ``df_filtered`` is empty.
        """
        if self.df_filtered is None or self.df_filtered.empty:
            return None

        df = self.df_filtered.copy()

        def ensure_col(name, col):
            if col and col in df.columns:
                return col
            df[name] = ""
            return name

        c_location = ensure_col("Location", self.col_location)
        c_category = ensure_col("Category", self.col_category)
        c_subcat = ensure_col("METRC-6", self.col_metrc6)
        c_brand = ensure_col("Brand", self.col_brand)
        c_product = self.col_product
        c_exp = self.col_exp

        if self.col_available and self.col_available in df.columns:
            c_avail = self.col_available
        else:
            df["Available Qty"] = ""
            c_avail = "Available Qty"

        keep = [c_location, c_category, c_subcat, c_brand, c_product, c_exp, "Days To Expire", c_avail]
        keep = [c for c in keep if c in df.columns]
        out = df[keep].copy()

        out[c_exp] = pd.to_datetime(out[c_exp], errors="coerce")
        out["_exp_sort"] = out[c_exp].fillna(pd.Timestamp.max)

        out = out.sort_values(
            by=[c_location, c_category, c_subcat, "_exp_sort", c_product],
            ascending=[True, True, True, True, True],
        ).drop(columns=["_exp_sort"])

        return out

    def build_bucket_summary(
        self,
        max_days: int,
        include_expired: bool,
        include_missing: bool,
        buckets: list[tuple[str, int, int]],
    ):
        """
        Group ``df_filtered`` into time buckets and return a summary table + per-bucket DataFrames.

        Parameters
        ----------
        max_days : int
            Upper bound on ``Days To Expire`` for inclusion in date buckets.
        include_expired : bool
            If ``True``, a special ``"Expired"`` bucket is prepended for items
            with ``Days To Expire < 0``.
        include_missing : bool
            If ``True``, a ``"Missing Expiration"`` bucket is appended for
            items with no expiration date.
        buckets : list[tuple[str, int, int]]
            Each element is ``(label, lo_days, hi_days)`` defining the
            day-range for one bucket row (inclusive bounds).

        Returns
        -------
        tuple[pd.DataFrame, dict[str, pd.DataFrame]]
            ``(summary_df, bucket_map)`` where *summary_df* has columns
            ``["Bucket", "Items", "Available Qty (sum)", "Earliest Exp",
            "Latest Exp"]`` and *bucket_map* maps each bucket label to the
            corresponding filtered DataFrame.  Returns ``(empty DataFrame, {})``
            if no filtered data exists.
        """
        if self.df_filtered is None:
            return pd.DataFrame(), {}

        df = self.df_filtered.copy()

        avail_col = self.col_available if (self.col_available and self.col_available in df.columns) else None
        df["_avail_num"] = _to_number(df[avail_col]).fillna(0) if avail_col else 0

        missing = df[df[self.col_exp].isna()].copy()
        dated = df[df[self.col_exp].notna()].copy()

        bucket_map = {}
        rows = []

        if include_expired:
            exp_df = dated[dated["Days To Expire"] < 0].copy()
            if not exp_df.empty:
                bucket_map["Expired"] = exp_df.drop(columns=["_avail_num"])
                rows.append({
                    "Bucket": "Expired",
                    "Items": len(exp_df),
                    "Available Qty (sum)": float(exp_df["_avail_num"].sum()),
                    "Earliest Exp": exp_df[self.col_exp].min(),
                    "Latest Exp": exp_df[self.col_exp].max(),
                })
            dated = dated[dated["Days To Expire"] >= 0].copy()

        dated = dated[dated["Days To Expire"] <= int(max_days)].copy()

        for label, lo, hi in buckets:
            bdf = dated[(dated["Days To Expire"] >= lo) & (dated["Days To Expire"] <= hi)].copy()
            bucket_map[label] = bdf.drop(columns=["_avail_num"])
            rows.append({
                "Bucket": label,
                "Items": len(bdf),
                "Available Qty (sum)": float(bdf["_avail_num"].sum()),
                "Earliest Exp": bdf[self.col_exp].min() if not bdf.empty else pd.NaT,
                "Latest Exp": bdf[self.col_exp].max() if not bdf.empty else pd.NaT,
            })

        if include_missing:
            bucket_map["Missing Expiration"] = missing.drop(columns=["_avail_num"])
            rows.append({
                "Bucket": "Missing Expiration",
                "Items": len(missing),
                "Available Qty (sum)": float(missing["_avail_num"].sum()),
                "Earliest Exp": pd.NaT,
                "Latest Exp": pd.NaT,
            })

        summary_df = pd.DataFrame(rows)
        return summary_df, bucket_map


# ----------------------------
# UI Widgets
# ----------------------------

class TableView(ttk.Frame):
    """
    Scrollable treeview widget with sortable columns and auto-width sizing.

    Wraps a ``ttk.Treeview`` with vertical and horizontal scrollbars.  Heading
    clicks toggle ascending/descending sort (numeric-aware: tries
    ``pd.to_numeric`` first, falls back to string sort).  Column widths are
    auto-sized from the column name length.  Truncation via *trunc_map* is
    applied at render time so the underlying data is not mutated.
    """

    def __init__(self, parent):
        super().__init__(parent)
        self.tree = ttk.Treeview(self, columns=(), show="headings")
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        # Sort state
        self._raw_df: pd.DataFrame | None = None
        self._columns: list[str] = []
        self._trunc_map: dict[str, int] = {}
        self._sort_col: str | None = None
        self._sort_asc: bool = True

        # Define tags for highlighting
        self.tree.tag_configure("critical", background="#FFE5E5", foreground="#8B0000")  # Red - expired or <=7 days
        self.tree.tag_configure("warning", background="#FFF4E0", foreground="#8B4500")  # Orange - 8-14 days
        self.tree.tag_configure("normal", background=MOVEUP_THEME["tree_bg"])

    def render(self, df: pd.DataFrame, columns: list[str], trunc_map: dict[str, int] | None = None):
        # Store for re-use by _sort_by_col
        self._raw_df = df
        self._columns = list(columns)
        self._trunc_map = trunc_map

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = columns

        for c in columns:
            indicator = (" ▲" if self._sort_asc else " ▼") if c == self._sort_col else ""
            self.tree.heading(c, text=c + indicator, command=lambda col=c: self._sort_by_col(col))

            col_name = str(c).lower()

            # Center the Available Qty column (and similar variants)
            if ("available" in col_name and "qty" in col_name) or col_name in {"qty", "quantity", "count", "on hand"}:
                self.tree.column(c, width=max(120, min(520, len(c) * 12)), anchor="center")
            else:
                self.tree.column(c, width=max(120, min(520, len(c) * 12)), anchor="w")

        if df is None or df.empty or not columns:
            return

        out = df[columns].copy()

        trunc_map = trunc_map or {}
        for col, lim in trunc_map.items():
            if col in out.columns:
                out[col] = out[col].map(lambda v: truncate_text(v, lim))

        if "Days To Expire" in out.columns:
            out["Days To Expire"] = out["Days To Expire"].apply(
                lambda v: str(int(v)) if pd.notna(v) else ""
            )

        if "Available Qty (sum)" in out.columns:
            out["Available Qty (sum)"] = out["Available Qty (sum)"].apply(
                lambda v: str(int(v)) if pd.notna(v) else ""
            )

        for c in out.columns:
            if pd.api.types.is_datetime64_any_dtype(out[c]):
                out[c] = out[c].dt.strftime("%Y-%m-%d")
            else:
                out[c] = out[c].apply(
                    lambda v: v.strftime("%Y-%m-%d") if isinstance(v, pd.Timestamp) else v
                )
        out = out.fillna("")

        # Find the "Days To Expire" column index if it exists
        days_col_idx = None
        if "Days To Expire" in columns:
            days_col_idx = columns.index("Days To Expire")

        for idx, row in out.iterrows():
            values = [row[c] for c in columns]

            # Determine tag based on days to expire
            tag = "normal"
            if days_col_idx is not None:
                try:
                    days_val = df.loc[idx, "Days To Expire"]  # Get from original df (not string version)
                    if pd.notna(days_val):
                        days = float(days_val)
                        if days < 0:
                            tag = "critical"  # Already expired
                        elif days <= 7:
                            tag = "critical"  # Expires within a week
                        elif days <= 14:
                            tag = "warning"  # Expires within 2 weeks
                except (ValueError, KeyError):
                    pass

            self.tree.insert("", "end", values=values, tags=(tag,))

    def _sort_by_col(self, col: str):
        if self._raw_df is None or col not in self._raw_df.columns:
            return
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True

        df = self._raw_df.copy()
        series = df[col]

        # Prefer numeric sort (catches Days To Expire, qty, etc.); fall back to case-insensitive string
        numeric = pd.to_numeric(series, errors="coerce")
        if numeric.notna().any():
            df["_sk"] = numeric
        else:
            df["_sk"] = series.astype(str).str.lower()

        df = df.sort_values("_sk", ascending=self._sort_asc, na_position="last").drop(columns=["_sk"])
        self.render(df, self._columns, self._trunc_map)

# ----------------------------
# Main App
# ----------------------------

class App(tk.Toplevel):
    """
    Expiring Items viewer — Toplevel window for expiration date analysis.

    Loads an inventory file independently (not shared with the main app window)
    via ``DataModel.load_file()``.  Displays items grouped into urgency buckets
    (e.g. "0-7 days", "8-14 days", "15-30 days") selected by clicking the
    bucket summary table.

    Features:
    - Dynamic day-range threshold (spinbox) with live recalculation
    - Include/exclude expired items and items with missing expiration dates
    - Location (room) filter
    - Free-text search with debounce
    - Keyword exclusion list (e.g. ``"accessory"``)
    - CSV, Excel, and PDF export (normal or kawaii palette)
    - Auto-opens PDF on Windows after export
    """

    def __init__(self, master):
        super().__init__(master)
        self.title(APP_TITLE)
        self.geometry("1340x820")

        self.model = DataModel()

        self._debounce_job: str | None = None
        self._last_dir: str | None = None

        self.days_var = tk.IntVar(value=30)
        self.days_var.trace_add("write", lambda *_: self._debounce(300, self.apply_filter))
        self.include_expired_var = tk.BooleanVar(value=False)
        self.include_missing_var = tk.BooleanVar(value=True)

        self.exclude_accessory_var = tk.BooleanVar(value=True)
        self.exclusions_var = tk.StringVar(value="Accessory, Accessories")

        self.pdf_kawaii_var = tk.BooleanVar(value=True)

        self.search_var = tk.StringVar(value="")
        self.info_var = tk.StringVar(value="Load a file to begin.")

        self.exp_cols = []
        self.action_df = None
        self.action_cols = []

        self.bucket_summary_df = None
        self.bucket_map = {}
        self.bucket_detail_df = None
        self.bucket_detail_cols = []

        self._build_ui()
        self.apply_subtle_theme()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        """Cancel pending debounce timers before destroying the window."""
        if self._debounce_job is not None:
            try:
                self.after_cancel(self._debounce_job)
            except (tk.TclError, ValueError):
                pass
            self._debounce_job = None
        self.destroy()

    def _set_export_buttons_state(self, state: str):
        for btn in getattr(self, "_export_buttons", []):
            btn.configure(state=state)

    def _build_trunc_map(self) -> dict[str, int]:
        m = {}
        if self.model.col_product:
            m[self.model.col_product] = TRUNCATE_PRODUCT_TO
        if self.model.col_metrc6 and self.model.col_metrc6 in (self.model.df_raw.columns if self.model.df_raw is not None else []):
            m[self.model.col_metrc6] = TRUNCATE_SUBCATEGORY_TO
        return m

    def apply_subtle_theme(self):
        apply_theme(self, "moveup_match")

    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Button(top, text="Open Sweed File (XLSX/CSV)", command=self.open_file).pack(side="left")
        self.bind("<Control-o>", lambda e: self.open_file())

        ttk.Label(top, text="Days (expiring within):").pack(side="left", padx=(12, 6))
        ttk.Spinbox(top, from_=1, to=365, textvariable=self.days_var, width=6).pack(side="left")

        ttk.Checkbutton(top, text="Include already expired", variable=self.include_expired_var, command=self.apply_filter)\
            .pack(side="left", padx=(12, 0))

        ttk.Checkbutton(top, text="Include missing expiration", variable=self.include_missing_var, command=self.apply_filter)\
            .pack(side="left", padx=(12, 0))

        ttk.Checkbutton(top, text="Kawaii PDF", variable=self.pdf_kawaii_var)\
            .pack(side="left", padx=(12, 0))

        _btn_pdf = ttk.Button(top, text="Export PDF (Current Tab)", command=self.export_pdf, state="disabled")
        _btn_pdf.pack(side="right")
        _btn_xlsx = ttk.Button(top, text="Export Excel (Current Tab)", command=self.export_xlsx, state="disabled")
        _btn_xlsx.pack(side="right", padx=(0, 8))
        _btn_csv = ttk.Button(top, text="Export CSV (Current Tab)", command=self.export_csv, state="disabled")
        _btn_csv.pack(side="right", padx=(0, 8))
        self._export_buttons = [_btn_pdf, _btn_xlsx, _btn_csv]

        ttk.Label(self, textvariable=self.info_var).pack(fill="x", padx=10, pady=(0, 6))

        exrow = ttk.Frame(self, padding=(10, 0, 10, 6))
        exrow.pack(fill="x")

        ttk.Checkbutton(
            exrow,
            text="Exclude accessory items",
            variable=self.exclude_accessory_var,
            command=self.apply_filter
        ).pack(side="left")

        ttk.Label(exrow, text="Extra exclusions (comma-separated):").pack(side="left", padx=(12, 6))
        ttk.Entry(exrow, textvariable=self.exclusions_var).pack(side="left", fill="x", expand=True)
        self.exclusions_var.trace_add("write", lambda *_: self._debounce(300, self.apply_filter))

        # Location filter row — hidden until a file with a location column is loaded
        self.location_frame = ttk.Frame(self, padding=(10, 0, 10, 6))
        ttk.Label(self.location_frame, text="Room filter:").pack(side="left")
        lb_wrap = ttk.Frame(self.location_frame)
        lb_wrap.pack(side="left", padx=(6, 0), fill="x", expand=True)
        self.location_listbox = tk.Listbox(
            lb_wrap,
            selectmode=tk.EXTENDED,
            height=4,
            exportselection=False,
            activestyle="none",
            background=MOVEUP_THEME["tree_bg"],
            foreground=MOVEUP_THEME["label_fg"],
            selectbackground=MOVEUP_THEME["tree_sel"],
            selectforeground=MOVEUP_THEME["btn_fg"],
            highlightthickness=1,
            highlightcolor=MOVEUP_THEME["btn_border"],
            highlightbackground=MOVEUP_THEME["btn_border"],
        )
        lb_vsb = ttk.Scrollbar(lb_wrap, orient="vertical", command=self.location_listbox.yview)
        self.location_listbox.configure(yscrollcommand=lb_vsb.set)
        self.location_listbox.pack(side="left", fill="both", expand=True)
        lb_vsb.pack(side="left", fill="y")
        self.location_listbox.bind("<<ListboxSelect>>", lambda e: self.apply_filter())
        ttk.Button(
            self.location_frame, text="Select All", command=self._select_all_locations
        ).pack(side="left", padx=(8, 0))

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.nb = nb

        tab1 = ttk.Frame(nb)
        nb.add(tab1, text="Expiration View")

        search_row = ttk.Frame(tab1, padding=(0, 8, 0, 8))
        search_row.pack(fill="x")
        ttk.Label(search_row, text="Search:").pack(side="left")
        self.search_var.trace_add("write", lambda *_: self._debounce(300, self.apply_filter))
        ttk.Entry(search_row, textvariable=self.search_var).pack(side="left", fill="x", expand=True, padx=(6, 0))

        self.exp_table = TableView(tab1)
        self.exp_table.pack(fill="both", expand=True)

        tab2 = ttk.Frame(nb)
        nb.add(tab2, text="Action List")
        ttk.Label(tab2, text="Pull list sorted by Location → Category → METRC-6 → Expiration → Product.")\
            .pack(anchor="w", padx=2, pady=(8, 8))
        self.action_table = TableView(tab2)
        self.action_table.pack(fill="both", expand=True)

        tab3 = ttk.Frame(nb)
        nb.add(tab3, text="Bucket View")

        ttk.Label(
            tab3,
            text="Single-click a bucket to view its items below. Buckets expand automatically as you increase Days."
        ).pack(anchor="w", padx=2, pady=(8, 8))

        bucket_split = ttk.Panedwindow(tab3, orient="vertical")
        bucket_split.pack(fill="both", expand=True)

        top_frame = ttk.Frame(bucket_split)
        bottom_frame = ttk.Frame(bucket_split)
        bucket_split.add(top_frame, weight=1)
        bucket_split.add(bottom_frame, weight=2)

        self.bucket_table = TableView(top_frame)
        self.bucket_table.pack(fill="both", expand=True)

        ttk.Label(bottom_frame, text="Bucket Details").pack(anchor="w", padx=2, pady=(6, 6))
        self.bucket_detail_table = TableView(bottom_frame)
        self.bucket_detail_table.pack(fill="both", expand=True)

        self.bucket_table.tree.bind("<<TreeviewSelect>>", self.on_bucket_select)

    def open_file(self, path: str | None = None):
        if path is None:
            path = filedialog.askopenfilename(
                title="Select Sweed Inventory Export",
                filetypes=[
                    ("All Supported", "*.xlsx *.xls *.xlsm *.xlsb *.ods *.csv *.tsv *.txt *.tab"),
                    ("Excel", "*.xlsx *.xls *.xlsm *.xlsb"),
                    ("CSV / Text", "*.csv *.tsv *.txt *.tab"),
                    ("OpenDocument", "*.ods"),
                    ("All Files", "*.*"),
                ],
                initialdir=self._last_dir,
            )
        if not path:
            return
        self._last_dir = os.path.dirname(path)
        try:
            self.model.load_file(path)
            self.title(f"{APP_TITLE} — {os.path.basename(path)}")
            self._populate_location_filter()
            self.info_var.set(
                f"Loaded {len(self.model.df_raw):,} rows | "
                f"Expiration: '{self.model.col_exp}' | Product: '{self.model.col_product}' | "
                f"METRC: '{self.model.col_metrc or 'N/A'}' | "
                f"Available: '{self.model.col_available or 'N/A'}'"
            )
            self.apply_filter()
            self._set_export_buttons_state("normal")
        except Exception as e:
            messagebox.showerror("Load Error", str(e))

    def _populate_location_filter(self):
        """Fill the location listbox from the loaded file, then show or hide the row."""
        self.location_listbox.delete(0, tk.END)
        col = self.model.col_location
        if col and self.model.df_raw is not None and col in self.model.df_raw.columns:
            locations = sorted(
                self.model.df_raw[col].dropna().astype(str).unique().tolist()
            )
            for loc in locations:
                self.location_listbox.insert(tk.END, loc)
            self._select_all_locations(trigger_filter=False)
            self.location_frame.pack(fill="x", before=self.nb)
        else:
            self.location_frame.pack_forget()

    def _select_all_locations(self, trigger_filter=True):
        self.location_listbox.selection_set(0, tk.END)
        if trigger_filter:
            self.apply_filter()

    def _get_selected_locations(self) -> list[str] | None:
        """Return selected room names, or None when all (or none) are selected — meaning no filter."""
        total = self.location_listbox.size()
        if total == 0:
            return None
        selected_indices = self.location_listbox.curselection()
        if not selected_indices or len(selected_indices) == total:
            return None  # All selected — no filtering needed
        return [self.location_listbox.get(i) for i in selected_indices]

    def _debounce(self, ms: int, func):
        if self._debounce_job is not None:
            self.after_cancel(self._debounce_job)
        def _run():
            self._debounce_job = None
            func()
        self._debounce_job = self.after(ms, _run)

    def _current_exclusions(self) -> list[str]:
        extras = _clean_keyword_list(self.exclusions_var.get())
        if self.exclude_accessory_var.get():
            return list(dict.fromkeys(DEFAULT_EXCLUSIONS + extras))
        return extras

    def apply_filter(self):
        try:
            days = self.days_var.get()
        except tk.TclError:
            return  # User is mid-typing; wait for a valid integer

        df = self.model.filter_expiring(
            days=days,
            include_expired=self.include_expired_var.get(),
            include_missing=self.include_missing_var.get(),
            search=self.search_var.get(),
            exclude_keywords=self._current_exclusions(),
            location_filter=self._get_selected_locations(),
        )
        if df is None:
            return

        trunc_map = self._build_trunc_map()

        preferred = []
        for key in ["location", "category"]:
            c = getattr(self.model, f"col_{key}", None)
            if c and c in df.columns:
                preferred.append(c)

        if self.model.col_metrc6 and self.model.col_metrc6 in df.columns:
            preferred.append(self.model.col_metrc6)

        if self.model.col_brand and self.model.col_brand in df.columns:
            preferred.append(self.model.col_brand)

        if self.model.col_product and self.model.col_product in df.columns:
            preferred.append(self.model.col_product)

        if self.model.col_available and self.model.col_available in df.columns:
            preferred.append(self.model.col_available)

        if self.model.col_exp and self.model.col_exp in df.columns:
            preferred.append(self.model.col_exp)
        if "Days To Expire" in df.columns:
            preferred.append("Days To Expire")

        cols, seen = [], set()
        for c in preferred:
            if c and c in df.columns and c not in seen:
                cols.append(c)
                seen.add(c)
        if not cols:
            cols = list(df.columns[:12])

        self.exp_cols = cols
        self.exp_table.render(df, cols, trunc_map=trunc_map)

        action_df = self.model.build_action_list()
        if action_df is not None and not action_df.empty:
            self.action_df = action_df
            self.action_cols = list(action_df.columns)
            self.action_table.render(action_df, self.action_cols, trunc_map=trunc_map)
        else:
            self.action_df = None
            self.action_cols = []
            self.action_table.render(pd.DataFrame(columns=self.action_cols or []), self.action_cols or [])

        buckets = build_dynamic_buckets(days)
        summary_df, bucket_map = self.model.build_bucket_summary(
            max_days=days,
            include_expired=self.include_expired_var.get(),
            include_missing=self.include_missing_var.get(),
            buckets=buckets
        )
        self.bucket_summary_df = summary_df
        self.bucket_map = bucket_map

        bucket_cols = ["Bucket", "Items", "Available Qty (sum)", "Earliest Exp", "Latest Exp"]
        self.bucket_table.render(summary_df, bucket_cols)

        self.bucket_detail_df = None
        self.bucket_detail_cols = []
        self.bucket_detail_table.render(pd.DataFrame(columns=[]), [])

        with_date = df[self.model.col_exp].notna().sum()
        missing = df[self.model.col_exp].isna().sum()
        self.info_var.set(
            f"Showing {len(df):,} rows | With date: {with_date:,} | Missing expiration: {missing:,} "
            f"| Window: {days} days | Exclusions: {', '.join(self._current_exclusions()) or 'None'}"
        )

    def on_bucket_select(self, event):
        sel = self.bucket_table.tree.selection()
        if not sel:
            return
        item_id = sel[0]
        vals = self.bucket_table.tree.item(item_id, "values")
        if not vals:
            return
        bucket_label = vals[0]
        if bucket_label not in self.bucket_map:
            return

        df = self.bucket_map[bucket_label].copy()
        self.bucket_detail_df = df

        if df.empty:
            self.bucket_detail_cols = []
            self.bucket_detail_table.render(pd.DataFrame(columns=[]), [])
            return

        cols = []
        for key in ["col_location", "col_category"]:
            c = getattr(self.model, key, None)
            if c and c in df.columns:
                cols.append(c)

        if self.model.col_metrc6 and self.model.col_metrc6 in df.columns:
            cols.append(self.model.col_metrc6)

        if self.model.col_brand and self.model.col_brand in df.columns:
            cols.append(self.model.col_brand)

        if self.model.col_product and self.model.col_product in df.columns:
            cols.append(self.model.col_product)

        if self.model.col_available and self.model.col_available in df.columns:
            cols.append(self.model.col_available)

        if self.model.col_exp and self.model.col_exp in df.columns:
            cols.append(self.model.col_exp)
        if "Days To Expire" in df.columns:
            cols.append("Days To Expire")

        self.bucket_detail_cols = cols
        self.bucket_detail_table.render(df, cols, trunc_map=self._build_trunc_map())

    def _current_export_df_and_cols(self):
        if self.model.df_filtered is None:
            return None, None, None

        current_tab = self.nb.tab(self.nb.select(), "text")

        if current_tab == "Action List":
            df = self.action_df if self.action_df is not None else self.model.build_action_list()
            cols = self.action_cols if self.action_cols else (list(df.columns) if df is not None else [])
            title = "Sweed Action List"
            return df, cols, title

        if current_tab == "Bucket View":
            if self.bucket_detail_df is not None and not self.bucket_detail_df.empty and self.bucket_detail_cols:
                return self.bucket_detail_df, self.bucket_detail_cols, "Sweed Bucket Details"
            cols = ["Bucket", "Items", "Available Qty (sum)", "Earliest Exp", "Latest Exp"]
            return self.bucket_summary_df, cols, "Sweed Bucket Summary"

        df = self.model.df_filtered
        cols = self.exp_cols if self.exp_cols else list(df.columns[:12])
        return df, cols, "Sweed Expiration View"

    def export_csv(self):
        df, cols, _ = self._current_export_df_and_cols()
        if df is None or df.empty:
            messagebox.showinfo("Export", "Nothing to export.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        export_cols = [c for c in cols if c in df.columns] if cols else list(df.columns)
        try:
            df[export_cols].to_csv(path, index=False)
            messagebox.showinfo("Export", f"Saved:\n{path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not save file:\n{e}")

    def export_xlsx(self):
        df, cols, _ = self._current_export_df_and_cols()
        if df is None or df.empty:
            messagebox.showinfo("Export", "Nothing to export.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        export_cols = [c for c in cols if c in df.columns] if cols else list(df.columns)
        try:
            df[export_cols].to_excel(path, index=False)
            messagebox.showinfo("Export", f"Saved:\n{path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not save file:\n{e}")

    def export_pdf(self):
        df, cols, title = self._current_export_df_and_cols()
        if df is None or df.empty:
            messagebox.showinfo("Export", "Nothing to export.")
            return

        default_name = f"{title.replace(' ', '_')}_{datetime.now().strftime('%Y-%m-%d_%H%M')}.pdf"
        path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=default_name, filetypes=[("PDF", "*.pdf")])
        if not path:
            return

        subtitle = (
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')} | "
            f"Window: {int(self.days_var.get())} days | "
            f"Include expired: {self.include_expired_var.get()} | "
            f"Include missing: {self.include_missing_var.get()} | "
            f"Kawaii PDF: {self.pdf_kawaii_var.get()} | "
            f"Exclusions: {', '.join(self._current_exclusions()) or 'None'}"
        )

        try:
            pdf_cols = cols.copy()
            if self.model.col_brand and self.model.col_brand in pdf_cols:
                pdf_cols = [c for c in pdf_cols if c != self.model.col_brand]

            df_to_pdf(
                path,
                title=title,
                subtitle=subtitle,
                df=df,
                columns=pdf_cols,
                kawaii_pdf=bool(self.pdf_kawaii_var.get()),
                product_col=self.model.col_product,
                metrc6_col=self.model.col_metrc6,
            )
            messagebox.showinfo("Export", f"Saved:\n{path}")
            open_file_with_default_app(path)
        except Exception as e:
            messagebox.showerror("PDF Export Error", str(e))


def open_expiring_window(parent, file_path: str | None = None) -> App:
    """Open the Expiring Soon viewer as a child Toplevel of parent.
    If file_path is supplied the file is loaded automatically."""
    win = App(parent)
    if file_path:
        win.after(100, lambda: win.open_file(file_path))
    return win


if __name__ == "__main__":
    _root = tk.Tk()
    _root.withdraw()
    open_expiring_window(_root)
    _root.mainloop()
