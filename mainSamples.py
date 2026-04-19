"""
Sample inventory manager window.

Toplevel window that loads a METRC export independently, identifies sample
items, and provides tabbed views (by brand, by type, summary) with PDF
export. Uses a standalone DataModel for file loading and column detection.
"""
# Sample Manager — identifies "sample" items (Wholesale Cost <= $0.01)
# and provides filtering, summarizing, and exporting.
#
# Architecture mirrors mainExpiring.py:
#   - Subtle/tasteful kawaii UI theme (always on)
#   - PDF export: Normal by default, optional Kawaii PDF toggle
#   - Notebook with 4 tabs (Inventory, Summary Dashboard, Distribution List, Distributor Report)
#   - Room/Type filter listboxes
#   - Debounced search
#   - Click-to-sort TableView

import math
import os
import re
import sys
import subprocess
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

# PDF
from reportlab.lib import colors
from pdf_common import build_section_pdf, PALETTE_KAWAII, PALETTE_PLAIN

from data_core import load_raw_df, automap_columns, normalize_rooms, ellipses, truncate_text
from themes import MOVEUP_THEME, apply_theme


APP_TITLE = "Sample Manager"
SAMPLE_COST_THRESHOLD = 0.01

SAMPLE_ROOMS = {"sample sales floor", "sample vault"}
BACKSTOCK_ROOMS = {"backstock"}
EXCLUDED_TYPES = {"accessories"}  # Always excluded from sample inventory

COL_PATTERNS = {
    "wholesale_cost": [r"wholesale\s*cost", r"wholesale\s*price", r"\bwholesale\b", r"\bcost\b",
                       r"unit\s*cost", r"vendor\s*cost", r"supplier\s*cost", r"buy\s*price"],
    "unit_price": [r"unit\s*price", r"retail\s*price", r"\bprice\b", r"sell\s*price",
                   r"selling\s*price", r"\bretail\b", r"\bmsrp\b", r"\bsrp\b"],
    "product": [r"product\s*name", r"\bproduct\b"],
    "brand": [r"\bbrand\b"],
    "type": [r"\btype\b", r"\bcategory\b"],
    "room": [r"\broom\b", r"\blocation\b"],
    "qty": [r"available\s*qty", r"qty\s*on\s*hand", r"\bavailable\b", r"\bquantity\b", r"\bqty\b",
            r"on\s*hand", r"\bcount\b"],
    "metrc": [r"\bmetrc\b", r"\btag\b", r"rfid", r"package\s*tag", r"package\s*id", r"metrc\s*tag"],
    "distributor": [r"\bdistributor\b", r"\bvendor\b", r"\bsupplier\b", r"\bproducer\b"],
    "expiry": [r"expir\w*\s*date", r"expir\w+", r"best\s*by", r"use\s*by",
               r"exp\s*date", r"\bexp\b", r"sell\s*by"],
    "received_date": [r"reception\s*date", r"\breception\b", r"receiv\w*\s*date",
                      r"date\s*receiv\w*", r"receipt\s*date",
                      r"packaged?\s*date", r"created?\s*date", r"date\s*packaged?"],
}

# ----------------------------
# Truncation rules
# ----------------------------
TRUNCATE_PRODUCT_TO = 60
TRUNCATE_METRC6_TO = 18


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
        print(f"[samples] Could not auto-open file: {e}")


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


def _fmt_currency(val) -> str:
    """Format a numeric value as '$X,XXX.XX'. Returns '' for non-numeric."""
    try:
        v = float(val)
        if not math.isfinite(v):
            return ""
        return f"${v:,.2f}"
    except (ValueError, TypeError):
        return ""


# ----------------------------
# PDF Export
# ----------------------------

def _sample_total_row_style(columns, table_data):
    """Extra TableStyle: bold + tinted background for TOTAL rows."""
    cmds = []
    if len(table_data) > 1:
        last_row = table_data[-1]
        if last_row and str(last_row[0]).upper().startswith("TOTAL"):
            last_idx = len(table_data) - 1
            cmds.append(("FONTNAME", (0, last_idx), (-1, last_idx), "Helvetica-Bold"))
            cmds.append(("BACKGROUND", (0, last_idx), (-1, last_idx), colors.Color(0.92, 0.90, 0.96)))
    return cmds


def _sample_pdf_export(
    path: str,
    title: str,
    subtitle: str,
    sections: list,
    kawaii_pdf: bool = False,
):
    """Export one or more tables to a single landscape PDF via pdf_common."""
    import pandas as _pd

    palette = PALETTE_KAWAII if kawaii_pdf else PALETTE_PLAIN

    # Convert (section_title, df, columns) → (section_title, columns, data_rows)
    converted = []
    for section_title, df, columns in sections:
        if df is None or df.empty:
            converted.append((section_title, list(columns), []))
        else:
            out = df[columns].copy()
            for c in out.columns:
                if _pd.api.types.is_datetime64_any_dtype(out[c]):
                    out[c] = out[c].dt.strftime("%Y-%m-%d")
            out = out.fillna("")
            converted.append((section_title, list(columns), out.values.tolist()))

    build_section_pdf(
        path, title, subtitle, converted,
        palette=palette,
        extra_style_fn=_sample_total_row_style,
    )


# ----------------------------
# Data Model
# ----------------------------

class SampleDataModel:
    """
    Data model for sample inventory: loads a METRC file, detects columns, filters sample items.

    Standalone from main.py's data pipeline because SampleApp runs as an independent
    Toplevel window with its own file import. Does not share state with MoveUpGUI.
    """

    def __init__(self):
        """
        Initialise a blank SampleDataModel with no data loaded.

        All ``col_*`` attributes start as ``None``.  Call ``load_file()`` to
        populate them.  ``df_raw`` holds the full loaded DataFrame;
        ``df_samples`` holds the last result of ``filter_samples()``.
        """
        self.df_raw: pd.DataFrame | None = None
        self.df_samples: pd.DataFrame | None = None

        self.col_wholesale_cost: str | None = None
        self.col_unit_price: str | None = None
        self.col_product: str | None = None
        self.col_brand: str | None = None
        self.col_type: str | None = None
        self.col_room: str | None = None
        self.col_qty: str | None = None
        self.col_metrc: str | None = None
        self.col_metrc6: str | None = None
        self.col_distributor: str | None = None
        self.col_expiry: str | None = None
        self.col_received_date: str | None = None

    def load_file(self, path: str):
        """
        Load an inventory file via data_core.load_raw_df + automap_columns,
        then detect Wholesale Cost and create the METRC-6 derived column.
        Raises ValueError if no Wholesale Cost column is found.
        """
        df = load_raw_df(path)
        df, _rename_map = automap_columns(df)

        # Detect columns via mapped names first, then fallback to regex
        self.col_wholesale_cost = "Wholesale Cost" if "Wholesale Cost" in df.columns else \
            first_matching_col(df.columns, COL_PATTERNS["wholesale_cost"])
        if not self.col_wholesale_cost or self.col_wholesale_cost not in df.columns:
            raise ValueError(
                "No Wholesale Cost column found.\n"
                "Expected something like: 'Wholesale Cost', 'Cost', 'Unit Cost', etc."
            )

        # Ensure numeric
        df[self.col_wholesale_cost] = pd.to_numeric(df[self.col_wholesale_cost], errors="coerce").fillna(0.0)

        self.col_unit_price = "Unit Price" if "Unit Price" in df.columns else \
            first_matching_col(df.columns, COL_PATTERNS["unit_price"])

        self.col_product = "Product Name" if "Product Name" in df.columns else \
            first_matching_col(df.columns, COL_PATTERNS["product"])
        self.col_brand = "Brand" if "Brand" in df.columns else \
            first_matching_col(df.columns, COL_PATTERNS["brand"])
        self.col_type = "Type" if "Type" in df.columns else \
            first_matching_col(df.columns, COL_PATTERNS["type"])
        self.col_room = "Room" if "Room" in df.columns else \
            first_matching_col(df.columns, COL_PATTERNS["room"])
        self.col_qty = "Qty On Hand" if "Qty On Hand" in df.columns else \
            first_matching_col(df.columns, COL_PATTERNS["qty"])
        self.col_metrc = "Package Barcode" if "Package Barcode" in df.columns else \
            first_matching_col(df.columns, COL_PATTERNS["metrc"])
        self.col_distributor = "Distributor" if "Distributor" in df.columns else \
            first_matching_col(df.columns, COL_PATTERNS["distributor"])

        self.col_expiry = first_matching_col(df.columns, COL_PATTERNS["expiry"])
        if self.col_expiry and self.col_expiry in df.columns:
            df[self.col_expiry] = pd.to_datetime(df[self.col_expiry], errors="coerce", format="mixed")

        self.col_received_date = "Reception Date" if "Reception Date" in df.columns else \
            "Received Date" if "Received Date" in df.columns else \
            first_matching_col(df.columns, COL_PATTERNS["received_date"])
        if self.col_received_date and self.col_received_date in df.columns:
            df[self.col_received_date] = pd.to_datetime(df[self.col_received_date], errors="coerce", format="mixed")

        # Create derived METRC-6 column
        if self.col_metrc and self.col_metrc in df.columns:
            df["METRC-6"] = df[self.col_metrc].map(metrc_last6)
        else:
            df["METRC-6"] = ""
        self.col_metrc6 = "METRC-6"

        self.df_raw = df
        self.df_samples = None

    def _get_sample_mask(self, df: pd.DataFrame) -> pd.Series:
        """Return boolean mask for rows where Wholesale Cost <= threshold."""
        return df[self.col_wholesale_cost] <= SAMPLE_COST_THRESHOLD

    def filter_samples(self, search: str = "", room_filter=None, type_filter=None):
        """
        Filter to sample items (Wholesale Cost <= 0.01), then apply
        search / room / type filters. Sets self.df_samples.
        """
        if self.df_raw is None:
            self.df_samples = None
            return None

        df = self.df_raw.copy()

        # Core sample filter
        df = df[self._get_sample_mask(df)].copy()

        # Exclude accessories
        if self.col_type and self.col_type in df.columns:
            df = df[~df[self.col_type].astype(str).str.strip().str.lower().isin(EXCLUDED_TYPES)].copy()

        # Room filter
        if room_filter and self.col_room and self.col_room in df.columns:
            df = df[df[self.col_room].astype(str).isin(room_filter)].copy()

        # Type filter
        if type_filter and self.col_type and self.col_type in df.columns:
            df = df[df[self.col_type].astype(str).isin(type_filter)].copy()

        # Text search across all text columns
        q = (search or "").strip().lower()
        if q:
            text_cols = [c for c in df.columns if df[c].dtype == "object" or df[c].dtype.name == "string"]
            if text_cols:
                big = df[text_cols].fillna("").astype(str).agg(" | ".join, axis=1).str.lower()
                df = df.loc[big.str.contains(q, na=False)].copy()

        # Sort: Type -> Brand -> Product (room shown via color coding, not sort order)
        sort_cols = []
        for col in [self.col_type, self.col_brand, self.col_product]:
            if col and col in df.columns:
                sort_cols.append(col)
        if sort_cols:
            df = df.sort_values(by=sort_cols, ascending=True, na_position="last")

        self.df_samples = df
        return df

    def get_all_rooms_from_raw(self) -> list[str]:
        """Unique room values from sample rows (before room/type filter)."""
        if self.df_raw is None or self.col_room is None or self.col_room not in self.df_raw.columns:
            return []
        sample_df = self.df_raw[self._get_sample_mask(self.df_raw)]
        return sorted(sample_df[self.col_room].dropna().astype(str).unique().tolist())

    def get_all_types_from_raw(self) -> list[str]:
        """Unique type values from sample rows (before room/type filter)."""
        if self.df_raw is None or self.col_type is None or self.col_type not in self.df_raw.columns:
            return []
        sample_df = self.df_raw[self._get_sample_mask(self.df_raw)]
        return sorted(sample_df[self.col_type].dropna().astype(str).unique().tolist())

    def _build_summary(self, label: str, group_col_attr: str,
                        include_retail: bool = False) -> pd.DataFrame:
        """Generic summary: group by a column, count samples, sum qty, optionally sum retail value."""
        base_cols = [label, "Sample Count", "Total Qty"]
        if include_retail:
            base_cols.append("Total Retail Value")

        if self.df_samples is None or self.df_samples.empty:
            return pd.DataFrame(columns=base_cols)

        df = self.df_samples.copy()
        src_col = getattr(self, group_col_attr, None)
        src_col = src_col if (src_col and src_col in df.columns) else None
        qty_col = self.col_qty if (self.col_qty and self.col_qty in df.columns) else None

        df["_grp"] = df[src_col].astype(str).fillna("Unknown") if src_col else "Unknown"
        df["_qty"] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0) if qty_col else 0

        aggs = {"sample_count": ("_grp", "size"), "total_qty": ("_qty", "sum")}
        if include_retail:
            price_col = self.col_unit_price if (self.col_unit_price and self.col_unit_price in df.columns) else None
            df["_price"] = pd.to_numeric(df[price_col], errors="coerce").fillna(0.0) if price_col else 0.0
            df["_retail_val"] = df["_qty"] * df["_price"]
            aggs["total_retail"] = ("_retail_val", "sum")

        grp = df.groupby("_grp", sort=True).agg(**aggs).reset_index()

        rows = []
        for _, r in grp.iterrows():
            row = {label: r["_grp"], "Sample Count": int(r["sample_count"]),
                   "Total Qty": int(r["total_qty"])}
            if include_retail:
                row["Total Retail Value"] = _fmt_currency(r["total_retail"])
            rows.append(row)

        total_row = {label: "TOTAL", "Sample Count": int(grp["sample_count"].sum()),
                     "Total Qty": int(grp["total_qty"].sum())}
        if include_retail:
            total_row["Total Retail Value"] = _fmt_currency(grp["total_retail"].sum())
        rows.append(total_row)

        return pd.DataFrame(rows)

    def build_brand_summary(self) -> pd.DataFrame:
        return self._build_summary("Brand", "col_brand", include_retail=True)

    def build_type_summary(self) -> pd.DataFrame:
        return self._build_summary("Type", "col_type")

    def build_room_summary(self) -> pd.DataFrame:
        return self._build_summary("Room", "col_room")

    def build_action_list(self) -> pd.DataFrame:
        """Samples in Backstock only, sorted Room -> Type -> Brand -> Product."""
        if self.df_samples is None or self.df_samples.empty:
            return pd.DataFrame()

        df = self.df_samples.copy()

        if self.col_room and self.col_room in df.columns:
            mask = df[self.col_room].astype(str).str.strip().str.lower().isin(BACKSTOCK_ROOMS)
            df = df[mask].copy()
        else:
            return pd.DataFrame()

        if df.empty:
            return pd.DataFrame()

        sort_cols = []
        for col in [self.col_type, self.col_brand, self.col_product]:
            if col and col in df.columns:
                sort_cols.append(col)
        if sort_cols:
            df = df.sort_values(by=sort_cols, ascending=True, na_position="last")

        return df

    def build_distributor_report(self) -> pd.DataFrame:
        """By Distributor with sample count, total qty, brands list + TOTAL row."""
        if self.df_samples is None or self.df_samples.empty:
            return pd.DataFrame(columns=["Distributor", "Sample Count", "Total Qty", "Brands"])

        df = self.df_samples.copy()
        dist_col = self.col_distributor if (self.col_distributor and self.col_distributor in df.columns) else None
        qty_col = self.col_qty if (self.col_qty and self.col_qty in df.columns) else None
        brand_col = self.col_brand if (self.col_brand and self.col_brand in df.columns) else None

        if not dist_col:
            return pd.DataFrame(columns=["Distributor", "Sample Count", "Total Qty", "Brands"])

        df["_dist"] = df[dist_col].astype(str).fillna("Unknown").str.strip()
        df["_qty"] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0) if qty_col else 0
        df["_brand"] = df[brand_col].astype(str).fillna("") if brand_col else ""

        rows = []
        for dist, grp in df.groupby("_dist", sort=True):
            brands = sorted(grp["_brand"].unique().tolist())
            brands_str = ", ".join(b for b in brands if b and b.lower() != "nan")
            rows.append({
                "Distributor": dist,
                "Sample Count": int(len(grp)),
                "Total Qty": int(grp["_qty"].sum()),
                "Brands": brands_str,
            })

        total_count = sum(r["Sample Count"] for r in rows)
        total_qty = sum(r["Total Qty"] for r in rows)
        rows.append({
            "Distributor": "TOTAL",
            "Sample Count": total_count,
            "Total Qty": total_qty,
            "Brands": "",
        })

        return pd.DataFrame(rows)


# ----------------------------
# UI Widgets
# ----------------------------

class TableView(ttk.Frame):
    """Scrollable treeview widget with sortable columns and auto-width sizing."""

    def __init__(self, parent):
        super().__init__(parent)
        self.tree = ttk.Treeview(self, columns=(), show="headings", selectmode="extended")
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        # Row data for multi-select export  {iid -> {col: value}}
        self._iid_to_row_data: dict = {}

        # Sort state
        self._raw_df: pd.DataFrame | None = None
        self._columns: list[str] = []
        self._trunc_map: dict[str, int] = {}
        self._room_col: str | None = None
        self._color_by_room: bool = False
        self._expiry_col: str | None = None
        self._group_by_col: str | None = None
        self._sort_col: str | None = None
        self._sort_asc: bool = True

        # Define tags for highlighting
        self.tree.tag_configure("normal", background=MOVEUP_THEME["tree_bg"])
        self.tree.tag_configure("expiring_soon", background="#FFCDD2")    # light red — expires within 30 days
        self.tree.tag_configure("group_header", background="#D1C4E9", foreground="#311B92",
                                font=("TkDefaultFont", 10, "bold"))
        self.tree.tag_configure("totals_row", foreground="#3A2869")       # bold purple for totals

    def render(self, df: pd.DataFrame, columns: list[str],
               trunc_map: dict[str, int] | None = None,
               room_col: str | None = None,
               color_by_room: bool = False,
               group_by_col: str | None = None,
               expiry_col: str | None = None):
        # Store for re-use by _sort_by_col
        self._iid_to_row_data = {}
        self._raw_df = df
        self._columns = list(columns)
        self._trunc_map = trunc_map or {}
        self._room_col = room_col
        self._color_by_room = color_by_room
        self._expiry_col = expiry_col
        self._group_by_col = group_by_col

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = columns

        for c in columns:
            indicator = (" \u25b2" if self._sort_asc else " \u25bc") if c == self._sort_col else ""
            self.tree.heading(c, text=c + indicator, command=lambda col=c: self._sort_by_col(col))

            col_name = str(c).lower()

            # Center qty-like and count columns
            if ("qty" in col_name) or col_name in {"quantity", "count", "on hand", "sample count", "total qty"}:
                self.tree.column(c, width=max(120, min(520, len(c) * 12)), anchor="center")
            elif "cost" in col_name or "price" in col_name or "value" in col_name or "retail" in col_name:
                self.tree.column(c, width=max(120, min(520, len(c) * 12)), anchor="e")
            else:
                self.tree.column(c, width=max(120, min(520, len(c) * 12)), anchor="w")

        if df is None or df.empty or not columns:
            return

        out = df[columns].copy()

        trunc_map = self._trunc_map
        for col, lim in trunc_map.items():
            if col in out.columns:
                out[col] = out[col].map(lambda v, _lim=lim: truncate_text(v, _lim))

        for c in out.columns:
            if pd.api.types.is_datetime64_any_dtype(out[c]):
                out[c] = out[c].dt.strftime("%Y-%m-%d")
            else:
                out[c] = out[c].apply(
                    lambda v: v.strftime("%Y-%m-%d") if isinstance(v, pd.Timestamp) else v
                )
        out = out.fillna("")

        last_group = None
        for idx, row in out.iterrows():
            values = [row[c] for c in columns]

            # Insert group separator row when category changes
            if group_by_col and group_by_col in df.columns:
                try:
                    group_val = str(df.loc[idx, group_by_col]).strip()
                except (KeyError, ValueError):
                    group_val = ""
                if group_val and group_val != last_group:
                    last_group = group_val
                    sep_values = [f"── {group_val} ──"] + [""] * (len(columns) - 1)
                    self.tree.insert("", "end", values=sep_values, tags=("group_header",))

            tag = "normal"

            # Light red if expiring within 30 days
            if expiry_col and expiry_col in df.columns:
                try:
                    exp_val = df.loc[idx, expiry_col]
                    if pd.notna(exp_val):
                        exp_date = exp_val if hasattr(exp_val, "date") else pd.to_datetime(exp_val, errors="coerce")
                        if pd.notna(exp_date):
                            days_left = (exp_date.date() - datetime.now().date()).days
                            if days_left <= 30:
                                tag = "expiring_soon"
                except (KeyError, ValueError, TypeError, AttributeError):
                    pass

            # Detect TOTAL rows
            if values and str(values[0]).upper().startswith("TOTAL"):
                tag = "totals_row"

            iid = self.tree.insert("", "end", values=values, tags=(tag,))
            # Store original (non-truncated) values for selection export
            orig_row = {c: (df.loc[idx, c] if c in df.columns else "") for c in columns}
            self._iid_to_row_data[iid] = orig_row

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

        # Prefer numeric sort; fall back to case-insensitive string
        numeric = pd.to_numeric(series, errors="coerce")
        if numeric.notna().any():
            df["_sk"] = numeric
        else:
            df["_sk"] = series.astype(str).str.lower()

        # If grouping is active, keep category groups together and sort within them
        if self._group_by_col and self._group_by_col in df.columns and self._group_by_col != col:
            df["_gk"] = df[self._group_by_col].astype(str).str.lower()
            df = df.sort_values(
                ["_gk", "_sk"], ascending=[True, self._sort_asc], na_position="last"
            ).drop(columns=["_sk", "_gk"])
        else:
            df = df.sort_values("_sk", ascending=self._sort_asc, na_position="last").drop(columns=["_sk"])

        self.render(df, self._columns, self._trunc_map, self._room_col, self._color_by_room,
                    self._group_by_col, expiry_col=self._expiry_col)

    def get_selected_df(self, columns: list | None = None) -> pd.DataFrame:
        """Return a DataFrame of currently selected rows (group headers excluded)."""
        selected = self.tree.selection()
        rows = [self._iid_to_row_data[iid] for iid in selected if iid in self._iid_to_row_data]
        if not rows:
            return pd.DataFrame()
        cols = columns or list(rows[0].keys())
        return pd.DataFrame(rows, columns=cols)

    def selected_count(self) -> int:
        """Count of selected data rows (excludes group-header separators)."""
        return sum(1 for iid in self.tree.selection() if iid in self._iid_to_row_data)


# ----------------------------
# Main App
# ----------------------------

class SampleApp(tk.Toplevel):
    """
    Sample inventory manager — Toplevel window for tracking sample items.

    Loads a METRC file independently via ``SampleDataModel.load_file()`` and
    identifies samples by Wholesale Cost ≤ ``SAMPLE_COST_THRESHOLD`` (0.01).

    Four tabs:
    - **Inventory**: full sample item list with room/type filters and live
      search.  Supports adding rows to the Distribution List via double-click
      or the context menu.
    - **Summary Dashboard**: brand/type/room aggregates with sample counts,
      total qty, and retail value.
    - **Distribution List**: user-curated list of items to distribute, with
      basket grouping and an "Action List" (backstock samples only) sub-tab.
    - **Distributor Report**: per-distributor breakdown with brand lists.

    PDF export uses ``pdf_common.build_section_pdf()`` with kawaii/plain
    palette toggle.  CSV and Excel exports are also available.
    """

    def __init__(self, master):
        super().__init__(master)
        self.title(APP_TITLE)
        self.geometry("1340x820")

        self.model = SampleDataModel()

        self._debounce_job: str | None = None
        self._last_dir: str | None = None

        self.pdf_kawaii_var = tk.BooleanVar(value=True)
        self.search_var = tk.StringVar(value="")
        self.info_var = tk.StringVar(value="Load a file to begin.")
        self.selection_count_var = tk.StringVar(value="Distribution list is empty")

        # Persistent distribution basket (survives filter/search changes)
        self._dist_basket: list[dict] = []   # ordered list of row dicts
        self._dist_basket_keys: set = set()  # for fast dedup

        # Tab data caches
        self.inv_cols: list[str] = []
        self.brand_summary_df: pd.DataFrame | None = None
        self.type_summary_df: pd.DataFrame | None = None
        self.room_summary_df: pd.DataFrame | None = None
        self.dist_df: pd.DataFrame | None = None

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
        if self.model.col_metrc6 and self.model.df_raw is not None and self.model.col_metrc6 in self.model.df_raw.columns:
            m[self.model.col_metrc6] = TRUNCATE_METRC6_TO
        return m

    def apply_subtle_theme(self):
        apply_theme(self, "sample_mgr_theme")

    def _build_ui(self):
        # ---------- Top bar ----------
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Button(top, text="Open Inventory File", command=self.open_file).pack(side="left")
        self.bind("<Control-o>", lambda e: self.open_file())

        ttk.Checkbutton(top, text="Kawaii PDF", variable=self.pdf_kawaii_var)\
            .pack(side="left", padx=(12, 0))

        _btn_pdf = ttk.Button(top, text="Export PDF (Current Tab)", command=self.export_pdf, state="disabled")
        _btn_pdf.pack(side="right")
        _btn_xlsx = ttk.Button(top, text="Export Excel (Current Tab)", command=self.export_xlsx, state="disabled")
        _btn_xlsx.pack(side="right", padx=(0, 8))
        _btn_csv = ttk.Button(top, text="Export CSV (Current Tab)", command=self.export_csv, state="disabled")
        _btn_csv.pack(side="right", padx=(0, 8))
        self._export_buttons = [_btn_pdf, _btn_xlsx, _btn_csv]

        # ---------- Status bar ----------
        ttk.Label(self, textvariable=self.info_var).pack(fill="x", padx=10, pady=(0, 6))

        # ---------- Filter row ----------
        filter_row = ttk.Frame(self, padding=(10, 0, 10, 6))
        filter_row.pack(fill="x")

        # Room filter
        ttk.Label(filter_row, text="Room filter:").pack(side="left")
        room_lb_wrap = ttk.Frame(filter_row)
        room_lb_wrap.pack(side="left", padx=(6, 0))
        self.room_listbox = tk.Listbox(
            room_lb_wrap,
            selectmode=tk.EXTENDED,
            height=4,
            width=22,
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
        room_vsb = ttk.Scrollbar(room_lb_wrap, orient="vertical", command=self.room_listbox.yview)
        self.room_listbox.configure(yscrollcommand=room_vsb.set)
        self.room_listbox.pack(side="left", fill="both", expand=True)
        room_vsb.pack(side="left", fill="y")
        self.room_listbox.bind("<<ListboxSelect>>", lambda e: self.apply_filter())
        ttk.Button(filter_row, text="All Rooms", command=self._select_all_rooms).pack(side="left", padx=(4, 12))

        # Type filter
        ttk.Label(filter_row, text="Type filter:").pack(side="left")
        type_lb_wrap = ttk.Frame(filter_row)
        type_lb_wrap.pack(side="left", padx=(6, 0))
        self.type_listbox = tk.Listbox(
            type_lb_wrap,
            selectmode=tk.EXTENDED,
            height=4,
            width=22,
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
        type_vsb = ttk.Scrollbar(type_lb_wrap, orient="vertical", command=self.type_listbox.yview)
        self.type_listbox.configure(yscrollcommand=type_vsb.set)
        self.type_listbox.pack(side="left", fill="both", expand=True)
        type_vsb.pack(side="left", fill="y")
        self.type_listbox.bind("<<ListboxSelect>>", lambda e: self.apply_filter())
        ttk.Button(filter_row, text="All Types", command=self._select_all_types).pack(side="left", padx=(4, 0))

        # ---------- Notebook ----------
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.nb = nb

        # Tab 1: Sample Inventory
        tab1 = ttk.Frame(nb)
        nb.add(tab1, text="Sample Inventory")

        search_row = ttk.Frame(tab1, padding=(0, 8, 0, 8))
        search_row.pack(fill="x")
        ttk.Label(search_row, text="Search:").pack(side="left")
        self.search_var.trace_add("write", lambda *_: self._debounce(300, self.apply_filter))
        ttk.Entry(search_row, textvariable=self.search_var).pack(side="left", fill="x", expand=True, padx=(6, 0))

        self.inv_table = TableView(tab1)
        self.inv_table.pack(fill="both", expand=True)
        self.inv_table.tree.bind("<Button-3>", self._show_inv_context_menu)

        # Distribution action bar — below Tab 1 inventory table
        dist_bar = ttk.Frame(tab1, padding=(4, 6, 4, 4))
        dist_bar.pack(fill="x")

        ttk.Label(
            dist_bar, textvariable=self.selection_count_var,
            foreground=MOVEUP_THEME["label_fg"], font=("TkDefaultFont", 9, "italic")
        ).pack(side="left")

        ttk.Label(dist_bar, text="  |  ", foreground=MOVEUP_THEME["label_fg"]).pack(side="left")
        ttk.Label(dist_bar, text="Click a row, then click Add — or right-click a row",
                  foreground="#888888", font=("TkDefaultFont", 8)).pack(side="left")

        ttk.Button(
            dist_bar, text="Add to Distribution  +",
            command=self._add_selected_to_distribution,
        ).pack(side="right", padx=(4, 0))

        # Tab 2: Summary Dashboard
        tab2 = ttk.Frame(nb)
        nb.add(tab2, text="Summary Dashboard")

        summary_split = ttk.Panedwindow(tab2, orient="vertical")
        summary_split.pack(fill="both", expand=True)

        brand_frame = ttk.Frame(summary_split)
        type_frame = ttk.Frame(summary_split)
        room_frame = ttk.Frame(summary_split)
        summary_split.add(brand_frame, weight=1)
        summary_split.add(type_frame, weight=1)
        summary_split.add(room_frame, weight=1)

        ttk.Label(brand_frame, text="By Brand").pack(anchor="w", padx=2, pady=(6, 4))
        self.brand_table = TableView(brand_frame)
        self.brand_table.pack(fill="both", expand=True)

        ttk.Label(type_frame, text="By Type").pack(anchor="w", padx=2, pady=(6, 4))
        self.type_table = TableView(type_frame)
        self.type_table.pack(fill="both", expand=True)

        ttk.Label(room_frame, text="By Room").pack(anchor="w", padx=2, pady=(6, 4))
        self.room_table = TableView(room_frame)
        self.room_table.pack(fill="both", expand=True)

        # Tab 3: Distribution List
        tab3 = ttk.Frame(nb)
        nb.add(tab3, text="Distribution List")

        dist_list_header = ttk.Frame(tab3, padding=(4, 8, 4, 4))
        dist_list_header.pack(fill="x")
        ttk.Label(
            dist_list_header,
            text="Items added for distribution  \u2014  use Tab 1 to add items",
            font=("TkDefaultFont", 9, "italic"),
            foreground=MOVEUP_THEME["label_fg"],
        ).pack(side="left")

        ttk.Button(
            dist_list_header, text="Clear All",
            command=self._clear_distribution,
        ).pack(side="right", padx=(4, 0))
        ttk.Button(
            dist_list_header, text="Remove Selected  \u2212",
            command=self._remove_selected_from_distribution,
        ).pack(side="right", padx=(4, 0))

        self.dist_list_table = TableView(tab3)
        self.dist_list_table.pack(fill="both", expand=True)
        self.dist_list_table.tree.bind("<Button-3>", self._show_dist_context_menu)

        # Tab 4: Distributor Report
        tab4 = ttk.Frame(nb)
        nb.add(tab4, text="Distributor Report")
        ttk.Label(
            tab4,
            text="Sample items grouped by distributor, with brand breakdown."
        ).pack(anchor="w", padx=2, pady=(8, 8))
        self.dist_table = TableView(tab4)
        self.dist_table.pack(fill="both", expand=True)

    # ---------- Filter helpers ----------

    def _select_all_rooms(self, trigger_filter=True):
        self.room_listbox.selection_set(0, tk.END)
        if trigger_filter:
            self.apply_filter()

    def _select_all_types(self, trigger_filter=True):
        self.type_listbox.selection_set(0, tk.END)
        if trigger_filter:
            self.apply_filter()

    def _get_selected_rooms(self) -> list[str] | None:
        total = self.room_listbox.size()
        if total == 0:
            return None
        selected = self.room_listbox.curselection()
        if not selected or len(selected) == total:
            return None
        return [self.room_listbox.get(i) for i in selected]

    def _get_selected_types(self) -> list[str] | None:
        total = self.type_listbox.size()
        if total == 0:
            return None
        selected = self.type_listbox.curselection()
        if not selected or len(selected) == total:
            return None
        return [self.type_listbox.get(i) for i in selected]

    def _debounce(self, ms: int, func):
        if self._debounce_job is not None:
            self.after_cancel(self._debounce_job)

        def _run():
            self._debounce_job = None
            func()

        self._debounce_job = self.after(ms, _run)

    def _populate_filters(self):
        """Fill room and type listboxes from sample rows."""
        self.room_listbox.delete(0, tk.END)
        rooms = self.model.get_all_rooms_from_raw()
        for r in rooms:
            self.room_listbox.insert(tk.END, r)
        self._select_all_rooms(trigger_filter=False)

        self.type_listbox.delete(0, tk.END)
        types = self.model.get_all_types_from_raw()
        for t in types:
            self.type_listbox.insert(tk.END, t)
        self._select_all_types(trigger_filter=False)

    # ---------- File open ----------

    def open_file(self, path: str | None = None):
        if path is None:
            path = filedialog.askopenfilename(
                title="Select Inventory Export",
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
        # Clear the distribution basket — items from the previous file are stale
        self._dist_basket.clear()
        self._dist_basket_keys.clear()
        self._update_distribution_count()
        self._refresh_distribution_list()
        try:
            self.model.load_file(path)
            self.title(f"{APP_TITLE} \u2014 {os.path.basename(path)}")
            self._populate_filters()
            self.info_var.set(
                f"Loaded {len(self.model.df_raw):,} rows | "
                f"Wholesale Cost: '{self.model.col_wholesale_cost}' | "
                f"Product: '{self.model.col_product or 'N/A'}' | "
                f"Room: '{self.model.col_room or 'N/A'}' | "
                f"METRC: '{self.model.col_metrc or 'N/A'}'"
            )
            self.apply_filter()
            self._set_export_buttons_state("normal")
        except Exception as e:
            messagebox.showerror("Load Error", str(e))

    # ---------- Apply filter (renders all tabs) ----------

    def apply_filter(self):
        df = self.model.filter_samples(
            search=self.search_var.get(),
            room_filter=self._get_selected_rooms(),
            type_filter=self._get_selected_types(),
        )
        if df is None:
            return

        trunc_map = self._build_trunc_map()

        # --- Tab 1: Sample Inventory ---
        preferred = []
        for col_attr in [self.model.col_room, self.model.col_type, self.model.col_brand]:
            if col_attr and col_attr in df.columns:
                preferred.append(col_attr)

        if self.model.col_metrc6 and self.model.col_metrc6 in df.columns:
            preferred.append(self.model.col_metrc6)

        if self.model.col_product and self.model.col_product in df.columns:
            preferred.append(self.model.col_product)

        if self.model.col_qty and self.model.col_qty in df.columns:
            preferred.append(self.model.col_qty)

        if self.model.col_wholesale_cost and self.model.col_wholesale_cost in df.columns:
            preferred.append(self.model.col_wholesale_cost)

        if self.model.col_unit_price and self.model.col_unit_price in df.columns:
            preferred.append(self.model.col_unit_price)

        if self.model.col_received_date and self.model.col_received_date in df.columns:
            preferred.append(self.model.col_received_date)

        if self.model.col_expiry and self.model.col_expiry in df.columns:
            preferred.append(self.model.col_expiry)

        cols, seen = [], set()
        for c in preferred:
            if c and c in df.columns and c not in seen:
                cols.append(c)
                seen.add(c)
        if not cols:
            cols = list(df.columns[:12])

        self.inv_cols = cols
        self.inv_table.render(df, cols, trunc_map=trunc_map,
                              group_by_col=self.model.col_type,
                              expiry_col=self.model.col_expiry)

        # --- Tab 2: Summary Dashboard ---
        self.brand_summary_df = self.model.build_brand_summary()
        brand_cols = ["Brand", "Sample Count", "Total Qty", "Total Retail Value"]
        brand_cols = [c for c in brand_cols if c in self.brand_summary_df.columns]
        self.brand_table.render(self.brand_summary_df, brand_cols)

        self.type_summary_df = self.model.build_type_summary()
        type_cols = ["Type", "Sample Count", "Total Qty"]
        type_cols = [c for c in type_cols if c in self.type_summary_df.columns]
        self.type_table.render(self.type_summary_df, type_cols)

        self.room_summary_df = self.model.build_room_summary()
        room_cols = ["Room", "Sample Count", "Total Qty"]
        room_cols = [c for c in room_cols if c in self.room_summary_df.columns]
        self.room_table.render(self.room_summary_df, room_cols)

        # --- Tab 3: Distribution List (live view of Tab 1 selection) ---
        # Refresh after inventory re-renders (selection is cleared by Treeview on re-render)
        self._refresh_distribution_list()

        # --- Tab 4: Distributor Report ---
        self.dist_df = self.model.build_distributor_report()
        dist_cols = ["Distributor", "Sample Count", "Total Qty", "Brands"]
        dist_cols = [c for c in dist_cols if c in self.dist_df.columns]
        self.dist_table.render(self.dist_df, dist_cols)

        # Update status bar
        sample_count = len(df) if df is not None else 0
        room_info = ""
        if self.model.col_room and self.model.col_room in df.columns:
            unique_rooms = df[self.model.col_room].dropna().astype(str).unique()
            room_info = f" | Rooms: {len(unique_rooms)}"
        self.info_var.set(
            f"Samples: {sample_count:,} (Wholesale Cost \u2264 ${SAMPLE_COST_THRESHOLD:.2f})"
            f"{room_info}"
            f" | Search: '{self.search_var.get() or ''}'"
        )

    # ---------- Selection helpers ----------

    # ---------- Distribution basket helpers ----------

    def _basket_key(self, row: dict) -> str:
        """Unique key for a row — uses METRC code if available, else hashes all values."""
        for col in [self.model.col_metrc, self.model.col_metrc6]:
            if col and col in row and str(row[col]).strip():
                return str(row[col]).strip()
        return str(hash(tuple(sorted((str(k), str(v)) for k, v in row.items()))))

    def _add_selected_to_distribution(self):
        """Add currently highlighted Tab 1 row(s) to the persistent basket."""
        rows_df = self.inv_table.get_selected_df(self.inv_cols if self.inv_cols else None)
        if rows_df.empty:
            self.selection_count_var.set("⬆ select a row first, then click Add")
            return
        added = 0
        for _, row in rows_df.iterrows():
            row_dict = row.to_dict()
            key = self._basket_key(row_dict)
            if key not in self._dist_basket_keys:
                self._dist_basket.append(row_dict)
                self._dist_basket_keys.add(key)
                added += 1
        self._update_distribution_count()
        self._refresh_distribution_list()

    def _remove_selected_from_distribution(self):
        """Remove rows selected in Tab 3 from the basket."""
        selected = self.dist_list_table.tree.selection()
        remove_keys = {
            self._basket_key(self.dist_list_table._iid_to_row_data[iid])
            for iid in selected
            if iid in self.dist_list_table._iid_to_row_data
        }
        if not remove_keys:
            return
        self._dist_basket = [r for r in self._dist_basket if self._basket_key(r) not in remove_keys]
        self._dist_basket_keys -= remove_keys
        self._update_distribution_count()
        self._refresh_distribution_list()

    def _clear_distribution(self):
        """Empty the entire basket."""
        self._dist_basket.clear()
        self._dist_basket_keys.clear()
        self._update_distribution_count()
        self._refresh_distribution_list()

    def _update_distribution_count(self):
        n = len(self._dist_basket)
        if n == 0:
            self.selection_count_var.set("Distribution list is empty")
        elif n == 1:
            self.selection_count_var.set("1 item in distribution list")
        else:
            self.selection_count_var.set(f"{n} items in distribution list")

    def _refresh_distribution_list(self):
        """Re-render Tab 3 from the persistent basket."""
        trunc_map = self._build_trunc_map()
        if not self._dist_basket:
            self.dist_list_table.render(pd.DataFrame(), self.inv_cols or [])
            return
        df = pd.DataFrame(self._dist_basket)
        cols = [c for c in (self.inv_cols or []) if c in df.columns] or list(df.columns)
        self.dist_list_table.render(
            df, cols, trunc_map=trunc_map,
            group_by_col=self.model.col_type,
            expiry_col=self.model.col_expiry,
        )

    def _show_inv_context_menu(self, event):
        """Right-click on Tab 1 inventory table."""
        iid = self.inv_table.tree.identify_row(event.y)
        if not iid or iid not in self.inv_table._iid_to_row_data:
            return
        self.inv_table.tree.selection_set(iid)
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Add to Distribution List", command=self._add_selected_to_distribution)
        menu.tk_popup(event.x_root, event.y_root)

    def _show_dist_context_menu(self, event):
        """Right-click on Tab 3 distribution table."""
        iid = self.dist_list_table.tree.identify_row(event.y)
        if not iid or iid not in self.dist_list_table._iid_to_row_data:
            return
        self.dist_list_table.tree.selection_set(iid)
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Remove from Distribution List",
                         command=self._remove_selected_from_distribution)
        menu.tk_popup(event.x_root, event.y_root)

    def export_selected_pdf(self):
        if not self._dist_basket:
            messagebox.showinfo("Empty Distribution List",
                                "Add items to the distribution list first using Tab 1.")
            return
        df = pd.DataFrame(self._dist_basket)
        export_cols = [c for c in (self.inv_cols or []) if c in df.columns] or list(df.columns)
        n = len(df)
        title = "Sample Distribution List"
        subtitle = (
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')} | "
            f"{n} item{'s' if n != 1 else ''}"
        )
        default_name = f"Sample_Distribution_{datetime.now().strftime('%Y-%m-%d_%H%M')}.pdf"
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf", initialfile=default_name, filetypes=[("PDF", "*.pdf")]
        )
        if not path:
            return
        try:
            _sample_pdf_export(
                path, title=title, subtitle=subtitle,
                sections=[(None, df, export_cols)],
                kawaii_pdf=bool(self.pdf_kawaii_var.get()),
            )
            messagebox.showinfo("Export", f"Saved:\n{path}")
            open_file_with_default_app(path)
        except Exception as e:
            messagebox.showerror("PDF Export Error", str(e))

    def export_selected_xlsx(self):
        if not self._dist_basket:
            messagebox.showinfo("Empty Distribution List",
                                "Add items to the distribution list first using Tab 1.")
            return
        df = pd.DataFrame(self._dist_basket)
        export_cols = [c for c in (self.inv_cols or []) if c in df.columns] or list(df.columns)
        default_name = f"Sample_Distribution_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile=default_name, filetypes=[("Excel", "*.xlsx")]
        )
        if not path:
            return
        try:
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df[export_cols].to_excel(writer, sheet_name="Distribution List", index=False)
            messagebox.showinfo("Export", f"Saved:\n{path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not save file:\n{e}")

    # ---------- Exports ----------

    def _current_export_df_and_cols(self):
        if self.model.df_samples is None:
            return None, None, None

        current_tab = self.nb.tab(self.nb.select(), "text")

        if current_tab == "Sample Inventory":
            df = self.model.df_samples
            cols = self.inv_cols if self.inv_cols else list(df.columns[:12])
            return df, cols, "Sample Inventory"

        if current_tab == "Summary Dashboard":
            # Return brand summary as default; XLSX export handles all three
            return self.brand_summary_df, ["Brand", "Sample Count", "Total Qty", "Total Retail Value"], "Sample Summary - By Brand"

        if current_tab == "Distribution List":
            if not self._dist_basket:
                return None, [], "Sample Distribution List"
            df = pd.DataFrame(self._dist_basket)
            cols = [c for c in (self.inv_cols or []) if c in df.columns] or list(df.columns)
            return df, cols, "Sample Distribution List"

        if current_tab == "Distributor Report":
            return self.dist_df, ["Distributor", "Sample Count", "Total Qty", "Brands"], "Sample Distributor Report"

        # Fallback
        return self.model.df_samples, self.inv_cols, "Sample Inventory"

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
        current_tab = self.nb.tab(self.nb.select(), "text")

        if current_tab == "Summary Dashboard":
            # Export all 3 summary tables as separate sheets
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
            if not path:
                return
            try:
                with pd.ExcelWriter(path, engine="openpyxl") as writer:
                    if self.brand_summary_df is not None and not self.brand_summary_df.empty:
                        brand_cols = [c for c in ["Brand", "Sample Count", "Total Qty", "Total Retail Value"]
                                      if c in self.brand_summary_df.columns]
                        self.brand_summary_df[brand_cols].to_excel(writer, sheet_name="By Brand", index=False)

                    if self.type_summary_df is not None and not self.type_summary_df.empty:
                        type_cols = [c for c in ["Type", "Sample Count", "Total Qty"]
                                     if c in self.type_summary_df.columns]
                        self.type_summary_df[type_cols].to_excel(writer, sheet_name="By Type", index=False)

                    if self.room_summary_df is not None and not self.room_summary_df.empty:
                        room_cols = [c for c in ["Room", "Sample Count", "Total Qty"]
                                     if c in self.room_summary_df.columns]
                        self.room_summary_df[room_cols].to_excel(writer, sheet_name="By Room", index=False)

                messagebox.showinfo("Export", f"Saved:\n{path}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Could not save file:\n{e}")
            return

        # All other tabs: single sheet
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
        current_tab = self.nb.tab(self.nb.select(), "text")

        subtitle = (
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')} | "
            f"Threshold: Wholesale Cost \u2264 ${SAMPLE_COST_THRESHOLD:.2f} | "
            f"Kawaii PDF: {self.pdf_kawaii_var.get()}"
        )

        if current_tab == "Summary Dashboard":
            title = "Sample Summary Dashboard"
            default_name = f"Sample_Summary_{datetime.now().strftime('%Y-%m-%d_%H%M')}.pdf"
            path = filedialog.asksaveasfilename(
                defaultextension=".pdf", initialfile=default_name, filetypes=[("PDF", "*.pdf")]
            )
            if not path:
                return

            sections = []
            if self.brand_summary_df is not None and not self.brand_summary_df.empty:
                brand_cols = [c for c in ["Brand", "Sample Count", "Total Qty", "Total Retail Value"]
                              if c in self.brand_summary_df.columns]
                sections.append(("By Brand", self.brand_summary_df, brand_cols))
            if self.type_summary_df is not None and not self.type_summary_df.empty:
                type_cols = [c for c in ["Type", "Sample Count", "Total Qty"]
                             if c in self.type_summary_df.columns]
                sections.append(("By Type", self.type_summary_df, type_cols))
            if self.room_summary_df is not None and not self.room_summary_df.empty:
                room_cols = [c for c in ["Room", "Sample Count", "Total Qty"]
                             if c in self.room_summary_df.columns]
                sections.append(("By Room", self.room_summary_df, room_cols))

            if not sections:
                messagebox.showinfo("Export", "Nothing to export.")
                return

            try:
                _sample_pdf_export(
                    path, title=title, subtitle=subtitle,
                    sections=sections, kawaii_pdf=bool(self.pdf_kawaii_var.get()),
                )
                messagebox.showinfo("Export", f"Saved:\n{path}")
                open_file_with_default_app(path)
            except Exception as e:
                messagebox.showerror("PDF Export Error", str(e))
            return

        # Other tabs: single-section PDF
        df, cols, title = self._current_export_df_and_cols()
        if df is None or df.empty:
            messagebox.showinfo("Export", "Nothing to export.")
            return

        default_name = f"{title.replace(' ', '_')}_{datetime.now().strftime('%Y-%m-%d_%H%M')}.pdf"
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf", initialfile=default_name, filetypes=[("PDF", "*.pdf")]
        )
        if not path:
            return

        export_cols = [c for c in cols if c in df.columns] if cols else list(df.columns)

        try:
            _sample_pdf_export(
                path, title=title, subtitle=subtitle,
                sections=[(None, df, export_cols)],
                kawaii_pdf=bool(self.pdf_kawaii_var.get()),
            )
            messagebox.showinfo("Export", f"Saved:\n{path}")
            open_file_with_default_app(path)
        except Exception as e:
            messagebox.showerror("PDF Export Error", str(e))


# ----------------------------
# Entry function
# ----------------------------

def open_sample_manager(parent, file_path: str | None = None) -> SampleApp:
    """Open the Sample Manager as a child Toplevel of parent.
    If file_path is supplied the file is loaded automatically."""
    win = SampleApp(parent)
    if file_path:
        win.after(100, lambda: win.open_file(file_path))
    return win


if __name__ == "__main__":
    _root = tk.Tk()
    _root.withdraw()
    open_sample_manager(_root)
    _root.mainloop()
