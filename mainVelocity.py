"""
Velocity tracker window.

Toplevel window that visualizes inventory movement patterns across successive
METRC imports. Shows item-level velocity scores, sell rates, slow movers,
and sold-out items. Data comes from the main app's velocity_df and
VelocityHistoryManager snapshots.

HOW IT WORKS:
=============
Each time the user imports a METRC file, main.py saves a "snapshot" of every
item's barcode, room, and qty to velocity_history.json. This window compares
those snapshots over time to answer "which items are selling?" and "which
items are sitting on the shelf?"

TABS:
  - Overview: Dashboard with headline numbers (items tracked, avg sell rate, etc.)
  - Item Detail: Every item with velocity status, sell rate, qty changes
  - Slow Movers: Items flagged Slow/Stale (unchanged qty across imports)
  - Sold Out: Items from history that disappeared from current inventory
  - History: List of all saved snapshots (import timestamps + file names)

ARCHITECTURE:
  - TableView: reusable treeview widget with click-to-sort and row color coding
  - VelocityApp: the Toplevel window that orchestrates everything
  - _velocity_pdf_export(): inline PDF export (does NOT use pdf_common.py yet —
    could be refactored to use it in the future)
"""
# Velocity Tracker — tracks inventory movement across successive imports.
# Room movement, qty deltas, stock age, slow-mover flagging.
#
# Architecture mirrors mainSamples.py:
#   - Subtle/tasteful kawaii UI theme (always on)
#   - PDF export: Normal by default, optional Kawaii PDF toggle
#   - Notebook with 4 tabs (Overview, Item Detail, Slow Movers, History)
#   - Debounced search
#   - Click-to-sort TableView (reused from mainSamples)

import os
import sys
import subprocess
from datetime import datetime, timedelta
from typing import Dict, List, Optional

import tkinter as tk
from tkinter import messagebox, ttk

import pandas as pd

# PDF
from reportlab.lib import colors
from pdf_common import build_section_pdf, PALETTE_KAWAII, PALETTE_PLAIN

from data_core import (
    COLUMNS_TO_USE,
    compute_velocity_metrics,
    compute_slow_movers,
    compute_sold_out,
    VELOCITY_SLOW_THRESHOLD_DEFAULT,
    ellipses,
    truncate_text,
)
from themes import MOVEUP_THEME, apply_theme

APP_TITLE = "Velocity Tracker"

TRUNCATE_PRODUCT_TO = 60

# Friendly column name mapping (internal → display)
_FRIENDLY_NAMES = {
    "velocity_label": "Status",
    "sell_rate": "Sells/Day",
    "qty_delta": "Qty Change",
    "room_changes": "Room Moves",
    "stock_age_days": "Age (Days)",
    "qty_unchanged_streak": "Unchanged Imports",
    "velocity_score": "Score",
    "last_qty": "Last Qty",
    "last_seen": "Last Seen",
}


def _add_friendly_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add human-readable column name aliases alongside the raw velocity columns.

    For each ``(raw_name, friendly_name)`` pair in ``_FRIENDLY_NAMES``, if
    *raw_name* is present in *df* and *friendly_name* is not already a column,
    a copy of the column is inserted under the friendly name.  The original raw
    column is preserved so both names are available for filtering/sorting.

    Returns
    -------
    pd.DataFrame
        The same DataFrame with additional friendly columns added in place.
        The DataFrame is mutated and also returned for chaining.
    """
    for raw, friendly in _FRIENDLY_NAMES.items():
        if raw in df.columns and friendly not in df.columns:
            df[friendly] = df[raw]
    return df


# ----------------------------
# PDF Export
# ----------------------------
def _velocity_style_extras(columns, table_data):
    """
    Return additional ReportLab TableStyle commands for velocity status coloring.

    Searches *columns* for a velocity label column (any name matching
    ``"velocity"``, ``"velocity_label"``, ``"label"``, or ``"status"``).
    For each body row, colors the cell text:

    - ``"Slow"`` / ``"Stale"`` → goldenrod (needs attention)
    - ``"Fast"`` → forest green (healthy)

    Designed to be passed as the ``extra_style_fn`` argument to
    ``build_section_pdf()`` / ``build_table_style()``.

    Parameters
    ----------
    columns : list[str]
        Column names in the same order as *table_data*.
    table_data : list[list]
        Full table data including the header row at index 0.

    Returns
    -------
    list[tuple]
        ReportLab TableStyle command tuples (may be empty if no velocity
        column is found or the data has fewer than 2 rows).
    """
    cmds = []
    # Find the velocity status column
    vel_col_idx = None
    for i, c in enumerate(columns):
        if str(c).lower() in ("velocity", "velocity_label", "label", "status"):
            vel_col_idx = i
            break
    if vel_col_idx is not None and len(table_data) > 1:
        for row_i in range(1, len(table_data)):
            val = str(table_data[row_i][vel_col_idx]).strip()
            if val in ("Slow", "Stale"):
                cmds.append(("TEXTCOLOR", (vel_col_idx, row_i), (vel_col_idx, row_i), colors.Color(0.72, 0.53, 0.04)))
            elif val == "Fast":
                cmds.append(("TEXTCOLOR", (vel_col_idx, row_i), (vel_col_idx, row_i), colors.Color(0.13, 0.55, 0.13)))
    return cmds


def _velocity_pdf_export(
    path: str,
    title: str,
    subtitle: str,
    sections: list,
    kawaii_pdf: bool = False,
):
    """
    Export velocity data tables to a landscape PDF via ``pdf_common``.

    Accepts sections in ``(section_title, df, columns)`` form and converts
    them to the ``(section_title, columns, data_rows)`` form expected by
    ``build_section_pdf()``.  Uses ``PALETTE_KAWAII`` for color output and
    ``PALETTE_PLAIN`` for plain/greyscale output.  Applies
    ``_velocity_style_extras`` for velocity label color-coding.

    Parameters
    ----------
    path : str
        Output file path for the generated PDF.
    title : str
        Document title (rendered as large header on first page).
    subtitle : str
        Subtitle line below the title (typically the generation timestamp).
    sections : list[tuple]
        Each element is ``(section_title, df, columns)`` where *df* is a
        DataFrame and *columns* is the list of column names to include.
    kawaii_pdf : bool
        If ``True``, uses ``PALETTE_KAWAII`` (pink/purple tints); otherwise
        ``PALETTE_PLAIN`` (neutral grey).
    """
    palette = PALETTE_KAWAII if kawaii_pdf else PALETTE_PLAIN

    # Convert (section_title, df, columns) → (section_title, columns, data_rows)
    converted = []
    for section_title, df, columns in sections:
        if df is None or df.empty:
            converted.append((section_title, list(columns), []))
        else:
            out = df[columns].copy().fillna("")
            converted.append((section_title, list(columns), out.values.tolist()))

    build_section_pdf(
        path, title, subtitle, converted,
        palette=palette,
        extra_style_fn=_velocity_style_extras,
    )


# ----------------------------
# TableView (reused from mainSamples)
# ----------------------------
class TableView(ttk.Frame):
    """
    Reusable treeview wrapper with click-to-sort, scroll bars, and row data storage.

    Wraps a ``ttk.Treeview`` with vertical and horizontal scrollbars.  Heading
    clicks toggle ascending/descending sort (numeric-aware: tries
    ``pd.to_numeric`` first, falls back to case-insensitive string sort).

    Six velocity-status color tags are pre-configured:
    ``"normal"``, ``"vel_slow"`` (goldenrod), ``"vel_stale"`` (red),
    ``"vel_fast"`` (green), ``"vel_new"`` (grey), ``"vel_sold_out"``
    (grey italic).

    Row original data is stored in ``self._iid_to_row_data`` keyed by the
    treeview item ID, for use by callers that need to look up the full row on
    selection events.
    """

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

        self._iid_to_row_data: dict = {}   # treeview item ID → original row dict
        self._raw_df: pd.DataFrame | None = None  # full unfiltered DataFrame
        self._columns: list[str] = []    # currently displayed column names
        self._trunc_map: dict[str, int] = {}  # column → max display chars
        self._sort_col: str | None = None  # currently sorted column
        self._sort_asc: bool = True        # sort direction

        # Velocity status color-coding for treeview rows
        self.tree.tag_configure("normal", background=MOVEUP_THEME["tree_bg"])
        self.tree.tag_configure("vel_slow", foreground="#B8860B")    # goldenrod — needs attention
        self.tree.tag_configure("vel_stale", foreground="#cc2222")   # red — stagnant
        self.tree.tag_configure("vel_fast", foreground="#228B22")    # forest green — healthy
        self.tree.tag_configure("vel_new", foreground="#666666")     # grey — not enough data
        self.tree.tag_configure("vel_sold_out", foreground="#999999", font=("TkDefaultFont", 9, "italic"))

    def render(self, df: pd.DataFrame, columns: list[str],
               trunc_map: dict[str, int] | None = None,
               vel_col: str | None = None):
        """
        Populate the treeview with *df* data, displaying only *columns*.

        Clears all existing rows and column definitions before rendering.
        Column width is heuristically sized based on keyword matching in the
        column name (numeric keywords → center-aligned narrower; others →
        left-aligned wider).

        Parameters
        ----------
        df : pd.DataFrame
            Source data.  Must contain all columns listed in *columns*.
        columns : list[str]
            Column names to display (in order).
        trunc_map : dict[str, int] | None
            Maps column name → maximum display character length.  Values are
            truncated with ``truncate_text()`` before rendering.  ``None``
            applies no truncation.
        vel_col : str | None
            Column whose string value determines row color tag
            (``"Slow"`` → ``vel_slow``, ``"Stale"`` → ``vel_stale``,
            ``"Fast"`` → ``vel_fast``, ``"New"`` → ``vel_new``,
            ``"Sold Out"`` → ``vel_sold_out``).  ``None`` → all rows
            tagged ``"normal"``.
        """
        self._iid_to_row_data = {}
        self._raw_df = df
        self._columns = list(columns)
        self._trunc_map = trunc_map or {}

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = columns

        for c in columns:
            indicator = (" \u25b2" if self._sort_asc else " \u25bc") if c == self._sort_col else ""
            self.tree.heading(c, text=c + indicator, command=lambda col=c: self._sort_by_col(col))

            col_name = str(c).lower()
            if any(k in col_name for k in ("qty", "count", "change", "score", "sells", "age", "unchanged")):
                self.tree.column(c, width=max(100, min(520, len(c) * 12)), anchor="center")
            else:
                self.tree.column(c, width=max(120, min(520, len(c) * 12)), anchor="w")

        if df is None or df.empty or not columns:
            return

        out = df[columns].copy()
        for col, lim in self._trunc_map.items():
            if col in out.columns:
                out[col] = out[col].map(lambda v, _lim=lim: truncate_text(v, _lim))
        out = out.fillna("")

        for idx, row in out.iterrows():
            values = [row[c] for c in columns]

            tag = "normal"
            if vel_col and vel_col in df.columns:
                try:
                    label = str(df.loc[idx, vel_col]).strip()
                    if label == "Slow":
                        tag = "vel_slow"
                    elif label == "Stale":
                        tag = "vel_stale"
                    elif label == "Fast":
                        tag = "vel_fast"
                    elif label == "New":
                        tag = "vel_new"
                    elif label == "Sold Out":
                        tag = "vel_sold_out"
                except (KeyError, ValueError):
                    pass

            iid = self.tree.insert("", "end", values=values, tags=(tag,))
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
        numeric = pd.to_numeric(series, errors="coerce")
        if numeric.notna().any():
            df["_sk"] = numeric
        else:
            df["_sk"] = series.astype(str).str.lower()

        df = df.sort_values("_sk", ascending=self._sort_asc, na_position="last").drop(columns=["_sk"])
        self.render(df, self._columns, self._trunc_map, vel_col="Status")


# ----------------------------
# VelocityApp
# ----------------------------
class VelocityApp(tk.Toplevel):
    """
    Toplevel window for inventory velocity tracking.

    Displays per-item movement metrics computed by ``compute_velocity_metrics()``
    across successive METRC imports.  Receives pre-computed DataFrames from the
    main app rather than running the pipeline itself.

    Five tabs:
    - **Overview**: dashboard with headline numbers (items tracked, avg sell
      rate, fast/slow/stale/new/sold-out counts, snapshot count and date range).
    - **Item Detail**: full per-item velocity table with live search bar and
      velocity label color coding.
    - **Slow Movers**: items flagged Slow/Stale by ``compute_slow_movers()``,
      with a configurable "unchanged imports" threshold spinbox.
    - **Sold Out**: items from velocity history that no longer appear in the
      current import, via ``compute_sold_out()``.
    - **History**: list of all saved velocity snapshots with purge controls.

    PDF export uses ``_velocity_pdf_export()`` with kawaii/plain toggle.
    The search bar uses ``_debounce()`` to avoid re-filtering on every keystroke.
    """

    def __init__(self, master, current_df, velocity_df, velocity_mgr):
        super().__init__(master)
        self.title(APP_TITLE)
        self.geometry("1340x820")

        self.current_df = current_df
        self.velocity_df = velocity_df
        self.velocity_mgr = velocity_mgr

        self._debounce_job: str | None = None
        self.pdf_kawaii_var = tk.BooleanVar(value=True)
        self.search_var = tk.StringVar(value="")
        self.info_var = tk.StringVar(value="")
        self.threshold_var = tk.IntVar(value=VELOCITY_SLOW_THRESHOLD_DEFAULT)

        # Merged detail data
        self._detail_df: pd.DataFrame | None = None
        self._slow_df: pd.DataFrame | None = None
        self._sold_out_df: pd.DataFrame | None = None

        self._build_ui()
        self.apply_subtle_theme()
        self._refresh_all()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        if self._debounce_job is not None:
            try:
                self.after_cancel(self._debounce_job)
            except (tk.TclError, ValueError):
                pass
            self._debounce_job = None
        self.destroy()

    def apply_subtle_theme(self):
        apply_theme(self, "velocity_theme")

    def _build_ui(self):
        # ---------- Top bar ----------
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Button(top, text="Export PDF", command=self._export_pdf).pack(side="left")
        ttk.Checkbutton(
            top, text="Kawaii PDF", variable=self.pdf_kawaii_var,
        ).pack(side="left", padx=(10, 0))

        ttk.Label(top, textvariable=self.info_var).pack(side="right")

        # ---------- Search ----------
        search_frm = ttk.Frame(self, padding=(10, 0, 10, 6))
        search_frm.pack(fill="x")
        ttk.Label(search_frm, text="Search:").pack(side="left")
        search_entry = ttk.Entry(search_frm, textvariable=self.search_var, width=40)
        search_entry.pack(side="left", padx=(6, 0))
        self.search_var.trace_add("write", lambda *_: self._debounce(300, self._apply_search))

        # ---------- Notebook ----------
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Tab 1: Overview
        self.frm_overview = ttk.Frame(self.notebook)
        self.notebook.add(self.frm_overview, text="Overview")
        self._build_overview_tab()

        # Tab 2: Item Detail
        self.frm_detail = ttk.Frame(self.notebook)
        self.notebook.add(self.frm_detail, text="Item Detail")
        self.detail_table = TableView(self.frm_detail)
        self.detail_table.pack(fill="both", expand=True)

        # Tab 3: Slow Movers
        self.frm_slow = ttk.Frame(self.notebook)
        self.notebook.add(self.frm_slow, text="Slow Movers")
        self._build_slow_tab()

        # Tab 4: Sold Out
        self.frm_sold_out = ttk.Frame(self.notebook)
        self.notebook.add(self.frm_sold_out, text="Sold Out")
        self.sold_out_table = TableView(self.frm_sold_out)
        self.sold_out_table.pack(fill="both", expand=True)

        # Tab 5: History
        self.frm_history = ttk.Frame(self.notebook)
        self.notebook.add(self.frm_history, text="History")
        self._build_history_tab()

    def _build_overview_tab(self):
        container = ttk.Frame(self.frm_overview, padding=20)
        container.pack(fill="both", expand=True)

        self._overview_labels: Dict[str, tk.StringVar] = {}

        # --- Big numbers row ---
        big_row = ttk.Frame(container)
        big_row.pack(fill="x", pady=(0, 20))

        big_metrics = [
            ("total_tracked", "Items Tracked"),
            ("avg_sell_rate", "Avg Sells/Day"),
            ("sold_out_count", "Sold Out"),
            ("slow_count", "Need Attention"),
        ]

        for i, (key, label_text) in enumerate(big_metrics):
            var = tk.StringVar(value="—")
            self._overview_labels[key] = var

            card = ttk.Frame(big_row, padding=12)
            card.pack(side="left", fill="both", expand=True, padx=(0 if i == 0 else 8, 0))

            ttk.Label(card, textvariable=var, font=("TkDefaultFont", 24, "bold"),
                      foreground=MOVEUP_THEME["label_fg"]).pack(anchor="center")
            ttk.Label(card, text=label_text, font=("TkDefaultFont", 10),
                      foreground="#666666").pack(anchor="center")

        # --- Details row ---
        detail_row = ttk.Frame(container)
        detail_row.pack(fill="x", pady=(0, 10))

        detail_metrics = [
            ("fast_count", "Fast"),
            ("moderate_count", "Moderate"),
            ("stale_count", "Stale"),
            ("new_count", "New"),
        ]

        for i, (key, label_text) in enumerate(detail_metrics):
            var = tk.StringVar(value="—")
            self._overview_labels[key] = var

            frm = ttk.Frame(detail_row)
            frm.pack(side="left", padx=(0 if i == 0 else 16, 0))

            ttk.Label(frm, textvariable=var, font=("TkDefaultFont", 14, "bold")).pack(side="left")
            ttk.Label(frm, text=f"  {label_text}", font=("TkDefaultFont", 10)).pack(side="left")

        # --- History info ---
        history_row = ttk.Frame(container)
        history_row.pack(fill="x", pady=(10, 0))

        for key, label_text in [("snapshot_count", "Snapshots"), ("date_range", "Date Range")]:
            var = tk.StringVar(value="—")
            self._overview_labels[key] = var

            frm = ttk.Frame(history_row)
            frm.pack(side="left", padx=(0, 20))

            ttk.Label(frm, text=f"{label_text}:", font=("TkDefaultFont", 9, "bold")).pack(side="left")
            ttk.Label(frm, textvariable=var, font=("TkDefaultFont", 9)).pack(side="left", padx=(4, 0))

    def _build_slow_tab(self):
        ctrl = ttk.Frame(self.frm_slow, padding=(10, 6))
        ctrl.pack(fill="x")

        ttk.Label(ctrl, text="Slow threshold (imports unchanged):").pack(side="left")
        spin = ttk.Spinbox(
            ctrl, from_=1, to=50, width=5, textvariable=self.threshold_var,
        )
        spin.pack(side="left", padx=(6, 10))
        ttk.Button(ctrl, text="Recompute", command=self._recompute_slow).pack(side="left")

        self.slow_table = TableView(self.frm_slow)
        self.slow_table.pack(fill="both", expand=True)

    def _build_history_tab(self):
        ctrl = ttk.Frame(self.frm_history, padding=(10, 6))
        ctrl.pack(fill="x")

        ttk.Button(ctrl, text="Purge All History", command=self._purge_all).pack(side="left")
        ttk.Button(ctrl, text="Purge Older Than 30 Days", command=self._purge_old).pack(side="left", padx=(10, 0))

        self.history_table = TableView(self.frm_history)
        self.history_table.pack(fill="both", expand=True)

    # ---------- Refresh ----------

    def _refresh_all(self):
        self._refresh_overview()
        self._refresh_detail()
        self._refresh_slow()
        self._refresh_sold_out()
        self._refresh_history()

    def _refresh_overview(self):
        labels = self._overview_labels
        vel = self.velocity_df
        snaps = self.velocity_mgr.get_snapshots()

        labels["snapshot_count"].set(str(len(snaps)))

        if snaps:
            timestamps = [s.get("timestamp", "") for s in snaps]
            timestamps = [t for t in timestamps if t]
            if timestamps:
                labels["date_range"].set(f"{timestamps[0][:10]}  to  {timestamps[-1][:10]}")
            else:
                labels["date_range"].set("—")
        else:
            labels["date_range"].set("No history yet")

        if vel is None or vel.empty:
            for k in ("total_tracked", "avg_sell_rate", "fast_count", "moderate_count",
                       "slow_count", "stale_count", "new_count", "sold_out_count"):
                labels[k].set("0" if k != "avg_sell_rate" else "—")
            self.info_var.set("No velocity data. Import inventory files to build history.")
            return

        labels["total_tracked"].set(str(len(vel)))

        # Avg sell rate (exclude sold-out and new for meaningful average)
        active = vel[~vel["velocity_label"].isin(["New", "Sold Out"])]
        if not active.empty and "sell_rate" in active.columns:
            labels["avg_sell_rate"].set(f"{active['sell_rate'].mean():.2f}")
        else:
            labels["avg_sell_rate"].set("—")

        counts = vel["velocity_label"].value_counts()
        labels["fast_count"].set(str(counts.get("Fast", 0)))
        labels["moderate_count"].set(str(counts.get("Moderate", 0)))
        slow_n = counts.get("Slow", 0) + counts.get("Stale", 0)
        labels["slow_count"].set(str(slow_n))
        labels["stale_count"].set(str(counts.get("Stale", 0)))
        labels["new_count"].set(str(counts.get("New", 0)))
        labels["sold_out_count"].set(str(counts.get("Sold Out", 0)))

        sold_n = counts.get("Sold Out", 0)
        self.info_var.set(
            f"{len(vel)} items | {sold_n} sold out | {slow_n} slow/stale"
        )

    def _refresh_detail(self):
        vel = self.velocity_df
        cur = self.current_df

        if vel is None or vel.empty or cur is None or cur.empty:
            self._detail_df = None
            self.detail_table.render(pd.DataFrame(), [])
            return

        # Merge velocity with product info
        merged = vel.merge(
            cur.drop_duplicates(subset=["Package Barcode"]),
            on="Package Barcode", how="left",
        )

        # Friendly column renames
        merged = _add_friendly_columns(merged)

        detail_cols = []
        for c in ["Product Name", "Room", "Qty On Hand"]:
            if c in merged.columns:
                detail_cols.append(c)
        detail_cols += ["Status", "Sells/Day", "Qty Change"]

        self._detail_df = merged
        trunc = {"Product Name": TRUNCATE_PRODUCT_TO}
        self.detail_table.render(merged, detail_cols, trunc, vel_col="Status")

    def _refresh_slow(self):
        vel = self.velocity_df
        cur = self.current_df

        if vel is None or vel.empty:
            self._slow_df = None
            self.slow_table.render(pd.DataFrame(), [])
            return

        slow = compute_slow_movers(vel, cur)
        slow = _add_friendly_columns(slow)
        self._slow_df = slow

        if slow.empty:
            self.slow_table.render(pd.DataFrame(), [])
            return

        cols = []
        for c in ["Product Name", "Room", "Qty On Hand"]:
            if c in slow.columns:
                cols.append(c)
        cols += ["Status", "Sells/Day", "Unchanged Imports", "Age (Days)"]

        trunc = {"Product Name": TRUNCATE_PRODUCT_TO}
        self.slow_table.render(slow, cols, trunc, vel_col="Status")

    def _refresh_sold_out(self):
        vel = self.velocity_df
        if vel is None or vel.empty:
            self._sold_out_df = None
            self.sold_out_table.render(pd.DataFrame(), [])
            return

        sold = compute_sold_out(vel, self.velocity_mgr.get_snapshots())
        sold = _add_friendly_columns(sold)
        self._sold_out_df = sold

        if sold.empty:
            self.sold_out_table.render(pd.DataFrame(), [])
            return

        cols = ["Package Barcode"]
        if "Room" in sold.columns:
            cols.append("Room")
        cols += ["Last Qty", "Last Seen", "Sells/Day"]

        self.sold_out_table.render(sold, cols, vel_col="Status")

    def _refresh_history(self):
        snaps = self.velocity_mgr.get_snapshots()
        if not snaps:
            self.history_table.render(pd.DataFrame(), [])
            return

        rows = []
        for s in snaps:
            rows.append({
                "Timestamp": s.get("timestamp", "")[:19],
                "File": s.get("file_name", ""),
                "Items": len(s.get("entries", [])),
            })

        df = pd.DataFrame(rows)
        self.history_table.render(df, ["Timestamp", "File", "Items"])

    # ---------- Actions ----------

    def _recompute_slow(self):
        """Recompute velocity with user-defined threshold."""
        try:
            thresh = self.threshold_var.get()
        except (tk.TclError, ValueError):
            thresh = VELOCITY_SLOW_THRESHOLD_DEFAULT

        if self.current_df is None:
            return

        self.velocity_df = compute_velocity_metrics(
            self.current_df,
            self.velocity_mgr.get_snapshots(),
            slow_threshold=thresh,
        )
        self._refresh_all()

    def _apply_search(self):
        """Filter detail table by search term."""
        term = self.search_var.get().strip().lower()
        if not term or self._detail_df is None or self._detail_df.empty:
            self._refresh_detail()
            return

        mask = self._detail_df.apply(
            lambda row: any(term in str(v).lower() for v in row), axis=1,
        )
        filtered = self._detail_df[mask]

        detail_cols = []
        for c in ["Product Name", "Room", "Qty On Hand"]:
            if c in filtered.columns:
                detail_cols.append(c)
        detail_cols += ["Status", "Sells/Day", "Qty Change"]

        trunc = {"Product Name": TRUNCATE_PRODUCT_TO}
        self.detail_table.render(filtered, detail_cols, trunc, vel_col="Status")

    def _debounce(self, ms: int, func):
        """
        Schedule *func* to run after *ms* milliseconds, cancelling any pending call.

        Used to rate-limit the search bar so filtering only runs once the user
        pauses typing.  If called again before the previous timer fires, the
        previous ``after()`` job is cancelled and a new one is scheduled.

        Parameters
        ----------
        ms : int
            Delay in milliseconds.
        func : callable
            Zero-argument callable to invoke after the delay.
        """
        if self._debounce_job is not None:
            self.after_cancel(self._debounce_job)

        def _run():
            self._debounce_job = None
            func()
        self._debounce_job = self.after(ms, _run)

    def _purge_all(self):
        if not messagebox.askyesno(
            "Purge All History",
            "This will delete ALL velocity snapshots.\n\nAre you sure?",
            parent=self,
        ):
            return
        removed = self.velocity_mgr.purge_all()
        self.velocity_df = None
        self._refresh_all()
        messagebox.showinfo("Purged", f"Removed {removed} snapshots.", parent=self)

    def _purge_old(self):
        cutoff = (datetime.now() - timedelta(days=30)).isoformat()
        removed = self.velocity_mgr.purge_before(cutoff)
        if removed:
            # Recompute velocity with remaining history
            if self.current_df is not None:
                self.velocity_df = compute_velocity_metrics(
                    self.current_df,
                    self.velocity_mgr.get_snapshots(),
                )
            self._refresh_all()
        messagebox.showinfo(
            "Purged", f"Removed {removed} snapshots older than 30 days.", parent=self,
        )

    def _export_pdf(self):
        """Export velocity report PDF."""
        if self.velocity_df is None or self.velocity_df.empty:
            messagebox.showwarning("No Data", "No velocity data to export.", parent=self)
            return

        # Determine output dir
        app_dir = self.velocity_mgr.app_dir
        out_dir = os.path.join(app_dir, "generated", f"velocity_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        os.makedirs(out_dir, exist_ok=True)

        path = os.path.join(out_dir, "velocity_report.pdf")
        kawaii = self.pdf_kawaii_var.get()

        sections = []

        # Section 1: Overview summary as small table
        overview_data = []
        labels = self._overview_labels
        for key, var in labels.items():
            overview_data.append({"Metric": key.replace("_", " ").title(), "Value": var.get()})
        overview_df = pd.DataFrame(overview_data)
        sections.append(("Overview", overview_df, ["Metric", "Value"]))

        # Section 2: Slow movers
        if self._slow_df is not None and not self._slow_df.empty:
            slow_cols = []
            for c in ["Product Name", "Room", "Qty On Hand"]:
                if c in self._slow_df.columns:
                    slow_cols.append(c)
            slow_cols += ["Status", "Sells/Day", "Unchanged Imports", "Age (Days)"]
            slow_cols = [c for c in slow_cols if c in self._slow_df.columns]
            sections.append(("Slow Movers", self._slow_df, slow_cols))

        # Section 3: Full detail
        if self._detail_df is not None and not self._detail_df.empty:
            det_cols = []
            for c in ["Product Name", "Room", "Qty On Hand"]:
                if c in self._detail_df.columns:
                    det_cols.append(c)
            det_cols += ["Status", "Sells/Day", "Qty Change"]
            det_cols = [c for c in det_cols if c in self._detail_df.columns]
            sections.append(("All Items", self._detail_df, det_cols))

        try:
            _velocity_pdf_export(
                path,
                "Velocity Report",
                f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                sections,
                kawaii_pdf=kawaii,
            )

            # Auto-open on Windows
            if os.name == "nt":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])

        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export PDF:\n\n{e}", parent=self)


# ----------------------------
# Public entry point
# ----------------------------
def open_velocity_window(
    parent,
    current_df: pd.DataFrame | None = None,
    velocity_df: pd.DataFrame | None = None,
    velocity_mgr=None,
) -> VelocityApp:
    """
    Open the Velocity Tracker window as a child of *parent*.

    Parameters
    ----------
    parent : tk.Tk | tk.Toplevel
        Parent window.
    current_df : pd.DataFrame | None
        The currently loaded METRC DataFrame (used for product name/room joins
        in the Item Detail and Slow Movers tabs).  ``None`` renders empty tabs.
    velocity_df : pd.DataFrame | None
        Pre-computed velocity DataFrame from ``compute_velocity_metrics()``.
        ``None`` renders the Overview with all-zero counts.
    velocity_mgr : VelocityHistoryManager | None
        The live manager instance; used by the Sold Out tab and History tab,
        and for purge actions.

    Returns
    -------
    VelocityApp
        The newly created window instance.
    """
    win = VelocityApp(parent, current_df, velocity_df, velocity_mgr)
    return win


if __name__ == "__main__":
    _root = tk.Tk()
    _root.withdraw()
    from velocity_history import VelocityHistoryManager
    mgr = VelocityHistoryManager()
    mgr.load()
    open_velocity_window(_root, None, None, mgr)
    _root.mainloop()
