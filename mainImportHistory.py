"""
Import History window.

Toplevel window that displays a timeline of all METRC imports with
aggregate statistics, trend sparklines, and snapshot comparison.
Data comes from ImportHistoryManager (aggregate stats) and
VelocityHistoryManager (per-item snapshots for comparison).

TABS:
  - Timeline: Scrollable table of all imports with stats columns
  - Trends: Canvas-based sparklines for key metrics over time
  - Compare: Select two imports and diff their inventory snapshots
"""

import tkinter as tk
from datetime import datetime, timedelta
from tkinter import messagebox, ttk
from typing import Any, Dict, List, Optional

import pandas as pd

from import_history import ImportHistoryManager
from velocity_history import VelocityHistoryManager
from themes import MOVEUP_THEME, apply_theme

APP_TITLE = "Import History"


# ------------------------------------------------------------------ #
# TableView (same pattern as mainVelocity.py)
# ------------------------------------------------------------------ #
class _TableView(ttk.Frame):
    """
    Lightweight treeview wrapper with click-to-sort and color-tagged rows.

    Wraps a ``ttk.Treeview`` with vertical and horizontal scrollbars, a
    click-sortable heading that toggles ascending/descending, and three
    color tags: ``"normal"`` (default), ``"positive"`` (green, for positive
    deltas), and ``"negative"`` (red, for negative deltas).

    Used for the Timeline, and all four Compare sub-tabs.
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

        self._raw_df: Optional[pd.DataFrame] = None
        self._columns: List[str] = []
        self._sort_col: Optional[str] = None
        self._sort_asc: bool = True

        # Color tags for comparison results
        self.tree.tag_configure("normal", background=MOVEUP_THEME["tree_bg"])
        self.tree.tag_configure("positive", foreground="#228B22")   # green — gained
        self.tree.tag_configure("negative", foreground="#cc2222")   # red — lost

    def render(self, df: pd.DataFrame, columns: List[str], tag_col: Optional[str] = None):
        """
        Populate the treeview with *df* data, displaying only *columns*.

        Resets all existing rows and column definitions before rendering.
        Column widths are heuristically sized based on column name keywords
        (numeric keywords → center-aligned narrower; others → left-aligned wider).

        Parameters
        ----------
        df : pd.DataFrame
            Source data.  Only the columns in *columns* are displayed.
            Empty or ``None`` DataFrames produce an empty treeview.
        columns : list[str]
            Ordered list of column names to display (must be present in *df*).
        tag_col : str | None
            Optional column in *df* whose numeric value determines row coloring.
            Positive values → ``"positive"`` tag (green); negative → ``"negative"``
            tag (red); zero / non-numeric → ``"normal"`` tag.
        """
        self._raw_df = df
        self._columns = list(columns)

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = columns

        for c in columns:
            indicator = (" \u25b2" if self._sort_asc else " \u25bc") if c == self._sort_col else ""
            self.tree.heading(c, text=c + indicator, command=lambda col=c: self._sort_by_col(col))
            col_lower = str(c).lower()
            if any(k in col_lower for k in ("qty", "count", "rows", "brands", "types", "moves", "delta", "move-up")):
                self.tree.column(c, width=max(90, min(400, len(c) * 11)), anchor="center")
            else:
                self.tree.column(c, width=max(120, min(400, len(c) * 11)), anchor="w")

        if df is None or df.empty or not columns:
            return

        out = df[columns].copy().fillna("")
        for idx, row in out.iterrows():
            values = [row[c] for c in columns]
            tag = "normal"
            if tag_col and tag_col in df.columns:
                try:
                    val = df.loc[idx, tag_col]
                    if isinstance(val, (int, float)):
                        if val > 0:
                            tag = "positive"
                        elif val < 0:
                            tag = "negative"
                except (KeyError, ValueError):
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
        numeric = pd.to_numeric(series, errors="coerce")
        if numeric.notna().any():
            df["_sk"] = numeric
        else:
            df["_sk"] = series.astype(str).str.lower()

        df = df.sort_values("_sk", ascending=self._sort_asc, na_position="last").drop(columns=["_sk"])
        self.render(df, self._columns)


# ------------------------------------------------------------------ #
# Sparkline metrics config
# ------------------------------------------------------------------ #
_SPARKLINE_METRICS = [
    ("total_rows", "Total Rows"),
    ("moveup_count", "Move-Up Count"),
    ("unique_brands", "Unique Brands"),
    ("unique_types", "Unique Types"),
    ("total_qty", "Total Qty"),
    ("candidate_pool", "Candidate Pool"),
    ("removed_as_on_sf", "Already on Sales Floor"),
    ("bisa_moveups", "Bisa Moves"),
]


# ------------------------------------------------------------------ #
# ImportHistoryWindow
# ------------------------------------------------------------------ #
class ImportHistoryWindow(tk.Toplevel):
    """
    Toplevel window displaying a complete timeline of METRC file imports.

    Three tabs:
    - **Timeline**: scrollable table of all recorded imports with aggregate
      stats (total rows, move-up count, brands, types, qty, Bisa moves).
      Supports purging old entries (all / >30 days).
    - **Trends**: canvas-based sparkline charts for eight key metrics over
      time, arranged in a 2-column scrollable grid.
    - **Compare**: select any two imports from dropdowns and diff their
      velocity snapshots to see new items, removed items, qty changes, and
      room moves.  Requires velocity history to be present for both imports.

    Data sources:
    - ``ImportHistoryManager`` — aggregate stats for the Timeline and Trends tabs.
    - ``VelocityHistoryManager`` — per-barcode snapshots for the Compare tab.
    """

    def __init__(self, master, import_history_mgr: ImportHistoryManager,
                 velocity_mgr: VelocityHistoryManager):
        super().__init__(master)
        self.title(APP_TITLE)
        self.geometry("1200x780")

        self.history_mgr = import_history_mgr
        self.velocity_mgr = velocity_mgr

        self._sparkline_canvases: List[Dict[str, Any]] = []  # [{canvas, label_var, key}, ...]
        self._sparkline_values: Dict[str, List[float]] = {}

        self._build_ui()
        apply_theme(self, "import_history_theme")
        self._refresh_all()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        self.destroy()

    # ================================================================
    # UI Build
    # ================================================================
    def _build_ui(self):
        # Info bar
        self.info_var = tk.StringVar(value="")
        ttk.Label(self, textvariable=self.info_var, padding=(10, 6)).pack(fill="x")

        # Notebook
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        self.frm_timeline = ttk.Frame(self.notebook)
        self.frm_trends = ttk.Frame(self.notebook)
        self.frm_compare = ttk.Frame(self.notebook)

        self.notebook.add(self.frm_timeline, text="  Timeline  ")
        self.notebook.add(self.frm_trends, text="  Trends  ")
        self.notebook.add(self.frm_compare, text="  Compare  ")

        self._build_timeline_tab()
        self._build_trends_tab()
        self._build_compare_tab()

    # ----------------------------------------------------------------
    # Tab 1: Timeline
    # ----------------------------------------------------------------
    def _build_timeline_tab(self):
        ctrl = ttk.Frame(self.frm_timeline, padding=(10, 6))
        ctrl.pack(fill="x")

        ttk.Button(ctrl, text="Purge All", command=self._purge_all).pack(side="left")
        ttk.Button(ctrl, text="Purge > 30 Days", command=self._purge_old).pack(side="left", padx=(10, 0))

        self.timeline_count_var = tk.StringVar(value="")
        ttk.Label(ctrl, textvariable=self.timeline_count_var).pack(side="right")

        self.timeline_table = _TableView(self.frm_timeline)
        self.timeline_table.pack(fill="both", expand=True)

    def _refresh_timeline(self):
        entries = self.history_mgr.get_entries()
        if not entries:
            self.timeline_table.render(pd.DataFrame(), [])
            self.timeline_count_var.set("No imports recorded yet.")
            self.info_var.set("No import history. Import a METRC file to start tracking.")
            return

        rows = []
        for e in entries:
            rows.append({
                "Timestamp": e.get("timestamp", "")[:16].replace("T", " "),
                "File": e.get("file_name", ""),
                "Total Rows": e.get("total_rows", 0),
                "Move-Up": e.get("moveup_count", 0),
                "Brands": e.get("unique_brands", 0),
                "Types": e.get("unique_types", 0),
                "Total Qty": e.get("total_qty", 0),
                "Bisa Moves": e.get("bisa_moveups", 0),
            })

        df = pd.DataFrame(rows)
        cols = ["Timestamp", "File", "Total Rows", "Move-Up", "Brands", "Types", "Total Qty", "Bisa Moves"]
        self.timeline_table.render(df, cols)

        first_ts = entries[0].get("timestamp", "")[:10]
        last_ts = entries[-1].get("timestamp", "")[:10]
        self.timeline_count_var.set(f"{len(entries)} imports")
        self.info_var.set(f"{len(entries)} imports recorded  |  {first_ts} to {last_ts}")

    # ----------------------------------------------------------------
    # Tab 2: Trends (sparklines)
    # ----------------------------------------------------------------
    def _build_trends_tab(self):
        # Scrollable frame
        canvas_outer = tk.Canvas(self.frm_trends, highlightthickness=0,
                                 bg=MOVEUP_THEME["bg"])
        vsb = ttk.Scrollbar(self.frm_trends, orient="vertical", command=canvas_outer.yview)
        canvas_outer.configure(yscrollcommand=vsb.set)

        vsb.pack(side="right", fill="y")
        canvas_outer.pack(side="left", fill="both", expand=True)

        self._trends_inner = ttk.Frame(canvas_outer)
        canvas_outer.create_window((0, 0), window=self._trends_inner, anchor="nw")
        self._trends_inner.bind("<Configure>",
                                lambda e: canvas_outer.configure(scrollregion=canvas_outer.bbox("all")))

        # Build sparkline cards in 2-column grid
        self._sparkline_canvases = []
        for i, (key, label) in enumerate(_SPARKLINE_METRICS):
            row_idx = i // 2
            col_idx = i % 2

            card = ttk.Frame(self._trends_inner, padding=8)
            card.grid(row=row_idx, column=col_idx, padx=10, pady=6, sticky="ew")

            label_var = tk.StringVar(value=f"{label}: —")
            ttk.Label(card, textvariable=label_var,
                      font=("TkDefaultFont", 10, "bold")).pack(anchor="w")

            spark_canvas = tk.Canvas(
                card, width=280, height=50,
                bg=MOVEUP_THEME["tree_bg"],
                highlightthickness=0,
            )
            spark_canvas.pack(fill="x", expand=True, pady=(4, 0))

            self._sparkline_canvases.append({
                "canvas": spark_canvas,
                "label_var": label_var,
                "key": key,
                "label": label,
            })

            # Redraw on resize
            spark_canvas.bind("<Configure>",
                              lambda e, k=key: self._redraw_sparkline(k))

        # Make both columns expand equally
        self._trends_inner.columnconfigure(0, weight=1)
        self._trends_inner.columnconfigure(1, weight=1)

    def _refresh_trends(self):
        entries = self.history_mgr.get_entries()

        # Extract metric values from entries
        self._sparkline_values = {}
        for key, label in _SPARKLINE_METRICS:
            values = []
            for e in entries:
                if key in ("candidate_pool", "removed_as_on_sf"):
                    val = e.get("diag", {}).get(key, 0)
                else:
                    val = e.get(key, 0)
                try:
                    values.append(float(val))
                except (TypeError, ValueError):
                    values.append(0.0)
            self._sparkline_values[key] = values

        # Update labels and draw sparklines
        for item in self._sparkline_canvases:
            key = item["key"]
            label = item["label"]
            values = self._sparkline_values.get(key, [])
            latest = int(values[-1]) if values else 0
            item["label_var"].set(f"{label}: {latest}")
            self._draw_sparkline(item["canvas"], values)

    def _redraw_sparkline(self, key: str):
        values = self._sparkline_values.get(key, [])
        for item in self._sparkline_canvases:
            if item["key"] == key:
                self._draw_sparkline(item["canvas"], values)
                break

    def _draw_sparkline(self, canvas: tk.Canvas, values: List[float],
                        color: str = "#7251A8"):
        canvas.delete("all")
        w = canvas.winfo_width() or 280
        h = canvas.winfo_height() or 50
        pad = 4

        if not values or len(values) < 2:
            canvas.create_text(
                w // 2, h // 2, text="Not enough data",
                fill="#999999", font=("TkDefaultFont", 8),
            )
            return

        min_v = min(values)
        max_v = max(values)
        val_range = max_v - min_v or 1

        n = len(values)
        x_step = (w - 2 * pad) / max(n - 1, 1)

        points = []
        for i, v in enumerate(values):
            x = pad + i * x_step
            y = h - pad - ((v - min_v) / val_range) * (h - 2 * pad)
            points.append((x, y))

        # Fill area under line (subtle)
        fill_points = list(points) + [(points[-1][0], h - pad), (points[0][0], h - pad)]
        flat = [coord for pt in fill_points for coord in pt]
        canvas.create_polygon(flat, fill="#DED9F0", outline="", stipple="gray25")

        # Line
        for i in range(len(points) - 1):
            canvas.create_line(
                points[i][0], points[i][1],
                points[i + 1][0], points[i + 1][1],
                fill=color, width=2, smooth=True,
            )

        # Dots
        for x, y in points:
            r = 3
            canvas.create_oval(x - r, y - r, x + r, y + r, fill=color, outline="")

        # Min/max labels
        canvas.create_text(pad + 2, h - pad - 2, text=str(int(min_v)),
                           anchor="sw", fill="#999", font=("TkDefaultFont", 7))
        canvas.create_text(w - pad - 2, pad + 2, text=str(int(max_v)),
                           anchor="ne", fill="#999", font=("TkDefaultFont", 7))

    # ----------------------------------------------------------------
    # Tab 3: Compare
    # ----------------------------------------------------------------
    def _build_compare_tab(self):
        # Top controls
        ctrl = ttk.Frame(self.frm_compare, padding=(10, 8))
        ctrl.pack(fill="x")

        ttk.Label(ctrl, text="Import A:").pack(side="left")
        self.combo_a = ttk.Combobox(ctrl, state="readonly", width=50)
        self.combo_a.pack(side="left", padx=(4, 16))

        ttk.Label(ctrl, text="Import B:").pack(side="left")
        self.combo_b = ttk.Combobox(ctrl, state="readonly", width=50)
        self.combo_b.pack(side="left", padx=(4, 16))

        ttk.Button(ctrl, text="Compare", command=self._do_compare).pack(side="left")

        # Summary bar
        self._summary_frame = ttk.Frame(self.frm_compare, padding=(10, 4))
        self._summary_frame.pack(fill="x")

        self._summary_vars: Dict[str, tk.StringVar] = {}
        for key, label in [("new", "New Items"), ("removed", "Removed"),
                           ("qty_changed", "Qty Changed"), ("room_moved", "Room Moved"),
                           ("net_qty", "Net Qty")]:
            var = tk.StringVar(value="—")
            self._summary_vars[key] = var
            frm = ttk.Frame(self._summary_frame)
            frm.pack(side="left", padx=(0, 20))
            ttk.Label(frm, text=f"{label}:", font=("TkDefaultFont", 9, "bold")).pack(side="left")
            ttk.Label(frm, textvariable=var, font=("TkDefaultFont", 9)).pack(side="left", padx=(4, 0))

        # Sub-notebook for comparison results
        self._compare_nb = ttk.Notebook(self.frm_compare)
        self._compare_nb.pack(fill="both", expand=True, padx=6, pady=(4, 6))

        self.frm_new = ttk.Frame(self._compare_nb)
        self.frm_removed = ttk.Frame(self._compare_nb)
        self.frm_qty = ttk.Frame(self._compare_nb)
        self.frm_room = ttk.Frame(self._compare_nb)

        self._compare_nb.add(self.frm_new, text="  New Items  ")
        self._compare_nb.add(self.frm_removed, text="  Removed  ")
        self._compare_nb.add(self.frm_qty, text="  Qty Changes  ")
        self._compare_nb.add(self.frm_room, text="  Room Moves  ")

        self.new_table = _TableView(self.frm_new)
        self.new_table.pack(fill="both", expand=True)

        self.removed_table = _TableView(self.frm_removed)
        self.removed_table.pack(fill="both", expand=True)

        self.qty_table = _TableView(self.frm_qty)
        self.qty_table.pack(fill="both", expand=True)

        self.room_table = _TableView(self.frm_room)
        self.room_table.pack(fill="both", expand=True)

    def _populate_compare_dropdowns(self):
        entries = self.history_mgr.get_entries()
        items = []
        self._compare_ts_map: Dict[str, str] = {}  # display string → raw timestamp

        for e in entries:
            ts_raw = e.get("timestamp", "")
            ts_display = ts_raw[:16].replace("T", " ")
            fname = e.get("file_name", "")
            display = f"{ts_display}  —  {fname}"
            items.append(display)
            self._compare_ts_map[display] = ts_raw

        self.combo_a["values"] = items
        self.combo_b["values"] = items

        if len(items) >= 2:
            self.combo_a.current(len(items) - 2)
            self.combo_b.current(len(items) - 1)
        elif len(items) == 1:
            self.combo_a.current(0)
            self.combo_b.current(0)

    def _do_compare(self):
        sel_a = self.combo_a.get()
        sel_b = self.combo_b.get()

        if not sel_a or not sel_b:
            messagebox.showwarning("Compare", "Select two imports to compare.", parent=self)
            return

        ts_a = self._compare_ts_map.get(sel_a, "")
        ts_b = self._compare_ts_map.get(sel_b, "")

        if not ts_a or not ts_b:
            messagebox.showwarning("Compare", "Could not resolve timestamps.", parent=self)
            return

        result = self._compute_comparison(ts_a, ts_b)
        if result is None:
            return
        self._render_comparison(result)

    def _compute_comparison(self, ts_a: str, ts_b: str) -> Optional[Dict]:
        """
        Diff two velocity snapshots identified by their ISO-8601 timestamps.

        Looks up both snapshots in ``self.velocity_mgr`` by exact timestamp
        match.  Computes four change categories by barcode:

        - **new_items**: barcodes present in B but not in A.
        - **removed_items**: barcodes present in A but not in B.
        - **qty_changes**: barcodes in both with different ``qty`` values;
          includes a signed ``Delta`` column.
        - **room_moves**: barcodes in both where the ``room`` value changed.

        If either snapshot is not found, shows a warning messagebox and returns
        ``None``.  Callers should check the return value before rendering.

        Returns
        -------
        dict | None
            ``{"new_items", "removed_items", "qty_changes", "room_moves",
            "summary"}`` where *summary* is a flat dict with counts and
            ``net_qty``, or ``None`` if snapshots are missing.
        """
        snaps = self.velocity_mgr.get_snapshots()

        # Find matching snapshots — match on the timestamp prefix since
        # both import_history and velocity use the same import_ts variable
        snap_a = None
        snap_b = None
        for s in snaps:
            s_ts = s.get("timestamp", "")
            if s_ts == ts_a:
                snap_a = s
            if s_ts == ts_b:
                snap_b = s

        if not snap_a or not snap_b:
            messagebox.showwarning(
                "Missing Snapshot",
                "One or both velocity snapshots were not found.\n"
                "Comparison requires velocity history for both imports.\n\n"
                "This can happen if velocity history was purged.",
                parent=self,
            )
            return None

        # Build barcode → {room, qty} dicts
        map_a = {e["barcode"]: e for e in snap_a.get("entries", [])}
        map_b = {e["barcode"]: e for e in snap_b.get("entries", [])}

        barcodes_a = set(map_a.keys())
        barcodes_b = set(map_b.keys())

        # New items (in B, not in A)
        new_items = [
            {"Barcode": bc, "Room": map_b[bc].get("room", ""), "Qty": map_b[bc].get("qty", 0)}
            for bc in sorted(barcodes_b - barcodes_a)
        ]

        # Removed items (in A, not in B)
        removed_items = [
            {"Barcode": bc, "Room": map_a[bc].get("room", ""), "Qty": map_a[bc].get("qty", 0)}
            for bc in sorted(barcodes_a - barcodes_b)
        ]

        # Items in both — check for changes
        qty_changes = []
        room_moves = []
        for bc in sorted(barcodes_a & barcodes_b):
            a = map_a[bc]
            b = map_b[bc]
            a_qty = a.get("qty", 0)
            b_qty = b.get("qty", 0)
            a_room = a.get("room", "")
            b_room = b.get("room", "")

            if a_qty != b_qty:
                delta = b_qty - a_qty if isinstance(b_qty, (int, float)) and isinstance(a_qty, (int, float)) else 0
                qty_changes.append({
                    "Barcode": bc,
                    "Room": b_room,
                    "Qty (A)": a_qty,
                    "Qty (B)": b_qty,
                    "Delta": delta,
                })
            if str(a_room).strip().lower() != str(b_room).strip().lower():
                room_moves.append({
                    "Barcode": bc,
                    "From Room": a_room,
                    "To Room": b_room,
                    "Qty": b_qty,
                })

        net_qty = sum(c["Delta"] for c in qty_changes)

        return {
            "new_items": new_items,
            "removed_items": removed_items,
            "qty_changes": qty_changes,
            "room_moves": room_moves,
            "summary": {
                "new": len(new_items),
                "removed": len(removed_items),
                "qty_changed": len(qty_changes),
                "room_moved": len(room_moves),
                "net_qty": net_qty,
            },
        }

    def _render_comparison(self, result: Dict):
        summary = result["summary"]
        self._summary_vars["new"].set(str(summary["new"]))
        self._summary_vars["removed"].set(str(summary["removed"]))
        self._summary_vars["qty_changed"].set(str(summary["qty_changed"]))
        self._summary_vars["room_moved"].set(str(summary["room_moved"]))
        net = summary["net_qty"]
        sign = "+" if net > 0 else ""
        self._summary_vars["net_qty"].set(f"{sign}{net}")

        # New items
        if result["new_items"]:
            df_new = pd.DataFrame(result["new_items"])
            self.new_table.render(df_new, ["Barcode", "Room", "Qty"])
        else:
            self.new_table.render(pd.DataFrame(), [])

        # Removed items
        if result["removed_items"]:
            df_rem = pd.DataFrame(result["removed_items"])
            self.removed_table.render(df_rem, ["Barcode", "Room", "Qty"])
        else:
            self.removed_table.render(pd.DataFrame(), [])

        # Qty changes
        if result["qty_changes"]:
            df_qty = pd.DataFrame(result["qty_changes"])
            self.qty_table.render(df_qty, ["Barcode", "Room", "Qty (A)", "Qty (B)", "Delta"],
                                  tag_col="Delta")
        else:
            self.qty_table.render(pd.DataFrame(), [])

        # Room moves
        if result["room_moves"]:
            df_room = pd.DataFrame(result["room_moves"])
            self.room_table.render(df_room, ["Barcode", "From Room", "To Room", "Qty"])
        else:
            self.room_table.render(pd.DataFrame(), [])

    # ----------------------------------------------------------------
    # Refresh all
    # ----------------------------------------------------------------
    def _refresh_all(self):
        self._refresh_timeline()
        self._refresh_trends()
        self._populate_compare_dropdowns()

    # ----------------------------------------------------------------
    # Actions
    # ----------------------------------------------------------------
    def _purge_all(self):
        if not messagebox.askyesno("Purge All", "Delete ALL import history entries?", parent=self):
            return
        removed = self.history_mgr.purge_all()
        self._refresh_all()
        messagebox.showinfo("Purged", f"Removed {removed} entries.", parent=self)

    def _purge_old(self):
        cutoff = (datetime.now() - timedelta(days=30)).isoformat()
        removed = self.history_mgr.purge_before(cutoff)
        self._refresh_all()
        if removed:
            messagebox.showinfo("Purged", f"Removed {removed} entries older than 30 days.", parent=self)
        else:
            messagebox.showinfo("Purge", "No entries older than 30 days.", parent=self)


# ------------------------------------------------------------------ #
# Public entry point
# ------------------------------------------------------------------ #
def open_import_history_window(parent, import_history_mgr, velocity_mgr):
    """
    Open the Import History window.

    Creates an ``ImportHistoryWindow`` as a child of *parent* and returns it.
    Called from the main toolbar in ``main.py``.

    Parameters
    ----------
    parent : tk.Tk | tk.Toplevel
        Parent window.
    import_history_mgr : ImportHistoryManager
        The live manager instance from the main app (already loaded).
    velocity_mgr : VelocityHistoryManager
        The live velocity manager instance (used by the Compare tab).

    Returns
    -------
    ImportHistoryWindow
        The newly created window instance.
    """
    win = ImportHistoryWindow(parent, import_history_mgr, velocity_mgr)
    return win
