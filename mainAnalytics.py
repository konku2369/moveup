"""
Store Analytics window for Bisa Inventory Utility.

Standalone satellite window that performs deep analysis on one or two
store inventory files. Can be launched from the main toolbar or from
the Multi-Store Comparison window.

Analysis includes:
  - Overview: product counts, unit totals, catalog overlap
  - Category breakdown by Type
  - Top brands with gap analysis
  - Stock imbalances (ratio >= 2x AND diff >= 3)
  - Transfer recommendations with priority scoring
  - Exclusive product highlights by category

This is a standalone Toplevel window — no dependency on MultiStoreWindow.
"""

import os
from typing import Optional, Dict, List, Tuple

import pandas as pd

from tkinter import (
    Toplevel, StringVar, filedialog, messagebox, ttk
)
import tkinter as tk

from data_core import (
    load_raw_df,
    automap_columns,
)
from inventory_analysis import (
    build_product_map,
    compute_imbalances,
    compute_transfer_recs,
    category_breakdown,
)


def open_analytics_window(
    parent,
    current_file_path: Optional[str] = None,
    store_a_data: Optional[Tuple[pd.DataFrame, str]] = None,
    store_b_data: Optional[Tuple[pd.DataFrame, str]] = None,
):
    """Launch the Analytics window.

    Args:
        parent: Tk parent window.
        current_file_path: path to the currently loaded file in main app
            (used as Store A fallback if store_a_data is not provided).
        store_a_data: optional (DataFrame, store_name) tuple — pre-loaded.
        store_b_data: optional (DataFrame, store_name) tuple — pre-loaded.
    """
    AnalyticsWindow(parent, current_file_path, store_a_data, store_b_data)


class AnalyticsWindow:
    """
    Toplevel window for deep inventory analytics.

    Supports single-store mode (overview, category breakdown, brand ranking,
    room distribution, low/zero stock alerts) and two-store comparison mode
    (catalog overlap, imbalances, transfer recommendations, exclusive products).

    The window can be launched three ways:
    1. From ``main.py`` toolbar with the current file pre-loaded as Store A.
    2. From ``MultiStoreWindow`` with both stores already loaded (auto-runs
       the analysis immediately).
    3. Standalone with no pre-loaded data (user imports both files manually).
    """

    def __init__(
        self,
        parent,
        current_file_path: Optional[str] = None,
        store_a_data: Optional[Tuple[pd.DataFrame, str]] = None,
        store_b_data: Optional[Tuple[pd.DataFrame, str]] = None,
    ):
        self.parent = parent

        self.win = Toplevel(parent)
        self.win.title("Store Analytics")
        self.win.geometry("960x720")
        self.win.transient(parent)

        self.store_a_df: Optional[pd.DataFrame] = None
        self.store_b_df: Optional[pd.DataFrame] = None
        self.store_a_name = StringVar(value="Carol Stream")
        self.store_b_name = StringVar(value="Joliet")
        self.store_a_file = StringVar(value="No file loaded")
        self.store_b_file = StringVar(value="No file loaded")
        self.status = StringVar(value="Import at least one store to analyze.")

        self._a_products: Dict = {}
        self._b_products: Dict = {}

        self._build_ui()

        # Pre-load from arguments
        if store_a_data is not None:
            df, name = store_a_data
            self.store_a_df = df
            self.store_a_name.set(name)
            self.store_a_file.set(f"({len(df)} items)")
        elif current_file_path and os.path.isfile(current_file_path):
            try:
                raw = load_raw_df(current_file_path)
                mapped, _ = automap_columns(raw)
                self.store_a_df = mapped
                self.store_a_file.set(os.path.basename(current_file_path))
            except Exception as e:
                print(f"[moveup] Analytics auto-load failed: {e}")

        if store_b_data is not None:
            df, name = store_b_data
            self.store_b_df = df
            self.store_b_name.set(name)
            self.store_b_file.set(f"({len(df)} items)")

        # Auto-run if both stores are loaded (launched from Multi-Store)
        if self.store_a_df is not None and self.store_b_df is not None:
            self._run_analysis()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        """Build the static UI layout (store import controls, text area)."""
        # ── Top: store import controls ──
        frm_top = ttk.Frame(self.win, padding=10)
        frm_top.pack(fill="x")

        # Store A
        frm_a = ttk.LabelFrame(frm_top, text="Store A", padding=6)
        frm_a.pack(side="left", fill="x", expand=True, padx=(0, 4))
        ttk.Label(frm_a, text="Name:").pack(side="left")
        ttk.Entry(frm_a, textvariable=self.store_a_name, width=16).pack(side="left", padx=4)
        ttk.Button(frm_a, text="Import…", command=self._import_store_a).pack(side="left", padx=4)
        ttk.Label(frm_a, textvariable=self.store_a_file, foreground="#555").pack(side="left", padx=6)

        # Store B (optional)
        frm_b = ttk.LabelFrame(frm_top, text="Store B (optional)", padding=6)
        frm_b.pack(side="left", fill="x", expand=True, padx=(4, 0))
        ttk.Label(frm_b, text="Name:").pack(side="left")
        ttk.Entry(frm_b, textvariable=self.store_b_name, width=16).pack(side="left", padx=4)
        ttk.Button(frm_b, text="Import…", command=self._import_store_b).pack(side="left", padx=4)
        ttk.Label(frm_b, textvariable=self.store_b_file, foreground="#555").pack(side="left", padx=6)

        # Action bar
        frm_act = ttk.Frame(self.win, padding=(10, 4))
        frm_act.pack(fill="x")
        ttk.Button(frm_act, text="Run Analysis", command=self._run_analysis).pack(side="left", padx=4)
        ttk.Button(frm_act, text="Export Text Report…", command=self._export_text).pack(side="left", padx=4)
        ttk.Label(frm_act, textvariable=self.status, foreground="#333").pack(side="left", padx=12)

        # ── Analysis text area ──
        frm_text = ttk.Frame(self.win, padding=(10, 4, 10, 10))
        frm_text.pack(fill="both", expand=True)

        self.text = tk.Text(frm_text, wrap="word", font=("Consolas", 11), padx=12, pady=12)
        sb = ttk.Scrollbar(frm_text, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=sb.set)
        self.text.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.text.insert("1.0", "Import at least one store file, then click Run Analysis.")
        self.text.config(state="disabled")

        # Text tags
        self.text.tag_configure("header", font=("Consolas", 13, "bold"))
        self.text.tag_configure("subheader", font=("Consolas", 11, "bold"))
        self.text.tag_configure("highlight", foreground="#cc6600")

    # ------------------------------------------------------------------
    # Import
    # ------------------------------------------------------------------
    def _import_file(self):
        """
        Prompt the user to select an inventory file and return a mapped DataFrame.

        Opens a file dialog, reads the file via ``load_raw_df()``, and runs
        ``automap_columns()`` to map METRC columns.  On success returns
        ``(df, base_filename)``; on any error shows a messagebox and returns
        ``(None, None)``.
        """
        path = filedialog.askopenfilename(
            parent=self.win,
            title="Select Inventory File",
            filetypes=[
                ("All Supported", "*.xlsx *.xls *.xlsm *.xlsb *.ods *.csv *.tsv *.txt *.tab"),
                ("Excel", "*.xlsx *.xls *.xlsm *.xlsb"),
                ("CSV / Text", "*.csv *.tsv *.txt *.tab"),
                ("OpenDocument", "*.ods"),
            ],
        )
        if not path:
            return None, None
        try:
            raw = load_raw_df(path)
            mapped, _ = automap_columns(raw)
            return mapped, os.path.basename(path)
        except Exception as e:
            messagebox.showerror("Import Error", str(e), parent=self.win)
            return None, None

    def _import_store_a(self):
        df, name = self._import_file()
        if df is not None:
            self.store_a_df = df
            self.store_a_file.set(name)
            self.status.set(f"Store A: {len(df)} items loaded.")

    def _import_store_b(self):
        df, name = self._import_file()
        if df is not None:
            self.store_b_df = df
            self.store_b_file.set(name)
            self.status.set(f"Store B: {len(df)} items loaded.")

    # ------------------------------------------------------------------
    # Analysis engine
    # ------------------------------------------------------------------
    def _run_analysis(self):
        """
        Run the analysis engine and render results into the text widget.

        Builds product maps via ``build_product_map()`` for whichever stores
        are loaded.  If both stores are loaded, delegates to
        ``_render_two_store()``; otherwise renders a single-store report via
        ``_render_single_store()``.  The text widget is put into ``"normal"``
        state for writing and ``"disabled"`` afterward to prevent editing.
        """
        if self.store_a_df is None:
            messagebox.showinfo("Analytics", "Import at least Store A.", parent=self.win)
            return

        a_name = self.store_a_name.get() or "Store A"
        self._a_products = build_product_map(self.store_a_df)

        two_stores = self.store_b_df is not None
        if two_stores:
            b_name = self.store_b_name.get() or "Store B"
            self._b_products = build_product_map(self.store_b_df)
        else:
            b_name = ""
            self._b_products = {}

        t = self.text
        t.config(state="normal")
        t.delete("1.0", "end")

        if two_stores:
            self._render_two_store(t, a_name, b_name)
        else:
            self._render_single_store(t, a_name)

        t.config(state="disabled")  # tk.Text is editable by default; disable prevents user from corrupting the formatted report
        self.status.set("Analysis complete.")

    # ------------------------------------------------------------------
    # Single-store analysis
    # ------------------------------------------------------------------
    def _render_single_store(self, t: tk.Text, name: str):
        """
        Write a single-store analytics report into text widget *t*.

        Sections rendered: Overview, Breakdown by Category, Top Brands,
        Room Distribution, Low Stock Alerts (qty ≤ 2), Zero Stock (qty = 0).
        All product data comes from ``self._a_products`` (built during
        ``_run_analysis()``).
        """
        products = self._a_products
        total_units = sum(info["qty"] for info in products.values())
        unique_products = len(products)

        t.insert("end", f"Store Analytics — {name}\n", "header")
        t.insert("end", "=" * 50 + "\n\n")

        # Overview
        t.insert("end", "OVERVIEW\n", "subheader")
        t.insert("end", f"  Unique Products:  {unique_products:,d}\n")
        t.insert("end", f"  Total Units:      {total_units:,d}\n")
        if unique_products:
            avg = total_units / unique_products
            t.insert("end", f"  Avg Units/Product: {avg:.1f}\n")
        t.insert("end", "\n")

        # Category breakdown
        t.insert("end", "BREAKDOWN BY CATEGORY (TYPE)\n", "subheader")
        type_stats = category_breakdown(products, "type")
        t.insert("end", f"  {'Category':24s} {'Products':>10s} {'Units':>10s} {'Avg Qty':>10s}\n")
        t.insert("end", f"  {'-' * 56}\n")
        for cat in sorted(type_stats.keys()):
            s = type_stats[cat]
            avg = s["qty"] / max(s["count"], 1)
            t.insert("end", f"  {cat:24s} {s['count']:>10,d} {s['qty']:>10,d} {avg:>10.1f}\n")
        t.insert("end", "\n")

        # Top brands
        t.insert("end", "TOP BRANDS (by total units)\n", "subheader")
        brand_stats = category_breakdown(products, "brand")
        top_brands = sorted(brand_stats.items(), key=lambda x: -x[1]["qty"])[:20]
        t.insert("end", f"  {'Brand':28s} {'Products':>10s} {'Units':>10s}\n")
        t.insert("end", f"  {'-' * 50}\n")
        for brand, s in top_brands:
            t.insert("end", f"  {brand:28s} {s['count']:>10,d} {s['qty']:>10,d}\n")
        t.insert("end", "\n")

        # Room distribution
        t.insert("end", "ROOM DISTRIBUTION\n", "subheader")
        room_stats: Dict[str, Dict] = {}
        for info in products.values():
            for room in info["rooms"].split(", "):
                room = room.strip()
                if not room:
                    continue
                if room not in room_stats:
                    room_stats[room] = {"count": 0, "qty": 0}
                room_stats[room]["count"] += 1
                room_stats[room]["qty"] += info["qty"]
        for room in sorted(room_stats.keys()):
            s = room_stats[room]
            t.insert("end", f"  {room:28s} {s['count']:>6d} products, {s['qty']:>6d} units\n")
        t.insert("end", "\n")

        # Low stock alerts
        t.insert("end", "LOW STOCK ALERTS (qty <= 2)\n", "subheader")
        low_stock = [(k, v) for k, v in products.items() if 0 < v["qty"] <= 2]
        low_stock.sort(key=lambda x: x[1]["qty"])
        if low_stock:
            t.insert("end", f"  {len(low_stock)} products at 2 or fewer units:\n")
            for (brand, pname), info in low_stock[:25]:
                t.insert("end", f"    [{info['qty']}] {brand} — {pname}\n")
            if len(low_stock) > 25:
                t.insert("end", f"    ... and {len(low_stock) - 25} more\n")
        else:
            t.insert("end", "  No low-stock products.\n")
        t.insert("end", "\n")

        # Zero stock
        zero_stock = [(k, v) for k, v in products.items() if v["qty"] == 0]
        t.insert("end", "ZERO STOCK (qty = 0)\n", "subheader")
        if zero_stock:
            t.insert("end", f"  {len(zero_stock)} products showing 0 units on hand.\n")
            for (brand, pname), info in sorted(zero_stock)[:15]:
                t.insert("end", f"    {brand} — {pname} ({info['type']})\n")
            if len(zero_stock) > 15:
                t.insert("end", f"    ... and {len(zero_stock) - 15} more\n")
        else:
            t.insert("end", "  No zero-stock products.\n")

    # ------------------------------------------------------------------
    # Two-store comparison analysis
    # ------------------------------------------------------------------
    def _render_two_store(self, t: tk.Text, a_name: str, b_name: str):
        """
        Write a two-store comparison report into text widget *t*.

        Sections rendered: Overview (overlap %, exclusive counts), Category
        Breakdown side by side, Top Brands with gap analysis, Stock Imbalances
        (via ``compute_imbalances()``), Transfer Recommendations (via
        ``compute_transfer_recs()``), and Exclusive Product Highlights for each
        store.
        """
        a_keys = set(self._a_products.keys())
        b_keys = set(self._b_products.keys())
        only_a = sorted(a_keys - b_keys)
        only_b = sorted(b_keys - a_keys)
        both = sorted(a_keys & b_keys)

        total_a = sum(info["qty"] for info in self._a_products.values())
        total_b = sum(info["qty"] for info in self._b_products.values())

        t.insert("end", "Multi-Store Comparison Analysis\n", "header")
        t.insert("end", "=" * 50 + "\n\n")

        # ── Overview ──
        t.insert("end", "OVERVIEW\n", "subheader")
        t.insert("end", f"  {'':30s} {a_name:>14s} {b_name:>14s}\n")
        t.insert("end", f"  {'Unique Products':30s} {len(self._a_products):>14,d} {len(self._b_products):>14,d}\n")
        t.insert("end", f"  {'Total Units':30s} {total_a:>14,d} {total_b:>14,d}\n")
        t.insert("end", f"  {'Exclusive Products':30s} {len(only_a):>14,d} {len(only_b):>14,d}\n")
        overlap_pct = len(both) / max(len(self._a_products), len(self._b_products), 1) * 100
        t.insert("end", f"  {'Shared Products':30s} {len(both):>14,d}\n")
        t.insert("end", f"  {'Catalog Overlap':30s} {overlap_pct:>13.1f}%\n\n")

        # ── Category Breakdown by Type ──
        t.insert("end", "BREAKDOWN BY CATEGORY (TYPE)\n", "subheader")
        type_stats_a = category_breakdown(self._a_products, "type")
        type_stats_b = category_breakdown(self._b_products, "type")
        all_types = sorted(set(type_stats_a.keys()) | set(type_stats_b.keys()))

        t.insert("end", f"  {'Category':22s} {a_name + ' (products)':>18s} {a_name + ' (units)':>16s}"
                        f" {b_name + ' (products)':>18s} {b_name + ' (units)':>16s}\n")
        t.insert("end", f"  {'-' * 92}\n")
        for cat in all_types:
            sa = type_stats_a.get(cat, {"count": 0, "qty": 0})
            sb = type_stats_b.get(cat, {"count": 0, "qty": 0})
            t.insert("end", f"  {cat:22s} {sa['count']:>18,d} {sa['qty']:>16,d}"
                            f" {sb['count']:>18,d} {sb['qty']:>16,d}\n")
        t.insert("end", "\n")

        # ── Brand Coverage ──
        t.insert("end", "TOP BRANDS (by total units across both stores)\n", "subheader")
        brand_totals: Dict[str, int] = {}
        for products in (self._a_products, self._b_products):
            for (brand, _), info in products.items():
                brand_totals[brand] = brand_totals.get(brand, 0) + info["qty"]
        top_brands = sorted(brand_totals.items(), key=lambda x: -x[1])[:15]

        brand_a = category_breakdown(self._a_products, "brand")
        brand_b = category_breakdown(self._b_products, "brand")

        t.insert("end", f"  {'Brand':28s} {a_name:>14s} {b_name:>14s}   {'Gap':>8s}\n")
        t.insert("end", f"  {'-' * 68}\n")
        for brand, total in top_brands:
            qa = brand_a.get(brand, {"qty": 0})["qty"]
            qb = brand_b.get(brand, {"qty": 0})["qty"]
            gap = qa - qb
            gap_str = f"+{gap}" if gap > 0 else str(gap)
            flag = " !!!" if abs(gap) > total * 0.4 and abs(gap) >= 5 else ""
            t.insert("end", f"  {brand:28s} {qa:>14,d} {qb:>14,d}   {gap_str:>8s}{flag}\n")
        t.insert("end", "\n")

        # ── Imbalances ──
        imbalances = compute_imbalances(both, self._a_products, self._b_products, a_name, b_name)
        t.insert("end", "STOCK IMBALANCES\n", "subheader")
        if imbalances:
            heavy_a = sum(1 for r in imbalances if r["overstocked"] == a_name)
            heavy_b = sum(1 for r in imbalances if r["overstocked"] == b_name)
            t.insert("end", f"  {len(imbalances)} products have significantly uneven stock:\n")
            t.insert("end", f"    {heavy_a} overstocked at {a_name}\n")
            t.insert("end", f"    {heavy_b} overstocked at {b_name}\n")
            t.insert("end", f"\n  Worst imbalances:\n")
            for r in imbalances[:5]:
                t.insert("end", f"    {r['brand']} — {r['name']}: "
                                f"{a_name}={r['qty_a']}, {b_name}={r['qty_b']} ({r['ratio']})\n")
        else:
            t.insert("end", "  No significant imbalances found. Stock levels are well-balanced!\n")
        t.insert("end", "\n")

        # ── Transfer Recommendations ──
        transfers = compute_transfer_recs(
            only_a, only_b, both,
            self._a_products, self._b_products, a_name, b_name,
        )
        t.insert("end", "TRANSFER RECOMMENDATIONS\n", "subheader")
        if transfers:
            high = [r for r in transfers if r["priority"] == "High"]
            med = [r for r in transfers if r["priority"] == "Medium"]
            low = [r for r in transfers if r["priority"] == "Low"]
            t.insert("end", f"  {len(transfers)} total recommendations:\n")
            t.insert("end", f"    HIGH priority:   {len(high):>4d}  (product missing at one store, qty >= 3)\n")
            t.insert("end", f"    MEDIUM priority: {len(med):>4d}  (missing with low qty, or severe imbalance)\n")
            t.insert("end", f"    LOW priority:    {len(low):>4d}  (moderate imbalance)\n\n")

            a_to_b = [r for r in transfers if r["from"] == a_name]
            b_to_a = [r for r in transfers if r["from"] == b_name]
            a_to_b_units = sum(r["qty"] for r in a_to_b)
            b_to_a_units = sum(r["qty"] for r in b_to_a)
            t.insert("end", f"  Direction breakdown:\n")
            t.insert("end", f"    {a_name} -> {b_name}: {len(a_to_b)} products ({a_to_b_units} units)\n")
            t.insert("end", f"    {b_name} -> {a_name}: {len(b_to_a)} products ({b_to_a_units} units)\n")
        else:
            t.insert("end", "  No transfer recommendations. Inventories look aligned!\n")
        t.insert("end", "\n")

        # ── Exclusive Product Highlights ──
        t.insert("end", f"EXCLUSIVE TO {a_name.upper()} ({len(only_a)} products)\n", "subheader")
        if only_a:
            excl_types: Dict[str, int] = {}
            for key in only_a:
                ptype = self._a_products[key]["type"]
                excl_types[ptype] = excl_types.get(ptype, 0) + 1
            for ptype, count in sorted(excl_types.items(), key=lambda x: -x[1]):
                t.insert("end", f"  {ptype}: {count} products\n")
        else:
            t.insert("end", "  None — every product at this store also exists at the other.\n")
        t.insert("end", "\n")

        t.insert("end", f"EXCLUSIVE TO {b_name.upper()} ({len(only_b)} products)\n", "subheader")
        if only_b:
            excl_types = {}
            for key in only_b:
                ptype = self._b_products[key]["type"]
                excl_types[ptype] = excl_types.get(ptype, 0) + 1
            for ptype, count in sorted(excl_types.items(), key=lambda x: -x[1]):
                t.insert("end", f"  {ptype}: {count} products\n")
        else:
            t.insert("end", "  None — every product at this store also exists at the other.\n")

    # ------------------------------------------------------------------
    # Export
    # ------------------------------------------------------------------
    def _export_text(self):
        """Export the analysis text to a .txt file."""
        content = self.text.get("1.0", "end-1c")
        if not content.strip() or "Import at least" in content:
            messagebox.showinfo("Export", "Run analysis first.", parent=self.win)
            return

        path = filedialog.asksaveasfilename(
            parent=self.win,
            title="Save Analysis Report",
            defaultextension=".txt",
            filetypes=[("Text file", "*.txt")],
            initialfile="store_analytics.txt",
        )
        if not path:
            return

        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
            self.status.set(f"Exported: {os.path.basename(path)}")
            messagebox.showinfo("Export", f"Saved to:\n{path}", parent=self.win)
        except Exception as e:
            messagebox.showerror("Export Error", str(e), parent=self.win)
