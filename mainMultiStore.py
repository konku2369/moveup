"""
Multi-Store Comparison window for Bisa Inventory Utility.

Lets users import inventory files from two stores (e.g., Carol Stream and
Joliet) and compare them side-by-side. Analysis includes:
  - Products at Store A only (transfer candidates → Store B)
  - Products at Store B only (transfer candidates → Store A)
  - Products at both stores (with qty comparison + imbalance flags)
  - Stock imbalances (one store has way more than the other)
  - Category breakdown (by Type and Brand)
  - Transfer recommendations with priority scoring
  - Summary dashboard

This is a standalone Toplevel window launched from the main toolbar.

Architecture:
  - Uses load_raw_df() + automap_columns() from data_core.py for each file
  - Comparison is by (Brand, Product Name) tuple — same logic as move-up
  - No config persistence needed (comparison is ephemeral)
"""

import os
from typing import Optional, Dict, List, Tuple

import pandas as pd

from tkinter import (
    Toplevel, StringVar, filedialog, messagebox, ttk
)
import tkinter as tk

from data_core import (
    COLUMNS_TO_USE,
    load_raw_df,
    automap_columns,
)
from inventory_analysis import (
    build_product_map,
    compute_imbalances,
    compute_transfer_recs,
)
from tree_ops import make_scrollable_tree


def open_multi_store_window(parent, current_file_path: Optional[str] = None):
    """
    Open the Multi-Store Comparison window.

    Creates a ``MultiStoreWindow`` as a child of *parent*.  If
    *current_file_path* is provided and is a valid file, it is pre-loaded as
    Store A so the user only needs to import Store B.

    Called from the main toolbar in ``main.py``.
    """
    MultiStoreWindow(parent, current_file_path)


class MultiStoreWindow:
    """
    Multi-Store Comparison window for comparing two store inventories.

    Provides a tabbed analysis of the catalog and stock differences between
    two METRC exports.  Analysis includes:
    - Products exclusive to Store A (transfer candidates to Store B)
    - Products exclusive to Store B (transfer candidates to Store A)
    - Products shared by both stores with qty comparison and imbalance flags
    - Stock imbalances (one store has ≥2× more units, diff ≥ 3)
    - Category breakdown (by Type and Brand)
    - Transfer recommendations with High/Medium/Low priority scoring
    - Summary dashboard with headline numbers

    Comparison key is ``(Brand, Product Name)`` — same logic as the move-up
    algorithm.  Product identity is not by barcode because different lots of
    the same product should be compared together.

    After running a comparison, the "Open in Analytics" button passes both
    DataFrames directly to ``AnalyticsWindow`` for deeper analysis.
    """

    def __init__(self, parent, current_file_path: Optional[str] = None):
        self.parent = parent
        self.current_file_path = current_file_path

        self.win = Toplevel(parent)
        self.win.title("Multi-Store Comparison")
        self.win.geometry("1200x750")
        self.win.transient(parent)

        self.store_a_df: Optional[pd.DataFrame] = None
        self.store_b_df: Optional[pd.DataFrame] = None
        self.store_a_name = StringVar(value="Carol Stream")
        self.store_b_name = StringVar(value="Joliet")
        self.store_a_file = StringVar(value="No file loaded")
        self.store_b_file = StringVar(value="No file loaded")
        self.status = StringVar(value="Import files from two stores to compare inventory.")

        # Cached analysis results (populated by _run_comparison)
        self._a_products: Dict = {}
        self._b_products: Dict = {}

        self._build_ui()

        # If a file is already loaded in main, use it for Store A
        if current_file_path and os.path.isfile(current_file_path):
            try:
                raw = load_raw_df(current_file_path)
                mapped, _ = automap_columns(raw)
                self.store_a_df = mapped
                self.store_a_file.set(os.path.basename(current_file_path))
                self.status.set(f"Store A loaded from main app ({len(mapped)} items). Import Store B to compare.")
            except Exception as e:
                print(f"[moveup] Multi-Store auto-load failed: {e}")

    def _build_ui(self):
        # ── Top: store import controls ──
        frm_top = ttk.Frame(self.win, padding=10)
        frm_top.pack(fill="x")

        # Store A
        frm_a = ttk.LabelFrame(frm_top, text="Store A", padding=8)
        frm_a.pack(side="left", fill="x", expand=True, padx=(0, 4))

        ttk.Label(frm_a, text="Name:").pack(side="left")
        ttk.Entry(frm_a, textvariable=self.store_a_name, width=18).pack(side="left", padx=4)
        ttk.Button(frm_a, text="Import…", command=self._import_store_a).pack(side="left", padx=4)
        ttk.Label(frm_a, textvariable=self.store_a_file, foreground="#555").pack(side="left", padx=6)

        # Store B
        frm_b = ttk.LabelFrame(frm_top, text="Store B", padding=8)
        frm_b.pack(side="left", fill="x", expand=True, padx=(4, 0))

        ttk.Label(frm_b, text="Name:").pack(side="left")
        ttk.Entry(frm_b, textvariable=self.store_b_name, width=18).pack(side="left", padx=4)
        ttk.Button(frm_b, text="Import…", command=self._import_store_b).pack(side="left", padx=4)
        ttk.Label(frm_b, textvariable=self.store_b_file, foreground="#555").pack(side="left", padx=6)

        # Compare + export + analytics buttons
        frm_cmp = ttk.Frame(self.win, padding=(10, 4))
        frm_cmp.pack(fill="x")
        ttk.Button(frm_cmp, text="Compare", command=self._run_comparison).pack(side="left", padx=4)
        ttk.Button(frm_cmp, text="Export Full Report (Excel)", command=self._export_excel).pack(side="left", padx=4)
        ttk.Button(frm_cmp, text="Open Analytics…", command=self._open_analytics).pack(side="left", padx=4)
        ttk.Label(frm_cmp, textvariable=self.status, foreground="#333").pack(side="left", padx=12)

        # ── Notebook with comparison tabs ──
        self.nb = ttk.Notebook(self.win)
        self.nb.pack(fill="both", expand=True, padx=10, pady=(4, 10))

        self.tab_a_only = ttk.Frame(self.nb)
        self.tab_b_only = ttk.Frame(self.nb)
        self.tab_both = ttk.Frame(self.nb)
        self.tab_imbalance = ttk.Frame(self.nb)
        self.tab_transfers = ttk.Frame(self.nb)

        self.nb.add(self.tab_a_only, text="Store A Only")
        self.nb.add(self.tab_b_only, text="Store B Only")
        self.nb.add(self.tab_both, text="Both Stores")
        self.nb.add(self.tab_imbalance, text="Imbalances")
        self.nb.add(self.tab_transfers, text="Transfer Recs")

        # Treeviews
        cols_basic = ["Type", "Brand", "Product Name", "Room(s)", "Qty"]
        cols_both = ["Type", "Brand", "Product Name", "Rooms A", "Qty A", "Rooms B", "Qty B", "Diff"]
        cols_imbalance = ["Type", "Brand", "Product Name", "Qty A", "Qty B", "Ratio", "Overstocked At"]
        cols_transfer = ["Priority", "Type", "Brand", "Product Name", "From", "To", "Qty Available", "Reason"]

        self.tree_a_only = make_scrollable_tree(self.tab_a_only, cols_basic)
        self.tree_b_only = make_scrollable_tree(self.tab_b_only, cols_basic)
        self.tree_both = make_scrollable_tree(self.tab_both, cols_both)
        self.tree_imbalance = make_scrollable_tree(self.tab_imbalance, cols_imbalance)
        self.tree_transfers = make_scrollable_tree(self.tab_transfers, cols_transfer)

        # Tag colors for imbalance tree
        self.tree_imbalance.tag_configure("heavy_a", foreground="#cc2222")
        self.tree_imbalance.tag_configure("heavy_b", foreground="#2255cc")
        self.tree_both.tag_configure("imbalanced", foreground="#cc6600")

        # Tag colors for transfer tree
        self.tree_transfers.tag_configure("high", foreground="#cc2222")
        self.tree_transfers.tag_configure("medium", foreground="#cc6600")
        self.tree_transfers.tag_configure("low", foreground="#228B22")


    def _import_file(self):
        """Common file import — returns mapped DataFrame or None."""
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
    # Core comparison
    # ------------------------------------------------------------------
    def _run_comparison(self):
        if self.store_a_df is None or self.store_b_df is None:
            messagebox.showinfo("Compare", "Import files for both stores first.", parent=self.win)
            return

        a_name = self.store_a_name.get() or "Store A"
        b_name = self.store_b_name.get() or "Store B"

        # Build product maps
        self._a_products = build_product_map(self.store_a_df)
        self._b_products = build_product_map(self.store_b_df)

        a_keys = set(self._a_products.keys())
        b_keys = set(self._b_products.keys())

        only_a_keys = sorted(a_keys - b_keys)
        only_b_keys = sorted(b_keys - a_keys)
        both_keys = sorted(a_keys & b_keys)

        # ── Tab 1: Store A Only ──
        self._clear_tree(self.tree_a_only)
        for key in only_a_keys:
            info = self._a_products[key]
            self.tree_a_only.insert("", "end", values=(
                info["type"], key[0], key[1], info["rooms"], info["qty"],
            ))

        # ── Tab 2: Store B Only ──
        self._clear_tree(self.tree_b_only)
        for key in only_b_keys:
            info = self._b_products[key]
            self.tree_b_only.insert("", "end", values=(
                info["type"], key[0], key[1], info["rooms"], info["qty"],
            ))

        # ── Tab 3: Both Stores ──
        self._clear_tree(self.tree_both)
        for key in both_keys:
            a_info = self._a_products[key]
            b_info = self._b_products[key]
            diff = a_info["qty"] - b_info["qty"]
            diff_str = f"+{diff}" if diff > 0 else str(diff)
            tags = ()
            # Both conditions: percentage alone flags 2 vs 1 (1 unit difference); abs diff alone misses proportional gaps on high stock.
            if abs(diff) > max(a_info["qty"], b_info["qty"]) * 0.5 and abs(diff) >= 3:
                tags = ("imbalanced",)
            self.tree_both.insert("", "end", values=(
                a_info["type"], key[0], key[1],
                a_info["rooms"], a_info["qty"],
                b_info["rooms"], b_info["qty"],
                diff_str,
            ), tags=tags)

        # ── Tab 4: Imbalances ──
        imbalances = compute_imbalances(both_keys, self._a_products, self._b_products, a_name, b_name)
        self._clear_tree(self.tree_imbalance)
        for row in imbalances:
            tag = "heavy_a" if row["overstocked"] == a_name else "heavy_b"
            self.tree_imbalance.insert("", "end", values=(
                row["type"], row["brand"], row["name"],
                row["qty_a"], row["qty_b"],
                row["ratio"], row["overstocked"],
            ), tags=(tag,))

        # ── Tab 5: Transfer Recommendations ──
        transfers = compute_transfer_recs(
            only_a_keys, only_b_keys, both_keys,
            self._a_products, self._b_products, a_name, b_name,
        )
        self._clear_tree(self.tree_transfers)
        for rec in transfers:
            tag = rec["priority"].lower()
            self.tree_transfers.insert("", "end", values=(
                rec["priority"], rec["type"], rec["brand"], rec["name"],
                rec["from"], rec["to"], rec["qty"], rec["reason"],
            ), tags=(tag,))

        # Update tab titles
        self.nb.tab(self.tab_a_only, text=f"{a_name} Only ({len(only_a_keys)})")
        self.nb.tab(self.tab_b_only, text=f"{b_name} Only ({len(only_b_keys)})")
        self.nb.tab(self.tab_both, text=f"Both Stores ({len(both_keys)})")
        self.nb.tab(self.tab_imbalance, text=f"Imbalances ({len(imbalances)})")
        self.nb.tab(self.tab_transfers, text=f"Transfer Recs ({len(transfers)})")

        self.status.set(
            f"Compared: {len(only_a_keys)} {a_name}-only, "
            f"{len(only_b_keys)} {b_name}-only, {len(both_keys)} shared, "
            f"{len(imbalances)} imbalanced, {len(transfers)} transfer recs."
        )


    def _open_analytics(self):
        """Launch the Analytics window with pre-loaded store data."""
        try:
            from mainAnalytics import open_analytics_window
            store_a_data = None
            store_b_data = None
            if self.store_a_df is not None:
                store_a_data = (self.store_a_df, self.store_a_name.get() or "Store A")
            if self.store_b_df is not None:
                store_b_data = (self.store_b_df, self.store_b_name.get() or "Store B")
            open_analytics_window(
                self.win, store_a_data=store_a_data, store_b_data=store_b_data,
            )
        except Exception as e:
            messagebox.showerror("Analytics", f"Could not open window:\n\n{e}", parent=self.win)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _clear_tree(self, tree: ttk.Treeview):
        for i in tree.get_children():
            tree.delete(i)

    # ------------------------------------------------------------------
    # Export
    # ------------------------------------------------------------------
    def _export_excel(self):
        """Export full comparison report to Excel with multiple sheets."""
        if not self._a_products or not self._b_products:
            messagebox.showinfo("Export", "Run a comparison first.", parent=self.win)
            return

        path = filedialog.asksaveasfilename(
            parent=self.win,
            title="Save Comparison Report",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="store_comparison.xlsx",
        )
        if not path:
            return

        a_name = self.store_a_name.get() or "Store A"
        b_name = self.store_b_name.get() or "Store B"

        a_keys = set(self._a_products.keys())
        b_keys = set(self._b_products.keys())

        def _to_rows(keys, product_map):
            rows = []
            for key in sorted(keys):
                info = product_map[key]
                rows.append({
                    "Type": info["type"],
                    "Brand": key[0],
                    "Product Name": key[1],
                    "Room(s)": info["rooms"],
                    "Qty": info["qty"],
                })
            return pd.DataFrame(rows) if rows else pd.DataFrame(
                columns=["Type", "Brand", "Product Name", "Room(s)", "Qty"]
            )

        only_a_df = _to_rows(a_keys - b_keys, self._a_products)
        only_b_df = _to_rows(b_keys - a_keys, self._b_products)

        # Both stores sheet
        both_keys = sorted(a_keys & b_keys)
        both_rows = []
        for key in both_keys:
            a_info = self._a_products[key]
            b_info = self._b_products[key]
            diff = a_info["qty"] - b_info["qty"]
            both_rows.append({
                "Type": a_info["type"],
                "Brand": key[0],
                "Product Name": key[1],
                f"{a_name} Room(s)": a_info["rooms"],
                f"{a_name} Qty": a_info["qty"],
                f"{b_name} Room(s)": b_info["rooms"],
                f"{b_name} Qty": b_info["qty"],
                "Difference": diff,
            })
        both_df = pd.DataFrame(both_rows) if both_rows else pd.DataFrame()

        # Imbalances sheet
        imbalances = compute_imbalances(both_keys, self._a_products, self._b_products, a_name, b_name)
        imb_rows = [{
            "Type": r["type"], "Brand": r["brand"], "Product Name": r["name"],
            f"{a_name} Qty": r["qty_a"], f"{b_name} Qty": r["qty_b"],
            "Ratio": r["ratio"], "Overstocked At": r["overstocked"],
        } for r in imbalances]
        imb_df = pd.DataFrame(imb_rows) if imb_rows else pd.DataFrame()

        # Transfer recs sheet
        transfers = compute_transfer_recs(
            sorted(a_keys - b_keys), sorted(b_keys - a_keys),
            both_keys, self._a_products, self._b_products, a_name, b_name,
        )
        xfer_rows = [{
            "Priority": r["priority"], "Type": r["type"],
            "Brand": r["brand"], "Product Name": r["name"],
            "From Store": r["from"], "To Store": r["to"],
            "Qty to Transfer": r["qty"], "Reason": r["reason"],
        } for r in transfers]
        xfer_df = pd.DataFrame(xfer_rows) if xfer_rows else pd.DataFrame()

        try:
            with pd.ExcelWriter(path, engine="openpyxl") as w:
                only_a_df.to_excel(w, sheet_name=f"{a_name} Only"[:31], index=False)
                only_b_df.to_excel(w, sheet_name=f"{b_name} Only"[:31], index=False)
                if not both_df.empty:
                    both_df.to_excel(w, sheet_name="Both Stores", index=False)
                if not imb_df.empty:
                    imb_df.to_excel(w, sheet_name="Imbalances", index=False)
                if not xfer_df.empty:
                    xfer_df.to_excel(w, sheet_name="Transfer Recs", index=False)

            self.status.set(f"Exported: {os.path.basename(path)}")
            messagebox.showinfo("Export", f"Saved to:\n{path}", parent=self.win)
        except Exception as e:
            messagebox.showerror("Export Error", str(e), parent=self.win)
