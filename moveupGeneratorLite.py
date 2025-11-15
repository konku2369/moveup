"""
Move-Up Live v3.7.0 — Minimal (No Presets, Manual Column Mapping, 30-per-Page PDF)
Author: Konrad Kubica (+ ChatGPT)
Date: 2025-11-15  (Milestone: No Vault Low-Stock section)

Build:
    pyinstaller --onefile --noconsole --name "MoveUp Live 3.7.0 Minimal" moveup_live_v3_7_0_minimal.py

    pyinstaller --onefile --noconsole --distpath="C:\Users\kubic\Desktop" --name="moveup-generator 3.8" moveupGeneratorLite.py

Highlights
- Presets removed for simplicity.
- Manual column-mapping GUI to override auto-detect for input files.
- PDF pagination: exactly N items per page (default 30) to align with Avery 30-up sticker sheets.
- Rooms multi-select, brand multi-select, manual remove/hide removed.
- Robust CSV sniffing (delimiter + encoding) and Windows Protected-View unblock.
- Filters edited in a separate dialog window to keep the main UI clean.
- Vault low-stock logic removed: single Move-Up list only.
"""

import os, re, sys, json, csv
from io import TextIOWrapper
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd

# GUI
from tkinter import (
    Tk, Toplevel, StringVar, IntVar, BooleanVar, filedialog, messagebox,
    ttk
)
from tkinter import Listbox, MULTIPLE, END

# PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# ------------------------------
# Constants
# ------------------------------
COLUMNS_TO_USE = ["Type", "Brand", "Product Name", "Package Barcode", "Room", "Qty On Hand"]
COLUMN_WIDTHS = [50, 365, 70, 45, 35]  # widths for [Type, Product, Barcode, Loc, Qty]
BASE_PDF_FILENAME = "Print_me_Filtered_Move_Up"
DATE_FORMAT = "%B %d, %Y"

ALT_NAME_CANDIDATES = {
    "Type": ["type", "product type", "category", "item type", "class"],
    "Brand": ["brand", "brand name", "manufacturer", "mfr"],
    "Product Name": ["product name", "product", "item name", "name", "title", "item"],
    "Package Barcode": [
        "barcode", "package barcode", "package id", "upc", "ean", "gtin",
        "barcode", "metrc code", "metrc barcode", "package upc", "package ean"
    ],
    "Room": ["room", "location", "stock location", "bin", "area", "warehouse location", "site location"],
    "Qty On Hand": ["available qty", "qty on hand", "quantity on hand", "on hand", "quantity", "qoh", "stock", "stock qty", "current quantity", "current qty"],
}

SALES_FLOOR_ALIASES = {
    "sales floor", "floor", "salesfloor", "front of house",
    "foh", "front", "front of shop", "retail"
}

DEFAULT_ROOM_ALIASES = {
    "back room": "Overstock",
    "backroom": "Overstock",
    "back stock": "Overstock",
    "backstock": "Overstock",
    "stockroom": "Overstock",
    "stock room": "Overstock",
    "back": "Overstock",
    "over stock": "Overstock",
    "incoming": "Incoming Deliveries",
    "receiving": "Incoming Deliveries",
    "receiving area": "Incoming Deliveries",
    "delivery": "Incoming Deliveries",
    "deliveries": "Incoming Deliveries",
    "safe": "Vault",
    "safe room": "Vault",
}

# ------------------------------
# Utilities
# ------------------------------

def sanitize_prefix(pfx: str) -> str:
    if not pfx:
        return pfx
    pfx = pfx.strip()
    pfx = re.sub(r'[\\/:*?"<>|]+', "_", pfx)
    pfx = re.sub(r"\s+", "_", pfx)
    return pfx


def _lower_strip_cols(columns):
    return [str(c).strip().lower() for c in columns]


def _find_source_for(target_key: str, lower_cols, mapping=ALT_NAME_CANDIDATES):
    wanted = [w.strip().lower() for w in mapping.get(target_key, [])]
    for idx, lc in enumerate(lower_cols):
        if lc in wanted:
            return idx
    return None


def _build_room_map(user_aliases: dict) -> Dict[str, str]:
    final_map = {k.casefold(): v for k, v in DEFAULT_ROOM_ALIASES.items()}
    if user_aliases:
        final_map.update({(k or "").casefold(): v for k, v in user_aliases.items()})
    return final_map


def _normalize_rooms(df: pd.DataFrame, user_aliases: dict):
    if df is None or df.empty or "Room" not in df.columns:
        return df
    norm_map = _build_room_map(user_aliases)
    out = df.copy()
    out["Room"] = out["Room"].map(lambda v: norm_map.get(str(v).strip().casefold(), str(v).strip()))
    return out


def _windows_unblock_file(path: str):
    if os.name != "nt":
        return
    try:
        ads_path = path + ":Zone.Identifier"
        if os.path.exists(ads_path):
            os.remove(ads_path)
    except Exception:
        pass


def _read_csv_smart(path: str, skiprows: int) -> pd.DataFrame:
    def _attempt(encoding):
        with open(path, "rb") as raw:
            sample = raw.read(4096)
            raw.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample.decode(encoding, errors="ignore"))
                delim = dialect.delimiter
            except Exception:
                delim = ","
            return pd.read_csv(
                TextIOWrapper(raw, encoding=encoding, newline=""),
                skiprows=skiprows,
                dtype={"Barcode": "string", "Package Barcode": "string", "METRC Barcode": "string"},
                sep=delim,
                engine="python"
            )
    try:
        return _attempt("utf-8")
    except Exception:
        return _attempt("latin-1")


def _is_sweed_export(original_file: str, ext: str, sheet_name: str) -> bool:
    try:
        if ext == ".csv":
            head = pd.read_csv(original_file, header=None, nrows=1)
        else:
            head = pd.read_excel(original_file, sheet_name=sheet_name, header=None, nrows=1)
    except Exception:
        try:
            head = pd.read_excel(original_file, sheet_name=0, header=None, nrows=1)
        except Exception:
            return False
    first_cell = str(head.iloc[0, 0]).strip().lower()
    return first_cell.startswith("export date")


# ------------------------------
# Column Mapping
# ------------------------------

def automap_columns(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Try to produce required columns; return (mapped_df, rename_map_used).
    If missing, raise ValueError.
    """
    lower_cols = _lower_strip_cols(df.columns)
    out = df.copy()
    rename_map = {}

    for key in COLUMNS_TO_USE:
        if key in out.columns:
            continue
        idx = _find_source_for(key, lower_cols)
        if idx is not None:
            rename_map[out.columns[idx]] = key
    if rename_map:
        out = out.rename(columns=rename_map)

    missing = [c for c in COLUMNS_TO_USE if c not in out.columns]
    if missing:
        raise ValueError("Missing required column(s) after auto-mapping: " + ", ".join(missing))

    out["Package Barcode"] = out["Package Barcode"].astype("string").fillna("")
    out["Qty On Hand"] = pd.to_numeric(out["Qty On Hand"], errors="coerce").fillna(0).astype(int)
    for col in ["Product Name", "Brand", "Type", "Room"]:
        out[col] = out[col].astype(str)
    return out, rename_map


# ------------------------------
# Loading
# ------------------------------

def load_raw_df(original_file: str, sheet_name: str = "Inventory Adjustments") -> pd.DataFrame:
    _windows_unblock_file(original_file)
    ext = os.path.splitext(original_file)[1].lower()
    skiprows = 3 if _is_sweed_export(original_file, ext, sheet_name) else 0
    if ext == ".csv":
        return _read_csv_smart(original_file, skiprows=skiprows)
    else:
        try:
            return pd.read_excel(
                original_file,
                sheet_name=sheet_name,
                skiprows=skiprows,
                dtype={"Barcode": "string", "Package Barcode": "string", "METRC Barcode": "string"}
            )
        except Exception:
            return pd.read_excel(
                original_file,
                sheet_name=0,
                skiprows=skiprows,
                dtype={"Barcode": "string", "Package Barcode": "string", "METRC Barcode": "string"}
            )


# ------------------------------
# Core filtering
# ------------------------------

def compute_moveup_from_df(
    df: pd.DataFrame,
    candidate_rooms: List[str],
    room_alias_overrides: Optional[Dict[str, str]] = None,
    brand_filter: Optional[List[str]] = None,
    skip_sales_floor: bool = False,
) -> Tuple[pd.DataFrame, Dict[str, int]]:
    """
    Compute the Move-Up DataFrame only (no Vault low-stock split).
    Returns (move_up_df, diagnostics_dict).
    """
    diag = {
        "total_loaded": int(len(df) if df is not None else 0),
        "after_dropna": 0,
        "after_brand": 0,
        "after_type": 0,
        "candidate_pool": 0,
        "removed_as_on_sf": 0,
        "move_up": 0,
    }
    if df is None or df.empty:
        return pd.DataFrame(columns=COLUMNS_TO_USE), diag

    df = df.copy()
    for c in ["Product Name", "Brand", "Package Barcode", "Room", "Type"]:
        if c in df.columns:
            df[c] = df[c].astype(str)

    df = df.dropna(subset=["Product Name", "Brand", "Package Barcode", "Room"]).copy()
    diag["after_dropna"] = int(len(df))

    # Brand filter
    if brand_filter:
        bf = [str(b).strip() for b in brand_filter if str(b).strip()]
        is_all = any(b.upper() == "ALL" for b in bf)
        if not is_all:
            df = df[df["Brand"].astype(str).isin(bf)]
    diag["after_brand"] = int(len(df))

    # Room normalization
    df = _normalize_rooms(df, room_alias_overrides or {})

    # Exclude accessories
    if "Type" in df.columns:
        mask_accessory = df["Type"].astype(str).str.contains(r"accessor", case=False, na=False)
        df = df.loc[~mask_accessory].copy()
    diag["after_type"] = int(len(df))

    # Sales floor removal
    room_lower = df["Room"].astype(str).str.strip().str.lower()
    if not skip_sales_floor:
        sf_mask = room_lower.eq("sales floor") | room_lower.isin(SALES_FLOOR_ALIASES)
        sales_floor = df.loc[sf_mask, ["Brand", "Product Name"]].drop_duplicates()
    else:
        sales_floor = pd.DataFrame(columns=["Brand", "Product Name"])

    candidates = df.loc[df["Room"].isin(candidate_rooms), COLUMNS_TO_USE]
    diag["candidate_pool"] = int(len(candidates))

    if skip_sales_floor or sales_floor.empty:
        move_up_df = candidates.copy()
        diag["removed_as_on_sf"] = 0
    else:
        merged = candidates.merge(sales_floor.assign(on_sf=1), on=["Brand", "Product Name"], how="left")
        removed = merged["on_sf"].notna().sum()
        move_up_df = merged.loc[merged["on_sf"].isna()].drop(columns=["on_sf"])
        diag["removed_as_on_sf"] = int(removed)

    sort_cols = [c for c in ["Type", "Brand", "Product Name"] if c in move_up_df.columns]
    if sort_cols:
        move_up_df.sort_values(by=sort_cols, inplace=True, kind="stable")

    diag["move_up"] = int(len(move_up_df))
    return move_up_df, diag


# ------------------------------
# Exports (30-per-page pagination)
# ------------------------------

def _draw_footer(canvas, doc):
    canvas.saveState()
    w, h = letter
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


def _ellipses(s: str, n: int) -> str:
    s = str(s)
    return s if len(s) <= n else s[: max(0, n-3)] + "..."


def _prep_table_df(df: pd.DataFrame) -> pd.DataFrame:
    pdf_df = df.loc[:, COLUMNS_TO_USE].copy()
    sort_cols = [c for c in ["Type", "Brand", "Product Name"] if c in pdf_df.columns]
    if sort_cols:
        pdf_df.sort_values(by=sort_cols, inplace=True, kind="stable")
    pdf_df["Room"] = pdf_df["Room"].astype(str).fillna("").map(lambda s: _ellipses(s, 8))
    pdf_df["Type"] = pdf_df["Type"].astype(str).fillna("").map(lambda s: _ellipses(s, 8))
    pdf_df["Product Name"] = pdf_df["Product Name"].astype(str).fillna("").map(lambda s: _ellipses(s, 75))
    pdf_df["Package Barcode"] = pdf_df["Package Barcode"].map(lambda x: str(x)[-6:] if str(x) else "")
    pdf_df = pdf_df[["Type", "Product Name", "Package Barcode", "Room", "Qty On Hand"]]
    return pdf_df


def _build_page_elements(df_chunk: pd.DataFrame, title: str):
    styles = getSampleStyleSheet()
    elements = []
    # On each page, repeat section header for clarity
    elements.append(Paragraph(f"<b>{title}</b>", styles["Heading2"]))
    elements.append(Paragraph(datetime.now().strftime(DATE_FORMAT), styles["Normal"]))
    elements.append(Paragraph(" ", styles["Normal"]))

    headers = ["Type", "Product", "Barcode", "Loc", "Qty"]
    table_data = [headers] + df_chunk.values.tolist()
    table = Table(table_data, colWidths=COLUMN_WIDTHS)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
    ]))
    for i in range(1, len(table_data), 2):
        table.setStyle([("BACKGROUND", (0, i), (-1, i), colors.gainsboro)])
    elements.append(table)
    return elements


def export_pdf_paginated(
    move_up_df: pd.DataFrame,
    base_dir: str,
    timestamp: bool,
    prefix: Optional[str],
    auto_open: bool,
    items_per_page: int = 30,
):
    parts = [BASE_PDF_FILENAME]
    if timestamp:
        parts.append(datetime.now().strftime("%Y-%m-%d_%H-%M"))
    pdf_filename = "_".join(parts) + ".pdf"
    if prefix:
        prefix = sanitize_prefix(prefix)
        pdf_filename = f"{prefix}_{pdf_filename}"
    output_path = os.path.join(base_dir, pdf_filename)

    doc = SimpleDocTemplate(output_path, pagesize=letter)
    elements = []

    # Move-Up pages only
    mu_df = _prep_table_df(move_up_df)
    if not mu_df.empty:
        for start in range(0, len(mu_df), items_per_page):
            chunk = mu_df.iloc[start:start+items_per_page]
            elements += _build_page_elements(chunk, "Move-Up Inventory List")
            if start + items_per_page < len(mu_df):
                elements.append(PageBreak())

    doc.build(elements, onFirstPage=_draw_footer, onLaterPages=_draw_footer)

    if auto_open and os.name == "nt":
        try:
            os.startfile(output_path)
        except Exception:
            pass
    return output_path


def export_excel(
    move_up_df: pd.DataFrame,
    base_dir: str,
    timestamp: bool,
    prefix: Optional[str],
):
    parts = ["Sticker_Sheet_Filtered_Move_Up"]
    if timestamp:
        parts.append(datetime.now().strftime("%Y-%m-%d_%H-%M"))
    xlsx = "_".join(parts) + ".xlsx"
    if prefix:
        prefix = sanitize_prefix(prefix)
        xlsx = f"{prefix}_{xlsx}"
    out = os.path.join(base_dir, xlsx)
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        move_up_df.to_excel(w, sheet_name="Move_Up_Items", index=False)
    return out


# ------------------------------
# GUI
# ------------------------------

class MoveUpGUI:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title("Move-Up Live v3.7.0 — Minimal")
        self.root.geometry("1180x860")

        # State
        self.rooms_var = StringVar(value="Incoming Deliveries, Vault, Overstock")
        self.prefix_var = StringVar(value="")
        self.timestamp_var = BooleanVar(value=True)
        self.auto_open_var = BooleanVar(value=(os.name == "nt"))
        self.skip_sales_floor_var = BooleanVar(value=False)
        self.hide_removed_var = BooleanVar(value=True)
        self.page_items_var = IntVar(value=30)  # Avery sync

        self.room_alias_map: Dict[str, str] = {}
        self.raw_df: Optional[pd.DataFrame] = None     # loaded raw
        self.current_df: Optional[pd.DataFrame] = None # mapped + normalized df ready for compute
        self.col_mapping_override: Dict[str, str] = {} # src->dst mapping chosen by user
        self.moveup_df: Optional[pd.DataFrame] = None
        self.excluded_barcodes: set = set()

        # Filter state (independent of any window)
        self.selected_rooms: List[str] = []
        self.selected_brands: List[str] = []

        # Filters window handle
        self.filters_window: Optional[Toplevel] = None

        self._build_ui()

    # ---------- UI ----------
    def _build_ui(self):
        frm_top = ttk.Frame(self.root, padding=10)
        frm_top.pack(fill="x")
        ttk.Button(frm_top, text="Import File…", command=self.import_file).pack(side="left", padx=5)
        ttk.Button(frm_top, text="Map Columns…", command=self.map_columns_dialog).pack(side="left", padx=5)
        btn_update = ttk.Button(frm_top, text="★ Update Table", command=self._recompute_from_current)
        btn_update.pack(side="left", padx=12)
        ttk.Button(frm_top, text="Export PDF", command=self.do_export_pdf).pack(side="left", padx=5)
        ttk.Button(frm_top, text="Export Excel", command=self.do_export_xlsx).pack(side="left", padx=5)
        ttk.Button(frm_top, text="Filters…", command=self.open_filters_window).pack(side="left", padx=15)

        ttk.Checkbutton(frm_top, text="Timestamp", variable=self.timestamp_var).pack(side="left", padx=5)
        ttk.Checkbutton(frm_top, text="Auto-open PDF", variable=self.auto_open_var).pack(side="left", padx=5)
        ttk.Checkbutton(frm_top, text="Skip Sales-Floor Removal", variable=self.skip_sales_floor_var).pack(side="left", padx=10)
        ttk.Checkbutton(frm_top, text="Hide removed", variable=self.hide_removed_var).pack(side="left", padx=10)

        # Paging control
        frm_page = ttk.Frame(self.root)
        frm_page.pack(fill="x", padx=10)
        ttk.Label(frm_page, text="Items per page").pack(side="left")
        ttk.Spinbox(frm_page, from_=10, to=200, textvariable=self.page_items_var, width=6).pack(side="left", padx=6)

        self.status = StringVar(value="Ready.")
        ttk.Label(self.root, textvariable=self.status, anchor="w").pack(fill="x", padx=10)

        # Row count display
        self.rowcount_var = StringVar(value="Items loaded: 0")
        ttk.Label(self.root, textvariable=self.rowcount_var, anchor="w").pack(fill="x", padx=10)

        # Move-Up count display
        self.moveupcount_var = StringVar(value="Move-Up items: 0")
        ttk.Label(self.root, textvariable=self.moveupcount_var, anchor="w").pack(fill="x", padx=10)

        # Filters summary (compact)
        self.filters_summary_var = StringVar(value="Filters: default")
        ttk.Label(self.root, textvariable=self.filters_summary_var, anchor="w").pack(fill="x", padx=10, pady=(0, 5))

        # Results table
        self.tree = ttk.Treeview(self.root, columns=tuple(COLUMNS_TO_USE), show="headings", height=18)
        for col in COLUMNS_TO_USE:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150 if col != "Product Name" else 420, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)
        try:
            self.tree.tag_configure("excluded", foreground="#888")
        except Exception:
            pass

        # Remove controls
        frm_remove = ttk.Frame(self.root)
        frm_remove.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(frm_remove, text="Remove Selected", command=self._remove_selected).pack(side="left")
        ttk.Button(frm_remove, text="Toggle Remove", command=self._toggle_remove_selected).pack(side="left", padx=6)
        ttk.Button(frm_remove, text="Clear Removed", command=self._clear_removed).pack(side="left", padx=6)
        self.tree.bind("<Delete>", lambda e: self._remove_selected())
        self.tree.bind("<Double-1>", lambda e: self._toggle_remove_selected())

        # Diagnostics
        self.diag_var = StringVar(value="")
        ttk.Label(self.root, textvariable=self.diag_var, anchor="w", foreground="#555").pack(fill="x", padx=10, pady=(0, 10))

    # ---------- File Import + Mapping ----------
    def _update_rowcount(self, df: Optional[pd.DataFrame]):
        n = 0 if df is None else len(df)
        self.rowcount_var.set(f"Items loaded: {n}")

    def _update_moveupcount(self, df: Optional[pd.DataFrame]):
        n = 0 if df is None else len(df)
        self.moveupcount_var.set(f"Move-Up items: {n}")

    def import_file(self):
        path = filedialog.askopenfilename(title="Select Inventory File", filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if not path:
            return
        try:
            self.status.set(f"Loading {os.path.basename(path)}…")
            raw = load_raw_df(path)
            self.raw_df = raw
            # Try auto mapping immediately
            mapped, used = automap_columns(raw)
            self.current_df = mapped
            self.status.set(f"Loaded {len(mapped)} rows. Auto-mapped columns. You may adjust via 'Map Columns…'.")
            self._update_rowcount(mapped)
            self._post_load_housekeeping(mapped)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status.set(f"Error: {e}")

    def _post_load_housekeeping(self, df: pd.DataFrame):
        # Reset filter selections when a new file is loaded
        self.selected_rooms = []
        self.selected_brands = []
        self._recompute_from_current()

    def map_columns_dialog(self):
        if self.raw_df is None or self.raw_df.empty:
            messagebox.showinfo("Map Columns", "Import a file first.")
            return
        src_cols = list(self.raw_df.columns)

        # Pre-fill with auto mapping
        auto_df, auto_map = None, {}
        try:
            auto_df, auto_map = automap_columns(self.raw_df)
        except Exception:
            auto_map = {}

        win = Toplevel(self.root)
        win.title("Map Columns")
        win.geometry("620x360")
        ttk.Label(win, text="Choose which source column maps to each required field.").pack(anchor="w", padx=10, pady=10)

        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=10)
        combos = {}
        for i, target in enumerate(COLUMNS_TO_USE):
            ttk.Label(frame, text=target + ":").grid(row=i, column=0, sticky="e", pady=4)
            var = StringVar(value="")
            cb = ttk.Combobox(frame, textvariable=var, values=src_cols, width=40, state="readonly")
            # preselect based on auto_map (inverse lookup)
            pre = next((src for src, dst in auto_map.items() if dst == target), None)
            if pre and pre in src_cols:
                var.set(pre)
            cb.grid(row=i, column=1, sticky="w", pady=4)
            combos[target] = var

        btns = ttk.Frame(win)
        btns.pack(fill="x", pady=10)

        def _apply_mapping():
            # Build mapping src->target
            mapping = {}
            for target, var in combos.items():
                src = var.get().strip()
                if not src:
                    messagebox.showerror("Missing", f"Please choose a source for '{target}'.")
                    return
                mapping[src] = target
            try:
                df = self.raw_df.rename(columns=mapping)
                missing = [c for c in COLUMNS_TO_USE if c not in df.columns]
                if missing:
                    raise ValueError("After mapping, still missing: " + ", ".join(missing))
                # canonicalize types
                df["Package Barcode"] = df["Package Barcode"].astype("string").fillna("")
                df["Qty On Hand"] = pd.to_numeric(df["Qty On Hand"], errors="coerce").fillna(0).astype(int)
                for col in ["Product Name", "Brand", "Type", "Room"]:
                    df[col] = df[col].astype(str)
                self.col_mapping_override = mapping
                self.current_df = df
                self._update_rowcount(df)
                self._post_load_housekeeping(df)
                win.destroy()
                self.status.set("Column mapping applied.")
            except Exception as e:
                messagebox.showerror("Mapping Error", str(e))

        ttk.Button(btns, text="Apply", command=_apply_mapping).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="left", padx=6)

    # ---------- Room helpers ----------
    def _get_all_rooms_normalized(self, df: pd.DataFrame) -> List[str]:
        if df is None or df.empty or "Room" not in df.columns:
            return []
        df_norm = _normalize_rooms(df, self.room_alias_map)
        rooms = sorted(set(str(x).strip() for x in df_norm["Room"].dropna().astype(str).tolist()))
        return rooms

    def _rooms_list_normalized(self) -> Tuple[List[str], List[str]]:
        if self.current_df is None or self.current_df.empty or "Room" not in self.current_df.columns:
            return [], []
        raw = sorted(set(str(x).strip() for x in self.current_df["Room"].dropna().astype(str).tolist()))
        df_norm = _normalize_rooms(self.current_df, self.room_alias_map)
        norm = sorted(set(str(x).strip() for x in df_norm["Room"].dropna().astype(str).tolist()))
        return raw, norm

    def _inspect_rooms(self):
        raw, norm = self._rooms_list_normalized()
        win = Toplevel(self.root)
        win.title("Room Inspector")
        win.geometry("700x420")
        ttk.Label(win, text="Raw rooms from file").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        ttk.Label(win, text="Normalized rooms (after aliases)").grid(row=0, column=1, padx=10, pady=10, sticky="w")
        raw_box = Listbox(win, height=16, width=35)
        norm_box = Listbox(win, height=16, width=35)
        raw_box.grid(row=1, column=0, padx=10, sticky="n")
        norm_box.grid(row=1, column=1, padx=10, sticky="n")
        for r in raw:
            raw_box.insert(END, r)
        for n in norm:
            norm_box.insert(END, n)
        msg = ttk.Label(
            win,
            text="Tip: If your rooms here don’t match what you expect,\n"
                 "adjust aliases in code or use the room selections in Filters.",
            foreground="#555"
        )
        msg.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="w")

    def _default_candidate_rooms(self, df: pd.DataFrame) -> List[str]:
        """Compute a reasonable default set of candidate rooms when none are selected."""
        rooms = self._get_all_rooms_normalized(df)
        if not rooms:
            return []
        want = {"Incoming Deliveries", "Overstock"}
        room_set = set(rooms)
        if want.issubset(room_set):
            return [r for r in rooms if r in want]
        # Fallback: everything except sales floor
        out = []
        for r in rooms:
            r_l = r.strip().lower()
            if r_l not in SALES_FLOOR_ALIASES and r_l != "sales floor":
                out.append(r)
        return out or rooms

    # ---------- Filters dialog ----------
    def open_filters_window(self):
        if self.current_df is None or self.current_df.empty:
            messagebox.showinfo("Filters", "Import a file first.")
            return

        if self.filters_window is not None and self.filters_window.winfo_exists():
            self.filters_window.lift()
            self.filters_window.focus_set()
            return

        df = self.current_df
        all_rooms = self._get_all_rooms_normalized(df)
        all_brands = sorted(
            pd.Series(df["Brand"], dtype=str).dropna().astype(str).unique().tolist()
        ) if "Brand" in df.columns else []

        # Initial selections
        initial_rooms = self.selected_rooms or self._default_candidate_rooms(df)
        initial_brands = self.selected_brands[:]  # may be empty => ALL

        win = Toplevel(self.root)
        self.filters_window = win
        win.title("Filters")
        win.geometry("900x480")
        win.transient(self.root)
        win.grab_set()

        main = ttk.Frame(win, padding=10)
        main.pack(fill="both", expand=True)

        # Top split: Rooms (left) and Brands (right)
        top = ttk.Frame(main)
        top.pack(fill="both", expand=True)

        # ---- Rooms ----
        frm_rooms = ttk.LabelFrame(top, text="Rooms (multi-select)", padding=10)
        frm_rooms.pack(side="left", fill="both", expand=True, padx=(0, 5))

        rooms_listbox = Listbox(frm_rooms, selectmode=MULTIPLE, height=12, exportselection=False)
        rooms_listbox.pack(fill="both", expand=True)

        for r in all_rooms:
            rooms_listbox.insert(END, r)

        # Pre-select rooms
        if initial_rooms:
            for i, r in enumerate(all_rooms):
                if r in initial_rooms:
                    rooms_listbox.selection_set(i)

        rm_btns = ttk.Frame(frm_rooms)
        rm_btns.pack(fill="x", pady=6)

        def _room_select_all():
            try:
                rooms_listbox.selection_set(0, rooms_listbox.size()-1)
            except Exception:
                pass

        def _room_select_none():
            rooms_listbox.selection_clear(0, END)

        def _room_use_all_excl_sf():
            rooms_listbox.selection_clear(0, END)
            for i, r in enumerate(all_rooms):
                rl = r.strip().lower()
                if rl not in SALES_FLOOR_ALIASES and rl != "sales floor":
                    rooms_listbox.selection_set(i)

        ttk.Button(rm_btns, text="Select all", command=_room_select_all).pack(side="left")
        ttk.Button(rm_btns, text="Select none", command=_room_select_none).pack(side="left", padx=6)
        ttk.Button(rm_btns, text="Use all (excl. Sales Floor)", command=_room_use_all_excl_sf).pack(side="left", padx=6)
        ttk.Button(rm_btns, text="Inspect Rooms…", command=self._inspect_rooms).pack(side="left", padx=6)

        # ---- Brands ----
        frm_brand = ttk.LabelFrame(top, text="Brand Filter (multi-select; select none = ALL)", padding=10)
        frm_brand.pack(side="left", fill="both", expand=True, padx=(5, 0))

        brand_listbox = Listbox(frm_brand, selectmode=MULTIPLE, height=12, exportselection=False)
        brand_listbox.pack(fill="both", expand=True)

        for b in all_brands:
            brand_listbox.insert(END, b)

        # Pre-select brands (if any)
        if initial_brands:
            for i, b in enumerate(all_brands):
                if b in initial_brands:
                    brand_listbox.selection_set(i)

        br_btns = ttk.Frame(frm_brand)
        br_btns.pack(fill="x", pady=6)

        def _brand_select_all():
            try:
                brand_listbox.selection_set(0, brand_listbox.size()-1)
            except Exception:
                pass

        def _brand_select_none():
            brand_listbox.selection_clear(0, END)

        ttk.Button(br_btns, text="Select all", command=_brand_select_all).pack(side="left")
        ttk.Button(br_btns, text="Select none (ALL)", command=_brand_select_none).pack(side="left", padx=6)
        ttk.Button(br_btns, text="Update Table", command=self._recompute_from_current).pack(side="left", padx=6)

        # ---- Other Filters ----
        frm_other = ttk.LabelFrame(main, text="Other Filters", padding=10)
        frm_other.pack(fill="x", pady=(10, 0))

        ttk.Label(frm_other, text="Filename prefix").grid(row=0, column=0, sticky="e")
        ttk.Entry(frm_other, textvariable=self.prefix_var, width=24).grid(row=0, column=1, sticky="w")
        ttk.Checkbutton(frm_other, text="Skip Sales-Floor Removal", variable=self.skip_sales_floor_var).grid(row=0, column=2, padx=10, sticky="w")

        # ---- Buttons ----
        btns = ttk.Frame(main)
        btns.pack(fill="x", pady=10)

        def _apply_filters():
            # Rooms
            room_sel = [rooms_listbox.get(i) for i in rooms_listbox.curselection()]
            # If user selects nothing, we'll later fall back to default rooms when computing
            self.selected_rooms = room_sel

            # Brands
            brand_sel = [brand_listbox.get(i) for i in brand_listbox.curselection()]
            # Empty => ALL
            self.selected_brands = brand_sel

            # Update rooms_var for display
            if self.selected_rooms:
                self.rooms_var.set(", ".join(self.selected_rooms))
            else:
                self.rooms_var.set("")

            win.destroy()
            self.filters_window = None
            self._recompute_from_current()

        def _cancel_filters():
            win.destroy()
            self.filters_window = None

        ttk.Button(btns, text="Apply", command=_apply_filters).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=_cancel_filters).pack(side="left", padx=6)

    # ---------- Render ----------
    def _render_tree(self, df: pd.DataFrame):
        for i in self.tree.get_children():
            self.tree.delete(i)
        if df is None or df.empty:
            return
        for _, row in df.iterrows():
            vals = [row.get(c, "") for c in COLUMNS_TO_USE]
            bc = str(row.get("Package Barcode", ""))
            tags = ()
            if bc and (bc in self.excluded_barcodes) and not self.hide_removed_var.get():
                tags = ("excluded",)
            self.tree.insert("", "end", values=vals, tags=tags)

    # -------- Manual remove --------
    def _remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Remove", "Select row(s) first.")
            return
        idx_bar = COLUMNS_TO_USE.index("Package Barcode")
        removed = 0
        for iid in sel:
            vals = self.tree.item(iid, "values")
            if not vals or len(vals) <= idx_bar:
                continue
            bc = str(vals[idx_bar]).strip()
            if bc:
                self.excluded_barcodes.add(bc)
                removed += 1
        self._recompute_from_current()
        self.status.set(f"Removed {removed} item(s) this session.")

    def _toggle_remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx_bar = COLUMNS_TO_USE.index("Package Barcode")
        toggled = 0
        for iid in sel:
            vals = self.tree.item(iid, "values")
            if not vals or len(vals) <= idx_bar:
                continue
            bc = str(vals[idx_bar]).strip()
            if not bc:
                continue
            if bc in self.excluded_barcodes:
                self.excluded_barcodes.remove(bc)
            else:
                self.excluded_barcodes.add(bc)
            toggled += 1
        self._recompute_from_current()
        self.status.set(f"Toggled remove on {toggled} item(s).")

    def _clear_removed(self):
        self.excluded_barcodes.clear()
        self._recompute_from_current()
        self.status.set("Cleared manually removed items.")

    # ---------- Core recompute ----------
    def _recompute_from_current(self):
        df = self.current_df
        if df is None or df.empty:
            self._render_tree(pd.DataFrame(columns=COLUMNS_TO_USE))
            self._update_rowcount(None)
            self._update_moveupcount(None)
            self.status.set("No data loaded.")
            self.diag_var.set("")
            self.filters_summary_var.set("Filters: none (no data)")
            return

        # Rooms
        rooms = self.selected_rooms[:] if self.selected_rooms else self._default_candidate_rooms(df)
        self.rooms_var.set(", ".join(rooms))

        # Brands
        brand_filter = self.selected_brands[:] if self.selected_brands else ["ALL"]

        move_up_df, diag = compute_moveup_from_df(
            df,
            rooms,
            self.room_alias_map,
            brand_filter,
            skip_sales_floor=self.skip_sales_floor_var.get()
        )
        # Apply manual exclusions depending on hide/show
        if self.excluded_barcodes and self.hide_removed_var.get():
            move_up_df = move_up_df[~move_up_df["Package Barcode"].astype(str).isin(self.excluded_barcodes)].copy()
        self.moveup_df = move_up_df
        self._render_tree(move_up_df)
        self._update_moveupcount(move_up_df)

        brand_note = "ALL" if brand_filter == ["ALL"] else f"{len(brand_filter)} brand(s)"
        self.status.set(f"Loaded {len(df)} rows; {brand_note}; Move-Up {len(move_up_df)}")
        self.diag_var.set(
            f"Diagnostics — after dropna: {diag['after_dropna']}, after brand: {diag['after_brand']}, "
            f"after type(accessories removed): {diag['after_type']}, candidate pool: {diag['candidate_pool']}, "
            f"removed as on Sales Floor: {diag['removed_as_on_sf']}."
        )

        # Update filters summary line
        rooms_count = len(rooms) if rooms else 0
        brand_note_short = "ALL" if brand_filter == ["ALL"] else f"{len(brand_filter)} brands"
        skip_sf = "Yes" if self.skip_sales_floor_var.get() else "No"
        self.filters_summary_var.set(
            f"Filters — Rooms: {rooms_count} | Brands: {brand_note_short} | Skip SF: {skip_sf}"
        )

    # ---------- Export ----------
    def do_export_pdf(self):
        if self.moveup_df is None:
            messagebox.showwarning("No data", "Import + Update Table first.")
            return
        base_dir = os.getcwd()
        # Always drop manually excluded items from exports
        if self.excluded_barcodes:
            mu_use = self.moveup_df[~self.moveup_df["Package Barcode"].astype(str).isin(self.excluded_barcodes)].copy()
        else:
            mu_use = self.moveup_df.copy()
        p = export_pdf_paginated(
            move_up_df=mu_use,
            base_dir=base_dir,
            timestamp=self.timestamp_var.get(),
            prefix=self.prefix_var.get() or None,
            auto_open=self.auto_open_var.get(),
            items_per_page=int(self.page_items_var.get() or 30),
        )
        self.status.set(f"PDF saved: {os.path.basename(p)}")

    def do_export_xlsx(self):
        if self.moveup_df is None:
            messagebox.showwarning("No data", "Import + Update Table first.")
            return
        base_dir = os.getcwd()
        if self.excluded_barcodes:
            mu_use = self.moveup_df[~self.moveup_df["Package Barcode"].astype(str).isin(self.excluded_barcodes)].copy()
        else:
            mu_use = self.moveup_df.copy()
        p = export_excel(
            move_up_df=mu_use,
            base_dir=base_dir,
            timestamp=self.timestamp_var.get(),
            prefix=self.prefix_var.get() or None,
        )
        self.status.set(f"Excel saved: {os.path.basename(p)}")


# ------------------------------
# Main
# ------------------------------

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Move-Up Live v3.7.0 — Minimal")
    parser.add_argument("--headless", action="store_true", help="Run once and export, no GUI")
    parser.add_argument("--prefix", default=None, help="Filename prefix (headless)")
    parser.add_argument("--rooms", nargs="+", default=["Incoming Deliveries", "Vault", "Overstock"], help="Candidate rooms (headless)")
    parser.add_argument("--brands", nargs="*", help="Brand filter (omit for ALL)")
    parser.add_argument("--input", help="Excel/CSV path for headless mode")
    parser.add_argument("--skip-sales-floor", action="store_true", help="Do not remove items already on Sales Floor (headless)")
    parser.add_argument("--items-per-page", type=int, default=30, help="PDF items per page (headless)")

    args = parser.parse_args()

    if args.headless:
        try:
            if not args.input:
                print("--input is required in headless mode", file=sys.stderr)
                sys.exit(2)
            raw = load_raw_df(args.input)
            mapped, _ = automap_columns(raw)
            mu, diag = compute_moveup_from_df(
                mapped,
                args.rooms,
                {},
                args.brands or ["ALL"],
                skip_sales_floor=args.skip_sales_floor
            )
            base_dir = os.getcwd()
            excel = export_excel(mu, base_dir, True, args.prefix)
            pdf = export_pdf_paginated(
                mu,
                base_dir,
                True,
                args.prefix,
                auto_open=False,
                items_per_page=args.items_per_page
            )
            print(json.dumps(diag, indent=2))
            print(excel)
            print(pdf)
            sys.exit(0)
        except Exception as e:
            print(f"[ERROR] {e}")
            sys.exit(1)

    # GUI
    root = Tk()
    gui = MoveUpGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

