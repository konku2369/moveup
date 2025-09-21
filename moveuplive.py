"""
Move‑Up Live v3.4.0 Lite — GUI + (optional) Sweed API
Author: Konrad Kubica (+ ChatGPT)
Date: 2025‑09‑21

What's new
- **No history / database**: simpler, zero side‑effects, no dedupe-by-history logic.
- **Big "Update Table" button**: recompute on demand after you tweak brands/rooms/flags.
- Keeps diagnostics panel, brand multi‑select (Select none = ALL), Skip Sales‑Floor toggle,
  "Inspect Rooms…" and "Use all rooms (excl. Sales Floor)".
- Fully offline with Excel/CSV; Sweed API remains optional.
"""

import os
import re
import sys
import json
import time
import queue
import threading
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests
from requests.adapters import HTTPAdapter, Retry

# GUI
from tkinter import (
    Tk, Toplevel, StringVar, IntVar, BooleanVar, filedialog, messagebox,
    ttk
)
from tkinter import Listbox, MULTIPLE, END

# PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# ------------------------------
# Constants
# ------------------------------
COLUMNS_TO_USE = ["Type", "Brand", "Product Name", "Package Barcode", "Room", "Qty On Hand"]
COLUMN_WIDTHS = [50, 365, 70, 45, 35]
BASE_PDF_FILENAME = "Print_me_Filtered_Move_Up"
DATE_FORMAT = "%B %d, %Y"

ALT_NAME_CANDIDATES = {
    "Type": ["type", "product type", "category", "item type", "class"],
    "Brand": ["brand", "brand name", "manufacturer", "mfr"],
    "Product Name": ["product name", "product", "item name", "name", "title", "item"],
    "Package Barcode": ["barcode", "package barcode", "package id", "upc", "ean", "gtin", "Barcode", "metrc code", "package upc", "package ean"],
    "Room": ["room", "location", "stock location", "bin", "area", "warehouse location", "site location"],
    "Qty On Hand": ["available qty", "qty on hand", "quantity on hand", "on hand", "quantity", "qoh", "stock", "stock qty", "current quantity", "current qty"],
}

SALES_FLOOR_ALIASES = {"sales floor", "floor", "salesfloor", "front of house", "foh", "front", "front of shop", "retail"}

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
    wanted = mapping.get(target_key, [])
    for idx, lc in enumerate(lower_cols):
        if lc in wanted:
            return idx
    return None


def _auto_map_columns(df: pd.DataFrame) -> pd.DataFrame:
    lower_cols = _lower_strip_cols(df.columns)
    if all(col in df.columns for col in COLUMNS_TO_USE):
        out = df.copy()
    else:
        out = df.copy()
        for key in COLUMNS_TO_USE:
            if key in out.columns:
                continue
            idx = _find_source_for(key, lower_cols)
            if idx is not None:
                out.rename(columns={out.columns[idx]: key}, inplace=True)

    missing = [c for c in COLUMNS_TO_USE if c not in out.columns]
    if missing:
        raise ValueError(
            "Missing required column(s) after auto-mapping: " + ", ".join(missing)
        )

    out["Package Barcode"] = out["Package Barcode"].astype("string").fillna("")
    out["Qty On Hand"] = pd.to_numeric(out["Qty On Hand"], errors="coerce").fillna(0).astype(int)
    for col in ["Product Name", "Brand", "Type", "Room"]:
        out[col] = out[col].astype(str)
    return out


def _normalize_rooms(df: pd.DataFrame, user_aliases: dict):
    if "Room" not in df.columns:
        return df
    final_map = {k.casefold(): v for k, v in DEFAULT_ROOM_ALIASES.items()}
    final_map.update({(k or "").casefold(): v for k, v in (user_aliases or {}).items()})

    def _map_room(val):
        s = str(val).strip()
        return final_map.get(s.casefold(), s)

    df = df.copy()
    df["Room"] = df["Room"].apply(_map_room)
    return df

# --- Sweed-style export detection & file loading (for offline mode) ---

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


def load_inventory_df(original_file: str, sheet_name: str = "Inventory Adjustments") -> pd.DataFrame:
    ext = os.path.splitext(original_file)[1].lower()
    skiprows = 3 if _is_sweed_export(original_file, ext, sheet_name) else 0
    try:
        if ext == ".csv":
            df = pd.read_csv(original_file, skiprows=skiprows, dtype={"Barcode": "string", "Package Barcode": "string"})
        else:
            try:
                df = pd.read_excel(original_file, sheet_name=sheet_name, skiprows=skiprows, dtype={"Barcode": "string", "Package Barcode": "string"})
            except Exception:
                df = pd.read_excel(original_file, sheet_name=0, skiprows=skiprows, dtype={"Barcode": "string", "Package Barcode": "string"})
    except Exception as e:
        raise RuntimeError(f"Could not read file '{original_file}': {e}")

    df = _auto_map_columns(df)
    df = df.loc[:, [c for c in COLUMNS_TO_USE if c in df.columns]]
    df["Package Barcode"] = df["Package Barcode"].fillna("")
    return df

# ------------------------------
# Sweed API Client (optional)
# ------------------------------

@dataclass
class SweedConfig:
    base_url: str
    inventory_endpoint: str
    auth_endpoint: Optional[str] = None
    client_id: Optional[str] = None
    client_secret: Optional[str] = None
    api_key: Optional[str] = None
    token: Optional[str] = None
    page_size: int = 500


class SweedClient:
    def __init__(self, cfg: SweedConfig):
        self.cfg = cfg
        self.sess = requests.Session()
        retries = Retry(total=4, backoff_factor=0.5, status_forcelist=[429, 500, 502, 503, 504])
        self.sess.mount("https://", HTTPAdapter(max_retries=retries))
        self.sess.mount("http://", HTTPAdapter(max_retries=retries))
        self._token_expiry = datetime.now(timezone.utc)

    def _headers(self) -> Dict[str, str]:
        h = {"Accept": "application/json"}
        if self.cfg.api_key:
            h["Authorization"] = f"ApiKey {self.cfg.api_key}"
        elif self.cfg.token:
            h["Authorization"] = f"Bearer {self.cfg.token}"
        return h

    def _refresh_token_if_needed(self):
        if not self.cfg.auth_endpoint:
            return
        if datetime.now(timezone.utc) < self._token_expiry:
            return
        data = {
            "client_id": self.cfg.client_id,
            "client_secret": self.cfg.client_secret,
            "grant_type": "client_credentials",
        }
        r = self.sess.post(self.cfg.base_url + self.cfg.auth_endpoint, data=data, timeout=30)
        r.raise_for_status()
        js = r.json()
        self.cfg.token = js.get("access_token")
        ttl = int(js.get("expires_in", 3600))
        self._token_expiry = datetime.now(timezone.utc) + timedelta(seconds=ttl - 60)

    def fetch_inventory(self) -> pd.DataFrame:
        self._refresh_token_if_needed()
        records: List[dict] = []
        next_page = 1
        while True:
            params = {"page": next_page, "page_size": self.cfg.page_size}
            r = self.sess.get(self.cfg.base_url + self.cfg.inventory_endpoint, headers=self._headers(), params=params, timeout=60)
            r.raise_for_status()
            payload = r.json()
            items = payload.get("items") or payload.get("data") or payload
            if not isinstance(items, list):
                items = payload.get("results", [])
            records.extend(items)
            has_next = bool(payload.get("has_next")) or (len(items) == self.cfg.page_size)
            if not has_next:
                break
            next_page += 1
        if not records:
            return pd.DataFrame(columns=COLUMNS_TO_USE)
        raw = pd.DataFrame(records)
        rename_map = {
            "product_type": "Type",
            "brand": "Brand",
            "name": "Product Name",
            "barcode": "Package Barcode",
            "location": "Room",
            "available_qty": "Qty On Hand",
        }
        for k, v in rename_map.items():
            if k in raw.columns and v not in raw.columns:
                raw.rename(columns={k: v}, inplace=True)
        df = raw
        df = df[[c for c in df.columns if c in COLUMNS_TO_USE or c in rename_map.values()]]
        df = _auto_map_columns(df)
        return df

# ------------------------------
# Core filtering (with diagnostics & optional skip_sales_floor)
# ------------------------------

def compute_moveup_from_df(
    df: pd.DataFrame,
    candidate_rooms: List[str],
    lowstock_threshold: int,
    room_alias_overrides: Optional[Dict[str, str]] = None,
    brand_filter: Optional[List[str]] = None,
    skip_sales_floor: bool = False,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, int]]:
    diag = {
        "total_loaded": int(len(df) if df is not None else 0),
        "after_dropna": 0,
        "after_brand": 0,
        "after_type": 0,
        "candidate_pool": 0,
        "removed_as_on_sf": 0,
        "vault_low": 0,
        "move_up": 0,
    }
    if df is None or df.empty:
        return (
            pd.DataFrame(columns=COLUMNS_TO_USE),
            pd.DataFrame(columns=COLUMNS_TO_USE),
            pd.DataFrame(columns=COLUMNS_TO_USE),
            diag,
        )

    # Ensure strings
    df = df.copy()
    for c in ["Product Name", "Brand", "Package Barcode", "Room", "Type"]:
        if c in df.columns:
            df[c] = df[c].astype(str)

    df = df.dropna(subset=["Product Name", "Brand", "Package Barcode", "Room"]).copy()
    diag["after_dropna"] = int(len(df))

    # Brand filtering (ignore if includes ALL)
    if brand_filter and "ALL" not in [b.upper() for b in brand_filter]:
        df = df[df["Brand"].astype(str).isin(brand_filter)]
    diag["after_brand"] = int(len(df))

    df = _normalize_rooms(df, room_alias_overrides)

    # Remove accessories
    if "Type" in df.columns:
        mask_accessory = df["Type"].astype(str).str.strip().str.contains(r"accessor", case=False, na=False)
        df = df.loc[~mask_accessory].copy()
    diag["after_type"] = int(len(df))

    for col in ("Type", "Room"):
        if col in df.columns:
            df[col] = df[col].astype("category")

    room_lower = df["Room"].astype(str).str.strip().str.lower()
    sales_floor_mask = room_lower.eq("sales floor") | room_lower.isin(SALES_FLOOR_ALIASES)
    sales_floor = df.loc[sales_floor_mask, ["Brand", "Product Name"]].drop_duplicates()

    room_mask_candidates = df["Room"].isin(candidate_rooms)
    candidates = df.loc[room_mask_candidates, COLUMNS_TO_USE]
    diag["candidate_pool"] = int(len(candidates))

    if skip_sales_floor:
        filtered = candidates.copy()
        diag["removed_as_on_sf"] = 0
    else:
        merged = candidates.merge(
            sales_floor.assign(on_sf=1),
            on=["Brand", "Product Name"],
            how="left",
            indicator=False,
        )
        removed = merged["on_sf"].notna().sum()
        filtered = merged.loc[merged["on_sf"].isna()].drop(columns=["on_sf"])
        diag["removed_as_on_sf"] = int(removed)

    low_stock_df = filtered.loc[
        filtered["Room"].eq("Vault") & (filtered["Qty On Hand"] < lowstock_threshold)
    ].copy()
    diag["vault_low"] = int(len(low_stock_df))

    move_up_df = filtered.drop(index=low_stock_df.index, errors="ignore").copy()
    diag["move_up"] = int(len(move_up_df))

    sort_cols = [c for c in ["Type", "Brand", "Product Name"] if c in move_up_df.columns]
    if sort_cols:
        move_up_df.sort_values(by=sort_cols, inplace=True, kind="stable")
        low_stock_df.sort_values(by=sort_cols, inplace=True, kind="stable")

    return move_up_df, low_stock_df, df, diag

# ------------------------------
# Exports (PDF/Excel)
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


def _build_pdf_section(df, title):
    styles = getSampleStyleSheet()
    elements = [Paragraph(f"<b>{title}</b>", styles["Heading2"])]
    elements.append(Paragraph(datetime.now().strftime(DATE_FORMAT), styles["Normal"]))
    elements.append(Paragraph(" ", styles["Normal"]))

    pdf_df = df.loc[:, COLUMNS_TO_USE].copy()
    sort_cols = [c for c in ["Type", "Brand", "Product Name"] if c in pdf_df.columns]
    if sort_cols:
        pdf_df.sort_values(by=sort_cols, inplace=True, kind="stable")

    pdf_df["Room"] = pdf_df["Room"].astype(str).str.slice(0, 8).where(pdf_df["Room"].notna(), "")
    pdf_df["Type"] = pdf_df["Type"].astype(str).str.slice(0, 8).where(pdf_df["Type"].notna(), "")
    pdf_df["Product Name"] = pdf_df["Product Name"].astype(str).apply(lambda x: x if len(x) <= 75 else x[:72] + "...")
    pdf_df["Package Barcode"] = pdf_df["Package Barcode"].apply(lambda x: str(x)[-6:] if pd.notna(x) and str(x) else "")

    pdf_df = pdf_df[["Type", "Product Name", "Package Barcode", "Room", "Qty On Hand"]]

    headers = ["Type", "Product", "Barcode", "Loc", "Qty"]
    table_data = [headers] + pdf_df.values.tolist()

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


def export_pdf(move_up_df, low_stock_df, base_dir: str, include_lowstock: bool, timestamp: bool, prefix: Optional[str], auto_open: bool):
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
    elements += _build_pdf_section(move_up_df, "Move‑Up Inventory List")
    if include_lowstock and not low_stock_df.empty:
        from reportlab.platypus import PageBreak
        elements.append(PageBreak())
        elements += _build_pdf_section(low_stock_df, "Vault Low Stock")
    doc.build(elements, onFirstPage=_draw_footer, onLaterPages=_draw_footer)

    if auto_open and os.name == "nt":
        try:
            os.startfile(output_path)
        except Exception:
            pass
    return output_path


def export_excel(move_up_df, low_stock_df, base_dir: str, timestamp: bool, prefix: Optional[str]):
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
        low_stock_df.to_excel(w, sheet_name="Vault_Low_Stock", index=False)
    return out

# ------------------------------
# GUI
# ------------------------------

class MoveUpGUI:
    def __init__(self, root: Tk, client: Optional['SweedClient'] = None):
        self.root = root
        self.client = client
        self.root.title("Move‑Up Live v3.4.0 Lite — Sweed")
        self.root.geometry("1180x800")

        # State vars
        self.rooms_var = StringVar(value="Incoming Deliveries, Vault, Overstock")
        self.lowstock_var = IntVar(value=5)
        self.prefix_var = StringVar(value="")
        self.include_lowstock_var = BooleanVar(value=False)
        self.timestamp_var = BooleanVar(value=True)
        self.auto_open_var = BooleanVar(value=(os.name == "nt"))
        self.autorefresh_var = BooleanVar(value=False)
        self.autorefresh_secs_var = IntVar(value=300)
        self.skip_sales_floor_var = BooleanVar(value=False)

        self.room_alias_map: Dict[str, str] = {}
        self.current_df: Optional[pd.DataFrame] = None
        self.moveup_df: Optional[pd.DataFrame] = None
        self.low_df: Optional[pd.DataFrame] = None

        # Brand filter UI handle
        self.brand_listbox: Optional[Listbox] = None

        # Worker infra
        self.work_q: queue.Queue = queue.Queue()
        self.result_q: queue.Queue = queue.Queue()
        self.worker_thread = threading.Thread(target=self._worker_loop, daemon=True)
        self.worker_thread.start()

        self._build_ui()

    def _build_ui(self):
        frm_top = ttk.Frame(self.root, padding=10)
        frm_top.pack(fill="x")

        ttk.Button(frm_top, text="Refresh Now (API)", command=self.refresh_api).pack(side="left", padx=5)
        ttk.Button(frm_top, text="Import File…", command=self.import_file).pack(side="left", padx=5)

        # Big update button
        btn_update = ttk.Button(frm_top, text="★ Update Table", command=self._recompute_from_current)
        btn_update.pack(side="left", padx=12)
        btn_update.configure(style="Accent.TButton")
        try:
            # On some ttk themes, "Accent.TButton" exists (e.g., Tk 8.6 on Windows 11).
            style = ttk.Style(self.root)
            style.configure("Accent.TButton", padding=6)
        except Exception:
            pass

        ttk.Button(frm_top, text="Export PDF", command=self.do_export_pdf).pack(side="left", padx=5)
        ttk.Button(frm_top, text="Export Excel", command=self.do_export_xlsx).pack(side="left", padx=5)

        ttk.Checkbutton(frm_top, text="Include Vault Low‑Stock", variable=self.include_lowstock_var).pack(side="left", padx=15)
        ttk.Checkbutton(frm_top, text="Timestamp", variable=self.timestamp_var).pack(side="left", padx=5)
        ttk.Checkbutton(frm_top, text="Auto‑open PDF", variable=self.auto_open_var).pack(side="left", padx=5)
        ttk.Checkbutton(frm_top, text="Skip Sales‑Floor Removal", variable=self.skip_sales_floor_var).pack(side="left", padx=10)

        frm_mid = ttk.LabelFrame(self.root, text="Filters", padding=10)
        frm_mid.pack(fill="x", padx=10, pady=10)
        ttk.Label(frm_mid, text="Rooms (comma‑sep)").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm_mid, textvariable=self.rooms_var, width=60).grid(row=0, column=1, sticky="w", columnspan=3)
        ttk.Label(frm_mid, text="Vault low‑stock <").grid(row=0, column=4, sticky="e")
        ttk.Spinbox(frm_mid, from_=0, to=999, textvariable=self.lowstock_var, width=6).grid(row=0, column=5, sticky="w")
        ttk.Label(frm_mid, text="Filename prefix").grid(row=0, column=6, sticky="e")
        ttk.Entry(frm_mid, textvariable=self.prefix_var, width=20).grid(row=0, column=7, sticky="w")

        ttk.Label(frm_mid, text="Auto‑refresh (s)").grid(row=1, column=0, sticky="w", pady=(8,0))
        ttk.Spinbox(frm_mid, from_=30, to=3600, textvariable=self.autorefresh_secs_var, width=8).grid(row=1, column=1, sticky="w", pady=(8,0))
        ttk.Checkbutton(frm_mid, text="Enable Auto‑refresh", variable=self.autorefresh_var, command=self._toggle_autorefresh).grid(row=1, column=2, sticky="w", pady=(8,0))
        ttk.Button(frm_mid, text="Use all rooms (excl. Sales Floor)", command=self._use_all_rooms).grid(row=1, column=3, padx=10, pady=(8,0))
        ttk.Button(frm_mid, text="Inspect Rooms…", command=self._inspect_rooms).grid(row=1, column=4, padx=10, pady=(8,0))

        # Brand filter listbox
        frm_brand = ttk.LabelFrame(self.root, text="Brand Filter (multi‑select or select none = ALL)", padding=10)
        frm_brand.pack(fill="x", padx=10, pady=5)
        self.brand_listbox = Listbox(frm_brand, selectmode=MULTIPLE, height=6, exportselection=False)
        self.brand_listbox.pack(fill="x")
        # Start with ALL sentinel until data comes in
        self.brand_listbox.insert(END, "ALL")
        self.brand_listbox.selection_set(0)
        # Auto‑update table when brand selection changes (still supported)
        self.brand_listbox.bind('<<ListboxSelect>>', lambda e: None)  # We'll rely on big update; keep quiet on select

        # Controls under the list
        btn_row = ttk.Frame(frm_brand)
        btn_row.pack(fill="x", pady=6)
        ttk.Button(btn_row, text="Select none", command=self._brand_select_none).pack(side="left")
        ttk.Button(btn_row, text="Update Table", command=self._recompute_from_current).pack(side="left", padx=8)

        self.status = StringVar(value="Ready.")
        ttk.Label(self.root, textvariable=self.status, anchor="w").pack(fill="x", padx=10)

        self.tree = ttk.Treeview(self.root, columns=("Type","Brand","Product Name","Package Barcode","Room","Qty On Hand"), show="headings", height=18)
        for col in COLUMNS_TO_USE:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150 if col != "Product Name" else 420, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        # Diagnostics label
        self.diag_var = StringVar(value="")
        ttk.Label(self.root, textvariable=self.diag_var, anchor="w", foreground="#555").pack(fill="x", padx=10, pady=(0,10))

    # --------------- Worker ---------------
    def _worker_loop(self):
        while True:
            task = self.work_q.get()
            if task is None:
                break
            kind = task.get("kind")
            try:
                if kind == "api_refresh":
                    if not self.client:
                        raise RuntimeError("API client not configured. Use Import File or provide --config.")
                    df = self.client.fetch_inventory()
                elif kind == "import_file":
                    path = task["path"]
                    df = load_inventory_df(path)
                else:
                    df = pd.DataFrame(columns=COLUMNS_TO_USE)
                self.result_q.put({"ok": True, "df": df, "kind": kind})
            except Exception as e:
                self.result_q.put({ "ok": False, "err": str(e), "kind": kind })

    # --------------- Actions ---------------
    def refresh_api(self):
        self.status.set("Pulling from API…")
        self.work_q.put({ "kind": "api_refresh" })
        self.root.after(100, self._poll_results)

    def import_file(self):
        path = filedialog.askopenfilename(title="Select Inventory File", filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if not path:
            return
        self.status.set(f"Loading {os.path.basename(path)}…")
        self.work_q.put({ "kind": "import_file", "path": path })
        self.root.after(100, self._poll_results)

    def _poll_results(self):
        try:
            res = self.result_q.get_nowait()
        except queue.Empty:
            self.root.after(150, self._poll_results)
            return
        if not res.get("ok"):
            self.status.set(f"Error: {res.get('err')}")
            messagebox.showerror("Error", res.get("err"))
            return
        df = res["df"]
        self.current_df = df
        # refresh brand list
        self._update_brand_list(df)
        # don't auto recompute; let the user hit Update
        self.status.set(f"Loaded {len(df)} rows. Pick brands/rooms, then click Update Table.")

    def _update_brand_list(self, df: pd.DataFrame):
        if df is None or df.empty or self.brand_listbox is None:
            return
        brands = []
        if "Brand" in df.columns:
            brands = sorted(pd.Series(df["Brand"], dtype=str).dropna().astype(str).unique().tolist())
        self.brand_listbox.delete(0, END)
        self.brand_listbox.insert(END, "ALL")
        for b in brands:
            self.brand_listbox.insert(END, b)
        # Default to ALL
        self.brand_listbox.selection_set(0)

    def _render_tree(self, df: pd.DataFrame):
        for i in self.tree.get_children():
            self.tree.delete(i)
        if df is None or df.empty:
            return
        for _, row in df.iterrows():
            vals = [row.get(c, "") for c in COLUMNS_TO_USE]
            self.tree.insert("", "end", values=vals)

    # --- Brand handlers ---
    def _brand_select_none(self):
        if not self.brand_listbox: return
        self.brand_listbox.selection_clear(0, END)

    def _rooms_list_normalized(self) -> Tuple[List[str], List[str]]:
        if self.current_df is None or self.current_df.empty or "Room" not in self.current_df.columns:
            return [], []
        raw = sorted(set(str(x).strip() for x in self.current_df["Room"].dropna().astype(str).tolist()))
        df_norm = _normalize_rooms(self.current_df, self.room_alias_map)
        norm = sorted(set(str(x).strip() for x in df_norm["Room"].dropna().astype(str).tolist()))
        return raw, norm

    def _use_all_rooms(self):
        raw, norm = self._rooms_list_normalized()
        if not norm:
            messagebox.showinfo("Rooms", "No rooms detected yet. Import a file first.")
            return
        # Exclude sales floor aliases
        norm_lc = [r for r in norm if r.strip().lower() not in SALES_FLOOR_ALIASES and r.strip().lower() != "sales floor"]
        self.rooms_var.set(", ".join(norm_lc))

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
        msg = ttk.Label(win, text="Tip: If your rooms here don’t match the ‘Rooms’ field in Filters,\nclick ‘Use all rooms (excl. Sales Floor)’ or edit the list manually.", foreground="#555")
        msg.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="w")

    def _recompute_from_current(self):
        df = self.current_df
        if df is None or df.empty:
            # Clear table too
            self._render_tree(pd.DataFrame(columns=COLUMNS_TO_USE))
            self.status.set("No data loaded.")
            self.diag_var.set("")
            return
        rooms = [s.strip() for s in self.rooms_var.get().split(",") if s.strip()]
        sels = [self.brand_listbox.get(i) for i in self.brand_listbox.curselection()] if self.brand_listbox else []
        brand_filter = sels or ["ALL"]
        move_up_df, low_df, _norm, diag = compute_moveup_from_df(
            df, rooms, self.lowstock_var.get(), self.room_alias_map, brand_filter,
            skip_sales_floor=self.skip_sales_floor_var.get()
        )
        self.moveup_df, self.low_df = move_up_df, low_df
        self._render_tree(move_up_df)
        brand_note = "ALL" if brand_filter == ["ALL"] else f"{len(brand_filter)} brand(s)"
        self.status.set(f"Loaded {len(df)} rows; {brand_note}; Move‑Up {len(move_up_df)} | Low {len(low_df)}")
        # Diag summary
        self.diag_var.set(
            f"Diagnostics — after dropna: {diag['after_dropna']}, after brand: {diag['after_brand']}, "
            f"after type(accessories removed): {diag['after_type']}, candidate pool: {diag['candidate_pool']}, "
            f"removed as on Sales Floor: {diag['removed_as_on_sf']}, Vault Low-Stock: {diag['vault_low']}."
        )

    def _toggle_autorefresh(self):
        if self.autorefresh_var.get():
            self._schedule_autorefresh()

    def _schedule_autorefresh(self):
        if not self.autorefresh_var.get():
            return
        if self.client:
            self.refresh_api()
        secs = max(30, int(self.autorefresh_secs_var.get() or 300))
        self.root.after(secs * 1000, self._schedule_autorefresh)

    def do_export_pdf(self):
        if self.moveup_df is None:
            messagebox.showwarning("No data", "Pull or import data first, then click Update Table.")
            return
        base_dir = os.getcwd()
        low_df_use = self.low_df if self.low_df is not None else pd.DataFrame()
        p = export_pdf(
            move_up_df=self.moveup_df,
            low_stock_df=low_df_use,
            base_dir=base_dir,
            include_lowstock=self.include_lowstock_var.get(),
            timestamp=self.timestamp_var.get(),
            prefix=self.prefix_var.get() or None,
            auto_open=self.auto_open_var.get(),
        )
        self.status.set(f"PDF saved: {os.path.basename(p)}")

    def do_export_xlsx(self):
        if self.moveup_df is None:
            messagebox.showwarning("No data", "Pull or import data first, then click Update Table.")
            return
        base_dir = os.getcwd()
        low_df_use = self.low_df if self.low_df is not None else pd.DataFrame()
        p = export_excel(
            move_up_df=self.moveup_df,
            low_stock_df=low_df_use,
            base_dir=base_dir,
            timestamp=self.timestamp_var.get(),
            prefix=self.prefix_var.get() or None,
        )
        self.status.set(f"Excel saved: {os.path.basename(p)}")

# ------------------------------
# Config loader & Main
# ------------------------------

def load_config_from_json(path: str) -> 'SweedConfig':
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return SweedConfig(
        base_url=data["base_url"],
        inventory_endpoint=data["inventory_endpoint"],
        auth_endpoint=data.get("auth_endpoint"),
        client_id=data.get("client_id"),
        client_secret=data.get("client_secret"),
        api_key=data.get("api_key"),
        token=data.get("token"),
        page_size=int(data.get("page_size", 500)),
    )


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Move‑Up Live v3.4.0 Lite (Sweed API + GUI)")
    parser.add_argument("--config", help="Path to sweed_config.json (base_url, endpoints, credentials)")
    parser.add_argument("--headless", action="store_true", help="Run once and export, no GUI")
    parser.add_argument("--prefix", default=None, help="Filename prefix (headless)")
    parser.add_argument("--include-lowstock", action="store_true", help="Include low‑stock section in PDF (headless)")
    parser.add_argument("--rooms", nargs="+", default=["Incoming Deliveries", "Vault", "Overstock"], help="Candidate rooms (headless)")
    parser.add_argument("--lowstock-threshold", type=int, default=5, help="Vault threshold (headless)")
    parser.add_argument("--brands", nargs="*", help="Brand filter (omit for ALL)")
    parser.add_argument("--input", help="Optional Excel/CSV path for headless + offline mode")
    parser.add_argument("--skip-sales-floor", action="store_true", help="Do not remove items already on Sales Floor (headless)")

    args = parser.parse_args()

    client = None
    if args.config and os.path.exists(args.config):
        cfg = load_config_from_json(args.config)
        client = SweedClient(cfg)

    if args.headless:
        # Prefer API if configured; else load from file if provided
        if client is not None:
            df = client.fetch_inventory()
        elif args.input:
            df = load_inventory_df(args.input)
        else:
            print("No API config and no --input provided; nothing to do.")
            sys.exit(2)
        brand_filter = args.brands if args.brands else ["ALL"]
        mu, low, _, diag = compute_moveup_from_df(df, args.rooms, args.lowstock_threshold, {}, brand_filter, skip_sales_floor=args.skip_sales_floor)
        base_dir = os.getcwd()
        excel = export_excel(mu, low, base_dir, True, args.prefix)
        pdf = export_pdf(mu, low, base_dir, args.include_lowstock, True, args.prefix, auto_open=False)
        print(json.dumps(diag, indent=2))
        print(excel)
        print(pdf)
        return

    # GUI path
    root = Tk()
    gui = MoveUpGUI(root, client)
    root.mainloop()


if __name__ == "__main__":
    main()
