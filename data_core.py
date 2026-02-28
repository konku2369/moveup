# data_core.py
import os
import re
import csv
from io import TextIOWrapper
from typing import Dict, List, Optional, Tuple

import pandas as pd


# ------------------------------
# Single source constants
# ------------------------------
APP_VERSION = "3.10.0"

COLUMNS_TO_USE = ["Type", "Brand", "Product Name", "Package Barcode", "Room", "Qty On Hand"]
AUDIT_OPTIONAL_FIELDS = ["Distributor", "Store", "Size", "Received Date"]

TYPE_TRUNC_LEN = 7  # single source (GUI + PDF)

ALT_NAME_CANDIDATES = {
    "Type": ["type", "product type", "category", "item type", "class"],
    "Brand": ["brand", "brand name", "manufacturer", "mfr"],
    "Product Name": ["product name", "product", "item name", "name", "title", "item"],
    # NOTE: We still support legacy "barcode" candidates for OTHER barcode columns,
    # but METRC mapping is handled STRICTLY via detect_metrc_source_column().
    "Package Barcode": [
        "package barcode", "package id", "upc", "ean", "gtin",
        "barcode", "package upc", "package ean"
    ],
    "Room": ["room", "location", "stock location", "bin", "area", "warehouse location", "site location"],
    "Qty On Hand": [
        "available qty", "qty on hand", "quantity on hand", "on hand", "quantity", "qoh", "stock",
        "stock qty", "current quantity", "current qty"
    ],
    "Distributor": ["distributor", "vendor", "supplier", "producer", "wholesaler"],
    "Store": ["store", "store name", "storename", "location name", "site"],
    "Size": ["size", "product size", "package size", "unit size", "net weight", "weight", "volume"],
    "Received Date": [
        "received date", "receipt date", "date received", "reception date",
        "receive date", "date of receipt", "received on", "arrival date",
        "intake date", "package date", "packaged date", "harvest date",
    ],
}

SALES_FLOOR_ALIASES = {
    "sales floor", "floor", "salesfloor", "front of house",
    "foh", "front", "front of shop", "retail"
}


# ------------------------------
# Small utilities
# ------------------------------
def sanitize_prefix(pfx: str) -> str:
    if not pfx:
        return pfx
    pfx = pfx.strip()
    pfx = re.sub(r'[\\/:*?"<>|]+', "_", pfx)
    pfx = re.sub(r"\s+", "_", pfx)
    return pfx


def _lower_strip_cols(columns) -> List[str]:
    return [str(c).strip().lower() for c in columns]


def _find_source_for(target_key: str, lower_cols: List[str], mapping=ALT_NAME_CANDIDATES) -> Optional[int]:
    wanted = [w.strip().lower() for w in mapping.get(target_key, [])]
    for idx, lc in enumerate(lower_cols):
        if lc in wanted:
            return idx
    return None


def _build_room_map(user_aliases: dict) -> Dict[str, str]:
    if not user_aliases:
        return {}
    return {(k or "").casefold(): v for k, v in user_aliases.items()}


def normalize_rooms(df: pd.DataFrame, user_aliases: dict) -> pd.DataFrame:
    if df is None or df.empty or "Room" not in df.columns:
        return df
    out = df.copy()
    out["Room"] = out["Room"].astype(str).str.strip()
    norm_map = _build_room_map(user_aliases)
    if norm_map:
        out["Room"] = out["Room"].map(lambda v: norm_map.get(str(v).casefold(), v))
    return out


def windows_unblock_file(path: str):
    if os.name != "nt":
        return
    try:
        ads_path = path + ":Zone.Identifier"
        if os.path.exists(ads_path):
            os.remove(ads_path)
    except Exception:
        pass


def _read_csv_smart(path: str, skiprows: int) -> pd.DataFrame:
    def _attempt(encoding: str) -> pd.DataFrame:
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


def is_sweed_export(original_file: str, ext: str, sheet_name: str) -> bool:
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


def sort_with_backstock_priority(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty or "Room" not in df.columns:
        return df
    out = df.copy()
    out["__room_priority"] = out["Room"].astype(str).str.strip().str.lower().apply(
        lambda r: 0 if r == "backstock" else 1
    )
    sort_cols = ["__room_priority"]
    for col in ["Type", "Brand", "Product Name"]:
        if col in out.columns:
            sort_cols.append(col)

    out.sort_values(by=sort_cols, inplace=True, kind="stable")
    out.drop(columns=["__room_priority"], inplace=True)
    return out


def ellipses(s: str, n: int) -> str:
    s = str(s)
    return s if len(s) <= n else s[: max(0, n - 3)] + "..."

def aggregate_split_packages_by_room(df: pd.DataFrame) -> pd.DataFrame:
    """
    Option A: If the same METRC (Package Barcode) appears multiple times, aggregate
    duplicates that are the SAME product identity within the SAME room.

    Result: one row per (METRC, Room, Type, Brand, Product Name), Qty summed.

    This keeps split-lot-by-room visible, but eliminates duplicate rows that are
    just export duplication/noise.
    """
    if df is None or df.empty:
        return df

    required = ["Package Barcode", "Room", "Type", "Brand", "Product Name", "Qty On Hand"]
    for c in required:
        if c not in df.columns:
            # If it’s missing, do nothing—caller can handle mapping issues elsewhere.
            return df

    out = df.copy()

    # Normalize for grouping stability
    out["Package Barcode"] = out["Package Barcode"].astype("string").fillna("").str.strip()
    out["Room"] = out["Room"].astype(str).fillna("").str.strip()
    out["Type"] = out["Type"].astype(str).fillna("").str.strip()
    out["Brand"] = out["Brand"].astype(str).fillna("").str.strip()
    out["Product Name"] = out["Product Name"].astype(str).fillna("").str.strip()
    out["Qty On Hand"] = pd.to_numeric(out["Qty On Hand"], errors="coerce").fillna(0).astype(int)

    # Drop blanks for METRC because they are not a real package identifier
    out = out[out["Package Barcode"] != ""].copy()

    group_cols = ["Package Barcode", "Room", "Type", "Brand", "Product Name"]

    agg = (
        out.groupby(group_cols, as_index=False, sort=False)["Qty On Hand"]
        .sum()
    )

    # Re-attach any extra columns (e.g. Received Date) that aren't part of the
    # groupby key or the aggregated value. Take the first value per group.
    extra_cols = [c for c in out.columns if c not in group_cols and c != "Qty On Hand"]
    if extra_cols:
        first_vals = out.groupby(group_cols, as_index=False, sort=False)[extra_cols].first()
        agg = agg.merge(first_vals, on=group_cols, how="left")

    # Keep your normal ordering preference
    agg = sort_with_backstock_priority(agg)

    return agg


# ------------------------------
# STRICT METRC detection
# ------------------------------
def detect_metrc_source_column(df: pd.DataFrame) -> Optional[str]:
    """
    Strictly detect the METRC column.
    We ONLY accept columns that contain 'metrc' AND one of these tokens:
    code, id, tag, barcode, package

    Examples accepted:
      - "METRC Code"
      - "Metrc Package ID"
      - "METRC Barcode"
      - "METRC Tag"
    """
    if df is None or df.empty:
        return None

    required = "metrc"
    tokens = ("code", "id", "tag", "barcode", "package")

    for col in df.columns:
        name = str(col).strip().lower()
        if required in name and any(t in name for t in tokens):
            return col

    return None


# ------------------------------
# Column mapping
# ------------------------------
def automap_columns(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """
    Force Package Barcode to come from the METRC column.
    Maps required fields + optional audit fields if present.
    """
    lower_cols = _lower_strip_cols(df.columns)
    out = df.copy()
    rename_map: Dict[str, str] = {}

    metrc_src = detect_metrc_source_column(out)
    if not metrc_src:
        raise ValueError(
            "No METRC column found.\n"
            "Expected something like: 'METRC Code', 'METRC Package ID', or 'METRC Barcode'."
        )

    # Internal legacy alias: keep column name as Package Barcode for app compatibility.
    rename_map[metrc_src] = "Package Barcode"

    # Map remaining REQUIRED columns via candidates
    for key in COLUMNS_TO_USE:
        if key == "Package Barcode":
            continue
        if key in out.columns:
            continue
        idx = _find_source_for(key, lower_cols)
        if idx is not None:
            src_col = out.columns[idx]
            if src_col not in rename_map:
                rename_map[src_col] = key

    # Map OPTIONAL audit columns if present
    for opt in AUDIT_OPTIONAL_FIELDS:
        if opt in out.columns:
            continue
        idx = _find_source_for(opt, lower_cols)
        if idx is not None:
            src_col = out.columns[idx]
            if src_col not in rename_map:
                rename_map[src_col] = opt

    out = out.rename(columns=rename_map)

    missing = [c for c in COLUMNS_TO_USE if c not in out.columns]
    if missing:
        raise ValueError("Missing required column(s) after mapping: " + ", ".join(missing))

    # Normalize types
    out["Package Barcode"] = out["Package Barcode"].astype("string").fillna("").astype(str).str.strip()
    out["Qty On Hand"] = pd.to_numeric(out["Qty On Hand"], errors="coerce").fillna(0).astype(int)

    for col in ["Product Name", "Brand", "Type", "Room"]:
        if col in out.columns:
            out[col] = out[col].astype(str)

    if "Distributor" in out.columns:
        out["Distributor"] = out["Distributor"].astype(str).fillna("").str.strip()
    if "Store" in out.columns:
        out["Store"] = out["Store"].astype(str).fillna("").str.strip()
    if "Size" in out.columns:
        out["Size"] = out["Size"].astype(str).fillna("").str.strip()

    # Also catch the raw "Reception Date" column name if it wasn't remapped
    # (e.g. if automap missed it due to skiprows issues on Sweed exports)
    if "Reception Date" in out.columns and "Received Date" not in out.columns:
        out = out.rename(columns={"Reception Date": "Received Date"})

    if "Received Date" in out.columns:
        # Normalize to date-only string (YYYY-MM-DD), blank if unparseable.
        # Explicitly handle MM/DD/YYYY HH:MM:SS AM/PM format from Sweed exports.
        def _to_date_str(v):
            s = str(v).strip().lstrip("'")  # strip hidden leading apostrophe (Excel text-force trick)
            if s in ("", "nan", "NaT", "None"):
                return ""
            try:
                # Sweed export format: MM/DD/YYYY HH:MM:SS AM/PM (24h hour with AM/PM marker)
                return pd.to_datetime(s, format="%m/%d/%Y %H:%M:%S %p").strftime("%Y-%m-%d")
            except Exception:
                try:
                    return pd.to_datetime(s, format="mixed").strftime("%Y-%m-%d")
                except Exception:
                    return ""
        out["Received Date"] = out["Received Date"].map(_to_date_str)

    return out, rename_map


# ------------------------------
# Loading
# ------------------------------
def load_raw_df(original_file: str, sheet_name: str = "Inventory Adjustments") -> pd.DataFrame:
    windows_unblock_file(original_file)
    ext = os.path.splitext(original_file)[1].lower()
    skiprows = 3 if is_sweed_export(original_file, ext, sheet_name) else 0

    if ext == ".csv":
        return _read_csv_smart(original_file, skiprows=skiprows)

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
    type_filter: Optional[List[str]] = None,
    skip_sales_floor: bool = False,
) -> Tuple[pd.DataFrame, Dict[str, int]]:
    diag = {
        "total_loaded": int(len(df) if df is not None else 0),
        "after_dropna": 0,
        "after_brand": 0,
        "after_type_filter": 0,
        "after_type": 0,
        "candidate_pool": 0,
        "removed_as_on_sf": 0,
        "move_up": 0,
    }

    if df is None or df.empty:
        return pd.DataFrame(columns=COLUMNS_TO_USE), diag

    work = df.copy()
    for c in ["Product Name", "Brand", "Package Barcode", "Room", "Type"]:
        if c in work.columns:
            work[c] = work[c].astype(str)

    work = work.dropna(subset=["Product Name", "Brand", "Package Barcode", "Room"]).copy()
    diag["after_dropna"] = int(len(work))

    if brand_filter:
        bf = [str(b).strip() for b in brand_filter if str(b).strip()]
        is_all = any(b.upper() == "ALL" for b in bf)
        if not is_all:
            work = work[work["Brand"].astype(str).isin(bf)]
    diag["after_brand"] = int(len(work))

    if type_filter and "Type" in work.columns:
        tf = [str(t).strip() for t in type_filter if str(t).strip()]
        is_all_type = any(t.upper() == "ALL" for t in tf)
        if not is_all_type:
            work = work[work["Type"].astype(str).isin(tf)]
    diag["after_type_filter"] = int(len(work))

    work = normalize_rooms(work, room_alias_overrides or {})

    # Exclude accessories from Move-Up logic
    if "Type" in work.columns:
        mask_accessory = work["Type"].astype(str).str.contains(r"accessor", case=False, na=False)
        work = work.loc[~mask_accessory].copy()
    diag["after_type"] = int(len(work))

    room_lower = work["Room"].astype(str).str.strip().str.lower()
    if not skip_sales_floor:
        sf_mask = room_lower.eq("sales floor") | room_lower.isin(SALES_FLOOR_ALIASES)
        sales_floor = work.loc[sf_mask, ["Brand", "Product Name"]].drop_duplicates()
    else:
        sales_floor = pd.DataFrame(columns=["Brand", "Product Name"])

    candidate_set = {str(r).strip() for r in (candidate_rooms or [])}
    # Keep all columns (including extras like Received Date) so they survive into move_up_df
    candidates = work.loc[work["Room"].astype(str).str.strip().isin(candidate_set)].copy()
    diag["candidate_pool"] = int(len(candidates))

    if skip_sales_floor or sales_floor.empty:
        move_up_df = candidates.copy()
        diag["removed_as_on_sf"] = 0
    else:
        merged = candidates.merge(sales_floor.assign(on_sf=1), on=["Brand", "Product Name"], how="left")
        removed = merged["on_sf"].notna().sum()
        move_up_df = merged.loc[merged["on_sf"].isna()].drop(columns=["on_sf"])
        diag["removed_as_on_sf"] = int(removed)

    diag["move_up"] = int(len(move_up_df))
    return move_up_df, diag