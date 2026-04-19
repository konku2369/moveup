"""
Core data logic for Bisa Inventory Utility.

Pure functions with no Tk dependency: column mapping, room normalization,
move-up computation, split-package aggregation, and velocity scoring.
All constants (COLUMNS_TO_USE, APP_VERSION, etc.) live here as the
single source of truth.

HOW THE DATA PIPELINE WORKS (read this first if you're new):
=============================================================

1. USER IMPORTS A FILE
   The user opens a METRC inventory export (.xlsx or .csv).
   load_raw_df() reads it into a pandas DataFrame.

2. COLUMN MAPPING
   METRC exports have messy column names. automap_columns() detects the
   METRC barcode column (strict regex) and maps all other columns to our
   6 standard names: Type, Brand, Product Name, Package Barcode, Room,
   Qty On Hand. If optional columns exist (Distributor, Size, etc.),
   those are mapped too.

3. ROOM NORMALIZATION
   normalize_rooms() applies user-defined aliases so "Vault 1" and
   "vault" both become "Vault" (or whatever the user configured).

4. MOVE-UP COMPUTATION
   compute_moveup_from_df() is the heart of the app. It answers:
   "Which backstock items have NO matching product on the Sales Floor?"

   Logic:
   a) Filter by brand/type selections
   b) Build a set of (Brand, Product Name) combos on the Sales Floor
   c) Find items in candidate rooms (Backstock, Vault, etc.)
   d) Anti-join: remove candidates whose product IS on the floor
   e) What's left = move-up candidates (stuff to bring out front)

5. SPLIT PACKAGE AGGREGATION
   aggregate_split_packages_by_room() handles when the same METRC
   barcode appears multiple times. It groups by (Barcode, Room, Type,
   Brand, Product Name) and sums Qty. Same SKU in different rooms =
   separate rows (intentional — they need separate move-up stickers).

6. VELOCITY TRACKING (optional)
   compute_velocity_metrics() compares inventory across successive
   imports to detect movement patterns. Items that never change qty
   are flagged Slow/Stale; items losing qty are scored as Fast.
"""

import os
import re
import csv
from io import TextIOWrapper
from typing import Dict, List, Optional, Tuple

import pandas as pd


# ------------------------------
# Single source constants
# ------------------------------
APP_VERSION = "4.1"
APP_NAME = "Konrad's Bisa Inventory Utility"

# The 6 required columns for move-up logic. Every imported file must map to these.
# Think of these as the "schema" — no matter what the METRC export calls them,
# automap_columns() will rename them to these 6 names so the rest of the app
# can always do df["Product Name"] without worrying about the source format.
COLUMNS_TO_USE = ["Type", "Brand", "Product Name", "Package Barcode", "Room", "Qty On Hand"]

# Optional columns that are included when present in the data (for audit/export features).
# These are NOT required — if the METRC export doesn't have them, the app still works.
# They just enable extra features like distributor-grouped audit PDFs or cost tracking.
AUDIT_OPTIONAL_FIELDS = ["Distributor", "Store", "Size", "Received Date", "Wholesale Cost", "Unit Price"]

# Max display length for Type column in both GUI treeviews and PDF exports.
# Cannabis product types like "Infused (edible)" get truncated to 7 chars → "Infused"
# to keep the UI columns compact.
TYPE_TRUNC_LEN = 7

# Fuzzy name candidates for auto-mapping imported columns to internal names.
# Different POS systems (Sweed, Dutchie, etc.) use different column names for
# the same data. This dict maps our internal name → list of names we've seen
# in the wild. automap_columns() walks the imported columns and matches them
# against these candidates.
#
# Example: A Sweed export might have "Available Qty" instead of "Qty On Hand".
#          automap_columns() finds "available qty" in the candidates for
#          "Qty On Hand" and renames the column automatically.
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
    "Wholesale Cost": [
        "wholesale cost", "wholesale price", "wholesale", "cost",
        "unit cost", "vendor cost", "supplier cost", "buy price",
    ],
    "Unit Price": [
        "unit price", "retail price", "price", "sell price",
        "selling price", "retail", "msrp", "srp",
    ],
}

# Room names that count as "Sales Floor" when checking if a product is already
# on the floor. All lowercase. If a room matches any of these, items from that
# room are excluded from move-up candidates (they're already out front).
SALES_FLOOR_ALIASES = {
    "sales floor", "floor", "salesfloor", "front of house",
    "foh", "front", "front of shop", "retail"
}


# ------------------------------
# Small utilities
# ------------------------------
def sanitize_prefix(pfx: str) -> str:
    """Strip and replace characters that are invalid in Windows/Unix filenames.

    Used when the user sets a filename prefix for PDF exports (e.g., "Bisa Lina").
    Removes characters like : * ? " < > | that would break Windows file paths.
    """
    if not pfx:
        return pfx
    pfx = pfx.strip()
    pfx = re.sub(r'[\\/:*?"<>|]+', "_", pfx)  # Replace Windows-invalid chars
    pfx = re.sub(r"\s+", "_", pfx)              # Replace whitespace with underscores
    return pfx


def _lower_strip_cols(columns) -> List[str]:
    """Return a lowercase, whitespace-stripped version of every column name.

    Used by automap_columns() to build a normalised index list before fuzzy
    matching against ALT_NAME_CANDIDATES. Operating on a pre-normalised list
    avoids repeating .strip().lower() on every column during the matching loop.
    """
    return [str(c).strip().lower() for c in columns]


_TOKEN_SETS: Dict[str, List[set]] = {
    "Type":             [{"type"}, {"category"}, {"product", "type"}, {"item", "type"}, {"class"}],
    "Brand":            [{"brand"}, {"manufacturer"}, {"mfr"}],
    "Product Name":     [{"product", "name"}, {"product"}, {"item", "name"}, {"item"}, {"title"}],
    "Room":             [{"room"}, {"location"}, {"bin"}, {"area"}],
    "Qty On Hand":      [{"qty"}, {"quantity"}, {"qoh"}, {"stock", "qty"}, {"on", "hand"}],
    "Distributor":      [{"distributor"}, {"vendor"}, {"supplier"}],
    "Store":            [{"store"}],
    "Size":             [{"size"}, {"weight"}, {"volume"}, {"net", "weight"}],
    "Received Date":    [{"received", "date"}, {"receipt", "date"}, {"received"}, {"arrival", "date"},
                         {"intake", "date"}, {"packaged", "date"}, {"harvest", "date"}],
    "Wholesale Cost":   [{"wholesale"}, {"cost"}, {"buy", "price"}],
    "Unit Price":       [{"unit", "price"}, {"retail", "price"}, {"sell", "price"}, {"msrp"}, {"srp"}],
}


def _find_source_for(target_key: str, lower_cols: List[str],
                     mapping=ALT_NAME_CANDIDATES,
                     used_indices: Optional[set] = None) -> Optional[int]:
    """Find the index of the first column that matches one of the candidate names for *target_key*.

    Uses two strategies:
      1. Exact match: column name equals a candidate string (e.g. "qty on hand")
      2. Token match: column name contains all required keyword tokens
         (e.g. "Product_Description" contains token "product" for Product Name)

    Exact match is tried first to avoid false positives.
    Columns in *used_indices* are skipped to prevent double-claiming.
    """
    skip = used_indices or set()
    wanted = [w.strip().lower() for w in mapping.get(target_key, [])]

    # Strategy 1: exact match (original behavior)
    for idx, lc in enumerate(lower_cols):
        if idx not in skip and lc in wanted:
            return idx

    # Strategy 2: token-based fuzzy match
    token_groups = _TOKEN_SETS.get(target_key)
    if not token_groups:
        return None

    for idx, lc in enumerate(lower_cols):
        if idx in skip:
            continue
        # Split on any non-alpha character to extract word tokens
        col_tokens = set(re.split(r"[^a-z]+", lc))
        for required_tokens in token_groups:
            if required_tokens.issubset(col_tokens):
                return idx

    return None


def _build_room_map(user_aliases: dict) -> Dict[str, str]:
    """Build a case-folded {lowercase_from: canonical_to} lookup from the user's aliases.

    The user configures aliases in the Filters dialog to normalise messy room
    names coming from METRC exports. For example:
        {"Vault 1": "Vault", "vault": "Vault", "SF": "Sales Floor"}

    Case-folding (str.casefold) is more aggressive than .lower() and handles
    Unicode edge cases (e.g. German ß → ss). This means alias lookups work
    regardless of how the user capitalised the "From" side.

    Returns an empty dict if user_aliases is None or empty, which makes
    normalize_rooms() a safe no-op when no aliases are configured.
    """
    if not user_aliases:
        return {}
    return {(k or "").casefold(): v for k, v in user_aliases.items()}


def normalize_rooms(df: pd.DataFrame, user_aliases: dict) -> pd.DataFrame:
    """Apply user-defined room name aliases to the DataFrame.

    The user can configure aliases in the Filters dialog, e.g.:
      "Vault 1" → "Vault"
      "vault"   → "Vault"
      "SF"      → "Sales Floor"

    This ensures that rooms with slightly different names in the METRC export
    are treated as the same room for move-up logic and filtering.
    The alias matching is case-insensitive.
    """
    if df is None or df.empty or "Room" not in df.columns:
        return df
    out = df.copy()
    out["Room"] = out["Room"].astype(str).str.strip()
    norm_map = _build_room_map(user_aliases)
    if norm_map:
        out["Room"] = out["Room"].map(lambda v: norm_map.get(str(v).casefold(), v))
    return out


def windows_unblock_file(path: str):
    """Remove the Windows "downloaded from internet" security flag from a file.

    When you download a file in Windows, it gets a hidden "Zone.Identifier"
    flag that can cause pandas/openpyxl to fail or show security warnings.
    This removes that flag so the file can be opened normally.
    Only runs on Windows — no-op on Mac/Linux.
    """
    if os.name != "nt":
        return
    try:
        ads_path = path + ":Zone.Identifier"
        if os.path.exists(ads_path):
            os.remove(ads_path)
    except Exception as e:
        print(f"[moveup] windows_unblock_file failed: {e}")


def _read_csv_smart(path: str, skiprows: int) -> pd.DataFrame:
    """Read a CSV file with automatic delimiter and encoding detection.

    Some METRC exports are comma-separated, some are tab-separated, and some
    use semicolons. This function sniffs the delimiter from the first 4KB of
    the file. It also tries UTF-8 first, then falls back to Latin-1 for
    files with special characters (common in product names like "açaí").
    """
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
    """Check if a file is a Sweed POS export (needs 3 header rows skipped).

    Sweed exports have "Export Date" in cell A1, then 2 more metadata rows
    before the actual column headers. Returns True if we need skiprows=3.
    """
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

    if head.empty or head.shape[1] == 0:
        return False
    first_cell = str(head.iloc[0, 0]).strip().lower()
    return first_cell.startswith("export date")


def sort_with_backstock_priority(df: pd.DataFrame) -> pd.DataFrame:
    """Sort items with Backstock room first, then alphabetically by Type/Brand/Product.

    This is the default sort order for move-up lists and PDF exports.
    Backstock items appear at the top because they're the most common
    move-up source and staff usually start there.
    """
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
    """Truncate *s* to *n* chars, adding '...' if truncated. Simple ASCII version."""
    s = str(s)
    if len(s) <= n:
        return s
    if n < 4:
        return s[:n]
    return s[: n - 3] + "..."


def truncate_text(val, max_len: int) -> str:
    """Truncate a value to *max_len* characters, appending ellipsis if needed.

    Handles None, NaN, and non-string values gracefully. This is the single
    canonical truncation function — satellite windows should import this
    instead of defining their own.
    """
    if val is None:
        return ""
    try:
        import math
        if isinstance(val, float) and math.isnan(val):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(val)
    if len(s) <= max_len:
        return s
    return s[: max(0, max_len - 1)] + "\u2026"

def aggregate_split_packages_by_room(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate duplicate inventory rows while preserving room-level visibility.

    WHY THIS EXISTS:
    METRC exports sometimes have the same barcode appearing multiple times.
    This can happen because:
      - A package was split across rooms (5 in Vault, 5 in Backstock)
      - Export duplication/noise from the POS system

    WHAT IT DOES:
    Groups by (Package Barcode, Room, Type, Brand, Product Name) and sums
    Qty On Hand. This means:
      - Same SKU in DIFFERENT rooms → stays as separate rows (correct!)
      - Same SKU in the SAME room with duplicate rows → merged into one row

    EXAMPLE:
    Input:
      SKU-123 | Vault     | 5
      SKU-123 | Backstock | 3
      SKU-123 | Vault     | 2  (duplicate — export noise)

    Output:
      SKU-123 | Vault     | 7  (5 + 2 merged)
      SKU-123 | Backstock | 3  (separate room, kept as-is)
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
def _looks_like_metrc_barcode(val) -> bool:
    """Check if a single value looks like a METRC tracking barcode.

    METRC barcodes are 24-char alphanumeric strings starting with '1A4'
    (or similar state prefixes like '1A3', '1A5').  We check for:
      - At least 16 characters long (some states use shorter codes)
      - Starts with '1A' followed by a digit
      - Mostly alphanumeric (allows a few dashes/spaces)
    """
    s = str(val).strip()
    if len(s) < 16:
        return False
    clean = re.sub(r"[\s\-]", "", s)
    if not clean.isalnum():
        return False
    return bool(re.match(r"1[Aa]\d", clean))


def _detect_metrc_by_content(df: pd.DataFrame, sample_size: int = 50) -> Optional[str]:
    """Fallback: detect the METRC column by scanning actual cell values.

    If the column name doesn't contain 'metrc', we sample up to *sample_size*
    non-null values from each column and check if the majority look like METRC
    barcodes.  The column with the highest hit rate wins (must be >= 50%).
    """
    if df is None or df.empty:
        return None

    best_col = None
    best_rate = 0.0

    for col in df.columns:
        vals = df[col].dropna().head(sample_size)
        if vals.empty:
            continue
        hits = sum(1 for v in vals if _looks_like_metrc_barcode(v))
        rate = hits / len(vals)
        if rate > best_rate:
            best_rate = rate
            best_col = col

    return best_col if best_rate >= 0.5 else None


def detect_metrc_source_column(df: pd.DataFrame) -> Optional[str]:
    """Detect the METRC barcode column, first by name pattern, then by content.

    Strategy 1 (strict name match):
      Requires BOTH 'metrc' in the column name AND one of: code, id, tag,
      barcode, package.  This prevents false positives like 'Metric Score'.

    Strategy 2 (content fallback):
      If no column name matches, scan actual cell values for METRC barcode
      patterns (24-char alphanumeric strings starting with '1A4...').
      The column where >= 50% of sampled values look like barcodes wins.

    Examples that match by NAME:
      'METRC Code', 'Metrc Package ID', 'METRC Barcode', 'METRC Tag'

    Examples that match by CONTENT (even if column is named 'Tracking #'):
      Values like '1A4060300007B9E000012345'
    """
    if df is None or df.empty:
        return None

    # Strategy 1: strict name matching
    required = "metrc"
    tokens = ("code", "id", "tag", "barcode", "package")

    for col in df.columns:
        name = str(col).strip().lower()
        if required in name and any(t in name for t in tokens):
            return col

    # Strategy 2: content-based fallback
    return _detect_metrc_by_content(df)


# ------------------------------
# Column mapping
# ------------------------------
def automap_columns(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Auto-detect and rename columns from a raw METRC export to our standard names.

    This is the critical step that makes the app work with any POS system.
    The function:
      1. Finds the METRC barcode column (strict regex via detect_metrc_source_column)
      2. Renames it to "Package Barcode" (our internal standard name)
      3. Maps remaining required columns (Type, Brand, etc.) using fuzzy name matching
      4. Optionally maps audit columns (Distributor, Store, Size, etc.)
      5. Normalizes data types (barcode→string, qty→int, dates→YYYY-MM-DD)

    Returns:
        (mapped_df, rename_map) — the cleaned DataFrame and a dict showing
        which source columns were renamed to which target names.

    Raises:
        ValueError: If no METRC column found, or required columns are missing
        after mapping.
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

    # Track which column indices are already claimed to prevent double-mapping
    used_indices: set = set()
    metrc_idx = list(out.columns).index(metrc_src)
    used_indices.add(metrc_idx)

    # Map remaining REQUIRED columns via candidates
    for key in COLUMNS_TO_USE:
        if key == "Package Barcode":
            continue
        if key in out.columns:
            continue
        idx = _find_source_for(key, lower_cols, used_indices=used_indices)
        if idx is not None:
            src_col = out.columns[idx]
            rename_map[src_col] = key
            used_indices.add(idx)

    # Map OPTIONAL audit columns if present
    for opt in AUDIT_OPTIONAL_FIELDS:
        if opt in out.columns:
            continue
        idx = _find_source_for(opt, lower_cols, used_indices=used_indices)
        if idx is not None:
            src_col = out.columns[idx]
            rename_map[src_col] = opt
            used_indices.add(idx)

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

    # --- Drop trailing junk / summary rows ---
    # Some POS exports append summary rows at the bottom (e.g. "Total: 1,250",
    # blank rows, disclaimers).  We detect junk rows as:
    #   - Barcode is empty/nan/none AND product name is empty/nan/none
    #   - OR barcode doesn't look like a real barcode (too short / no digits)
    #     AND product name is also empty/nan/none
    _EMPTY = {"", "nan", "none", "nat", "null"}
    _bc = out["Package Barcode"].astype(str).str.strip().str.lower()
    _pn = out["Product Name"].astype(str).str.strip().str.lower()
    bc_empty = _bc.isin(_EMPTY)
    pn_empty = _pn.isin(_EMPTY)
    # A barcode is "not real" if it has fewer than 6 digits (real METRCs are 24+ chars)
    bc_no_digits = _bc.map(lambda v: sum(c.isdigit() for c in v) < 6)
    junk_mask = (bc_empty & pn_empty) | (bc_no_digits & pn_empty)
    if junk_mask.any():
        print(f"[moveup] Dropped {junk_mask.sum()} junk/summary row(s) from import")
        out = out.loc[~junk_mask].reset_index(drop=True)

    if "Distributor" in out.columns:
        out["Distributor"] = out["Distributor"].astype(str).fillna("").str.strip()
    if "Store" in out.columns:
        out["Store"] = out["Store"].astype(str).fillna("").str.strip()
    if "Size" in out.columns:
        out["Size"] = out["Size"].astype(str).fillna("").str.strip()
    if "Wholesale Cost" in out.columns:
        out["Wholesale Cost"] = pd.to_numeric(out["Wholesale Cost"], errors="coerce").fillna(0.0)
    if "Unit Price" in out.columns:
        out["Unit Price"] = pd.to_numeric(out["Unit Price"], errors="coerce").fillna(0.0)

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
                # Sweed export format: MM/DD/YYYY HH:MM:SS AM/PM (12h hour with AM/PM marker)
                return pd.to_datetime(s, format="%m/%d/%Y %I:%M:%S %p").strftime("%Y-%m-%d")
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
def _detect_header_row(original_file: str, ext: str, sheet_name: str,
                       max_scan: int = 20) -> int:
    """Auto-detect which row contains column headers by scoring against known names.

    Reads the first *max_scan* rows as raw data (no header), then checks each
    row for how many cells match known column name candidates.  The row with
    the highest score wins.  Falls back to 0 if nothing looks like a header.

    This makes the app resilient to structural changes in METRC/POS exports --
    title rows, metadata rows, or extra blank rows at the top are skipped
    automatically.
    """
    # Build a flat set of all known column name keywords (lowercase)
    known: set = set()
    for candidates in ALT_NAME_CANDIDATES.values():
        known.update(candidates)
    # Also match the internal names themselves
    for name in list(COLUMNS_TO_USE) + list(AUDIT_OPTIONAL_FIELDS):
        known.add(name.strip().lower())
    # METRC-specific tokens
    known.update(("metrc code", "metrc barcode", "metrc package id", "metrc tag"))

    is_text = ext in {".csv", ".tsv", ".txt", ".tab"}
    engine = None
    if ext == ".xlsb":
        engine = "pyxlsb"
    elif ext == ".ods":
        engine = "odf"

    try:
        if is_text:
            with open(original_file, "rb") as raw:
                sample = raw.read(4096)
                try:
                    dialect = csv.Sniffer().sniff(sample.decode("utf-8", errors="ignore"))
                    delim = dialect.delimiter
                except Exception:
                    delim = "\t" if ext in (".tsv", ".tab") else ","
            try:
                preview = pd.read_csv(
                    original_file, header=None, nrows=max_scan,
                    sep=delim, encoding="utf-8", engine="python",
                )
            except Exception:
                preview = pd.read_csv(
                    original_file, header=None, nrows=max_scan,
                    sep=delim, encoding="latin-1", engine="python",
                )
        else:
            try:
                preview = pd.read_excel(
                    original_file, sheet_name=sheet_name,
                    header=None, nrows=max_scan, engine=engine,
                )
            except Exception:
                preview = pd.read_excel(
                    original_file, sheet_name=0,
                    header=None, nrows=max_scan, engine=engine,
                )
    except Exception:
        return 0

    best_row = 0
    best_score = 0
    for i in range(len(preview)):
        row_vals = [str(v).strip().lower() for v in preview.iloc[i] if pd.notna(v)]
        score = sum(1 for v in row_vals if v in known)
        # Also check for METRC pattern (column name containing "metrc" + a token)
        for v in row_vals:
            if "metrc" in v and any(t in v for t in ("code", "id", "tag", "barcode", "package")):
                score += 2  # Extra weight for METRC column
        if score > best_score:
            best_score = score
            best_row = i

    # Require at least 3 known column names to be confident
    return best_row if best_score >= 3 else 0


_TEXT_EXTS = {".csv", ".tsv", ".txt", ".tab"}
_EXCEL_EXTS = {".xlsx", ".xls", ".xlsm", ".xlsb", ".ods"}


def load_raw_df(original_file: str, sheet_name: str = "Inventory Adjustments") -> pd.DataFrame:
    """Load a raw inventory file (Excel or CSV) into a pandas DataFrame.

    Supported formats:
      - Excel: .xlsx, .xls, .xlsm, .xlsb, .ods
      - Text:  .csv, .tsv, .txt, .tab

    Handles:
      - Windows "downloaded from internet" security flags (auto-removed)
      - Auto-detects which row contains column headers (handles metadata rows,
        title rows, Sweed exports, or any structural format change)
      - Excel files with sheet name fallback
      - Text files with auto-detected delimiter and encoding
    """
    windows_unblock_file(original_file)
    ext = os.path.splitext(original_file)[1].lower()

    if ext not in _TEXT_EXTS and ext not in _EXCEL_EXTS:
        raise ValueError(
            f"Unsupported file type: '{ext}'\n\n"
            f"Supported: {', '.join(sorted(_TEXT_EXTS | _EXCEL_EXTS))}"
        )

    skiprows = _detect_header_row(original_file, ext, sheet_name)

    if ext in _TEXT_EXTS:
        return _read_csv_smart(original_file, skiprows=skiprows)

    # Pick the right engine for the Excel variant
    engine = None
    if ext == ".xlsb":
        engine = "pyxlsb"
    elif ext == ".ods":
        engine = "odf"

    barcode_dtypes = {"Barcode": "string", "Package Barcode": "string", "METRC Barcode": "string"}

    try:
        return pd.read_excel(
            original_file, sheet_name=sheet_name,
            skiprows=skiprows, dtype=barcode_dtypes,
            engine=engine,
        )
    except Exception:
        return pd.read_excel(
            original_file, sheet_name=0,
            skiprows=skiprows, dtype=barcode_dtypes,
            engine=engine,
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
    """THE CORE ALGORITHM — computes which items need to be moved to the sales floor.

    This is the main business logic of the entire app. Here's the step-by-step:

    1. Drop rows with missing critical fields (no barcode = can't track it)
    2. Apply brand filter (user picks which brands to include)
    3. Apply type filter (user picks which product types to include)
    4. Normalize room names using aliases
    5. Exclude accessories (rolling papers, lighters — they don't need move-up)
    6. Build a set of (Brand, Product Name) combos that are ON the Sales Floor
    7. Filter to items in candidate rooms only (Backstock, Vault, etc.)
    8. ANTI-JOIN: remove items whose product is already on the floor
    9. What's left = items that need to be brought out to the sales floor!

    Parameters:
        df: Mapped inventory DataFrame (output of automap_columns)
        candidate_rooms: Rooms to pull move-up candidates from (e.g., ["Backstock", "Vault"])
        room_alias_overrides: Room name aliases (e.g., {"vault 1": "Vault"})
        brand_filter: Only include these brands (None or ["ALL"] = include all)
        type_filter: Only include these types (None or ["ALL"] = include all)
        skip_sales_floor: If True, don't check Sales Floor — show ALL candidates

    Returns:
        (move_up_df, diagnostics_dict) — the filtered results and a dict with
        counts at each step (useful for debugging "why is my item missing?")
    """
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

    # Drop rows with missing critical fields BEFORE str conversion (astype(str) turns NaN → "nan")
    work = work.dropna(subset=["Product Name", "Brand", "Package Barcode", "Room"]).copy()

    for c in ["Product Name", "Brand", "Package Barcode", "Room", "Type"]:
        if c in work.columns:
            work[c] = work[c].astype(str)
    diag["after_dropna"] = int(len(work))

    # --- Brand filter ---
    # Special case: if "ALL" (case-insensitive) appears, skip filtering entirely.
    if brand_filter:
        bf = [str(b).strip() for b in brand_filter if str(b).strip()]
        is_all = any(b.upper() == "ALL" for b in bf)
        if not is_all:
            work = work[work["Brand"].astype(str).isin(bf)]
    diag["after_brand"] = int(len(work))

    # --- Type filter (same "ALL" convention as brand) ---
    if type_filter and "Type" in work.columns:
        tf = [str(t).strip() for t in type_filter if str(t).strip()]
        is_all_type = any(t.upper() == "ALL" for t in tf)
        if not is_all_type:
            work = work[work["Type"].astype(str).isin(tf)]
    diag["after_type_filter"] = int(len(work))

    work = normalize_rooms(work, room_alias_overrides or {})

    # Exclude accessories (e.g. rolling papers, lighters) — they don't need move-up.
    # Uses substring match: "accessor" catches "Accessories", "Accessory", etc.
    if "Type" in work.columns:
        mask_accessory = work["Type"].astype(str).str.contains(r"accessor", case=False, na=False)
        work = work.loc[~mask_accessory].copy()
    diag["after_type"] = int(len(work))

    # --- Build the Sales Floor set ---
    # Collect unique (Brand, Product Name) combos that exist on any sales-floor room.
    # These will be used to EXCLUDE candidates that don't need restocking.
    room_lower = work["Room"].astype(str).str.strip().str.lower()
    if not skip_sales_floor:
        sf_mask = room_lower.eq("sales floor") | room_lower.isin(SALES_FLOOR_ALIASES)
        sales_floor = work.loc[sf_mask, ["Brand", "Product Name"]].drop_duplicates()
    else:
        sales_floor = pd.DataFrame(columns=["Brand", "Product Name"])

    # --- Filter to candidate rooms only (e.g., Backstock, Vault) ---
    candidate_set = {str(r).strip() for r in (candidate_rooms or [])}
    candidates = work.loc[work["Room"].astype(str).str.strip().isin(candidate_set)].copy()
    diag["candidate_pool"] = int(len(candidates))

    # --- Anti-join: remove candidates whose product is already on the Sales Floor ---
    # We merge candidates with sales_floor, adding a marker column "on_sf".
    # Rows where on_sf is NOT null = product is on the floor → remove them.
    # Rows where on_sf IS null = product needs move-up → keep them.
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


# ------------------------------
# Velocity tracking
# ------------------------------
# How many consecutive imports an item must be unchanged before it's flagged.
# If a product's quantity hasn't changed for 3 consecutive imports → "Slow".
# If unchanged for 6+ imports (threshold × 2) → "Stale".
# Users can adjust this in the Velocity Tracker window's threshold spinner.
# Higher threshold = more lenient (more imports before flagging as slow).
VELOCITY_SLOW_THRESHOLD_DEFAULT = 3

VELOCITY_LABELS = {
    "new": "New",
    "fast": "Fast",
    "moderate": "Moderate",
    "slow": "Slow",
    "stale": "Stale",
    "sold_out": "Sold Out",
}


def _safe_int(val, default: int = 0) -> int:
    """Convert *val* to int, returning *default* on any conversion error.

    Used when reading qty/count fields from snapshot dicts where the value
    could be a str, float, None, or missing key. Avoids scattered try/except
    throughout the velocity computation loop.
    """
    try:
        return int(val)
    except (TypeError, ValueError):
        return default


def build_velocity_snapshot_entries(df: pd.DataFrame) -> List[Dict[str, str]]:
    """Extract a snapshot of the current inventory for velocity tracking.

    Each time the user imports a METRC file, we save a "snapshot" — a list of
    every item's barcode, room, qty, and received date at that moment in time.
    Later, compute_velocity_metrics() compares snapshots across imports to
    detect movement.

    Returns list of dicts: [{barcode, room, qty, received_date}, ...]
    Note: Does NOT store product_name/brand/type (those are looked up from
    current inventory when needed, saving ~60% storage space).
    """
    entries: List[Dict[str, str]] = []
    if df is None or df.empty:
        return entries
    if "Package Barcode" not in df.columns or "Room" not in df.columns:
        return entries

    has_qty = "Qty On Hand" in df.columns
    has_date = "Received Date" in df.columns

    for _, row in df.iterrows():
        bc = str(row["Package Barcode"]).strip()
        if not bc or bc.lower() == "nan":
            continue
        entries.append({
            "barcode": bc,
            "room": str(row["Room"]).strip().lower(),
            "qty": _safe_int(row["Qty On Hand"]) if has_qty else 0,
            "received_date": str(row.get("Received Date", "")).strip() if has_date else "",
        })
    return entries


def compute_velocity_metrics(
    current_df: pd.DataFrame,
    snapshots: List[Dict],
    slow_threshold: int = VELOCITY_SLOW_THRESHOLD_DEFAULT,
) -> pd.DataFrame:
    """Compute how fast each inventory item is moving (selling/being restocked).

    This is the analytics engine behind the Velocity Tracker window.
    For each barcode, it calculates:

    - room_changes: How many times the item moved between rooms across imports
    - qty_delta: Current qty minus first-seen qty (negative = sold units)
    - sell_rate: Units lost per day (positive = selling, negative = restocked)
    - stock_age_days: Days since Received Date (or first snapshot appearance)
    - qty_unchanged_streak: How many consecutive recent imports had same qty
    - velocity_score: 0-1 composite (0.7 × qty movement + 0.3 × room movement)
    - velocity_label: Human-readable status:
        "New"      — only seen in 1 import (not enough data)
        "Fast"     — score ≥ 0.25 and not stagnant
        "Moderate" — score < 0.25 but still changing
        "Slow"     — unchanged for ≥ slow_threshold consecutive imports
        "Stale"    — unchanged for ≥ 2× slow_threshold imports
        "Sold Out" — was in history but disappeared from current inventory

    The 70/30 weighting was chosen because in real cannabis retail data, room
    changes are rare (items usually stay in one room until sold). Quantity
    changes are the primary signal for sell-through velocity.
    """
    from datetime import datetime

    result_cols = [
        "Package Barcode", "room_changes", "qty_delta", "sell_rate",
        "stock_age_days", "qty_unchanged_streak",
        "velocity_score", "velocity_label",
    ]

    if current_df is None or current_df.empty or "Package Barcode" not in current_df.columns:
        return pd.DataFrame(columns=result_cols)

    # Build per-barcode history from snapshots (ordered by time)
    # {barcode: [(room, qty, received_date, timestamp), ...]}
    history: Dict[str, List[Tuple]] = {}
    for snap in snapshots:
        ts = snap.get("timestamp", "")
        for entry in snap.get("entries", []):
            bc = entry.get("barcode", "")
            if bc:
                history.setdefault(bc, []).append((
                    entry.get("room", ""),
                    int(entry.get("qty", 0)),
                    entry.get("received_date", ""),
                    ts,
                ))

    # Compute time span between first and last snapshot (for sell rate)
    history_days = 1.0
    if len(snapshots) >= 2:
        try:
            first_ts = pd.to_datetime(snapshots[0].get("timestamp", ""), errors="coerce")
            last_ts = pd.to_datetime(snapshots[-1].get("timestamp", ""), errors="coerce")
            if pd.notna(first_ts) and pd.notna(last_ts):
                history_days = max((last_ts - first_ts).days, 1)
        except Exception:
            pass

    today = datetime.now().date()
    rows = []
    barcodes_seen = set()

    for _, row in current_df.iterrows():
        bc = str(row["Package Barcode"]).strip()
        if not bc or bc.lower() == "nan" or bc in barcodes_seen:
            continue
        barcodes_seen.add(bc)

        hist = history.get(bc, [])
        n_snapshots = len(hist)

        # Room changes: count transitions
        room_changes = 0
        if n_snapshots >= 2:
            for i in range(1, n_snapshots):
                if hist[i][0] != hist[i - 1][0]:
                    room_changes += 1

        # Qty delta: first seen qty vs current
        current_qty = int(row.get("Qty On Hand", 0)) if "Qty On Hand" in current_df.columns else 0
        first_qty = hist[0][1] if hist else current_qty
        qty_delta = current_qty - first_qty

        # Sell rate: units lost per day (positive = selling, negative = restocked)
        sell_rate = round(-qty_delta / history_days, 2) if n_snapshots >= 2 else 0.0

        # Stock age: days since Received Date or first appearance
        stock_age_days = 0
        received = str(row.get("Received Date", "")).strip() if "Received Date" in current_df.columns else ""
        if received and received.lower() != "nan":
            try:
                rd = pd.to_datetime(received, errors="coerce")
                if pd.notna(rd):
                    stock_age_days = (today - rd.date()).days
            except Exception:
                pass
        if stock_age_days == 0 and hist:
            first_ts = hist[0][3]
            if first_ts:
                try:
                    ft = pd.to_datetime(first_ts, errors="coerce")
                    if pd.notna(ft):
                        stock_age_days = (today - ft.date()).days
                except Exception:
                    pass

        # Qty-only unchanged streak: consecutive tail where qty didn't change
        # (Room changes are too rare to be useful for slow detection)
        qty_unchanged_streak = 0
        if n_snapshots >= 2:
            last_qty = hist[-1][1]
            for i in range(n_snapshots - 2, -1, -1):
                if hist[i][1] == last_qty:
                    qty_unchanged_streak += 1
                else:
                    break

        # Velocity score (0-1 composite):
        #   70% qty sell-through, 30% room movement
        #   (Real data: room changes are rare, qty is the dominant signal)
        if n_snapshots <= 1:
            velocity_score = 0.0
            velocity_label = "New"
        else:
            max_possible_changes = max(n_snapshots - 1, 1)
            room_score = min(room_changes / max_possible_changes, 1.0)
            qty_score = min(abs(qty_delta) / max(abs(first_qty), 1), 1.0)
            velocity_score = round(room_score * 0.3 + qty_score * 0.7, 3)

            if qty_unchanged_streak >= slow_threshold * 2:
                velocity_label = "Stale"
            elif qty_unchanged_streak >= slow_threshold:
                velocity_label = "Slow"
            elif velocity_score >= 0.25:
                velocity_label = "Fast"
            else:
                velocity_label = "Moderate"

        rows.append({
            "Package Barcode": bc,
            "room_changes": room_changes,
            "qty_delta": qty_delta,
            "sell_rate": sell_rate,
            "stock_age_days": stock_age_days,
            "qty_unchanged_streak": qty_unchanged_streak,
            "velocity_score": velocity_score,
            "velocity_label": velocity_label,
        })

    # --- Also track items that disappeared (in history but not current) ---
    for bc, hist in history.items():
        if bc in barcodes_seen:
            continue
        n_snapshots = len(hist)
        if n_snapshots == 0:
            continue
        first_qty = hist[0][1]
        room_changes = 0
        if n_snapshots >= 2:
            for i in range(1, n_snapshots):
                if hist[i][0] != hist[i - 1][0]:
                    room_changes += 1

        rows.append({
            "Package Barcode": bc,
            "room_changes": room_changes,
            "qty_delta": -first_qty,
            "sell_rate": round(first_qty / history_days, 2),
            "stock_age_days": 0,
            "qty_unchanged_streak": 0,
            "velocity_score": 1.0,
            "velocity_label": "Sold Out",
        })

    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=result_cols)


def compute_room_movement_history(
    snapshots: List[Dict], barcode: str
) -> List[Dict[str, str]]:
    """Reconstruct the room-to-room movement timeline for a single barcode.

    Walks through all velocity snapshots in chronological order and records
    each transition where the item's room changed between consecutive imports.
    Useful for investigating "where has this product been?" in the Velocity
    Tracker's detail view.

    Each snapshot entry is assumed to have at most one row per barcode. Only
    the first matching entry per snapshot is used (breaks inner loop after match).

    Parameters
    ----------
    snapshots : list[dict]
        Ordered list of velocity snapshots from VelocityHistoryManager.get_snapshots().
        Each dict has "timestamp" (ISO-8601 str) and "entries" (list of barcode dicts).
    barcode : str
        The Package Barcode to trace.

    Returns
    -------
    list[dict]
        Ordered list of movement events. Each event is:
        {"timestamp": str, "from_room": str, "to_room": str}
        Empty list if the barcode never changed rooms or was only seen once.
    """
    movements = []
    prev_room = None
    for snap in snapshots:
        ts = snap.get("timestamp", "")
        for entry in snap.get("entries", []):
            if entry.get("barcode") == barcode:
                room = entry.get("room", "")
                if prev_room is not None and room != prev_room:
                    movements.append({
                        "timestamp": ts,
                        "from_room": prev_room,
                        "to_room": room,
                    })
                prev_room = room
                break
    return movements


def compute_slow_movers(
    velocity_df: pd.DataFrame,
    current_df: pd.DataFrame,
) -> pd.DataFrame:
    """Return slow/stale items enriched with product details from current inventory.

    Filters velocity_df to rows with velocity_label "Slow" or "Stale", then
    left-merges with current_df on Package Barcode to attach product name,
    brand, room, and quantity. The merge uses drop_duplicates on Package Barcode
    to avoid multiplying rows when the same SKU appears in multiple rooms.

    Powers the "Slow Movers" tab in the Velocity Tracker window. Items here
    have had the same quantity for VELOCITY_SLOW_THRESHOLD_DEFAULT (or more)
    consecutive imports, suggesting they are not selling. Staff can use this
    list to decide on markdowns, front-of-shelf placement, or distributor returns.

    Parameters
    ----------
    velocity_df : pd.DataFrame
        Output of compute_velocity_metrics(). Must have "velocity_label" column.
    current_df : pd.DataFrame
        The current mapped inventory DataFrame (for product details).

    Returns
    -------
    pd.DataFrame
        Merged DataFrame of slow/stale rows. Empty if no slow items or if
        either input is None/empty.
    """
    if velocity_df is None or velocity_df.empty:
        return pd.DataFrame()
    slow = velocity_df[velocity_df["velocity_label"].isin(["Slow", "Stale"])].copy()
    if slow.empty:
        return slow
    if current_df is None or current_df.empty:
        return slow

    # Merge with current_df to get product details
    merged = slow.merge(
        current_df.drop_duplicates(subset=["Package Barcode"]),
        on="Package Barcode",
        how="left",
    )
    return merged


def compute_sold_out(
    velocity_df: pd.DataFrame,
    snapshots: List[Dict],
) -> pd.DataFrame:
    """Return items that existed in past imports but are absent from current inventory.

    These are barcodes with velocity_label "Sold Out" — they appeared in at
    least one historical snapshot but are not in the current inventory DataFrame.
    This indicates the product fully sold through, was returned to the distributor,
    or was destroyed/wasted.

    For each sold-out barcode, the function recovers last-known metadata from
    the most recent snapshot that contained it:
      - Room: the last room the item was seen in
      - last_qty: the quantity when last seen
      - last_seen: date of the snapshot (YYYY-MM-DD, first 10 chars of ISO timestamp)

    Note: product name/brand are NOT recovered here because sold-out items are
    no longer in current_df. The velocity_df only has barcode + velocity metrics.
    The Velocity Tracker window handles the display of just the barcode + last info.

    Parameters
    ----------
    velocity_df : pd.DataFrame
        Output of compute_velocity_metrics(). Must have "velocity_label" column.
    snapshots : list[dict]
        All velocity snapshots from VelocityHistoryManager. Used to recover
        last-known room/qty for each sold-out barcode.

    Returns
    -------
    pd.DataFrame
        Rows from velocity_df where velocity_label == "Sold Out", with Room,
        last_qty, and last_seen columns added. Empty if none found.
    """
    if velocity_df is None or velocity_df.empty:
        return pd.DataFrame()
    sold = velocity_df[velocity_df["velocity_label"] == "Sold Out"].copy()
    if sold.empty or not snapshots:
        return sold

    # Recover product info from the last snapshot each barcode appeared in
    last_info: Dict[str, Dict] = {}
    for snap in snapshots:
        for entry in snap.get("entries", []):
            bc = entry.get("barcode", "")
            if bc:
                last_info[bc] = {
                    "Room": entry.get("room", ""),
                    "last_qty": entry.get("qty", 0),
                    "last_seen": snap.get("timestamp", "")[:10],
                }

    for col in ["Room", "last_qty", "last_seen"]:
        sold[col] = sold["Package Barcode"].map(
            lambda bc, c=col: last_info.get(bc, {}).get(c, "")
        )

    return sold