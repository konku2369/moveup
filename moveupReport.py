"""
Move-Up Inventory Tool — v2.4
Author: Konrad Kubica (+ ChatGPT)
Date: 2025-09-13

See README.md for usage, installation, and build instructions.
"""

import os
import re
import sys
import argparse
from datetime import datetime

import pandas as pd
from tkinter import Tk, filedialog, messagebox
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# --- CONSTANTS ---
COLUMNS_TO_USE = ["Type", "Brand", "Product Name", "Package Barcode", "Room", "Qty On Hand"]
# PDF renders 5 columns (Brand hidden). Widths MUST match the rendered columns.
COLUMN_WIDTHS = [50, 365, 70, 45, 35]  # Type, Product, Barcode, Loc, Qty
BASE_PDF_FILENAME = "Print_me_Filtered_Move_Up"
DATE_FORMAT = "%B %d, %Y"

# Alternate-input column name candidates (for Sweed or other exports).
# Keys are our canonical names; each list contains lowercase variants.
ALT_NAME_CANDIDATES = {
    "Type": ["type", "product type", "category", "item type", "class"],
    "Brand": ["brand", "brand name", "manufacturer", "mfr"],
    "Product Name": ["product name", "product", "item name", "name", "title", "item"],
    # Sweed uses "Barcode"
    "Package Barcode": ["barcode", "package barcode", "package id", "upc", "ean", "gtin", "Barcode","metrc code", "package upc", "package ean"],
    # Sweed uses "Location"
    "Room": ["room", "location", "stock location", "bin", "area", "warehouse location", "site location"],
    # Sweed uses "Available Qty"
    "Qty On Hand": ["available qty", "qty on hand", "quantity on hand", "on hand", "quantity", "qoh", "stock", "stock qty"],
}

# If the export marks sales floor differently, include aliases here
SALES_FLOOR_ALIASES = {"sales floor", "floor", "salesfloor", "front of house", "foh"}

# Default room aliases (left = vendor/site labels → right = canonical rooms)
# Matching is case-insensitive and trims whitespace.
DEFAULT_ROOM_ALIASES = {
    "back room": "Overstock",
    "stockroom": "Overstock",
    "stock room": "Overstock",
    "back": "Overstock",
    "over stock": "Overstock",
    "incoming": "Incoming Deliveries",
    "receiving": "Incoming Deliveries",
    "delivery": "Incoming Deliveries",
    "deliveries": "Incoming Deliveries",
    "safe": "Vault",
}

# --- HELPERS ---
def sanitize_prefix(pfx: str) -> str:
    """Make prefix safe for filenames (Windows-safe)."""
    if not pfx:
        return pfx
    pfx = pfx.strip()
    pfx = re.sub(r'[\\/:*?"<>|]+', "_", pfx)  # Replace invalid characters with underscores
    pfx = re.sub(r"\s+", "_", pfx)            # Collapse whitespace to single underscores
    return pfx


def pick_excel_file():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Inventory File",
        filetypes=[("Excel/CSV Files", "*.xlsx *.xls *.csv")]
    )
    return file_path


def _lower_strip_cols(columns):
    return [str(c).strip().lower() for c in columns]


def _find_source_for(target_key: str, lower_cols, mapping=ALT_NAME_CANDIDATES):
    """Return the source column index from lower_cols that matches our target_key candidates; else None."""
    wanted = mapping.get(target_key, [])
    for idx, lc in enumerate(lower_cols):
        if lc in wanted:
            return idx
    return None


def _auto_map_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Map arbitrary export column names (e.g., Sweed) to our canonical schema.
    Returns a DataFrame with columns exactly COLUMNS_TO_USE (extras preserved but ignored later).
    Raises ValueError if required mappings can't be found.
    """
    lower_cols = _lower_strip_cols(df.columns)

    # If already canonical, fast path
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

    # Validate presence
    missing = [c for c in COLUMNS_TO_USE if c not in out.columns]
    if missing:
        raise ValueError(
            "Missing required column(s) after auto-mapping: "
            + ", ".join(missing)
            + "\nDetected columns: "
            + ", ".join(map(str, df.columns))
            + "\nTip: If this is a Sweed export, ensure the first 3 rows are skipped "
              "(this script auto-detects and handles that). Otherwise, update ALT_NAME_CANDIDATES."
        )

    # Normalize field types
    out["Package Barcode"] = out["Package Barcode"].astype("string").fillna("")
    out["Qty On Hand"] = pd.to_numeric(out["Qty On Hand"], errors="coerce").fillna(0).astype(int)

    for col in ["Product Name", "Brand", "Type", "Room"]:
        out[col] = out[col].astype(str)

    return out


def _is_sweed_export(original_file: str, ext: str, sheet_name: str) -> bool:
    """Detect Sweed report header ('Export date:' in A1)."""
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


def load_inventory_df(original_file: str, sheet_name: str) -> pd.DataFrame:
    """
    Read Excel/CSV and auto-normalize columns so downstream logic can keep using COLUMNS_TO_USE.
    Auto-detects Sweed exports (skips first 3 rows).
    """
    ext = os.path.splitext(original_file)[1].lower()
    skiprows = 3 if _is_sweed_export(original_file, ext, sheet_name) else 0
    try:
        if ext == ".csv":
            df = pd.read_csv(original_file, skiprows=skiprows, dtype={"Barcode": "string", "Package Barcode": "string"})
        else:
            try:
                df = pd.read_excel(
                    original_file,
                    sheet_name=sheet_name,
                    skiprows=skiprows,
                    dtype={"Barcode": "string", "Package Barcode": "string"}
                )
            except Exception:
                df = pd.read_excel(
                    original_file,
                    sheet_name=0,
                    skiprows=skiprows,
                    dtype={"Barcode": "string", "Package Barcode": "string"}
                )
    except Exception as e:
        raise RuntimeError(f"Could not read file '{original_file}': {e}")

    # Auto-map to canonical schema
    df = _auto_map_columns(df)

    # Keep only the columns we use
    df = df.loc[:, [c for c in COLUMNS_TO_USE if c in df.columns]]

    # Ensure barcodes filled
    df["Package Barcode"] = df["Package Barcode"].fillna("")

    return df


def _parse_room_alias_flags(room_alias_flags):
    """
    Parse repeated --room-alias "From=To" flags into a dict {from_lower: To}.
    Later values override earlier ones and defaults.
    """
    result = {}
    if not room_alias_flags:
        return result
    for item in room_alias_flags:
        if "=" in item:
            left, right = item.split("=", 1)
            left = left.strip().casefold()
            right = right.strip()
            if left and right:
                result[left] = right
    return result


def _normalize_rooms(df: pd.DataFrame, user_aliases: dict):
    """
    Normalize df['Room'] using default + user-provided aliases.
    - Match on casefolded (case-insensitive) source values.
    - Replace with canonical 'To' value, preserving canonical casing from alias target.
    """
    if "Room" not in df.columns:
        return df
    # Build final alias map: defaults, then user overrides
    final_map = {k.casefold(): v for k, v in DEFAULT_ROOM_ALIASES.items()}
    final_map.update(user_aliases or {})

    def _map_room(val):
        s = str(val).strip()
        return final_map.get(s.casefold(), s)

    df = df.copy()
    df["Room"] = df["Room"].apply(_map_room)
    return df


# --- FILTER LOGIC (optimized) ---
def filter_inventory(original_file, sheet_name, candidate_rooms, lowstock_threshold, room_alias_overrides=None):
    try:
        df = load_inventory_df(original_file, sheet_name)
    except Exception as e:
        try:
            messagebox.showerror("Error", str(e))
        except Exception:
            print(str(e))
        return None, None, None

    # Clean essential fields (after mapping)
    df = df.dropna(subset=["Product Name", "Brand", "Package Barcode", "Room"])

    # Normalize room names using default + user-provided aliases
    df = _normalize_rooms(df, room_alias_overrides)

    # 🚫 Always remove accessories (broad match on Type; case-insensitive)
    if "Type" in df.columns:
        mask_accessory = df["Type"].astype(str).str.strip().str.contains(r"accessor", case=False, na=False)
        df = df.loc[~mask_accessory].copy()

    # Optional: categories for faster sorts on big data
    for col in ("Type", "Room"):
        if col in df.columns:
            df[col] = df[col].astype("category")

    # Sales Floor (Brand + Product Name) — allow for aliases
    room_lower = df["Room"].astype(str).str.strip().str.lower()
    sales_floor_mask = room_lower.eq("sales floor") | room_lower.isin(SALES_FLOOR_ALIASES)
    sales_floor = df.loc[sales_floor_mask, ["Brand", "Product Name"]].drop_duplicates()

    # Candidate pool: param-driven rooms (default: Incoming Deliveries, Vault, Overstock)
    room_mask_candidates = df["Room"].isin(candidate_rooms)
    candidates = df.loc[room_mask_candidates, COLUMNS_TO_USE]

    # Vectorized anti-join: keep candidates NOT on Sales Floor
    merged = candidates.merge(
        sales_floor.assign(on_sf=1),
        on=["Brand", "Product Name"],
        how="left",
        indicator=False
    )
    filtered = merged.loc[merged["on_sf"].isna()].drop(columns=["on_sf"])

    # Low stock only for Vault by default; but if user removed 'Vault' from candidate rooms,
    # the low_stock_df will simply be empty.
    low_stock_df = filtered.loc[
        filtered["Room"].eq("Vault") & (filtered["Qty On Hand"] < lowstock_threshold)
    ].copy()

    # Move-up = filtered minus Vault low-stock
    move_up_df = filtered.drop(index=low_stock_df.index, errors="ignore").copy()

    # Sort once for Excel readability
    sort_cols = [c for c in ["Type", "Brand", "Product Name"] if c in move_up_df.columns]
    if sort_cols:
        move_up_df.sort_values(by=sort_cols, inplace=True, kind="stable")
        low_stock_df.sort_values(by=sort_cols, inplace=True, kind="stable")

    return move_up_df, low_stock_df, df


# --- PDF HELPERS ---
def _draw_footer(canvas, doc):
    """Footer with date (left) and page number (right)."""
    canvas.saveState()
    w, h = letter
    y = 20  # distance from bottom
    # date (left)
    canvas.setFont("Helvetica", 8)
    canvas.drawString(40, y, datetime.now().strftime(DATE_FORMAT))
    # page number (right)
    page_text = f"Page {canvas.getPageNumber()}"
    text_width = canvas.stringWidth(page_text, "Helvetica", 8)
    canvas.drawString(w - 40 - text_width, y, page_text)
    # subtle divider line
    canvas.setLineWidth(0.25)
    canvas.setStrokeColor(colors.lightgrey)
    canvas.line(40, y + 10, w - 40, y + 10)
    canvas.restoreState()


def build_pdf_section(df, title):
    styles = getSampleStyleSheet()
    elements = [Paragraph(f"<b>{title}</b>", styles["Heading2"])]
    elements.append(Paragraph(datetime.now().strftime(DATE_FORMAT), styles["Normal"]))
    elements.append(Paragraph(" ", styles["Normal"]))  # Spacer

    # Projection for PDF; keep Excel data pristine
    pdf_df = df.loc[:, COLUMNS_TO_USE].copy()
    sort_cols = [c for c in ["Type", "Brand", "Product Name"] if c in pdf_df.columns]
    if sort_cols:
        pdf_df.sort_values(by=sort_cols, inplace=True, kind="stable")

    # Display tweaks (PDF only): drop Brand
    pdf_df["Room"] = pdf_df["Room"].astype(str).str.slice(0, 8).where(pdf_df["Room"].notna(), "")
    pdf_df["Type"] = pdf_df["Type"].astype(str).str.slice(0, 8).where(pdf_df["Type"].notna(), "")
    pdf_df["Product Name"] = pdf_df["Product Name"].astype(str).apply(lambda x: x if len(x) <= 75 else x[:72] + "...")
    # show last 6 of barcode; handle empty strings
    pdf_df["Package Barcode"] = pdf_df["Package Barcode"].apply(
        lambda x: str(x)[-6:] if pd.notna(x) and str(x) else ""
    )

    # Reorder & drop Brand (PDF only)
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

    # Zebra striping (even data rows)
    for i in range(1, len(table_data), 2):
        table.setStyle([("BACKGROUND", (0, i), (-1, i), colors.gainsboro)])

    elements.append(table)
    return elements


def generate_pdf(move_up_df, low_stock_df, source_path, include_lowstock, timestamp, prefix, auto_open):
    base_path = os.path.dirname(source_path) if source_path else os.getcwd()

    parts = [BASE_PDF_FILENAME]
    if timestamp:
        parts.append(datetime.now().strftime("%Y-%m-%d_%H-%M"))

    pdf_filename = "_".join(parts) + ".pdf"

    if prefix:
        prefix = sanitize_prefix(prefix)
        pdf_filename = f"{prefix}_{pdf_filename}"

    output_path = os.path.join(base_path, pdf_filename)

    doc = SimpleDocTemplate(output_path, pagesize=letter)
    elements = []

    # Always include Move-Up
    elements += build_pdf_section(move_up_df, "Move-Up Inventory List")

    # Optionally include Low Stock section
    if include_lowstock and not low_stock_df.empty:
        from reportlab.platypus import PageBreak
        elements.append(PageBreak())
        elements += build_pdf_section(low_stock_df, "Vault Low Stock")

    # Build with footer on every page
    doc.build(elements, onFirstPage=_draw_footer, onLaterPages=_draw_footer)

    # Auto-open on Windows if requested
    if auto_open and os.name == "nt":
        try:
            os.startfile(output_path)
        except Exception as e:
            print(f"Could not open PDF automatically: {e}")

    return output_path


# --- EXCEL OUTPUT ---
def save_filtered_excel(move_up_df, low_stock_df, original_path, timestamp, prefix):
    base_dir = os.path.dirname(original_path) if original_path else os.getcwd()
    parts = ["Sticker_Sheet_Filtered_Move_Up"]
    if timestamp:
        parts.append(datetime.now().strftime("%Y-%m-%d_%H-%M"))

    output_filename = "_".join(parts) + ".xlsx"

    if prefix:
        prefix = sanitize_prefix(prefix)
        output_filename = f"{prefix}_{output_filename}"

    output_path = os.path.join(base_dir, output_filename)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        move_up_df.to_excel(writer, sheet_name="Move_Up_Items", index=False)
        low_stock_df.to_excel(writer, sheet_name="Vault_Low_Stock", index=False)
    return output_path


# --- ARGPARSE / MAIN ---
def parse_args():
    default_auto_open = os.name == "nt"
    parser = argparse.ArgumentParser(description="Move-Up Inventory Tool (Excel/CSV → PDF/Excel exports)")

    # Inputs
    parser.add_argument("--input", "-i", help="Path to source Excel/CSV file. If omitted, a file picker will open.")
    parser.add_argument("--sheet", default="Inventory Adjustments", help="Sheet name to read for Excel files "
                        "(default: 'Inventory Adjustments'). Ignored for CSV.")

    # Output toggles
    parser.add_argument("--no-pdf", action="store_true", help="Skip generating the PDF.")
    parser.add_argument("--no-excel", action="store_true", help="Skip generating the Excel output.")
    parser.add_argument("--pdf-include-lowstock", action="store_true", help="Include 'Vault Low Stock' section in the PDF.")

    # Naming
    parser.add_argument("--timestamp", dest="timestamp", action="store_true", help="Add timestamp to filenames (default).")
    parser.add_argument("--no-timestamp", dest="timestamp", action="store_false", help="Do not add timestamp to filenames.")
    parser.set_defaults(timestamp=True)
    parser.add_argument("--prefix", default=None, help="Optional filename prefix (e.g., 'BisaLina' or 'EarthMed').")

    # Filtering
    parser.add_argument("--rooms", nargs="+", default=["Incoming Deliveries", "Vault", "Overstock"],
                        help="Candidate rooms to check (default: Incoming Deliveries Vault Overstock).")
    parser.add_argument("--lowstock-threshold", type=int, default=5, help="Vault low-stock threshold (default: 5).")
    parser.add_argument(
        "--room-alias",
        action="append",
        help='Map alternate room labels to your canonical ones; repeatable. Example: --room-alias "Back Room=Overstock"'
    )

    # Behavior
    parser.add_argument("--open", dest="auto_open", action="store_true", help="Auto-open PDF after generation (default on Windows).")
    parser.add_argument("--no-open", dest="auto_open", action="store_false", help="Do not auto-open the PDF.")
    parser.set_defaults(auto_open=default_auto_open)
    parser.add_argument("--quiet", action="store_true", help="Suppress GUI popups; print output paths to stdout only.")

    # Debug / inspection
    parser.add_argument("--list-rooms", action="store_true",
                        help="Print unique Room values before/after normalization, then exit (no files generated).")

    return parser.parse_args()


def main():
    args = parse_args()

    source_path = args.input
    if not source_path:
        source_path = pick_excel_file()
        if not source_path:
            sys.exit(0)

    # Build alias overrides from CLI flags
    room_alias_overrides = _parse_room_alias_flags(args.room_alias)

    # If listing rooms, do a quick load and show before/after, then exit
    if args.list_rooms:
        try:
            df_probe = load_inventory_df(source_path, args.sheet)
        except Exception as e:
            print(str(e))
            sys.exit(1)

        pre = sorted({str(x).strip() for x in df_probe.get("Room", pd.Series(dtype=str)).unique()})
        df_norm = _normalize_rooms(df_probe, room_alias_overrides)
        post = sorted({str(x).strip() for x in df_norm.get("Room", pd.Series(dtype=str)).unique()})

        print("Rooms (raw):")
        for r in pre:
            print(f"  - {r}")
        print("\nRooms (normalized):")
        for r in post:
            print(f"  - {r}")
        sys.exit(0)

    # Normal run
    move_up_df, low_stock_df, _ = filter_inventory(
        original_file=source_path,
        sheet_name=args.sheet,
        candidate_rooms=args.rooms,
        lowstock_threshold=args.lowstock_threshold,
        room_alias_overrides=room_alias_overrides,
    )
    if move_up_df is None:
        sys.exit(1)

    outputs = []

    if not args.no_excel:
        excel_output = save_filtered_excel(move_up_df, low_stock_df, source_path, args.timestamp, args.prefix)
        outputs.append(excel_output)

    if not args.no_pdf:
        pdf_output = generate_pdf(
            move_up_df=move_up_df,
            low_stock_df=low_stock_df,
            source_path=source_path,
            include_lowstock=args.pdf_include_lowstock,
            timestamp=args.timestamp,
            prefix=args.prefix,
            auto_open=args.auto_open,
        )
        outputs.append(pdf_output)

    # Friendly popup if running interactively and Tk is available (unless --quiet)
    if not args.quiet:
        try:
            message = "Files created successfully:\n\n" + "\n".join(os.path.basename(p) for p in outputs)
            messagebox.showinfo("Done", message)
        except Exception:
            for p in outputs:
                print(p)
    else:
        for p in outputs:
            print(p)


if __name__ == "__main__":
    main()
