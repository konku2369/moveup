"""
Move-Up Inventory Tool — v2.1 (with CLI flags)
Author: Konrad Kubica (+ ChatGPT)
Date: 2025-09-13

This file contains two parts:
1) The Python script (fully runnable)
2) A README section at the bottom (usage instructions & examples)

To run: python moveup.py [flags]
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


# --- HELPERS ---
def sanitize_prefix(pfx: str) -> str:
    """Make prefix safe for filenames (Windows-safe)."""
    if not pfx:
        return pfx
    pfx = pfx.strip()
    pfx = re.sub(r'[\\/:*?"<>|]+', "_", pfx)  # Replace invalid characters with underscores
    pfx = re.sub(r"\s+", "_", pfx)            # Collapse whitespace to single underscores
    return pfx


# --- UI HELPERS ---
def pick_excel_file():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Full Inventory Excel File",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    return file_path


# --- FILTER LOGIC (optimized) ---
def filter_inventory(original_file, sheet_name, candidate_rooms, lowstock_threshold):
    try:
        # Read only columns we need; keep barcodes as strings reliably
        df = pd.read_excel(
            original_file,
            sheet_name=sheet_name,
            usecols=COLUMNS_TO_USE,
            dtype={"Package Barcode": "string"}
        ).fillna({"Package Barcode": ""})
    except ValueError:
        # Fallback if sheet name varies: read first sheet
        try:
            df = pd.read_excel(
                original_file,
                sheet_name=0,
                usecols=COLUMNS_TO_USE,
                dtype={"Package Barcode": "string"}
            ).fillna({"Package Barcode": ""})
        except Exception as e:
            try:
                messagebox.showerror("Error", f"Could not read Excel file:\n{e}")
            except Exception:
                print(f"Could not read Excel file:\n{e}")
            return None, None, None
    except Exception as e:
        try:
            messagebox.showerror("Error", f"Could not read Excel file:\n{e}")
        except Exception:
            print(f"Could not read Excel file:\n{e}")
        return None, None, None

    # Clean essential fields
    df = df.dropna(subset=["Product Name", "Brand", "Package Barcode", "Room"])

    # Ensure Qty is numeric (coerce texty values to 0)
    if "Qty On Hand" in df.columns:
        df["Qty On Hand"] = pd.to_numeric(df["Qty On Hand"], errors="coerce").fillna(0).astype(int)

    # 🚫 Always remove accessories (broad match on Type; case-insensitive)
    # Matches: "Accessory", "Accessories", "Accessory Item", "Accessory - X", "ACCeSSorY/Parts", etc.
    if "Type" in df.columns:
        # Normalize to string, trim, then drop any row whose Type contains "accessor"
        mask_accessory = df["Type"].astype(str).str.strip().str.contains(r"accessor", case=False, na=False)
        df = df.loc[~mask_accessory].copy()

    # Optional: categories for faster sorts on big data
    for col in ("Type", "Room"):
        if col in df.columns:
            df[col] = df[col].astype("category")

    # Sales Floor (Brand + Product Name)
    sales_floor = df.loc[df["Room"].eq("Sales Floor"), ["Brand", "Product Name"]].drop_duplicates()

    # Candidate pool: param-driven rooms (default: Incoming Deliveries, Vault, Overstock)
    room_mask = df["Room"].isin(candidate_rooms)
    candidates = df.loc[room_mask, COLUMNS_TO_USE]

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


# --- PDF HELPERS (non-destructive, fast) ---
def build_pdf_section(df, title):
    styles = getSampleStyleSheet()
    elements = [Paragraph(f"<b>{title}</b>", styles["Heading2"])]
    elements.append(Paragraph(datetime.now().strftime(DATE_FORMAT), styles["Normal"]))
    elements.append(Paragraph(" ", styles["Normal"]))  # Spacer

    # Work on a projection for PDF; keep Excel data pristine
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

    doc.build(elements)

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
    parser = argparse.ArgumentParser(description="Move-Up Inventory Tool (Excel → PDF/Excel exports)")

    # Inputs
    parser.add_argument("--input", "-i", help="Path to source Excel file. If omitted, a file picker will open.")
    parser.add_argument("--sheet", default="Inventory Adjustments", help="Sheet name to read (default: 'Inventory Adjustments').")

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

    # Behavior
    parser.add_argument("--open", dest="auto_open", action="store_true", help="Auto-open PDF after generation (default on Windows).")
    parser.add_argument("--no-open", dest="auto_open", action="store_false", help="Do not auto-open the PDF.")
    parser.set_defaults(auto_open=default_auto_open)
    parser.add_argument("--quiet", action="store_true", help="Suppress GUI popups; print output paths to stdout only.")

    return parser.parse_args()


def main():
    args = parse_args()

    source_path = args.input
    if not source_path:
        source_path = pick_excel_file()
        if not source_path:
            sys.exit(0)

    move_up_df, low_stock_df, _ = filter_inventory(
        original_file=source_path,
        sheet_name=args.sheet,
        candidate_rooms=args.rooms,
        lowstock_threshold=args.lowstock_threshold,
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
