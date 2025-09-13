"""
Move-Up Inventory Tool — v2 (with CLI flags)
Author: Konrad Kubica (+ ChatGPT)
Date: 2025-09-03

This file contains two parts:
1) The Python script (fully runnable)
2) A README section at the bottom (usage instructions & examples)

To run: python moveup.py [flags]
"""

import os
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
# PDF drops Brand (kept in Excel), but we keep a six-length list for column widths consistency
COLUMN_WIDTHS = [45, 370, 50, 45, 25, 30]
BASE_PDF_FILENAME = "Print_me_Filtered_Move_Up"
DATE_FORMAT = "%B %d, %Y"


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
            converters={"Package Barcode": lambda x: "" if pd.isna(x) else str(x)}
        )
    except ValueError:
        # Fallback if sheet name varies: read first sheet
        try:
            df = pd.read_excel(
                original_file,
                sheet_name=0,
                usecols=COLUMNS_TO_USE,
                converters={"Package Barcode": lambda x: "" if pd.isna(x) else str(x)}
            )
        except Exception as e:
            messagebox.showerror("Error", f"Could not read Excel file:\n{e}")
            return None, None, None
    except Exception as e:
        messagebox.showerror("Error", f"Could not read Excel file:\n{e}")
        return None, None, None

    # Clean essential fields
    df = df.dropna(subset=["Product Name", "Brand", "Package Barcode", "Room"])

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
    pdf_df["Package Barcode"] = pdf_df["Package Barcode"].apply(lambda x: str(x)[-6:] if pd.notna(x) and str(x) else "")

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

    # Friendly popup if running interactively and Tk is available
    try:
        message = "Files created successfully:\n\n" + "\n".join(os.path.basename(p) for p in outputs)
        messagebox.showinfo("Done", message)
    except Exception:
        # Fallback to console output if Tk not available (e.g., headless run)
        for p in outputs:
            print(p)


if __name__ == "__main__":
    main()


# ==========================
# README — Usage & Examples
# ==========================
"""
README — Move-Up Inventory Tool (v2)

Overview
--------
This tool scans your inventory export, finds items in specified back rooms (Incoming Deliveries / Vault / Overstock by default)
that are **not** on the Sales Floor (matched by Brand + Product Name), and outputs:

- Excel workbook with two tabs:
  - Move_Up_Items (primary working list)
  - Vault_Low_Stock (Vault items under the threshold; reference only)
- PDF report of the Move-Up list (optionally also includes Vault Low Stock)

Highlights
---------
- Fast vectorized filters (no Python loops)
- Timestamped filenames by default
- Command-line flags to customize behavior
- Still works with a GUI file picker if you don’t pass --input

Install Requirements
--------------------
python -m pip install pandas openpyxl reportlab

Basic Usage
-----------
1) Double-click the script (Windows) or run without flags:
   - You’ll be prompted to pick an Excel file (sheet "Inventory Adjustments" by default).
   - Outputs timestamped Excel + PDF next to the source file.

2) From terminal, with explicit input:
   python moveup.py --input "C:\\path\\to\\inventory.xlsx"

Key Flags / Toggles
-------------------
Input / Sheet
- --input, -i <path>           : Path to the Excel file. If omitted, a file picker opens.
- --sheet <name>               : Sheet name to read (default: "Inventory Adjustments").

Outputs
- --no-pdf                     : Skip generating the PDF.
- --no-excel                   : Skip generating the Excel output.
- --pdf-include-lowstock       : Include a second section in the PDF for Vault Low Stock.

Naming
- --timestamp / --no-timestamp : Add or remove timestamp in filenames (default: timestamp on).
- --prefix <text>              : Prefix filenames (e.g., "BisaLina" → BisaLina_Filtered_Move_Up_YYYY-MM-DD_HH-MM.pdf).

Filtering
- --rooms <list>               : Candidate rooms to check (default: Incoming Deliveries Vault Overstock).
- --lowstock-threshold <N>     : Vault low-stock cutoff (default: 5).

Behavior
- --open / --no-open           : Auto-open the generated PDF (default: on in Windows; off elsewhere).

Examples
--------
1) Include low-stock in PDF and add a site prefix:
   python moveup.py -i inventory.xlsx --pdf-include-lowstock --prefix EarthMed

2) Excel only (no PDF), restrict to Vault + Overstock:
   python moveup.py -i inventory.xlsx --no-pdf --rooms Vault Overstock

3) Raise the Vault low-stock threshold to 10:
   python moveup.py -i inventory.xlsx --pdf-include-lowstock --lowstock-threshold 10

4) Produce fixed filenames (no timestamp) and don’t auto-open:
   python moveup.py -i inventory.xlsx --no-timestamp --no-open

Output Files
------------
- <basename>_Filtered_Move_Up_YYYY-MM-DD_HH-MM.xlsx
- Filtered_Move_Up_YYYY-MM-DD_HH-MM.pdf (or with your --prefix)

Notes
-----
- Barcodes are treated as strings and the PDF shows the **last 6** characters.
- Brand is hidden in the PDF for compact layout but remains in Excel.
- If your source sheet name changes, pass --sheet <name>.
- If you exclude "Vault" from --rooms, the Low Stock tab/section simply ends up empty.

"""
 #command to build app  pyinstaller --onefile --noconsole --name "MoveUp-Inventory V1.1" --icon ihjicon.ico moveupReport.py