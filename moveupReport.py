# Konrad Kubica github test upload 9.2.2025 — refreshed & optimized

import pandas as pd
import os
from tkinter import Tk, filedialog, messagebox
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# --- CONSTANTS ---
COLUMNS_TO_USE = ["Type", "Brand", "Product Name", "Package Barcode", "Room", "Qty On Hand"]
# PDF drops Brand (kept in Excel), but widths are kept for consistent layout
COLUMN_WIDTHS = [45, 370, 50, 45, 25, 30]
BASE_PDF_FILENAME = "Filtered_Move_Up"
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
def filter_inventory(original_file):
    try:
        # Read only columns we need; keep barcodes as strings reliably
        df = pd.read_excel(
            original_file,
            sheet_name="Inventory Adjustments",
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

    # Optional: categories for faster sorts on big data
    for col in ("Type", "Room"):
        if col in df.columns:
            df[col] = df[col].astype("category")

    # Sales Floor (Brand + Product Name)
    sales_floor = df.loc[df["Room"].eq("Sales Floor"), ["Brand", "Product Name"]].drop_duplicates()

    # Candidate pool: Incoming Deliveries, Vault, Overstock
    room_mask = df["Room"].isin(["Incoming Deliveries", "Vault", "Overstock"])
    candidates = df.loc[room_mask, COLUMNS_TO_USE]

    # Vectorized anti-join: keep candidates NOT on Sales Floor
    merged = candidates.merge(
        sales_floor.assign(on_sf=1),
        on=["Brand", "Product Name"],
        how="left",
        indicator=False
    )
    filtered = merged.loc[merged["on_sf"].isna()].drop(columns=["on_sf"])

    # Low stock only for Vault (<5) — stays for Excel only
    low_stock_df = filtered.loc[
        filtered["Room"].eq("Vault") & (filtered["Qty On Hand"] < 5)
    ].copy()

    # Move-up = filtered minus Vault low-stock
    move_up_df = filtered.drop(index=low_stock_df.index, errors="ignore").copy()

    # Sort once for Excel readability
    move_up_df.sort_values(by=["Type", "Brand", "Product Name"], inplace=True, kind="stable")
    low_stock_df.sort_values(by=["Type", "Brand", "Product Name"], inplace=True, kind="stable")

    return move_up_df, low_stock_df, df


# --- EXCEL OUTPUT (timestamped) ---
def save_filtered_excel(move_up_df, low_stock_df, original_path):
    base, _ = os.path.splitext(original_path)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    output_path = f"{base}_Filtered_Move_Up_{timestamp}.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        move_up_df.to_excel(writer, sheet_name="Move_Up_Items", index=False)
        low_stock_df.to_excel(writer, sheet_name="Vault_Low_Stock", index=False)
    return output_path


# --- PDF HELPERS (non-destructive, fast) ---
def truncate_string(value, max_len):
    if isinstance(value, str) and len(value) > max_len:
        return value[:max_len - 3] + "..."
    return value


def build_pdf_section(df, title):
    styles = getSampleStyleSheet()
    elements = [Paragraph(f"<b>{title}</b>", styles["Heading2"])]
    elements.append(Paragraph(datetime.now().strftime(DATE_FORMAT), styles["Normal"]))
    elements.append(Paragraph(" ", styles["Normal"]))  # Spacer

    # Work on a projection for PDF; keep Excel data pristine
    pdf_df = df.loc[:, COLUMNS_TO_USE].copy()
    pdf_df.sort_values(by=["Type", "Brand", "Product Name"], inplace=True, kind="stable")

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


# --- PDF OUTPUT (timestamped; Move-Up only) ---
def generate_pdf(move_up_df, low_stock_df, source_path):
    base_path = os.path.dirname(source_path)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    pdf_filename = f"{BASE_PDF_FILENAME}_{timestamp}.pdf"
    output_path = os.path.join(base_path, pdf_filename)

    doc = SimpleDocTemplate(output_path, pagesize=letter)

    elements = []
    # Only include Move-Up section in the PDF
    elements += build_pdf_section(move_up_df, "Move-Up Inventory List")

    doc.build(elements)

    # Open the PDF automatically (Windows only)
    try:
        os.startfile(output_path)
    except Exception as e:
        print(f"Could not open PDF automatically: {e}")

    return output_path


# --- MAIN PROGRAM ---
if __name__ == "__main__":
    file_path = pick_excel_file()
    if not file_path:
        exit()

    move_up_df, low_stock_df, _ = filter_inventory(file_path)
    if move_up_df is None:
        exit()

    excel_output = save_filtered_excel(move_up_df, low_stock_df, file_path)
    pdf_output = generate_pdf(move_up_df, low_stock_df, file_path)

    messagebox.showinfo(
        "Done",
        f"Files created successfully:\n\n"
        f"{os.path.basename(excel_output)}\n"
        f"{os.path.basename(pdf_output)}"
    )
