##Konrad Kubica github test upload 9.2.2025

import pandas as pd
import os
from tkinter import Tk, filedialog, messagebox
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# --- CONSTANTS ---
COLUMNS_TO_USE = ["Type", "Brand", "Product Name", "Package Barcode", "Room", "Qty On Hand"]
COLUMN_WIDTHS = [45, 370, 50, 45, 25, 30]
PDF_FILENAME = "Filtered_Move_Up.pdf"
EXCEL_SUFFIX = "_Filtered_Move_Up.xlsx"
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


# --- FILTER LOGIC ---
def truncate_string(value, max_len):
    if isinstance(value, str) and len(value) > max_len:
        return value[:max_len - 3] + "..."
    return value


def filter_inventory(original_file):
    try:
        df = pd.read_excel(original_file, sheet_name="Inventory Adjustments")
    except Exception as e:
        messagebox.showerror("Error", f"Could not read Excel file:\n{e}")
        return None, None, None

    df = df.dropna(subset=["Product Name", "Brand", "Package Barcode", "Room"])
    df = df.sort_values(by=["Type", "Brand", "Product Name"])

    sales_floor_keys = set(zip(
        df[df["Room"] == "Sales Floor"]["Brand"],
        df[df["Room"] == "Sales Floor"]["Product Name"]
    ))

    candidates = df[df["Room"].isin(["Incoming Deliveries", "Vault", "Overstock"])]
    filtered = candidates[candidates.apply(
        lambda row: (row["Brand"], row["Product Name"]) not in sales_floor_keys, axis=1
    )]

    low_stock_df = filtered[
        (filtered["Room"] == "Vault") & (filtered["Qty On Hand"] < 5)
        ]
    move_up_df = filtered.drop(low_stock_df.index)

    return move_up_df, low_stock_df, df


# --- EXCEL OUTPUT ---
def save_filtered_excel(move_up_df, low_stock_df, original_path):
    base, _ = os.path.splitext(original_path)
    output_path = base + EXCEL_SUFFIX
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        move_up_df.to_excel(writer, sheet_name="Move_Up_Items", index=False)
        low_stock_df.to_excel(writer, sheet_name="Vault_Low_Stock", index=False)
    return output_path


# --- PDF OUTPUT ---
def build_pdf_section(df, title):
    styles = getSampleStyleSheet()
    elements = [Paragraph(f"<b>{title}</b>", styles["Heading2"])]



    elements.append(Paragraph(datetime.now().strftime(DATE_FORMAT), styles["Normal"]))
    elements.append(Paragraph(" ", styles["Normal"]))  # Spacer


    df = df[COLUMNS_TO_USE].sort_values(by=["Type", "Brand", "Product Name"])
    df = df.drop(columns=["Brand"])

    # Truncate Room and Type for better PDF layout
    df["Room"] = df["Room"].apply(lambda x: truncate_string(x, 8))
    df["Type"] = df["Type"].apply(lambda x: truncate_string(x, 8))
    df["Product Name"] = df["Product Name"].apply(lambda x: truncate_string(x, 75))
    df["Package Barcode"] = df["Package Barcode"].apply(lambda x: str(x)[-6:] if pd.notna(x) else "")

    # Custom headers for PDF only
    headers = list(df.columns)
    headers = ["Type", "Product", "Barcode", "Loc", "Qty"]  # Skipping Brand already
    table_data = [headers] + df.values.tolist()

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

    # Add alternating row colors (zebra striping)
    row_count = len(table_data)
    for i in range(1, row_count):
        if i % 2 == 0:
            bg_color = colors.gainsboro
            table.setStyle([("BACKGROUND", (0, i), (-1, i), bg_color)])

    elements.append(table)
    return elements


def generate_pdf(move_up_df, low_stock_df, source_path):
    base_path = os.path.dirname(source_path)
    output_path = os.path.join(base_path, PDF_FILENAME)
    doc = SimpleDocTemplate(output_path, pagesize=letter)

    elements = []
    elements += build_pdf_section(move_up_df, "Move-Up Inventory List")
    elements.append(PageBreak())
    elements += build_pdf_section(low_stock_df, "Vault Low Stock (<5)")

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
        f"Files created successfully:\n\n{os.path.basename(excel_output)}\n{os.path.basename(pdf_output)}"
    )
