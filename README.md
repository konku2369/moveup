# Move-Up Inventory Tool (v2.4)

## Overview
This tool scans your inventory export (from **iHeartJane** or **Sweed**), finds items in specified back rooms (default: *Incoming Deliveries*, *Vault*, *Overstock*) that are **not** on the Sales Floor, and outputs:

- **Excel workbook** with two tabs:
  - `Move_Up_Items` (primary working list)
  - `Vault_Low_Stock` (Vault items under the threshold; reference only)
- **PDF report** of the Move-Up list  
  - Always includes Move-Up  
  - Optionally also includes Vault Low Stock  
  - Footer with date (left) and page number (right)

## Highlights
- Works with **iHeartJane** *and* **Sweed** exports:
  - Sweed files are auto-detected (`Export date:` at the top) and normalized
  - First 3 rows of Sweed files are skipped automatically
  - Columns are auto-mapped (e.g. *Location → Room*, *Available Qty → Qty On Hand*)
- Always excludes accessory items (any `Type` containing "accessor")
- Filename prefix + timestamp options
- Room aliasing system (map vendor-specific labels to your canonical room names)
- CLI-friendly, with optional GUI file picker if no input is provided
- Clean, compact PDF with zebra striping
- Page numbers and run date on every PDF page

## Install Requirements
```bash
python -m pip install pandas openpyxl reportlab
