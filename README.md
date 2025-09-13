# Move-Up Inventory Tool (v2.1)

## Overview
This tool scans your inventory export, finds items in specified back rooms (Incoming Deliveries / Vault / Overstock by default)  
that are **not** on the Sales Floor (matched by Brand + Product Name), and outputs:

- Excel workbook with two tabs:
  - `Move_Up_Items` (primary working list)
  - `Vault_Low_Stock` (Vault items under the threshold; reference only)
- PDF report of the Move-Up list (optionally also includes Vault Low Stock)

## Highlights
- Fast vectorized filters (no Python loops)
- Always excludes Accessories (any Type containing "accessor", case-insensitive)
- Timestamped filenames by default
- Command-line flags to customize behavior
- Works with GUI file picker if you don’t pass `--input`



## Build App code
pyinstaller --onefile --noconsole --name "MoveUp-Inventory v2.1" --icon ihjicon.ico moveup.py


## Install Requirements
```bash
python -m pip install -U pandas openpyxl reportlab
