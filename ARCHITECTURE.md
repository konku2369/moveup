# Bisa Inventory Utility -- Architecture & Core Logic

## What This App Does

Bisa Inventory Utility is a Tkinter desktop app for cannabis retail inventory management. It reads METRC inventory exports (Excel/CSV), determines which backstock items need to be "moved up" to the sales floor, and exports sticker-sheet PDFs and Excel reports. It also provides audit PDFs, velocity tracking, expiration analysis, sample tracking, and multi-store comparison.

---

## Data Pipeline

The core data flow follows five stages:

```
File Import --> Column Mapping --> Room Normalization --> Move-Up Computation --> Display/Export
```

### Stage 1: File Import (`data_core.load_raw_df`)

Reads a METRC Excel (.xlsx/.xls) or CSV file into a raw pandas DataFrame.

- **Sweed POS detection**: If the first cell is "export date", skips 3 header rows
- **CSV smart reading**: Sniffs delimiter (comma/tab/semicolon), tries UTF-8 then Latin-1
- **Excel sheet fallback**: Tries named sheet "Inventory Adjustments", falls back to sheet 0
- **Windows unblock**: Removes Zone.Identifier security flag from downloaded files

### Stage 2: Column Mapping (`data_core.automap_columns`)

Maps messy source column names to the app's internal schema:

**Required columns** (`COLUMNS_TO_USE`):
| Internal Name | Example Source Names |
|---|---|
| Type | "Product Type", "Category", "Item Type" |
| Brand | "Brand Name", "Manufacturer" |
| Product Name | "Product", "Item Name", "Title" |
| Package Barcode | Detected via strict METRC regex (must contain "metrc" + id/code/tag/barcode) |
| Room | "Location", "Stock Location", "Bin" |
| Qty On Hand | "Available Qty", "Quantity On Hand" |

**Optional columns** (enable extra features when present):
- Distributor, Store, Size -- used by Audit PDF grouping
- Received Date -- used by stock age calculation and velocity tracking
- Wholesale Cost, Unit Price -- used by Sample Manager

**Mapping logic**:
1. Detect the METRC barcode column using strict regex (requires both "metrc" AND one of: code/id/tag/barcode/package)
2. Fuzzy-match remaining columns using `ALT_NAME_CANDIDATES` lookup table
3. Normalize data types (barcodes to string, qty to int, dates to YYYY-MM-DD)
4. Raise `ValueError` with descriptive message if any required column is missing

### Stage 3: Room Normalization (`data_core.normalize_rooms`)

Applies user-defined room aliases (e.g., "Vault 1" -> "Vault") using case-insensitive matching. The alias map is persisted in `moveup_config.json` and edited via the Filters dialog.

### Stage 4: Move-Up Computation (`data_core.compute_moveup_from_df`)

This is the core algorithm. It determines which backstock items are NOT already on the sales floor and therefore need to be "moved up."

**Algorithm steps:**

1. **Drop incomplete rows** -- Remove rows missing Product Name, Brand, Package Barcode, or Room
2. **Apply brand filter** -- Keep only selected brands (or all if filter contains "ALL")
3. **Apply type filter** -- Keep only selected types (same "ALL" convention)
4. **Apply room aliases** -- Normalize room names via user alias map
5. **Remove accessories** -- Exclude rows where Type contains "accessor" (case-insensitive)
6. **Build Sales Floor set** -- Collect unique (Brand, Product Name) pairs from rooms matching `SALES_FLOOR_ALIASES` ("sales floor", "floor", "salesfloor", "front of house", "foh", etc.)
7. **Filter to candidate rooms** -- Keep only rows in user-selected candidate rooms (e.g., "Backstock", "Incoming Deliveries")
8. **Anti-join** -- Left-merge candidates with the Sales Floor set. Items already on the floor are removed; items NOT on the floor are the move-up candidates

**Key insight**: The match is by **(Brand, Product Name)**, not by barcode. If *any* unit of "Brand X / Product Y" is on the sales floor, then *all* backstock units of that same product are excluded from move-up.

**Diagnostics**: Returns a dict tracking row counts at each stage for debugging:
```python
{"total_loaded": 1250, "after_dropna": 1230, "after_brand": 800,
 "after_type_filter": 750, "after_type": 740, "candidate_pool": 150,
 "removed_as_on_sf": 63, "move_up": 87}
```

### Stage 5: Post-Processing

After the core computation:

1. **Split package aggregation** (`aggregate_split_packages_by_room`) -- Merges duplicate barcodes in the same room, summing qty
2. **Exclusion filtering** -- If `hide_removed=True`, strips user-excluded barcodes from the result
3. **Backstock priority sort** -- Backstock items float to top
4. **Velocity label injection** -- If history snapshots exist, computes and merges a "Velocity" column (Fast/Moderate/Slow/Stale/New)

---

## Velocity Tracking

Velocity tracks how inventory moves across successive imports.

### Snapshots

Each time the user imports a file, a "velocity snapshot" is saved -- a list of every item's barcode, room, qty, and received date at that moment. Snapshots are stored in `velocity_history.json`.

### Metrics (`data_core.compute_velocity_metrics`)

For each current item, the system compares against historical snapshots to compute:

| Metric | Description |
|---|---|
| room_changes | How many times the item moved rooms across snapshots |
| qty_delta | Current qty minus first-seen qty (negative = units sold) |
| sell_rate | Units lost per day (positive = selling) |
| stock_age_days | Days since Received Date (or first snapshot appearance) |
| qty_unchanged_streak | Consecutive recent snapshots where qty didn't change |
| velocity_score | Weighted score: 0.3 x room_score + 0.7 x qty_score |
| velocity_label | Human-readable label (see below) |

### Velocity Labels

| Label | Condition |
|---|---|
| New | Fewer than 2 snapshots of history |
| Fast | velocity_score >= 0.25 |
| Moderate | Score < 0.25 but not stagnant |
| Slow | Qty unchanged for N consecutive imports (default N=3) |
| Stale | Qty unchanged for 2N consecutive imports |
| Sold Out | Item was in history but absent from current inventory |

---

## GUI Architecture

### Main Window (`main.py` -- `MoveUpGUI`)

Four-tab notebook:

1. **Move Up** -- The primary view. Shows items that need restocking. Color-coded:
   - Red: backstock items
   - Pink: user-starred priority items
   - Grey: excluded items (when `hide_removed=False`)
   - Goldenrod: slow/stale velocity
   - Green: fast velocity

2. **Priority!** -- User-starred items (manually marked for urgency)

3. **Excluded** -- Items removed from the move-up list. Double-click to restore.

4. **All Items** -- Full inventory with live search across all columns

### State Management

| State | Type | Purpose |
|---|---|---|
| `current_df` | DataFrame | Full mapped inventory (all rooms, all items) |
| `moveup_df` | DataFrame | Filtered move-up candidates (what's displayed) |
| `velocity_df` | DataFrame | Velocity metrics for all items |
| `excluded_barcodes` | set | Barcodes hidden from move-up view |
| `kuntal_priority_barcodes` | set | Barcodes marked as priority |
| `room_alias_map` | dict | User-defined room name aliases |
| `selected_rooms/brands/types` | list | Active filter selections |

All state is persisted to `moveup_config.json` via `ConfigManager` (atomic writes with backup).

### Interaction Flow

- **Double-click in Move Up tab** -> toggles exclusion (adds/removes from `excluded_barcodes`)
- **Double-click in All Items tab** -> adds to priority
- **Double-click in Priority! tab** -> removes from priority
- **Double-click in Excluded tab** -> restores to move-up list
- **Filter changes** -> triggers full `_recompute_from_current()` pipeline
- **Import new file** -> resets pipeline from Stage 1

---

## Export

### PDF Export (`pdf_export.export_moveup_pdf_paginated`)

Generates paginated sticker-sheet PDFs:

1. Priority items first (marked with star prefix)
2. Backstock items next
3. Other rooms last
4. Paginated by `items_per_page` (default 35)

**Kawaii mode** adds decorative elements (daisies, paw prints, stars, cat faces) in page margins with configurable intensity and color hue. Settings persisted in `kawaii_pdf_settings.json`.

### Audit PDF Export (`pdf_export.export_audit_pdfs`)

Generates two PDFs:
- **Master**: Full audit with quantities filled in
- **Blank**: Same layout with empty qty column (for physical counting)

Grouped by Distributor, Brand, or Type with page breaks between groups.

### Excel Export (`main.py:export_excel`)

Two-sheet workbook:
- Sheet 1: Priority items
- Sheet 2: Move-up items

---

## Satellite Windows

| Window | File | Purpose |
|---|---|---|
| Expiring Items | `mainExpiring.py` | Detects items approaching expiration, groups into time buckets |
| Sample Manager | `mainSamples.py` | Identifies sample items (Wholesale Cost <= $0.01), tracks margins |
| Analytics | `mainAnalytics.py` | Deep inventory analysis: category breakdown, low stock alerts |
| Multi-Store | `mainMultiStore.py` | Compares two store inventories, finds imbalances and transfer opportunities |
| Velocity Tracker | `mainVelocity.py` | Visualizes velocity metrics, slow movers, sold-out items |
| Import History | `mainImportHistory.py` | Timeline of imports with trend sparklines and snapshot comparison |

All satellite windows are independent Toplevel windows that load their own data or receive DataFrames from the main window.

---

## Support Modules

| Module | Purpose |
|---|---|
| `config_manager.py` | Atomic JSON config persistence with backup to `~/.moveup/` |
| `tree_ops.py` | Treeview rendering, sorting, column configuration |
| `dialogs.py` | Column mapping, filters, audit export, manual-add dialogs |
| `kawaii_settings.py` | Kawaii PDF decoration settings and profile computation |
| `kawaii_preview.py` | Live canvas preview of kawaii decoration settings |
| `pdf_common.py` | Shared PDF utilities (palettes, section builder) |
| `inventory_analysis.py` | Pure functions for multi-store comparison (imbalances, transfer recs) |
| `velocity_history.py` | Velocity snapshot persistence (JSON with atomic writes) |
| `import_history.py` | Import aggregate persistence (JSON with atomic writes) |
| `bisa.py` | Bisa the cat -- animated ASCII companion with 10+ behaviors |
| `themes.py` | Shared lavender/purple UI theme for satellite windows |

---

## File Layout

```
moveup/
  main.py                  # Main GUI window
  data_core.py             # Core data pipeline (load, map, compute)
  config_manager.py        # Config persistence
  tree_ops.py              # Treeview rendering and sorting
  dialogs.py               # Dialog windows
  pdf_export.py            # PDF generation (move-up + audit)
  pdf_common.py            # Shared PDF utilities
  kawaii_settings.py        # Kawaii decoration settings
  kawaii_preview.py         # Kawaii live preview
  bisa.py                  # Bisa the cat companion
  themes.py                # UI theme definitions
  mainExpiring.py          # Expiring items window
  mainSamples.py           # Sample manager window
  mainAnalytics.py         # Analytics window
  mainMultiStore.py        # Multi-store comparison window
  mainVelocity.py          # Velocity tracker window
  mainImportHistory.py     # Import history window
  inventory_analysis.py    # Shared analysis functions
  velocity_history.py      # Velocity snapshot storage
  import_history.py        # Import history storage
  moveup_config.json       # User config (gitignored)
  kawaii_pdf_settings.json  # Kawaii settings (gitignored)
  velocity_history.json    # Velocity snapshots (gitignored)
  import_history.json      # Import history (gitignored)
  generated/               # Export output (gitignored)
```
