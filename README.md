# Bisa Inventory Utility

**v4.1** — A Tkinter desktop app for cannabis retail inventory management.

Reads METRC inventory exports (Excel/CSV), computes which backstock items need "move-up" to the sales floor, and exports sticker-sheet PDFs. Features a kawaii cat companion named Bisa, velocity tracking, sample management, and expiring item monitoring.

Built with love at Bisa Lina.

---

## Quick Start

### Requirements
- Python 3.10+
- Dependencies: `pandas`, `openpyxl`, `reportlab`

```bash
pip install pandas openpyxl reportlab
```

### Run
```bash
python main.py
```

### Build (Windows .exe)
```bash
pip install pyinstaller
pyinstaller "Bisa Inv Utility v4.15release.spec"
```

---

## What This Program Does (The Big Picture)

At a cannabis retail store, inventory lives in multiple rooms: **Sales Floor** (what customers see), **Backstock** (overflow storage), **Vault** (high-value items), etc. Throughout the day, items sell off the floor and need to be replaced from the back rooms.

This app answers: **"Which products in the back rooms are NOT currently on the Sales Floor?"** Those are your **move-up candidates** — items that need to be brought out front.

### The Workflow

1. **Export inventory** from your POS system (Sweed, Dutchie, etc.) or directly from METRC
2. **Import the file** into this app (drag-and-drop or File > Import)
3. The app **auto-detects columns** and maps them to a standard format
4. The **move-up algorithm** identifies items to bring to the floor
5. **Print sticker PDFs** — staff cuts them into strips and places them on products
6. Optionally: track **velocity** (which items are selling vs. sitting), manage **samples**, monitor **expiring items**

---

## How the Data Pipeline Works (data_core.py)

This is the heart of the program. Every step is a pure function with no GUI dependency.

### Step 1: Load the File (`load_raw_df`)

Reads an Excel or CSV file into a pandas DataFrame. Handles:
- Windows "downloaded from internet" security flags (auto-removed)
- Sweed POS exports (auto-detects and skips 3 metadata header rows)
- CSV delimiter auto-detection (comma, tab, semicolon)
- UTF-8 and Latin-1 encoding fallback

### Step 2: Column Mapping (`automap_columns`)

METRC exports have messy column names that differ between POS systems. This function:

1. **Finds the METRC barcode column** — uses strict regex detection (`detect_metrc_source_column`). The column must contain both "metrc" AND one of: code, id, tag, barcode, package. This prevents false positives.

2. **Maps remaining columns** using fuzzy name matching against `ALT_NAME_CANDIDATES`. For example, "Available Qty" → "Qty On Hand", "Item Name" → "Product Name".

3. **Normalizes data types** — barcodes become strings, qty becomes int, dates become YYYY-MM-DD.

The 6 required columns after mapping:
| Internal Name | What It Is |
|---|---|
| `Type` | Product category (Flower, Edible, Concentrate, etc.) |
| `Brand` | Manufacturer/brand name |
| `Product Name` | Full product name |
| `Package Barcode` | METRC tracking barcode (unique ID) |
| `Room` | Physical location (Sales Floor, Backstock, Vault, etc.) |
| `Qty On Hand` | Number of units in stock |

Optional columns (used by audit features): Distributor, Store, Size, Received Date, Wholesale Cost, Unit Price.

### Step 3: Room Normalization (`normalize_rooms`)

Applies user-defined aliases so messy room names get cleaned up:
- "Vault 1" → "Vault"
- "SF" → "Sales Floor"
- "vault" → "Vault" (case-insensitive)

### Step 4: Move-Up Computation (`compute_moveup_from_df`)

This is the core algorithm:

```
1. Filter by user's brand/type selections
2. Exclude accessories (rolling papers, lighters don't need move-up)
3. Build a set of (Brand, Product Name) combos that ARE on the Sales Floor
4. Filter to items in candidate rooms only (Backstock, Vault, etc.)
5. ANTI-JOIN: remove items whose product IS already on the floor
6. What's left = MOVE-UP CANDIDATES
```

**Example:**
- "Blue Dream 3.5g" by BrandX exists on the Sales Floor → any BrandX "Blue Dream 3.5g" in Backstock is NOT a move-up candidate (it's already out front)
- "Purple Haze 1g" by BrandY is in Backstock but NOT on the Sales Floor → this IS a move-up candidate

The function also returns a diagnostics dict showing how many items were filtered at each step — useful for debugging "why is my item missing from the list?"

### Step 5: Split Package Aggregation (`aggregate_split_packages_by_room`)

Handles when the same METRC barcode appears multiple times:
- Same SKU in the **same room** with duplicate rows → merged (qty summed)
- Same SKU in **different rooms** → kept as separate rows (intentional!)

This is important because a product in Vault and Backstock needs separate stickers.

### Step 6: Velocity Tracking (Optional)

Each import saves a snapshot of every item's barcode, room, and qty. Over successive imports, `compute_velocity_metrics()` compares snapshots to compute:

| Metric | What It Means |
|---|---|
| `sell_rate` | Units lost per day (positive = selling) |
| `velocity_label` | New, Fast, Moderate, Slow, Stale, or Sold Out |
| `qty_unchanged_streak` | Consecutive imports where qty didn't change |
| `room_changes` | How many times the item moved between rooms |

**Scoring:** 70% qty sell-through + 30% room movement. Room changes are weighted less because they're rare in practice — quantity change is the dominant signal.

**Labels:**
- **New**: Only seen in 1 import (not enough data yet)
- **Fast**: Score >= 0.25 and not stagnant
- **Moderate**: Score < 0.25 but still changing
- **Slow**: Qty unchanged for >= 3 consecutive imports
- **Stale**: Qty unchanged for >= 6 consecutive imports
- **Sold Out**: Was in history but disappeared from current inventory

---

## Architecture

```
main.py                 <- GUI entry point (MoveUpGUI class)
|-- data_core.py        <- Pure logic: column mapping, move-up computation,
|                          velocity scoring, room normalization
|-- config_manager.py   <- Config persistence (moveup_config.json)
|-- tree_ops.py         <- Treeview rendering and column management
|-- dialogs.py          <- Dialog windows (filters, column mapping, audit, manual add)
|-- pdf_export.py       <- Move-up + audit PDF export (with kawaii decorations)
|-- pdf_common.py       <- Reusable PDF table library (standalone, no app dependencies)
|-- kawaii_settings.py   <- Kawaii PDF decoration settings + profiles
|-- kawaii_preview.py    <- Live preview dialog for kawaii settings
|-- bisa.py             <- Bisa the cat companion widget
|-- mainExpiring.py     <- Expiring items window (Toplevel)
|-- mainSamples.py      <- Sample manager window (Toplevel)
+-- mainVelocity.py     <- Velocity tracker window (Toplevel)
```

### Key Design Decisions

- **`COLUMNS_TO_USE`** in `data_core.py` is the single source of truth for column names
- **`automap_columns()`** requires a METRC column (strict regex) -> mapped to "Package Barcode"
- **Config** persisted atomically (`.tmp` -> `os.replace()`) with backup to `~/.moveup/`
- **No global state** -- dialog windows communicate via explicit callbacks
- **Satellite windows** (`mainExpiring`, `mainSamples`, `mainVelocity`) are self-contained Toplevels
- **`pdf_common.py`** has zero app dependencies -- can be reused in any project
- **Silent exception handling** logs to `print()` with `[moveup]` prefix; never crashes the UI

### Data Flow Diagram

```
User imports .xlsx/.csv
        |
        v
  load_raw_df()          <- handles encoding, Sweed detection, Windows unblock
        |
        v
  automap_columns()      <- METRC detection + fuzzy column name matching
        |
        v
  normalize_rooms()      <- apply user-defined room aliases
        |
        v
  compute_moveup_from_df()  <- THE CORE: filter by room, anti-join against Sales Floor
        |
        v
  aggregate_split_packages_by_room()  <- merge duplicate rows per (barcode, room)
        |
        v
  [inject velocity labels]  <- if history exists, add Slow/Fast/etc. labels
        |
        v
  render in treeview / export PDF / export Excel
```

---

## GUI Structure (main.py)

The main window has:

**Toolbar:**
- Import File / Re-import / Map Columns
- Export PDF / Export Excel / Audit PDFs / Open Output Folder
- Filters / Kawaii Settings
- Expiring Items / Sample Manager / Velocity Tracker

**Notebook Tabs:**
1. **Move-Up** — Items that need to go to the Sales Floor. Backstock items shown in red with alert emoji. Priority items in pink with dog emoji.
2. **Priority!** — User-starred items. These appear first in the PDF export with a star icon. Add items by double-clicking in Move-Up or using Manual Add.
3. **Excluded** — Items the user removed from the Move-Up view. Double-click to restore.
4. **All Items** — Full inventory with live search. Double-click to add to Priority.

**Status Bar:**
- Total rows loaded, move-up count, filter summary
- Diagnostics line showing how many items were filtered at each pipeline step

---

## PDF Exports (pdf_export.py)

### Move-Up Sticker PDF
- Portrait letter-size, paginated (30-35 items per page)
- Priority items at the top (marked with star)
- Kawaii decorations: stars, daisies, paw prints, cat faces in margins
- B/W mode available for non-color printers
- Configurable filename prefix and timestamp

### Audit PDFs
- Generates TWO files: Master (with qty) and Blank (count column empty)
- Master: reference showing what SHOULD be on the shelf
- Blank: staff walks the floor and writes actual counts
- Grouped by distributor, brand, or type (configurable)
- One-click Accessory Audit mode

---

## Velocity Tracker (mainVelocity.py)

Opened from the main menu. Shows:

| Tab | Shows |
|---|---|
| Overview | Dashboard with headline numbers, date range of snapshots |
| Item Detail | Every item with status, sell rate, qty changes |
| Slow Movers | Items flagged Slow/Stale with adjustable threshold |
| Sold Out | Items from history that disappeared |
| History | List of all saved snapshots with purge controls |

Data lives in `velocity_history.json`. Each import adds a snapshot.

---

## Config System (config_manager.py)

All 22+ settings persist to `moveup_config.json` with atomic writes:
1. Write to `.tmp` file
2. `os.replace()` atomically swaps `.tmp` -> `.json`
3. Backup copy written to `~/.moveup/moveup_config_backup.json`

If the main config is missing but backup exists, user is prompted to restore.

Key config entries: room aliases, filter selections, excluded/priority barcodes, Bisa companion stats, last import directory, items per page, filename prefix.

---

## Kawaii System (kawaii_settings.py + kawaii_preview.py)

The kawaii decoration system for PDFs:

**Presets:** Minimal, Cute, Extra (controls base intensity)

**User Sliders:**
- `bg_hue_pct`: 0=pink, 100=lavender/purple
- `elem_intensity`: 0=corners only, 100=decorations everywhere

**Decorations (in page margins only):**
- Background tint wash (pink or lavender, adjustable alpha)
- Decorative border
- Daisies (10-petal flowers, pool of 15 positions)
- Paw prints (pool of 10 positions)
- Cat faces (randomly scattered)
- Sparkle stars (randomly scattered)

All positions are margin-safe: decorations never overlap the table content area.

---

## File Descriptions

| File | Purpose | Tk Dependency |
|------|---------|:---:|
| `main.py` | Main GUI, toolbar, treeviews, import/export workflows | Yes |
| `data_core.py` | Constants, column mapping, move-up logic, velocity scoring | No |
| `config_manager.py` | Load/save `moveup_config.json` with atomic writes | Optional |
| `tree_ops.py` | Treeview column setup, sorting, rendering | Yes (ttk only) |
| `dialogs.py` | Filter, column mapping, audit, manual add dialogs | Yes |
| `pdf_export.py` | Move-up sticker PDFs + audit PDFs with kawaii decorations | No |
| `pdf_common.py` | Reusable PDF table builder (palettes, widths, styles) | No |
| `kawaii_settings.py` | Kawaii decoration profiles + persistence | No |
| `kawaii_preview.py` | Live Tk preview of kawaii PDF settings | Yes |
| `bisa.py` | Bisa the cat companion widget | Yes |
| `mainExpiring.py` | Expiring items analysis + PDF export | Yes |
| `mainSamples.py` | Sample inventory manager + PDF export | Yes |
| `mainVelocity.py` | Velocity tracker + PDF export | Yes |
| `velocity_history.py` | Snapshot storage for velocity tracking | No |

---

## Config Files (auto-generated, gitignored)

- `moveup_config.json` -- App settings, filters, lifetime stats
- `kawaii_pdf_settings.json` -- PDF decoration preferences
- `velocity_history.json` -- Velocity snapshot history
- `generated/` -- PDF/Excel export output directory

---

## Using pdf_common.py in Another Project

`pdf_common.py` is fully standalone. Copy it into any project that needs clean PDF table reports. Only dependency: `reportlab`.

```python
from pdf_common import build_section_pdf, PALETTE_KAWAII, PALETTE_BW, PALETTE_PLAIN

sections = [
    ("Inventory", ["SKU", "Product", "Qty"], [
        ["A001", "Widget", 10],
        ["A002", "Gadget", 5],
    ]),
    ("Summary", ["Metric", "Value"], [
        ["Total Items", 15],
        ["Total SKUs", 2],
    ]),
]

# Basic usage
build_section_pdf("report.pdf", "Inventory Report", "March 2026", sections)

# With kawaii palette
build_section_pdf("kawaii.pdf", "Report", "Today", sections, palette=PALETTE_KAWAII)

# Custom column width overrides
build_section_pdf("custom.pdf", "Report", "Today", sections,
                  width_overrides={"Product": 4.0, "Qty": 0.5})

# Portrait orientation
build_section_pdf("portrait.pdf", "Report", "Today", sections, orientation="portrait")
```

### Available Functions

| Function | Purpose |
|---|---|
| `build_section_pdf()` | High-level: multi-section PDF with title, subtitle, tables |
| `compute_column_widths()` | Auto-size columns by name heuristics |
| `build_table_style()` | Full table style from a color palette |
| `auto_align_commands()` | Center/right-align columns by name keywords |
| `draw_footer()` | Page footer with timestamp + page number |
| `truncate_text()` | Truncate with ellipsis |

### Available Palettes

| Palette | Description |
|---|---|
| `PALETTE_KAWAII` | Pink/lavender (matches the kawaii theme) |
| `PALETTE_BW` | Greyscale (for B/W printers) |
| `PALETTE_PLAIN` | Neutral grey (good default for non-themed apps) |

---

## Bisa the Cat (bisa.py)

Bisa is the app's ASCII art cat companion. She lives in the bottom-right corner of the main window and reacts to user actions:
- Greets the user on startup
- Celebrates when move-ups are detected between imports
- Can be petted (click) and given treats (click blank space)
- Tracks lifetime stats (pets, treats, move-ups found)
- Has a catnip reward system
- Can be renamed by the user

She's secretly there for Daisy, who loves cats.

---

## Glossary

| Term | Meaning |
|---|---|
| **METRC** | Marijuana Enforcement Tracking Reporting Compliance — state-level cannabis tracking system. Every package gets a unique barcode. |
| **Move-up** | An item in a back room that should be brought to the Sales Floor because no matching product is currently out front. |
| **Candidate room** | A room that items are pulled FROM for move-up (e.g., Backstock, Vault). |
| **Sales Floor** | The room where customers can see/buy products. Items here are excluded from move-up candidates. |
| **Split package** | When the same METRC barcode has units in multiple rooms (e.g., 5 in Vault, 3 in Backstock). Each room gets its own row. |
| **Priority!** | User-starred items that appear first in the PDF export (the "bring these out FIRST" list). |
| **Kawaii** | Japanese for "cute" — the decorative pink/lavender theme with daisies, paws, and cat faces on PDFs. |
| **Velocity** | How fast an item is moving (selling). Computed from qty changes across successive imports. |
| **Snapshot** | A record of every item's barcode, room, and qty at one point in time. Used for velocity tracking. |
| **Anti-join** | Database operation: "give me rows from A that have NO match in B". Used to find items NOT on the Sales Floor. |
