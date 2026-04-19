"""
Help window -- explains how Bisa Inventory Utility works.

Scrollable Toplevel with sections covering the core workflow,
move-up logic, velocity tracking, and all satellite features.
"""

import tkinter as tk
from tkinter import ttk


# ---------------------------------------------------------------------------
# Help content -- each tuple is (section_title, body_text)
# ---------------------------------------------------------------------------
_SECTIONS = [
    ("Overview", """\
Bisa Inventory Utility reads METRC inventory exports (Excel or CSV) and \
figures out which backstock items need to be "moved up" to the sales floor.

It also generates sticker-sheet PDFs for printing, audit PDFs for physical \
counts, and tracks how inventory moves over time."""),

    ("Quick Start", """\
1.  Click "Import File..." and select a METRC export (.xlsx or .csv).
2.  The app auto-detects columns and computes the move-up list.
3.  Review the Move Up tab -- these are items NOT on the sales floor.
4.  Double-click any row to exclude it (or star it as priority).
5.  Click "Export PDF" to generate a printable sticker sheet."""),

    ("How Move-Up Works", """\
The core question the app answers: "Which products in backstock are NOT \
already on the sales floor?"

Step by step:
  1.  Load all inventory rows from the METRC export.
  2.  Apply your filters (brands, types, rooms).
  3.  Remove accessories (items with "Accessory" in the Type column).
  4.  Build a list of every (Brand + Product Name) combo that appears \
in a "Sales Floor" room.
  5.  Look at your candidate rooms (e.g. Backstock, Incoming Deliveries).
  6.  Remove any candidate that has the same Brand + Product Name as \
something already on the floor.
  7.  What's left = items that need to be moved up.

Important: the match is by product name, not barcode. If ANY unit of \
"Brand X / Blue Dream 1g" is on the sales floor, then ALL backstock \
units of that same product are excluded -- even if they have different \
METRC codes. This is intentional: you don't need to restock something \
that's already out."""),

    ("Split Lots", """\
Sometimes Sweed or METRC will split a single lot across rooms -- for \
example, 32 units of barcode XXX on the Sales Floor and 2 units of \
that same barcode XXX in Quarantine.

Since the product name IS on the sales floor, those 2 quarantine units \
will NOT appear in the move-up list. The app treats them as "already \
stocked" because the product is already available to customers."""),

    ("Tabs", """\
Move Up
  The main view. Shows items that need restocking. Color-coded:
    Red = backstock items
    Pink = items you starred as priority
    Grey = excluded items (when "hide removed" is off)
    Gold = slow/stale velocity
    Green = fast-moving items

Priority!
  Items you manually starred for urgency. These appear first in PDF \
exports (marked with a star). Double-click to un-star.

Excluded
  Items you removed from the move-up list. Double-click to restore. \
These items are always excluded from PDF/Excel exports.

All Items
  The full inventory from the imported file. Use the search bar to \
find anything. Double-click to add to Priority."""),

    ("Filters", """\
Click "Filters..." (in Advanced) to configure:

Rooms
  Select which rooms count as "candidate" rooms (where move-up items \
come from). Default: Backstock + Incoming Deliveries.

Brands / Types
  Filter the move-up list to specific brands or product types.

Room Aliases
  Map room names to canonical forms. Example: "Vault 1" and "Vault 2" \
can both map to "Vault". Aliases are case-insensitive.

Skip Sales Floor
  When checked, shows ALL candidate-room items regardless of whether \
the product is already on the floor. Useful for auditing."""),

    ("Velocity Tracking", """\
Each time you import a file, the app saves a "snapshot" of the current \
inventory. Over multiple imports, it compares snapshots to detect movement.

Velocity Labels:
  New       = Not enough history yet (fewer than 2 imports)
  Fast      = Product is actively selling or moving rooms
  Moderate  = Some movement, but not a lot
  Slow      = Quantity unchanged for 3+ consecutive imports
  Stale     = Quantity unchanged for 6+ consecutive imports
  Sold Out  = Was in previous imports but gone from current inventory

Open the Velocity window for detailed metrics: sell rate, stock age, \
room changes, and qty deltas."""),

    ("PDF Export", """\
Export PDF generates a paginated sticker sheet:
  1.  Priority items first (marked with a star)
  2.  Backstock items next
  3.  Other rooms last

Each page has up to N items (configurable via "Items per page").

Kawaii Mode
  Enable decorative elements (daisies, paw prints, stars, cat faces) \
in the page margins. Configure intensity, color hue, and B/W mode \
via "Kawaii PDF Settings...".

Audit PDFs
  Click "Audit PDFs..." to generate two documents:
    Master = full audit with quantities filled in
    Blank  = same layout with empty qty column (for physical counting)
  Group by Distributor, Brand, or Type with page breaks between groups."""),

    ("Excel Export", """\
Export Excel creates a .xlsx workbook with two sheets:
  Sheet 1: Priority items (if any)
  Sheet 2: Move-up items

Available via Advanced > Export Excel."""),

    ("Satellite Windows", """\
Expiring Soon
  Detects items approaching their expiration date. Groups into time \
buckets (0-7 days, 8-14 days, etc.) with urgency color-coding.

Sample Manager
  Identifies sample items (typically Wholesale Cost = $0). Tracks \
profit margins and provides distribution lists.

Velocity
  Detailed velocity metrics for all items. Shows slow movers, sold-out \
items, and full snapshot history.

Multi-Store
  Compare two store inventories side by side. Finds products exclusive \
to one store, quantity imbalances, and generates transfer recommendations.

Analytics
  Deep inventory analysis: category breakdowns, top brands, low/zero \
stock alerts. Supports single-store or two-store comparison.

History
  Timeline of all your imports with trend sparklines. Compare any two \
imports to see what changed (new items, removed items, qty changes, \
room moves)."""),

    ("Bisa", """\
Bisa is your cat companion! She lives in the top-right corner and \
reacts to what you do:
  - Click her to pet her
  - She celebrates when move-ups are found
  - She gets sad when you exclude items
  - She has seasonal outfits in October and December

Her lifetime stats (pets, treats, move-ups) are saved between sessions."""),

    ("Keyboard & Mouse", """\
Double-click a row    Toggle exclude (Move Up tab) / toggle priority \
(All Items tab) / restore (Excluded tab)
Click column header   Sort by that column (click again to reverse)
Search bar (All tab)  Live filter across all columns"""),
]


# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------

class HelpWindow(tk.Toplevel):
    """
    Scrollable Toplevel window containing the full app help text.

    Content is sourced from the ``_SECTIONS`` list of ``(title, body)`` tuples
    defined at module level.  Each section is rendered as a bold header,
    a horizontal separator, and a read-only ``tk.Text`` widget sized to its
    content.  Mouse-wheel scrolling is bound globally while the window is open
    and unbound on ``<Destroy>`` to prevent leaking the binding.
    """

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Help -- How It Works")
        self.geometry("720x640")
        self.minsize(480, 320)

        # Scrollable frame
        outer = ttk.Frame(self)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas, padding=16)

        self.inner.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Mouse-wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        self.bind("<Destroy>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # Populate
        self._build_content()

        # Close button
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", pady=(0, 10))
        ttk.Button(btn_frame, text="Close", command=self.destroy).pack()

    def _build_content(self):
        """Populate ``self.inner`` with all help sections from ``_SECTIONS``."""
        # Title
        title = ttk.Label(
            self.inner,
            text="Bisa Inventory Utility -- Help",
            font=("Segoe UI", 16, "bold"),
        )
        title.pack(anchor="w", pady=(0, 12))

        for section_title, body in _SECTIONS:
            # Section header
            header = ttk.Label(
                self.inner,
                text=section_title,
                font=("Segoe UI", 12, "bold"),
            )
            header.pack(anchor="w", pady=(14, 4))

            # Separator
            sep = ttk.Separator(self.inner, orient="horizontal")
            sep.pack(fill="x", pady=(0, 6))

            # Body text
            try:
                bg = self.inner.winfo_toplevel().cget("background")
            except Exception:
                bg = "#f0f0f0"
            txt = tk.Text(
                self.inner,
                wrap="word",
                font=("Segoe UI", 10),
                bg=bg,
                relief="flat",
                borderwidth=0,
                highlightthickness=0,
                padx=8,
                pady=4,
                cursor="arrow",
            )
            txt.insert("1.0", body.strip())
            txt.config(state="disabled")

            # Auto-height: measure how many display lines are needed
            txt.pack(fill="x", pady=(0, 4))
            self.update_idletasks()
            line_count = int(txt.index("end-1c").split(".")[0])
            txt.config(height=line_count)


def open_help_window(parent):
    """
    Open a new ``HelpWindow`` as a child of *parent*.

    Creates a new window each time it is called; if the user opens Help
    multiple times, multiple windows are allowed (each is independent).
    Called from the toolbar in ``main.py``.
    """
    HelpWindow(parent)
