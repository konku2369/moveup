"""
Treeview rendering and column management for the MoveUp GUI.

All functions in this module are pure operations on ttk.Treeview widgets and
pandas DataFrames. There is NO import of MoveUpGUI — every function receives
its inputs explicitly, which makes them independently testable and keeps the
rendering logic decoupled from the application state.

ARCHITECTURE OVERVIEW
=====================
The GUI has four treeview tabs, each rendered by a dedicated function:

  1. render_moveup_tree()   — Move-Up tab: candidates that need to go to the
                              sales floor. Rows are color-coded by room/priority/
                              velocity and optionally show excluded items greyed out.

  2. render_kuntal_tree()   — Priority! tab: user-starred items. These are items
                              the user has flagged with "Toggle Priority!" for
                              guaranteed inclusion on the sticker PDF.

  3. render_excluded_tree() — Excluded / Removed tab: items the user has clicked
                              "Toggle Remove" on. Shown here so they can be
                              restored, but hidden from the Move-Up tab.

  4. render_all_tree()      — All Items tab: every row in the full inventory
                              DataFrame, with live multi-token AND search.

COLOR-TAG SYSTEM
================
Rows in the Move-Up treeview are tagged at render time. Tags apply in priority
order — a higher-priority tag prevents lower-priority tags from overriding it:

  kuntal   (highest) — pink  #c0007a  → user-starred priority item (🐶🌼 prefix)
  backstock           — red   #cc2222  → item is physically in the Backstock room (🚨 prefix)
  excluded            — grey  #999999  → barcode is in excluded_barcodes set (hide_removed=False)
  vel_slow            — gold  #B8860B  → unchanged for ≥ slow_threshold consecutive imports
  vel_fast (lowest)   — green #228B22  → actively selling (velocity_label == "Fast")

COLUMN MANAGEMENT
=================
  configure_tree_columns() — sets widths and wires click-to-sort on all headings.
  sort_tree()              — in-place numeric/string sort, maintains per-column state.
  refresh_treeview_columns() — rebuilds column set when optional columns (Received Date,
                               Velocity) appear or disappear after a new import.
  get_display_cols()       — merges active_columns with any extra columns present in df.
"""

from typing import Dict, List, Optional

import pandas as pd
import tkinter as tk
from tkinter import ttk, StringVar

from data_core import COLUMNS_TO_USE, TYPE_TRUNC_LEN, ellipses


# ---------------------------------------------------------------------------
# Reusable scrollable treeview factory
# ---------------------------------------------------------------------------

# Default column-width heuristics by keyword. Satellite windows can override.
_COL_WIDTH_HINTS = {
    "product": 300,
    "reason": 140,
    "from": 140,
    "to": 140,
    "overstocked": 140,
    "priority": 70,
    "ratio": 70,
    "diff": 70,
}
_COL_WIDTH_DEFAULT = 110


def make_scrollable_tree(
    parent: tk.Widget,
    columns: List[str],
    *,
    height: int = 20,
    horizontal: bool = False,
    col_widths: Optional[Dict[str, int]] = None,
) -> ttk.Treeview:
    """Create a Treeview with a vertical (and optional horizontal) scrollbar inside *parent*.

    This factory consolidates what used to be 10+ identical copy-pasted blocks
    across the satellite window files. All treeviews in the app should be created
    through this function.

    Column widths are resolved in priority order:
      1. Explicit override in col_widths dict
      2. Keyword match in _COL_WIDTH_HINTS (e.g. "product" in col name → 300px)
      3. Default: _COL_WIDTH_DEFAULT (110px)

    Layout strategy:
      - horizontal=False: Treeview + VSB packed side by side. Simple and sufficient
        for most tables where horizontal scrolling is not needed.
      - horizontal=True: Treeview + VSB + HSB placed in a grid so both scrollbars
        track correctly. Required for wide tables (audit reports, multi-store).

    Parameters
    ----------
    parent : tk.Widget
        The container widget (Frame, tab, etc.) that will hold the treeview.
    columns : list[str]
        Column identifier strings. Each string is used both as the internal
        column ID and as the visible heading text.
    height : int
        Number of visible data rows before scrolling is needed (default 20).
        Does not cap the total number of rows; only affects the initial height.
    horizontal : bool
        Whether to attach a horizontal scrollbar (default False).
    col_widths : dict[str, int], optional
        Explicit per-column pixel widths that override the keyword heuristics.

    Returns
    -------
    ttk.Treeview
        The configured, laid-out treeview. Scrollbars are internal implementation
        details — callers only need to keep a reference to the returned Treeview.
    """
    frm = ttk.Frame(parent)
    frm.pack(fill="both", expand=True)

    tree = ttk.Treeview(frm, columns=tuple(columns), show="headings", height=height)
    vsb = ttk.Scrollbar(frm, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)

    if horizontal:
        hsb = ttk.Scrollbar(frm, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)
    else:
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

    overrides = col_widths or {}
    for col in columns:
        tree.heading(col, text=col)
        if col in overrides:
            w = overrides[col]
        else:
            low = col.lower()
            w = _COL_WIDTH_DEFAULT
            for keyword, kw in _COL_WIDTH_HINTS.items():
                if keyword in low:
                    w = kw
                    break
        tree.column(col, width=w, anchor="w")

    return tree


# ---------------------------------------------------------------------------
# Display-column helper
# ---------------------------------------------------------------------------

def get_display_cols(
    active_columns: List[str],
    df: Optional[pd.DataFrame],
    extra_columns: List[str],
) -> List[str]:
    """Return the ordered list of columns to display in the Move-Up treeview.

    The display columns are built from two sources:
      - active_columns: the user-configured subset of COLUMNS_TO_USE. The user
        can hide columns they don't care about (e.g. hide "Type" for a visual
        preference). Stored in config as "active_columns".
      - extra_columns: optional columns appended when they exist in *df*. These
        are currently ["Received Date", "Velocity"]. They appear only after an
        import that provides those fields — Received Date requires the METRC
        export to include it; Velocity requires at least 2 historical snapshots.

    Extra columns are checked via df.columns membership so that the treeview
    never shows a column header for data that isn't actually there.

    Parameters
    ----------
    active_columns : list[str]
        The base columns to display (user preference, subset of COLUMNS_TO_USE).
    df : DataFrame or None
        The current inventory DataFrame. If None or empty, no extras are added.
    extra_columns : list[str]
        Candidate extra columns to append (e.g. ["Received Date", "Velocity"]).

    Returns
    -------
    list[str]
        Ordered column list: active_columns + any extras present in df.
    """
    base = list(active_columns)
    if df is not None and not df.empty:
        for col in extra_columns:
            if col in df.columns and col not in base:
                base.append(col)
    return base


# ---------------------------------------------------------------------------
# Column configuration & sorting
# ---------------------------------------------------------------------------

def configure_tree_columns(
    tree: ttk.Treeview,
    cols: List[str],
    sort_state: Dict[str, Dict[str, bool]],
    sort_fn,
) -> None:
    """Set column widths, heading text, and wire click-to-sort on every heading.

    Called whenever the column set changes — either at initial build or after
    refresh_treeview_columns() detects that optional columns appeared or
    disappeared following a new import.

    Column widths are fixed heuristics: 150px for most columns, 440px for
    "Product Name" (the longest field). These are intentionally not configurable
    per-column in the UI because the layout works well for the typical METRC
    export structure.

    Each heading click is bound to sort_fn(tree, tree_id, col) so that clicking
    the same column twice toggles ascending/descending order.

    Parameters
    ----------
    tree : ttk.Treeview
        The treeview widget to configure.
    cols : list[str]
        Column identifiers (same strings passed to tree.config(columns=...)).
    sort_state : dict
        Shared mutable dict owned by MoveUpGUI._sort_state. Keyed by tree_id
        (str(id(tree))) → inner dict of col → bool (True = ascending).
        Initialised here for new tree IDs.
    sort_fn : callable
        The sort function to bind. Signature: sort_fn(tree, tree_id, col).
        Passed in from MoveUpGUI to keep tree_ops free of app imports.
    """
    tree_id = str(id(tree))
    if tree_id not in sort_state:
        sort_state[tree_id] = {}

    for col in cols:
        tree.heading(col, text=col)
        tree.column(
            col,
            width=150 if col != "Product Name" else 440,
            anchor="w",
        )
        tree.heading(
            col,
            command=lambda c=col, t=tree, tid=tree_id: sort_fn(t, tid, c),
        )


def sort_tree(
    tree: ttk.Treeview,
    tree_id: str,
    col: str,
    sort_state: Dict[str, Dict[str, bool]],
) -> None:
    """Sort a Treeview in-place by *col*, toggling ascending/descending on each call.

    Sort strategy (tried in order):
      1. Numeric — every non-empty cell is cast to float. Empty strings are
         mapped to -inf so blank cells sink to the bottom rather than the top.
      2. String — case-insensitive fallback when any value fails float conversion.
         This handles product names, brand names, room names, etc.

    After sorting, the heading for *col* gets a ▲ or ▼ appended. All other
    column headings have their existing ▲/▼ stripped so only the active sort
    column shows the indicator.

    State: sort_state[tree_id][col] stores True (ascending) or False (descending).
    First click on a new column → ascending. Subsequent clicks → toggle.
    The state persists as long as the widget exists, but resets if the treeview
    is rebuilt (e.g. after a refresh_treeview_columns call).

    Parameters
    ----------
    tree : ttk.Treeview
        The treeview widget to sort in-place.
    tree_id : str
        Opaque key for this treeview in sort_state. Use str(id(tree)).
    col : str
        Column identifier string (must exist in tree["columns"]).
    sort_state : dict
        Shared mutable dict owned by MoveUpGUI._sort_state.
    """
    state = sort_state.setdefault(tree_id, {})
    ascending = not state.get(col, True)   # first click → descending: default True, NOT True = False (ascending=False → reverse=True)
    state[col] = ascending

    rows = [(tree.set(iid, col), iid) for iid in tree.get_children("")]

    # Try numeric sort first. Empty cells get -inf so they sort to bottom.
    # If any value can't be parsed as float, fall back to case-insensitive string sort.
    try:
        rows.sort(
            key=lambda x: float(x[0]) if x[0] != "" else float("-inf"),
            reverse=not ascending,
        )
    except (ValueError, TypeError):
        rows.sort(key=lambda x: str(x[0]).lower(), reverse=not ascending)

    for pos, (_, iid) in enumerate(rows):
        tree.move(iid, "", pos)

    for c in tree["columns"]:
        current = tree.heading(c, "text")
        clean = current.removesuffix(" ▲").removesuffix(" ▼")
        if c == col:
            arrow = " ▲" if ascending else " ▼"
            tree.heading(c, text=clean + arrow)
        else:
            tree.heading(c, text=clean)


# ---------------------------------------------------------------------------
# Refresh all four treeviews when display columns change
# ---------------------------------------------------------------------------

def refresh_treeview_columns(
    moveup_tree: ttk.Treeview,
    k_tree: ttk.Treeview,
    x_tree: ttk.Treeview,
    all_tree: ttk.Treeview,
    active_columns: List[str],
    extra_columns: List[str],
    df: Optional[pd.DataFrame],
    sort_state: Dict[str, Dict[str, bool]],
    sort_fn,
) -> None:
    """Rebuild the column configuration of all four treeviews when the column set changes.

    Called by MoveUpGUI._recompute_from_current() after every import because
    optional columns ("Received Date", "Velocity") may appear or disappear
    depending on whether the new file has date data or whether velocity history
    has accumulated enough snapshots to compute labels.

    Column sets:
      - Move-Up tree: get_display_cols(active_columns, df, extra_columns)
        → user-configured subset + optional extras.
      - Priority! / Excluded / All Items trees: always show the full
        COLUMNS_TO_USE set + any extras (so users see all fields on those tabs).

    Rebuild is skipped if the column list hasn't changed — this avoids
    unnecessary flicker when the user refreshes without the column set changing.

    Parameters
    ----------
    moveup_tree, k_tree, x_tree, all_tree : ttk.Treeview
        The four main treeview widgets (Move-Up, Priority!, Excluded, All Items).
    active_columns : list[str]
        User-configured column subset for the Move-Up tab.
    extra_columns : list[str]
        Optional column names to append when present in df (e.g. ["Received Date", "Velocity"]).
    df : DataFrame or None
        Current inventory data (used to check which extras are present).
    sort_state : dict
        Shared sort-state dict passed through to configure_tree_columns.
    sort_fn : callable
        Sort function bound to heading clicks.
    """
    moveup_cols = get_display_cols(active_columns, df, extra_columns)

    # kuntal / excluded / all always show full COLUMNS_TO_USE + extras
    full_cols = list(COLUMNS_TO_USE)
    if df is not None and not df.empty:
        for col in extra_columns:
            if col in df.columns and col not in full_cols:
                full_cols.append(col)

    # Only rebuild if columns actually changed (avoids flicker)
    current_moveup = list(moveup_tree["columns"])
    if current_moveup != moveup_cols:
        moveup_tree.config(columns=tuple(moveup_cols))
        configure_tree_columns(moveup_tree, moveup_cols, sort_state, sort_fn)

    current_extra = list(k_tree["columns"])
    if current_extra != full_cols:
        for t in (k_tree, x_tree, all_tree):
            t.config(columns=tuple(full_cols))
            configure_tree_columns(t, full_cols, sort_state, sort_fn)


# ---------------------------------------------------------------------------
# Render helpers
# ---------------------------------------------------------------------------

def render_moveup_tree(
    tree: ttk.Treeview,
    df: pd.DataFrame,
    display_cols: List[str],
    kuntal_priority_barcodes: set,
    excluded_barcodes: set,
    hide_removed: bool,
) -> None:
    """Populate the Move-Up treeview with color-coded, prefixed rows.

    This is the most visually complex render function because it applies the
    full tag system. For each row the function determines:

      1. Whether the barcode is in kuntal_priority_barcodes → kuntal tag + 🐶🌼 prefix
      2. Whether the room is "backstock" → backstock tag + 🚨 prefix
      3. Whether the barcode is excluded AND hide_removed is False → excluded tag (grey)
      4. If none of the above, whether velocity label is Slow/Stale/Fast → vel_* tag

    The tag priority matters: a kuntal item that is also backstock gets the kuntal
    (pink) color, not the backstock (red) color, because kuntal is checked first.
    Velocity tags only apply when no higher-priority tag is set.

    The Room column value is prefixed with emoji only when the treeview is actually
    displaying that column — disp_idx_room is None if Room is hidden.

    Type values are truncated to TYPE_TRUNC_LEN (7 chars) to keep the column compact.
    Extra columns (Received Date, Velocity) are fetched from pre-extracted Series
    to avoid per-row .loc lookups on large DataFrames.

    Returns early (renders nothing) if:
      - df is None or empty
      - any required core column is missing from df.columns

    Parameters
    ----------
    tree : ttk.Treeview
        The Move-Up tab treeview widget. Cleared before repopulating.
    df : pd.DataFrame
        The computed move-up DataFrame (output of _recompute_from_current).
        When hide_removed=True, excluded rows have already been filtered out.
        When hide_removed=False, excluded rows are present and tagged grey.
    display_cols : list[str]
        Ordered column list from get_display_cols(). Determines which columns
        are shown and their order.
    kuntal_priority_barcodes : set[str]
        Barcodes marked Priority! by the user. These get pink + emoji prefix.
    excluded_barcodes : set[str]
        Barcodes the user has removed. Only relevant when hide_removed=False.
    hide_removed : bool
        When True, excluded rows were already filtered from df before this call.
        When False, excluded rows are in df and rendered grey here.
    """
    for i in tree.get_children():
        tree.delete(i)

    if df is None or df.empty:
        return

    core_cols = COLUMNS_TO_USE

    idx_bar  = core_cols.index("Package Barcode")
    idx_room = core_cols.index("Room")
    idx_type = core_cols.index("Type")

    disp_idx_bar  = display_cols.index("Package Barcode") if "Package Barcode" in display_cols else None
    disp_idx_room = display_cols.index("Room")            if "Room"            in display_cols else None
    disp_idx_type = display_cols.index("Type")            if "Type"            in display_cols else None

    extra_cols = [c for c in display_cols if c not in core_cols]
    extra_series = {
        c: df[c].reset_index(drop=True)
        for c in extra_cols if c in df.columns
    }

    core_missing = [c for c in core_cols if c not in df.columns]
    if core_missing:
        return

    for row_idx, full_row in enumerate(
        df[core_cols].itertuples(index=False, name=None)
    ):
        bc = str(full_row[idx_bar]).strip()
        room_lower = str(full_row[idx_room]).strip().lower()
        is_backstock = (room_lower == "backstock")
        is_kuntal = (bc in kuntal_priority_barcodes)

        vals = []
        for c in display_cols:
            if c in core_cols:
                vals.append(full_row[core_cols.index(c)])
            else:
                vals.append(
                    extra_series[c].iloc[row_idx] if c in extra_series else ""
                )

        if disp_idx_type is not None:
            vals[disp_idx_type] = ellipses(str(vals[disp_idx_type]), TYPE_TRUNC_LEN)

        prefix = ""
        if is_kuntal:
            prefix += "🐶🌼 "
        if is_backstock:
            prefix += "🚨 "
        if prefix and disp_idx_room is not None:
            vals[disp_idx_room] = f"{prefix}{vals[disp_idx_room]}"

        tags = []
        if bc and (bc in excluded_barcodes) and not hide_removed:
            tags.append("excluded")
        if is_backstock:
            tags.append("backstock")
        if is_kuntal:
            tags.append("kuntal")

        # Velocity color-coding (only if no higher-priority tag)
        if not tags and "Velocity" in display_cols:
            vel_idx = display_cols.index("Velocity")
            vel_val = str(vals[vel_idx]).strip() if vel_idx < len(vals) else ""
            if vel_val in ("Slow", "Stale"):
                tags.append("vel_slow")
            elif vel_val == "Fast":
                tags.append("vel_fast")

        tree.insert("", "end", values=vals, tags=tuple(tags))

    tree.tag_configure("excluded", foreground="#999999")
    tree.tag_configure("backstock", foreground="#cc2222")
    tree.tag_configure("kuntal", foreground="#c0007a")
    tree.tag_configure("vel_slow", foreground="#B8860B")
    tree.tag_configure("vel_fast", foreground="#228B22")


def render_kuntal_tree(
    tree: ttk.Treeview,
    df: pd.DataFrame,
    display_cols: List[str],
) -> None:
    """Populate the Priority! treeview with the user-starred items.

    The Priority! tab shows items the user has flagged with "Toggle Priority!".
    These items are guaranteed to appear on the exported PDF sticker sheet
    regardless of whether the move-up algorithm would include them.

    This is simpler than render_moveup_tree: no color tagging, no emoji prefixes,
    no velocity. The DataFrame passed in (*df*) is already pre-filtered to only
    include rows whose barcodes are in kuntal_priority_barcodes — filtering is
    done by MoveUpGUI._get_kuntal_priority_df() before calling this function.

    Returns early if df is None/empty or any core column is missing.

    Parameters
    ----------
    tree : ttk.Treeview
        The Priority! tab treeview. Cleared before repopulating.
    df : pd.DataFrame
        Pre-filtered DataFrame containing only priority items.
    display_cols : list[str]
        Full column set (COLUMNS_TO_USE + any extras present in data).
    """
    for i in tree.get_children():
        tree.delete(i)
    if df is None or df.empty:
        return

    core_cols = COLUMNS_TO_USE
    core_missing = [c for c in core_cols if c not in df.columns]
    if core_missing:
        return

    idx_type = core_cols.index("Type")
    extra_cols = [c for c in display_cols if c not in core_cols]
    extra_series = {
        c: df[c].reset_index(drop=True)
        for c in extra_cols if c in df.columns
    }

    for row_idx, row in enumerate(
        df[core_cols].itertuples(index=False, name=None)
    ):
        vals = list(row)
        vals[idx_type] = ellipses(str(vals[idx_type]), TYPE_TRUNC_LEN)
        for c in extra_cols:
            vals.append(extra_series[c].iloc[row_idx] if c in extra_series else "")
        tree.insert("", "end", values=vals)


def render_excluded_tree(
    tree: ttk.Treeview,
    df: pd.DataFrame,
    display_cols: List[str],
) -> None:
    """Populate the Excluded / Removed treeview.

    The Excluded tab shows items the user has removed from the Move-Up list
    via "Toggle Remove" or by double-clicking a row. Users can review them here
    and restore items by selecting and clicking "Toggle Remove" again.

    Each row gets a stable iid (f"x_{seq}") so that future references (e.g.
    for restoration) can identify the specific row reliably. The iid prefix
    "x_" avoids collisions with the integer iids used in other treeviews.

    No color tags are applied here — the purpose of this tab is simply to
    show what has been excluded, not to further analyze it.

    Returns early if df is None/empty or any core column is missing.

    Parameters
    ----------
    tree : ttk.Treeview
        The Excluded tab treeview. Cleared before repopulating.
    df : pd.DataFrame
        Pre-filtered DataFrame of only excluded items (from _get_excluded_df()).
    display_cols : list[str]
        Full column set to show.
    """
    for i in tree.get_children():
        tree.delete(i)
    if df is None or df.empty:
        return

    core_cols = COLUMNS_TO_USE
    core_missing = [c for c in core_cols if c not in df.columns]
    if core_missing:
        return

    idx_type = core_cols.index("Type")
    extra_cols = [c for c in display_cols if c not in core_cols]
    extra_series = {
        c: df[c].reset_index(drop=True)
        for c in extra_cols if c in df.columns
    }

    for seq, row in enumerate(
        df[core_cols].itertuples(index=False, name=None)
    ):
        vals = list(row)
        vals[idx_type] = ellipses(str(vals[idx_type]), TYPE_TRUNC_LEN)
        for c in extra_cols:
            vals.append(extra_series[c].iloc[seq] if c in extra_series else "")
        tree.insert("", "end", iid=f"x_{seq}", values=vals)


def render_all_tree(
    tree: ttk.Treeview,
    df: Optional[pd.DataFrame],
    display_cols: List[str],
    search_text: str,
    excluded_barcodes: set,
    kuntal_priority_barcodes: set,
    count_var: StringVar,
) -> None:
    """Populate the All Items treeview, optionally filtered by a search query.

    This tab shows every row in the full mapped inventory (current_df), not just
    move-up candidates. It's useful for looking up a specific product, verifying
    room contents, or finding items that aren't showing up in the Move-Up list.

    Search behavior:
      - Multi-token AND matching: search_text is split on whitespace. Every token
        must appear somewhere in the row (any column) for the row to be shown.
      - Case-insensitive. Example: "blue dream flower" matches rows containing
        all three words across any combination of columns.
      - Empty search shows all rows.
      - count_var is updated to "X of Y" while filtered, or "Y items" unfiltered.

    Color tags:
      - excluded_all (grey): barcode is in excluded_barcodes (excluded from Move-Up)
      - kuntal_all (pink): barcode is in kuntal_priority_barcodes (Priority! starred)
      These are intentionally lighter indicators than in render_moveup_tree —
      the All Items tab is informational, not action-oriented.

    Returns early (clears count, renders nothing) if df is None or empty.

    Parameters
    ----------
    tree : ttk.Treeview
        The All Items tab treeview. Cleared before repopulating.
    df : DataFrame or None
        The full inventory DataFrame (current_df).
    display_cols : list[str]
        Columns to display (full set including any extras).
    search_text : str
        Raw search string from the search Entry widget. May be empty.
    excluded_barcodes : set[str]
        For grey-tagging excluded rows.
    kuntal_priority_barcodes : set[str]
        For pink-tagging priority rows.
    count_var : StringVar
        Tk variable bound to the "Showing X of Y" label above the treeview.
    """
    for i in tree.get_children():
        tree.delete(i)

    if df is None or df.empty:
        count_var.set("")
        return

    core_cols = COLUMNS_TO_USE
    idx_type = core_cols.index("Type")
    idx_bar  = core_cols.index("Package Barcode")
    extra_cols = [c for c in display_cols if c not in core_cols]
    extra_series = {
        c: df[c].reset_index(drop=True)
        for c in extra_cols if c in df.columns
    }

    q = (search_text or "").strip().lower()
    tokens = q.split() if q else []

    shown = 0
    for row_idx, row in enumerate(
        df[core_cols].itertuples(index=False, name=None)
    ):
        vals = list(row)
        vals[idx_type] = ellipses(str(vals[idx_type]), TYPE_TRUNC_LEN)
        for c in extra_cols:
            vals.append(extra_series[c].iloc[row_idx] if c in extra_series else "")

        if tokens:
            haystack = " ".join(str(v).lower() for v in vals)
            if not all(t in haystack for t in tokens):
                continue

        bc = str(vals[idx_bar]).strip()
        tags = []
        if bc in excluded_barcodes:
            tags.append("excluded_all")
        if bc in kuntal_priority_barcodes:
            tags.append("kuntal_all")

        tree.insert("", "end", values=vals, tags=tuple(tags))
        shown += 1

    total = len(df)
    count_var.set(
        f"Showing {shown} of {total}" if tokens else f"{total} items"
    )
    tree.tag_configure("excluded_all", foreground="#999999")
    tree.tag_configure("kuntal_all", foreground="#c0007a")
