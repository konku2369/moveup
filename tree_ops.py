"""
Treeview rendering and column management for MoveUp.

All functions are pure operations on ttk.Treeview widgets + DataFrames.
No dependency on MoveUpGUI — each function receives its inputs explicitly.
"""

from typing import Dict, List, Optional

import pandas as pd
from tkinter import ttk, StringVar

from data_core import COLUMNS_TO_USE, TYPE_TRUNC_LEN, ellipses


# ---------------------------------------------------------------------------
# Display-column helper
# ---------------------------------------------------------------------------

def get_display_cols(
    active_columns: List[str],
    df: Optional[pd.DataFrame],
    extra_columns: List[str],
) -> List[str]:
    """
    Returns columns to show in treeviews: *active_columns* plus any
    *extra_columns* that are present in *df*.
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
    """Apply standard column widths and wire up click-to-sort on every heading."""
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
    """Sort a Treeview in-place by *col*, toggling asc/desc.  Updates heading arrows."""
    state = sort_state.setdefault(tree_id, {})
    ascending = not state.get(col, True)   # first click → ascending
    state[col] = ascending

    rows = [(tree.set(iid, col), iid) for iid in tree.get_children("")]

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
    """Reconfigure all four treeviews to reflect current display columns."""
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
    """Populate the main Move-Up treeview."""
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

        tree.insert("", "end", values=vals, tags=tuple(tags))

    tree.tag_configure("excluded", foreground="#999999")
    tree.tag_configure("backstock", foreground="#cc2222")
    tree.tag_configure("kuntal", foreground="#c0007a")


def render_kuntal_tree(
    tree: ttk.Treeview,
    df: pd.DataFrame,
    display_cols: List[str],
) -> None:
    """Populate the Priority! treeview."""
    for i in tree.get_children():
        tree.delete(i)
    if df is None or df.empty:
        return

    core_cols = COLUMNS_TO_USE
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
    """Populate the Excluded / Removed treeview."""
    for i in tree.get_children():
        tree.delete(i)
    if df is None or df.empty:
        return

    core_cols = COLUMNS_TO_USE
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
    """Populate the All Items treeview with live search filtering."""
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
