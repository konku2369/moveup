"""
Dialog windows and filter utilities for MoveUp.

Each dialog is a standalone function that creates a Toplevel window and
communicates back to the caller via explicit callbacks. This keeps the
main.py clean — dialog logic lives here, not in MoveUpGUI.

DIALOG FUNCTIONS:
=================
  open_map_columns_dialog()  — Manual column mapping override. Shown when
                                automap_columns() fails or the user wants
                                to remap columns manually.

  open_filters_window()      — Room/brand/type filter selection with room
                                alias management. This is how users control
                                which items appear in the Move-Up list.

  open_audit_window()        — Audit PDF export with distributor grouping,
                                sort mode selection, and Accessory Audit
                                one-click button.

  open_manual_add_dialog()   — Search inventory and manually add items to
                                the Priority! list (for items that aren't
                                in the move-up candidates but need stickers).

HELPER CLASSES:
  _FilterList                — Searchable multi-select Listbox with Select All /
                                Clear buttons. Used in the filters dialog for
                                rooms, brands, and types.

HELPER FUNCTIONS:
  get_all_rooms_normalized() — Unique room names after alias normalization
  get_all_brands()           — Unique brand names from the DataFrame
  get_all_types()            — Unique product type names from the DataFrame
  default_candidate_rooms()  — Smart default room selection (Backstock + Incoming,
                                or all non-floor rooms if those don't exist)
"""

import os
from datetime import datetime
from typing import Callable, Dict, List, Optional

import pandas as pd
import tkinter as tk
from tkinter import (
    Toplevel, StringVar, BooleanVar, messagebox, ttk,
    Listbox, MULTIPLE, END,
)

from data_core import (
    COLUMNS_TO_USE,
    AUDIT_OPTIONAL_FIELDS,
    TYPE_TRUNC_LEN,
    SALES_FLOOR_ALIASES,
    automap_columns,
    normalize_rooms,
    detect_metrc_source_column,
)
# pdf_export is lazy-imported inside export_audit_pdfs usage to avoid
# startup crash if reportlab is missing


# ---------------------------------------------------------------------------
# _FilterList — searchable, multi-select Listbox helper
# ---------------------------------------------------------------------------

class _FilterList:
    """
    Searchable multi-select Listbox with Select All / Clear buttons.

    Used in the Filters dialog for the Rooms, Brands, and Types panels.
    Supports multi-token AND search: typing ``"flower blue"`` shows only items
    whose name contains *both* ``"flower"`` and ``"blue"``.

    Selection is preserved across search refreshes — if an item is selected and
    then temporarily hidden by a search term, it remains selected when the
    search is cleared.  This is critical so that brand/type selections don't
    get silently dropped when the user searches, then applies filters.

    Internal state:
    - ``all_items``: the full (unfiltered) list of option strings.
    - ``filtered_idx``: list mapping current listbox positions to indices in
      ``all_items``, updated by every ``refresh()`` call.
    """

    def __init__(self, parent, title: str):
        """
        Build the filter list widget inside *parent*.

        Parameters
        ----------
        parent : tk.Widget
            Parent container (packed inside ``open_filters_window``).
        title : str
            Label text shown as the LabelFrame title.
        """
        self.title = title
        self.all_items: List[str] = []    # full list of options (unfiltered)
        self.filtered_idx: List[int] = [] # maps listbox position → index in all_items

        frm = ttk.Labelframe(parent, text=title, padding=8)
        frm.pack(side="left", fill="both", expand=True, padx=6, pady=6)

        self.search_var = StringVar(value="")
        ttk.Label(frm, text="Search").pack(anchor="w")
        self.ent = ttk.Entry(frm, textvariable=self.search_var)
        self.ent.pack(fill="x", pady=(0, 6))

        inner = ttk.Frame(frm)
        inner.pack(fill="both", expand=True)

        self.lb = Listbox(inner, selectmode=MULTIPLE, height=18, exportselection=False)
        self.lb.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(inner, orient="vertical", command=self.lb.yview)
        sb.pack(side="right", fill="y")
        self.lb.config(yscrollcommand=sb.set)

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(6, 0))
        ttk.Button(btns, text="Select All", command=self.select_all).pack(side="left")
        ttk.Button(btns, text="Clear", command=self.clear_selection).pack(side="left", padx=6)

        self.search_var.trace_add("write", lambda *_: self.refresh())  # trace_add passes (var, index, mode) — *_ discards all three

    def set_items(self, items: List[str]):
        """Replace the full option list and refresh the visible listbox."""
        self.all_items = list(items or [])
        self.refresh()

    def refresh(self):
        """Rebuild the visible listbox from search query, preserving selections."""
        q = (self.search_var.get() or "").strip().lower()
        selected_vals = set(self.get_selected_values())
        self.lb.delete(0, END)

        if not q:
            self.filtered_idx = list(range(len(self.all_items)))
        else:
            # Multi-token AND search: "flower blue" matches items containing BOTH words
            tokens = q.split()

            def match(i: int) -> bool:
                s = self.all_items[i].lower()
                return all(t in s for t in tokens)

            self.filtered_idx = [i for i in range(len(self.all_items)) if match(i)]

        for i in self.filtered_idx:
            self.lb.insert(END, self.all_items[i])

        # Re-select previously selected items that survived the filter
        for pos, i in enumerate(self.filtered_idx):
            if self.all_items[i] in selected_vals:
                self.lb.selection_set(pos)

    def select_all(self):
        """Select every visible item in the listbox."""
        self.lb.selection_set(0, END)

    def clear_selection(self):
        """Deselect every item in the listbox."""
        self.lb.selection_clear(0, END)

    def set_selected_values(self, values: List[str]):
        """
        Programmatically select items by value string.

        Clears the search bar (so all items are visible), then selects each
        item in *values* by matching against ``all_items``.  Items in *values*
        that are not in ``all_items`` are silently ignored.

        Parameters
        ----------
        values : list[str]
            Item strings to select.  ``None`` or empty → clear all selections.
        """
        values_set = set(values or [])
        self.lb.selection_clear(0, END)
        self.search_var.set("")
        self.refresh()
        for pos, i in enumerate(self.filtered_idx):
            if self.all_items[i] in values_set:
                self.lb.selection_set(pos)

    def get_selected_values(self) -> List[str]:
        """Return the actual item strings for all selected listbox positions."""
        sel = list(self.lb.curselection())
        out = []
        for pos in sel:
            if pos < 0 or pos >= len(self.filtered_idx):
                continue
            # Convert listbox position → original all_items index → item string
            i = self.filtered_idx[pos]
            out.append(self.all_items[i])
        return out


# ---------------------------------------------------------------------------
# Filter-helper utilities (formerly on MoveUpGUI)
# ---------------------------------------------------------------------------

def get_all_rooms_normalized(
    df: pd.DataFrame, room_alias_map: Dict[str, str]
) -> List[str]:
    """
    Return sorted unique room names after applying the user's alias map.

    Passes *df* through ``normalize_rooms()`` (which substitutes room names per
    *room_alias_map*) before collecting unique values, so the returned list
    reflects canonical names (e.g. ``"Vault"`` instead of ``"Vault 1"`` /
    ``"Vault 2"``).

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame with a ``"Room"`` column.
    room_alias_map : dict[str, str]
        Maps raw room names (lowercase) to canonical names.

    Returns
    -------
    list[str]
        Sorted unique room strings.  Empty list if *df* has no ``"Room"``
        column or is empty.
    """
    if df is None or df.empty or "Room" not in df.columns:
        return []
    df_norm = normalize_rooms(df, room_alias_map)
    return sorted(
        set(str(x).strip() for x in df_norm["Room"].dropna().astype(str).tolist())
    )


def get_all_brands(df: pd.DataFrame) -> List[str]:
    """
    Return sorted unique brand names from the ``"Brand"`` column of *df*.

    Strips whitespace and filters out blank/null values.  Returns ``[]`` if
    *df* is None, empty, or lacks a ``"Brand"`` column.
    """
    if df is None or df.empty or "Brand" not in df.columns:
        return []
    vals = sorted(
        set(str(x).strip() for x in df["Brand"].dropna().astype(str).tolist())
    )
    return [v for v in vals if v]


def get_all_types(df: pd.DataFrame) -> List[str]:
    """
    Return sorted unique product type names from the ``"Type"`` column of *df*.

    Strips whitespace and filters out blank/null values.  Returns ``[]`` if
    *df* is None, empty, or lacks a ``"Type"`` column.
    """
    if df is None or df.empty or "Type" not in df.columns:
        return []
    vals = sorted(
        set(str(x).strip() for x in df["Type"].dropna().astype(str).tolist())
    )
    return [v for v in vals if v]


def default_candidate_rooms(
    df: pd.DataFrame, room_alias_map: Dict[str, str]
) -> List[str]:
    """
    Determine sensible default candidate rooms for the move-up filter.

    Two-path logic:
    1. **Preferred**: if *both* ``"Incoming Deliveries"`` and ``"Backstock"``
       appear in the alias-normalised room list, return exactly those two.
       This covers the typical METRC store layout.
    2. **Fallback**: return all rooms that are *not* in ``SALES_FLOOR_ALIASES``
       and not literally ``"sales floor"`` (case-insensitive).  This handles
       stores with non-standard room names.  If the fallback produces an empty
       list (everything was a sales-floor alias), all rooms are returned.

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame with a ``"Room"`` column.
    room_alias_map : dict[str, str]
        User-defined room alias map (forwarded to ``get_all_rooms_normalized``).

    Returns
    -------
    list[str]
        Sorted list of room name strings to use as default candidates.  Empty
        if *df* has no room data.
    """
    rooms = get_all_rooms_normalized(df, room_alias_map)
    if not rooms:
        return []
    room_lookup = {r.strip().lower(): r for r in rooms}
    desired_keys = ["incoming deliveries", "backstock"]
    if all(k in room_lookup for k in desired_keys):
        return [room_lookup[k] for k in desired_keys]

    out = []
    for r in rooms:
        r_l = r.strip().lower()
        if r_l not in SALES_FLOOR_ALIASES and r_l != "sales floor":
            out.append(r)
    return out or rooms


# ---------------------------------------------------------------------------
# Dialog: Map Columns
# ---------------------------------------------------------------------------

def open_map_columns_dialog(
    parent,
    raw_df: pd.DataFrame,
    on_apply: Callable[[pd.DataFrame, dict], None],
    force: bool = False,
) -> None:
    """
    Open the manual column mapping dialog.

    Shown when ``automap_columns()`` fails (e.g. non-standard column names) or
    when the user manually opens the column mapping from the toolbar.  Allows
    the user to select which source column maps to each required field in
    ``COLUMNS_TO_USE`` (Type, Brand, Product Name, Package Barcode, Room,
    Qty On Hand) plus the METRC source column.

    Auto-detection runs first: ``automap_columns()`` pre-fills the dropdowns
    with its best guesses.  If it raises an exception, a red error label is
    shown in the dialog explaining why auto-detection failed, and the dropdowns
    default to empty so the user must select manually.

    When the user clicks Apply, the dialog validates that all required fields
    are mapped, performs the column rename, and calls
    *on_apply(mapped_df, mapping_dict)* with the remapped DataFrame and the
    ``{source_col: target_col}`` dict.

    Parameters
    ----------
    parent : tk.Widget
        Parent window for the Toplevel.
    raw_df : pd.DataFrame
        The raw loaded DataFrame (before any column renaming).
    on_apply : Callable[[pd.DataFrame, dict], None]
        Callback invoked with the mapped DataFrame and rename dict on Apply.
    force : bool
        Unused parameter kept for API compatibility.
    """
    src_cols = list(raw_df.columns)
    metrc_src_detected = detect_metrc_source_column(raw_df)

    auto_map = {}
    _automap_error = ""
    try:
        _auto_df, auto_map = automap_columns(raw_df)
    except Exception as e:
        auto_map = {}
        _automap_error = str(e)

    win = Toplevel(parent)
    win.title("Map Columns (Manual Override)")
    win.geometry("760x560")
    ttk.Label(
        win, text="Choose which source column maps to each required field."
    ).pack(anchor="w", padx=10, pady=10)
    if _automap_error:
        ttk.Label(
            win,
            text=f"Auto-detection failed — please map columns manually.\n({_automap_error[:140]})",
            foreground="#aa0000",
            wraplength=720,
        ).pack(anchor="w", padx=10, pady=(0, 8))

    frame = ttk.Frame(win)
    frame.pack(fill="both", expand=True, padx=10)

    combos: dict = {}

    ttk.Label(frame, text="METRC Source Column (required):").grid(
        row=0, column=0, sticky="e", pady=6
    )
    metrc_var = StringVar(value=metrc_src_detected or "")
    metrc_cb = ttk.Combobox(
        frame, textvariable=metrc_var, values=src_cols, width=52, state="readonly"
    )
    metrc_cb.grid(row=0, column=1, sticky="w", pady=6)
    ttk.Label(
        frame, text="⚠ Changing this resets the barcode key", foreground="#aa6600"
    ).grid(row=0, column=2, sticky="w", padx=(8, 0))

    def rebuild_non_metrc_dropdown_values():
        chosen = metrc_var.get().strip()
        non_metrc = [c for c in src_cols if c != chosen]
        for target, var in combos.items():
            cb = var["_cb"]
            if target == "Package Barcode":
                var["_var"].set(chosen)
                cb.configure(values=[chosen])
            else:
                cb.configure(values=non_metrc)
                if var["_var"].get().strip() == chosen:
                    var["_var"].set("")

    _metrc_prev = [metrc_var.get()]
    _metrc_warned = [False]

    def _on_metrc_changing(event):
        new_val = metrc_var.get()
        if new_val == _metrc_prev[0]:
            return
        if not _metrc_warned[0]:
            ok = messagebox.askokcancel(
                "Change METRC Column",
                "The METRC source column is used as the Package Barcode key.\n\n"
                "Changing it will re-map all barcodes and may clear your current\n"
                "excluded / Priority! lists if the barcodes no longer match.\n\n"
                "Are you sure you want to change it?",
                parent=win,
            )
            if not ok:
                metrc_var.set(_metrc_prev[0])
                metrc_cb.set(_metrc_prev[0])
                return
            _metrc_warned[0] = True
        _metrc_prev[0] = new_val
        rebuild_non_metrc_dropdown_values()

    metrc_cb.bind("<<ComboboxSelected>>", _on_metrc_changing)

    row_offset = 1
    for i, target in enumerate(COLUMNS_TO_USE):
        ttk.Label(frame, text=target + ":").grid(
            row=i + row_offset, column=0, sticky="e", pady=4
        )
        var = StringVar(value="")

        if target == "Package Barcode":
            var.set(metrc_var.get().strip())
            cb = ttk.Combobox(
                frame,
                textvariable=var,
                values=[metrc_var.get().strip()] if metrc_var.get().strip() else src_cols,
                width=52,
                state="disabled",
            )
        else:
            cb = ttk.Combobox(
                frame, textvariable=var, values=src_cols, width=52, state="readonly"
            )
            pre = next((src for src, dst in auto_map.items() if dst == target), None)
            if pre:
                var.set(pre)

        cb.grid(row=i + row_offset, column=1, sticky="w", pady=4)
        combos[target] = {"_var": var, "_cb": cb}

    rebuild_non_metrc_dropdown_values()

    ttk.Separator(frame, orient="horizontal").grid(
        row=len(COLUMNS_TO_USE) + row_offset,
        column=0, columnspan=2, sticky="ew", pady=10,
    )

    opt_start = len(COLUMNS_TO_USE) + row_offset + 1
    ttk.Label(
        frame, text="Optional (used by Audit PDFs):", font=("Helvetica", 9, "bold")
    ).grid(row=opt_start, column=0, sticky="e", pady=4)

    opt_vars: dict = {}
    for j, opt in enumerate(AUDIT_OPTIONAL_FIELDS):
        ttk.Label(frame, text=f"{opt} (optional):").grid(
            row=opt_start + 1 + j, column=0, sticky="e", pady=4
        )
        v = StringVar(value="")
        pre = next((src for src, dst in auto_map.items() if dst == opt), None)
        if pre:
            v.set(pre)
        cb = ttk.Combobox(
            frame, textvariable=v, values=[""] + src_cols, width=52, state="readonly"
        )
        cb.grid(row=opt_start + 1 + j, column=1, sticky="w", pady=4)
        opt_vars[opt] = v

    btns = ttk.Frame(win)
    btns.pack(fill="x", pady=10)

    def _apply_mapping():
        try:
            chosen_metrc = metrc_var.get().strip()
            if not chosen_metrc:
                messagebox.showerror(
                    "Missing", "Please choose the METRC source column (required)."
                )
                return

            mapping: dict = {}
            used_sources: set = set()

            mapping[chosen_metrc] = "Package Barcode"
            used_sources.add(chosen_metrc)

            for target in COLUMNS_TO_USE:
                if target == "Package Barcode":
                    continue
                src = combos[target]["_var"].get().strip()
                if not src:
                    messagebox.showerror(
                        "Missing", f"Please choose a source for '{target}'."
                    )
                    return
                if src in used_sources:
                    messagebox.showerror(
                        "Duplicate Source",
                        f"The source column '{src}' is used more than once.",
                    )
                    return
                used_sources.add(src)
                mapping[src] = target

            for opt in AUDIT_OPTIONAL_FIELDS:
                src_opt = opt_vars.get(opt).get().strip()
                if src_opt:
                    if src_opt in used_sources:
                        messagebox.showerror(
                            "Duplicate Source",
                            f"The source column '{src_opt}' is already used.",
                        )
                        return
                    used_sources.add(src_opt)
                    mapping[src_opt] = opt

            df = raw_df.rename(columns=mapping)

            missing = [c for c in COLUMNS_TO_USE if c not in df.columns]
            if missing:
                raise ValueError(
                    "After mapping, still missing: " + ", ".join(missing)
                )

            df["Package Barcode"] = df["Package Barcode"].astype("string").fillna("")
            df["Qty On Hand"] = (
                pd.to_numeric(df["Qty On Hand"], errors="coerce").fillna(0).astype(int)
            )
            for col in ["Product Name", "Brand", "Type", "Room"]:
                df[col] = df[col].astype(str)

            if "Distributor" in df.columns:
                df["Distributor"] = df["Distributor"].astype(str).fillna("").str.strip()
            if "Store" in df.columns:
                df["Store"] = df["Store"].astype(str).fillna("").str.strip()
            if "Size" in df.columns:
                df["Size"] = df["Size"].astype(str).fillna("").str.strip()

            on_apply(df, mapping)
            win.destroy()
        except Exception as e:
            messagebox.showerror("Mapping Error", str(e))

    ttk.Button(btns, text="Apply", command=_apply_mapping).pack(side="left", padx=6)
    ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="left", padx=6)


# ---------------------------------------------------------------------------
# Dialog: Filters Window
# ---------------------------------------------------------------------------

def open_filters_window(
    parent,
    current_df: pd.DataFrame,
    room_alias_map: Dict[str, str],
    selected_rooms: List[str],
    selected_brands: List[str],
    selected_types: List[str],
    on_apply: Callable[[List[str], List[str], List[str]], None],
    on_alias_changed: Callable[[], None],
    on_close: Callable[[], None],
) -> Toplevel:
    """
    Filters dialog with room aliases and room/brand/type selectors.

    *room_alias_map* is mutated in place (same as original behavior).
    *on_alias_changed()* is called after each alias add/remove so caller can
    persist config.
    *on_apply(rooms, brands, types)* is called when the user clicks Apply.
    *on_close()* is called when the window is closed (so caller can clear its
    reference).

    Returns the Toplevel window so the caller can store it for singleton checks.
    """
    win = Toplevel(parent)
    win.title("Filters")
    win.geometry("1180x860")
    win.transient(parent)
    win.grab_set()

    df = current_df

    top = ttk.Frame(win, padding=10)
    top.pack(fill="x")
    ttk.Label(
        top,
        text="Room Aliases (optional) — normalize messy room names into a clean canonical name.",
    ).pack(anchor="w")

    alias_row = ttk.Frame(top)
    alias_row.pack(fill="x", pady=6)

    alias_from = StringVar(value="")
    alias_to = StringVar(value="")
    ttk.Label(alias_row, text="From").pack(side="left")
    ttk.Entry(alias_row, textvariable=alias_from, width=22).pack(side="left", padx=6)
    ttk.Label(alias_row, text="To").pack(side="left")
    ttk.Entry(alias_row, textvariable=alias_to, width=22).pack(side="left", padx=6)

    alias_tree = ttk.Treeview(top, columns=("from", "to"), show="headings", height=4)
    alias_tree.heading("from", text="From")
    alias_tree.heading("to", text="To")
    alias_tree.column("from", width=260, anchor="w")
    alias_tree.column("to", width=260, anchor="w")
    alias_tree.pack(fill="x", pady=(6, 0))

    def refresh_alias_tree():
        for i in alias_tree.get_children():
            alias_tree.delete(i)
        for k, v in sorted(room_alias_map.items(), key=lambda kv: kv[0].lower()):
            alias_tree.insert("", "end", values=(k, v))

    def add_alias():
        f = (alias_from.get() or "").strip()
        t = (alias_to.get() or "").strip()
        if not f or not t:
            messagebox.showinfo("Alias", "Enter both From and To.")
            return
        if f.casefold() == t.casefold():
            messagebox.showinfo("Alias", "From and To cannot be the same room.")
            return
        if t in room_alias_map:
            messagebox.showwarning(
                "Alias",
                f'"{t}" is already aliased to "{room_alias_map[t]}". '
                f"This may cause unexpected results.",
            )
        room_alias_map[f] = t
        alias_from.set("")
        alias_to.set("")
        refresh_alias_tree()
        rooms_list.set_items(get_all_rooms_normalized(df, room_alias_map))
        on_alias_changed()

    def remove_alias():
        sel = alias_tree.selection()
        if not sel:
            return
        for iid in sel:
            vals = alias_tree.item(iid, "values")
            if vals and vals[0] in room_alias_map:
                del room_alias_map[vals[0]]
        refresh_alias_tree()
        rooms_list.set_items(get_all_rooms_normalized(df, room_alias_map))
        on_alias_changed()

    ttk.Button(alias_row, text="Add/Update", command=add_alias).pack(side="left", padx=6)
    ttk.Button(alias_row, text="Remove Selected", command=remove_alias).pack(
        side="left", padx=6
    )
    refresh_alias_tree()

    mid = ttk.Frame(win, padding=10)
    mid.pack(fill="both", expand=True)

    rooms_list = _FilterList(mid, "Rooms (Move-Up source rooms)")
    brands_list = _FilterList(mid, "Brands (empty = ALL)")
    types_list = _FilterList(mid, "Types (empty = ALL)")

    rooms = get_all_rooms_normalized(df, room_alias_map)
    brands = get_all_brands(df)
    types_ = get_all_types(df)

    rooms_list.set_items(rooms)
    brands_list.set_items(brands)
    types_list.set_items(types_)

    if selected_rooms:
        rooms_list.set_selected_values(selected_rooms)
    else:
        rooms_list.set_selected_values(default_candidate_rooms(df, room_alias_map))

    brands_list.set_selected_values(selected_brands)
    types_list.set_selected_values(selected_types)

    bot = ttk.Frame(win, padding=10)
    bot.pack(fill="x")

    def apply_filters():
        sel_rooms = rooms_list.get_selected_values()
        if not sel_rooms:
            sel_rooms = default_candidate_rooms(df, room_alias_map)
        on_apply(sel_rooms, brands_list.get_selected_values(), types_list.get_selected_values())
        win.destroy()

    def reset_defaults():
        rooms_list.set_selected_values(default_candidate_rooms(df, room_alias_map))
        brands_list.clear_selection()
        types_list.clear_selection()

    def _on_close():
        on_close()
        win.destroy()

    ttk.Button(bot, text="Apply", command=apply_filters).pack(side="left")
    ttk.Button(bot, text="Reset Defaults", command=reset_defaults).pack(
        side="left", padx=8
    )
    ttk.Button(bot, text="Close", command=_on_close).pack(side="left", padx=8)

    win.protocol("WM_DELETE_WINDOW", _on_close)
    return win


# ---------------------------------------------------------------------------
# Dialog: Audit Window
# ---------------------------------------------------------------------------

def open_audit_window(
    parent,
    current_df: pd.DataFrame,
    room_alias_map: Dict[str, str],
    export_run_dir: str,
    printer_bw: bool,
    auto_open: bool,
    on_status: Callable[[str], None],
    on_success: Optional[Callable[[str], None]] = None,
    on_error: Optional[Callable[[str], None]] = None,
) -> None:
    """
    Open the Audit PDF export dialog.

    Provides controls for generating two audit PDFs:
    - **Master audit**: full inventory with quantities — for reconciliation.
    - **Blank audit**: same layout with an empty qty column — for physical
      counting rounds.

    Export options:
    - **Group by**: Distributor, Brand, or Type (each group gets a page break).
    - **Sort by**: Distributor → Brand → Product, or Type → Brand → Product.
    - **Include optional fields**: Expiration Date, Received Date, Wholesale Cost,
      Room (configurable checkboxes).
    - **Accessory Audit**: one-click button that generates a separate accessory-only
      PDF (items whose Type contains ``"Accessory"``).

    Parameters
    ----------
    parent : tk.Widget
        Parent window.
    current_df : pd.DataFrame
        The full mapped inventory DataFrame.
    room_alias_map : dict[str, str]
        Used to normalise room names before grouping.
    export_run_dir : str
        Output directory for the generated PDFs.
    printer_bw : bool
        If ``True``, use the B/W palette; otherwise use the kawaii palette.
    auto_open : bool
        If ``True``, open the PDF in the system default viewer after export.
    on_status : Callable[[str], None]
        Status bar update callback.
    on_success : Callable[[str], None] | None
        Optional callback invoked with the output path on successful export.
    on_error : Callable[[str], None] | None
        Optional callback invoked with the error message on export failure.
    """
    df = current_df.copy()
    has_dist = "Distributor" in df.columns
    if has_dist:
        blanks = (df["Distributor"].astype(str).fillna("").str.strip() == "").sum()
        on_status(f"Audit: Distributor column present. Blank rows: {blanks}/{len(df)}")
    else:
        on_status("Audit: Distributor column NOT present (will show Unknown Distributor).")

    if "Distributor" not in df.columns:
        df["Distributor"] = ""
    if "Store" not in df.columns:
        df["Store"] = ""
    if "Size" not in df.columns:
        df["Size"] = ""

    df["Distributor"] = df["Distributor"].astype(str).fillna("").str.strip()
    df.loc[df["Distributor"] == "", "Distributor"] = "Unknown Distributor"

    df_norm = normalize_rooms(df, room_alias_map)

    types_ = sorted(set(df_norm["Type"].dropna().astype(str).str.strip().tolist()))
    types_ = [t for t in types_ if t]

    brands = sorted(set(df_norm["Brand"].dropna().astype(str).str.strip().tolist()))
    brands = [b for b in brands if b]

    rooms = get_all_rooms_normalized(df_norm, room_alias_map)

    dists = sorted(set(df_norm["Distributor"].dropna().astype(str).str.strip().tolist()))
    dists = [d for d in dists if d]

    dist_to_brands: Dict[str, set] = {}
    try:
        sub = df_norm[["Distributor", "Brand"]].copy()
        sub["Distributor"] = (
            sub["Distributor"].astype(str).str.strip().replace({"": "Unknown Distributor"})
        )
        sub["Brand"] = sub["Brand"].astype(str).str.strip()
        sub = sub[(sub["Distributor"] != "") & (sub["Brand"] != "")]
        for _, r in sub.drop_duplicates().iterrows():
            dist_to_brands.setdefault(r["Distributor"], set()).add(r["Brand"])
    except Exception:
        dist_to_brands = {}

    win = Toplevel(parent)
    win.title("Audit PDF Export (Distributor Groups)")
    win.geometry("1320x820")
    win.transient(parent)
    win.grab_set()

    pad = 10
    top = ttk.Frame(win, padding=pad)
    top.pack(fill="x")
    ttk.Label(
        top,
        text="Select filters for the Audit PDFs (Master + Blank). Page breaks follow the Sort Mode.",
        font=("Helvetica", 10, "bold"),
    ).pack(anchor="w")

    defaults = ttk.LabelFrame(
        win, text="Defaults (used if Store/Room missing)", padding=pad
    )
    defaults.pack(fill="x", padx=pad, pady=(0, pad))

    default_store_var = StringVar(value="Store")
    default_room_var = StringVar(value="Sales Floor")

    ttk.Label(defaults, text="Default Store:").pack(side="left")
    ttk.Entry(defaults, textvariable=default_store_var, width=26).pack(
        side="left", padx=(6, 18)
    )
    ttk.Label(defaults, text="Default Room:").pack(side="left")
    ttk.Entry(defaults, textvariable=default_room_var, width=26).pack(
        side="left", padx=6
    )

    title_row = ttk.Frame(win, padding=(pad, 0))
    title_row.pack(fill="x")
    title_var = StringVar(
        value=f"Inventory Audit \u2014 {datetime.now().strftime('%m-%d-%Y')}"
    )
    ttk.Label(title_row, text="Title").pack(side="left")
    ttk.Entry(title_row, textvariable=title_var).pack(
        side="left", fill="x", expand=True, padx=8
    )

    mid = ttk.Frame(win, padding=pad)
    mid.pack(fill="both", expand=True)

    def make_listbox(col_parent, title, items):
        frm = ttk.Labelframe(col_parent, text=title, padding=8)
        frm.pack(side="left", fill="both", expand=True, padx=6, pady=6)

        lb = tk.Listbox(frm, selectmode=tk.EXTENDED, exportselection=False, height=18)
        lb.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(frm, orient="vertical", command=lb.yview)
        sb.pack(side="right", fill="y")
        lb.config(yscrollcommand=sb.set)

        for it in items:
            lb.insert(tk.END, it)

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(6, 0))

        def sel_all():
            lb.select_set(0, tk.END)

        def sel_none():
            lb.select_clear(0, tk.END)

        ttk.Button(btns, text="All", command=sel_all).pack(side="left")
        ttk.Button(btns, text="None", command=sel_none).pack(side="left", padx=6)

        return lb, sel_all, sel_none

    lb_types, types_all, _ = make_listbox(mid, "Types (Category)", types_)
    lb_brands, brands_all, _brands_none = make_listbox(mid, "Brands", brands)
    lb_rooms, rooms_all, rooms_none = make_listbox(mid, "Rooms", rooms)
    lb_dists, dists_all, _ = make_listbox(mid, "Distributors", dists)

    # Select all types EXCEPT accessories (rarely audited with the rest)
    for i, t in enumerate(types_):
        if "accessor" not in str(t).lower():
            lb_types.select_set(i)
    brands_all()
    dists_all()

    if rooms:
        sf_idx = None
        for i, r in enumerate(rooms):
            if str(r).strip().lower() == "sales floor":
                sf_idx = i
                break
        rooms_none()
        if sf_idx is not None:
            lb_rooms.select_set(sf_idx)
        else:
            rooms_all()

    def selected_values(lb: tk.Listbox):
        return [lb.get(i) for i in lb.curselection()]

    def select_brands_for_selected_distributors(_event=None):
        sel_d = selected_values(lb_dists)
        if not sel_d:
            return
        union = set()
        for d in sel_d:
            union |= set(dist_to_brands.get(d, set()))
        if not union:
            return
        lb_brands.select_clear(0, tk.END)
        for i in range(lb_brands.size()):
            b = lb_brands.get(i)
            if b in union:
                lb_brands.select_set(i)

    lb_dists.bind("<<ListboxSelect>>", select_brands_for_selected_distributors)

    sort_mode_var = StringVar(value="distributor_type_size_product")

    frm_sort = ttk.LabelFrame(
        win, text="Sort Mode (controls page breaks)", padding=pad
    )
    frm_sort.pack(fill="x", padx=pad, pady=(0, pad))

    ttk.Radiobutton(
        frm_sort,
        text="Distributor → Type → Size → Product (page break by Distributor)",
        variable=sort_mode_var,
        value="distributor_type_size_product",
    ).pack(anchor="w")

    ttk.Radiobutton(
        frm_sort,
        text="Brand → Type → Product (page break by Brand)",
        variable=sort_mode_var,
        value="brand_type_product",
    ).pack(anchor="w")

    ttk.Radiobutton(
        frm_sort,
        text="Type → Brand → Product (page break by Type)",
        variable=sort_mode_var,
        value="type_brand_product",
    ).pack(anchor="w")

    bot = ttk.Frame(win, padding=pad)
    bot.pack(fill="x")

    use_barcode_var = tk.BooleanVar(value=False)
    has_barcode_col = "Barcode" in df_norm.columns

    def accessory_audit():
        lb_types.select_clear(0, tk.END)
        for i, t in enumerate(types_):
            if "accessor" in str(t).lower():
                lb_types.select_set(i)
        rooms_all()
        brands_all()
        dists_all()
        if has_barcode_col:
            use_barcode_var.set(True)
        title_var.set(
            f"Accessory Inventory Audit \u2014 {datetime.now().strftime('%m-%d-%Y')}"
        )
        sort_mode_var.set("type_brand_product")
        export_now()

    def export_now():
        sel_types = selected_values(lb_types)
        sel_brands = selected_values(lb_brands)
        sel_rooms = selected_values(lb_rooms)
        sel_dists = selected_values(lb_dists)

        if not sel_types:
            messagebox.showerror("Audit PDFs", "Pick at least one Type.")
            return
        if not sel_brands:
            messagebox.showerror("Audit PDFs", "Pick at least one Brand.")
            return
        if not sel_rooms:
            messagebox.showerror("Audit PDFs", "Pick at least one Room.")
            return
        if not sel_dists:
            messagebox.showerror("Audit PDFs", "Pick at least one Distributor.")
            return

        use = df_norm[
            df_norm["Type"].astype(str).isin(sel_types)
            & df_norm["Brand"].astype(str).isin(sel_brands)
            & df_norm["Room"].astype(str).isin(sel_rooms)
            & df_norm["Distributor"].astype(str).isin(sel_dists)
        ].copy()

        if use.empty:
            messagebox.showwarning("Audit PDFs", "Nothing matches your selections.")
            return

        try:
            from pdf_export import export_audit_pdfs
            master_path, blank_path = export_audit_pdfs(
                df=use,
                base_dir=export_run_dir,
                title_text=title_var.get().strip() or "Inventory Audit",
                sort_mode=sort_mode_var.get(),
                kawaii_pdf=True,
                printer_bw=printer_bw,
                auto_open=auto_open,
                default_store=default_store_var.get().strip() or "Store",
                default_room=default_room_var.get().strip() or "Sales Floor",
                type_trunc_len=TYPE_TRUNC_LEN,
                barcode_col="Barcode" if use_barcode_var.get() else None,
            )

            on_status(
                f"Audit PDFs saved: {os.path.basename(master_path)} + "
                f"{os.path.basename(blank_path)}"
            )
            if on_success:
                on_success("Audit PDFs ✅")
            win.destroy()
        except Exception as e:
            messagebox.showerror("Audit PDFs", str(e))
            if on_error:
                on_error("Audit failed 💥")

    ttk.Button(bot, text="Export Audit PDFs", command=export_now).pack(side="left")
    ttk.Button(bot, text="Accessory Audit", command=accessory_audit).pack(
        side="left", padx=(8, 0)
    )
    if has_barcode_col:
        ttk.Checkbutton(
            bot, text='Use "Barcode" column', variable=use_barcode_var
        ).pack(side="left", padx=(12, 0))
    ttk.Button(bot, text="Close", command=win.destroy).pack(side="right")


# ---------------------------------------------------------------------------
# Dialog: Manual Add to Priority!
# ---------------------------------------------------------------------------

def open_manual_add_dialog(
    parent,
    current_df: pd.DataFrame,
    kuntal_priority_barcodes: set,
    on_apply: Callable[[set], None],
) -> None:
    """
    Open the "Manual Add" dialog to add items directly to the Priority list.

    Displays the full inventory in a searchable treeview.  The user can type
    in the search bar to filter rows, then select one or more items and click
    "Add Selected to Priority!".

    Only items whose barcode is *not* already in *kuntal_priority_barcodes* are
    added; existing priority items are silently skipped.  The callback receives
    only the newly added barcodes so the caller can trigger a single recompute
    after the dialog closes.

    Useful for adding items to the sticker sheet that are on the sales floor
    (already restocked) but still need a label printed.

    Parameters
    ----------
    parent : tk.Widget
        Parent window for the Toplevel.
    current_df : pd.DataFrame
        The full mapped inventory DataFrame.  All of ``COLUMNS_TO_USE`` must
        be present; a messagebox error is shown and the dialog aborts if any
        are missing.
    kuntal_priority_barcodes : set
        Current set of priority barcode strings (used to skip duplicates).
    on_apply : Callable[[set], None]
        Called with the set of newly added barcode strings when the user
        clicks "Add Selected to Priority!".
    """
    df = current_df.copy()
    missing = [c for c in COLUMNS_TO_USE if c not in df.columns]
    if missing:
        messagebox.showerror(
            "Manual Add", "Missing required columns: " + ", ".join(missing)
        )
        return

    win = Toplevel(parent)
    win.title("Manual Add to Priority!")
    win.geometry("920x620")
    win.transient(parent)
    win.grab_set()

    ttk.Label(
        win, text="Search inventory. Select one or more, then Add."
    ).pack(anchor="w", padx=10, pady=(10, 6))
    search_var = StringVar(value="")
    ent = ttk.Entry(win, textvariable=search_var)
    ent.pack(fill="x", padx=10)

    frame = ttk.Frame(win)
    frame.pack(fill="both", expand=True, padx=10, pady=10)

    lb = Listbox(frame, selectmode=MULTIPLE, height=22, exportselection=False)
    lb.pack(side="left", fill="both", expand=True)

    sb = ttk.Scrollbar(frame, orient="vertical", command=lb.yview)
    sb.pack(side="right", fill="y")
    lb.config(yscrollcommand=sb.set)

    rows = df[COLUMNS_TO_USE].copy()
    rows["__bc"] = rows["Package Barcode"].astype(str).fillna("").str.strip()
    rows["__disp"] = rows.apply(
        lambda r: (
            f"{r['Brand']} | {r['Product Name']} | {r['Room']} "
            f"| Qty:{r['Qty On Hand']} | …{str(r['Package Barcode'])[-6:]}"
        ),
        axis=1,
    )
    rows = rows.sort_values(by=["Brand", "Product Name"], kind="stable").reset_index(
        drop=True
    )

    filtered_idx = list(range(len(rows)))

    def refresh_list(*_):
        nonlocal filtered_idx
        q = (search_var.get() or "").strip().lower()
        lb.delete(0, END)
        if not q:
            filtered_idx = list(range(len(rows)))
        else:
            tokens = q.split()

            def match(i):
                s = str(rows.loc[i, "__disp"]).lower() + " " + str(rows.loc[i, "__bc"]).lower()
                return all(t in s for t in tokens)

            filtered_idx = [i for i in range(len(rows)) if match(i)]

        for i in filtered_idx[:5000]:
            lb.insert(END, rows.loc[i, "__disp"])

    def do_add():
        sel = list(lb.curselection())
        if not sel:
            messagebox.showinfo("Manual Add", "Select at least one item.")
            return
        added: set = set()
        for pos in sel:
            i = filtered_idx[pos]
            bc = str(rows.loc[i, "__bc"]).strip()
            if bc and bc not in kuntal_priority_barcodes:
                added.add(bc)
        on_apply(added)
        win.destroy()

    btns = ttk.Frame(win)
    btns.pack(fill="x", padx=10, pady=(0, 10))
    ttk.Button(btns, text="Add Selected", command=do_add).pack(side="left")
    ttk.Button(btns, text="Close", command=win.destroy).pack(side="left", padx=8)

    search_var.trace_add("write", refresh_list)
    refresh_list()
    ent.focus_set()
