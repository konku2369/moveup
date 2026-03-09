import os
import sys
import subprocess
from datetime import datetime
from typing import Dict, List, Optional

import pandas as pd

# GUI
from tkinter import (
    Tk, Toplevel, StringVar, IntVar, BooleanVar, filedialog, messagebox,
    ttk
)
import tkinter as tk

# PDF exports live in pdf_export.py
from pdf_export import export_moveup_pdf_paginated

# Bisa (ASCII cat companion)
from bisa import AsciiDogWidget

# Config persistence
from config_manager import ConfigManager

# Core logic
from data_core import (
    APP_VERSION,
    APP_NAME,
    COLUMNS_TO_USE,
    SALES_FLOOR_ALIASES,
    load_raw_df,
    automap_columns,
    compute_moveup_from_df,
    normalize_rooms,
    sort_with_backstock_priority,
    sanitize_prefix,
    aggregate_split_packages_by_room,
)

# Extracted modules
from dialogs import (
    open_map_columns_dialog,
    open_filters_window as _dlg_filters,
    open_audit_window as _dlg_audit,
    open_manual_add_dialog,
    get_all_rooms_normalized,
    get_all_brands,
    get_all_types,
    default_candidate_rooms,
)
from tree_ops import (
    get_display_cols,
    configure_tree_columns,
    sort_tree,
    refresh_treeview_columns,
    render_moveup_tree,
    render_kuntal_tree,
    render_excluded_tree,
    render_all_tree,
)

# ------------------------------
# Excel export stays here for now
# ------------------------------
def export_excel(
    move_up_df: pd.DataFrame,
    priority_df: Optional[pd.DataFrame],
    base_dir: str,
    timestamp: bool,
    prefix: Optional[str],
):
    parts = ["Sticker_Sheet_Filtered_Move_Up"]
    if timestamp:
        parts.append(datetime.now().strftime("%Y-%m-%d_%H-%M"))
    xlsx = "_".join(parts) + ".xlsx"
    if prefix:
        prefix = sanitize_prefix(prefix)
        xlsx = f"{prefix}_{xlsx}"
    out = os.path.join(base_dir, xlsx)

    mu = move_up_df.copy() if move_up_df is not None else pd.DataFrame(columns=COLUMNS_TO_USE)
    prio = priority_df.copy() if priority_df is not None else pd.DataFrame(columns=COLUMNS_TO_USE)

    if not prio.empty and not mu.empty:
        prio_bcs = set(prio["Package Barcode"].astype(str).fillna("").str.strip().tolist())
        mu = mu[~mu["Package Barcode"].astype(str).fillna("").str.strip().isin(prio_bcs)].copy()

    prio = sort_with_backstock_priority(prio) if not prio.empty else prio
    mu = sort_with_backstock_priority(mu) if not mu.empty else mu

    with pd.ExcelWriter(out, engine="openpyxl") as w:
        if not prio.empty:
            prio.to_excel(w, sheet_name="Priority", index=False)
        mu.to_excel(w, sheet_name="Move_Up_Items", index=False)

    return out



# GUI
# ------------------------------
class MoveUpGUI:

    def __init__(self, root: Tk):
        self.root = root
        self.base_title = f"{APP_NAME} v{APP_VERSION}"
        self.root.title(self.base_title)
        self.root.geometry("1240x920")

        self.style = ttk.Style(self.root)
        self.base_theme = self.style.theme_use()

        # Persisted vars
        self.printer_bw_var = BooleanVar(value=False)
        self.skip_sales_floor_var = BooleanVar(value=False)
        self.hide_removed_var = BooleanVar(value=True)
        self.auto_open_var = BooleanVar(value=(os.name == "nt"))
        self.timestamp_var = BooleanVar(value=True)
        self.page_items_var = IntVar(value=35)
        self.prefix_var = StringVar(value="")
        self.show_advanced_var = BooleanVar(value=False)

        # Active display columns (may differ from COLUMNS_TO_USE at runtime)
        self.active_columns: List[str] = list(COLUMNS_TO_USE)
        self._sort_state: Dict[str, Dict[str, bool]] = {}  # {tree_id: {col: ascending}}

        self._button_registry = []
        self._create_kawaii_theme()

        self.cfg = ConfigManager(tk_root=self.root)
        self.app_dir = self.cfg.app_dir

        self.export_root = os.path.join(self.app_dir, "generated")
        os.makedirs(self.export_root, exist_ok=True)
        self._export_run_dir: Optional[str] = None  # created lazily on first export/open

        # Persistent filters + aliases
        self.room_alias_map: Dict[str, str] = {}
        self.selected_rooms: List[str] = []
        self.selected_brands: List[str] = []
        self.selected_types: List[str] = []

        self.last_import_dir: Optional[str] = None
        self.current_file_path: Optional[str] = None

        # Runtime state
        self.raw_df: Optional[pd.DataFrame] = None
        self.current_df: Optional[pd.DataFrame] = None
        self.col_mapping_override: Dict[str, str] = {}
        self.moveup_df: Optional[pd.DataFrame] = None
        self.excluded_barcodes: set = set()
        self.kuntal_priority_barcodes: set = set()

        self.filters_window: Optional[Toplevel] = None

        # Lifetime Bisa stats (loaded from config before UI is built)
        self._lifetime_pets: int = 0
        self._lifetime_treats: int = 0
        self._lifetime_moveups: int = 0
        self._bisa_name: str = "Bisa"
        self._catnip_redeemed: int = 0

        # Previous inventory snapshot for move-up detection (barcode → room, lowercase)
        self._prev_inventory_snapshot: Dict[str, str] = {}

        self._load_config()
        self._build_ui()

        # Push persisted totals into the widget now that it exists
        self.dog_widget.restore_state(
            pets=self._lifetime_pets,
            treats=self._lifetime_treats,
            moveups=self._lifetime_moveups,
            catnip_redeemed=self._catnip_redeemed,
            on_catnip_change=self._on_catnip_redeemed,
        )
        self.dog_widget.greet_startup()

        self._bind_window_treat()
        self._toggle_theme(initial=True)
        self._refresh_button_labels()
        self._update_kuntalcount()

        self.root.protocol("WM_DELETE_WINDOW", self._on_app_close)

    # ------------------------------
    # Config persistence
    # ------------------------------
    def _load_config(self):
        """Pull persisted config from disk via ConfigManager."""
        self.cfg.load(valid_columns=list(COLUMNS_TO_USE))
        c = self.cfg

        # Plain instance vars
        self.room_alias_map   = c["room_alias_map"]
        self.selected_rooms   = c["selected_rooms"]
        self.selected_brands  = c["selected_brands"]
        self.selected_types   = c["selected_types"]

        # Tk variables
        self.printer_bw_var.set(c["printer_bw"])
        self.skip_sales_floor_var.set(c["skip_sales_floor"])
        self.hide_removed_var.set(c["hide_removed"])
        self.auto_open_var.set(c["auto_open_pdf"])
        self.timestamp_var.set(c["timestamp"])
        try:
            self.page_items_var.set(c["items_per_page"])
        except Exception:
            pass
        self.prefix_var.set(c["prefix"])

        # Sets (stored as lists in JSON)
        self.excluded_barcodes         = set(c["excluded_barcodes"])
        self.kuntal_priority_barcodes  = set(c["kuntal_priority_barcodes"])

        # Active columns
        self.active_columns = c["active_columns"] or list(COLUMNS_TO_USE)

        # Validated paths
        if c["last_import_dir"]:
            self.last_import_dir = c["last_import_dir"]
        if c["current_file_path"]:
            self.current_file_path = c["current_file_path"]

        # Bisa stats
        self._lifetime_pets    = c["lifetime_pets"]
        self._lifetime_treats  = c["lifetime_treats"]
        self._lifetime_moveups = c["lifetime_moveups"]
        self._bisa_name        = c["bisa_name"]
        self._catnip_redeemed  = c["catnip_redeemed"]

        # Inventory snapshot
        self._prev_inventory_snapshot = c["prev_inventory_snapshot"]

    def _save_config(self):
        """Push GUI state into ConfigManager and write to disk."""
        c = self.cfg
        bs = self.dog_widget.get_state() if hasattr(self, "dog_widget") else {}

        c["room_alias_map"]           = self.room_alias_map
        c["selected_rooms"]           = self.selected_rooms
        c["selected_brands"]          = self.selected_brands
        c["selected_types"]           = self.selected_types
        c["printer_bw"]               = bool(self.printer_bw_var.get())
        c["skip_sales_floor"]         = bool(self.skip_sales_floor_var.get())
        c["hide_removed"]             = bool(self.hide_removed_var.get())
        c["auto_open_pdf"]            = bool(self.auto_open_var.get())
        c["timestamp"]                = bool(self.timestamp_var.get())
        c["items_per_page"]           = int(self.page_items_var.get() or 30)
        c["prefix"]                   = str(self.prefix_var.get() or "")
        c["last_import_dir"]          = self.last_import_dir or ""
        c["current_file_path"]        = self.current_file_path or ""
        c["excluded_barcodes"]        = sorted(list(self.excluded_barcodes))
        c["kuntal_priority_barcodes"] = sorted(list(self.kuntal_priority_barcodes))
        c["active_columns"]           = self.active_columns
        c["lifetime_pets"]            = bs.get("total_pets", self._lifetime_pets)
        c["lifetime_treats"]          = bs.get("total_treats", self._lifetime_treats)
        c["lifetime_moveups"]         = bs.get("total_moveups", self._lifetime_moveups)
        c["bisa_name"]                = bs.get("name", self._bisa_name)
        c["catnip_redeemed"]          = bs.get("catnip_redeemed", self._catnip_redeemed)
        c["prev_inventory_snapshot"]  = self._prev_inventory_snapshot

        c.save()

    def _on_bisa_renamed(self, new_name: str):
        """Callback from Bisa widget when the user renames her."""
        self._bisa_name = new_name
        self._save_config()

    def _on_catnip_redeemed(self, redeemed_count: int):
        """Callback from Bisa widget when catnip is redeemed."""
        self._catnip_redeemed = redeemed_count
        self._save_config()

    def _on_app_close(self):
        self._save_config()
        try:
            self.root.destroy()
        except Exception as e:
            print(f"[moveup] Warning: error during window close: {e}")

    # ------------------------------
    # Base helpers
    # ------------------------------
    @property
    def export_run_dir(self) -> str:
        """Lazily create the timestamped export directory on first actual use."""
        if self._export_run_dir is None:
            self._export_run_dir = os.path.join(
                self.export_root, datetime.now().strftime("%Y-%m-%d_%H-%M")
            )
            os.makedirs(self._export_run_dir, exist_ok=True)
        return self._export_run_dir

    def _create_kawaii_theme(self):
        if "kawaii_daisy" in self.style.theme_names():
            return
        self.style.theme_create(
            "kawaii_daisy",
            parent=self.base_theme,
            settings={
                "TFrame": {"configure": {"background": "#ede8f7"}},
                "TLabel": {"configure": {"background": "#ede8f7", "foreground": "#3b1f6e"}},
                "TButton": {
                    "configure": {"padding": 6, "relief": "raised", "background": "#d0c4ee", "foreground": "#3b1f6e"},
                    "map": {"background": [("active", "#bfb0e6"), ("pressed", "#ae9cde")]}
                },
                "Treeview": {
                    "configure": {"background": "#f5f2fb", "fieldbackground": "#f5f2fb", "foreground": "#333333",
                                  "rowheight": 20},
                    "map": {"background": [("selected", "#bfb0e6")], "foreground": [("selected", "#000000")]}
                },
                "TCheckbutton": {"configure": {"background": "#ede8f7"}},
                "TNotebook": {
                    "configure": {"background": "#ede8f7", "tabmargins": [2, 4, 2, 0]},
                },
                "TNotebook.Tab": {
                    "configure": {
                        "background": "#b8aad8",   # unselected: darker lavender
                        "foreground": "#5a3f8a",   # unselected: muted purple
                        "padding": [10, 4],
                    },
                    "map": {
                        "background": [
                            ("selected", "#f5f2fb"),  # selected: bright near-white, matches treeview
                            ("active",   "#cfc3ea"),  # hover
                        ],
                        "foreground": [
                            ("selected", "#3b1f6e"),  # selected: dark bold purple
                            ("active",   "#4a2a80"),
                        ],
                        "expand": [("selected", [2, 3, 2, 0])],  # lifts active tab up
                    },
                },
            }
        )

    def _register_button(self, btn, base_text: str):
        self._button_registry.append((btn, base_text))

    def _refresh_button_labels(self):
        for btn, base in self._button_registry:
            btn.config(text=f"🌼 {base} 🌼")

    def _toggle_theme(self, initial: bool = False):
        self.style.theme_use("kawaii_daisy")
        self.root.title(self.base_title + " 🌼🌼🌼")
        self._refresh_button_labels()

    def open_kawaii_settings(self):
        try:
            from kawaii_preview import open_kawaii_settings_window
            open_kawaii_settings_window(self.root)
        except Exception as e:
            messagebox.showerror("Kawaii PDF Settings", f"Could not open settings window:\n\n{e}")

    def open_expiring_window(self):
        try:
            from mainExpiring import open_expiring_window
            open_expiring_window(self.root, self.current_file_path)
        except Exception as e:
            messagebox.showerror("Expiring Soon", f"Could not open window:\n\n{e}")

    def open_sample_manager(self):
        try:
            from mainSamples import open_sample_manager
            open_sample_manager(self.root, self.current_file_path)
        except Exception as e:
            messagebox.showerror("Sample Manager", f"Could not open window:\n\n{e}")

    # ------------------------------
    # UI
    # ------------------------------
    def _build_ui(self):
        # ==============================
        # TOP ROW: controls (left) + Bisa natural-height (right)
        # ==============================
        frm_top_row = ttk.Frame(self.root, padding=(10, 8, 10, 4))
        frm_top_row.pack(fill="x")

        # ── Left: all controls ──
        frm_controls = ttk.LabelFrame(frm_top_row, text="Controls", padding=8)
        frm_controls.pack(side="left", fill="both", expand=True, padx=(0, 12))
        btn_row = ttk.Frame(frm_controls)
        btn_row.pack(fill="x", pady=(0, 4))

        btn_import = ttk.Button(btn_row, text="Import File…", command=self.import_file)
        btn_import.pack(side="left", padx=4)
        self._register_button(btn_import, "Import File…")

        btn_pdf = ttk.Button(btn_row, text="Export PDF", command=self.do_export_pdf)
        btn_pdf.pack(side="left", padx=4)
        self._register_button(btn_pdf, "Export PDF")

        btn_audit = ttk.Button(btn_row, text="Audit PDFs…", command=self.open_audit_window)
        btn_audit.pack(side="left", padx=4)
        self._register_button(btn_audit, "Audit PDFs…")

        btn_expiring = ttk.Button(
            btn_row, text="Expiring Soon…", command=self.open_expiring_window,
        )
        btn_expiring.pack(side="left", padx=4)
        self._register_button(btn_expiring, "Expiring Soon…")

        btn_samples = ttk.Button(
            btn_row, text="Sample Manager…", command=self.open_sample_manager,
        )
        btn_samples.pack(side="left", padx=4)
        self._register_button(btn_samples, "Sample Manager…")

        # Advanced toggle (ANCHOR target for frm_advanced)
        self.frm_adv_toggle = ttk.Frame(frm_controls)
        self.frm_adv_toggle.pack(fill="x", pady=(2, 0))
        self._adv_button = ttk.Button(
            self.frm_adv_toggle, text="▶ Advanced", command=self._toggle_advanced,
        )
        self._adv_button.pack(side="left")

        # Advanced controls (hidden, child of frm_controls so before= works)
        self.frm_advanced = ttk.Frame(frm_controls, padding=(0, 4, 0, 0))
        adv_row = ttk.Frame(self.frm_advanced)
        adv_row.pack(fill="x")

        btn_map = ttk.Button(adv_row, text="Map Columns…", command=self.map_columns_dialog)
        btn_map.pack(side="left", padx=4)
        self._register_button(btn_map, "Map Columns…")

        btn_xlsx = ttk.Button(adv_row, text="Export Excel", command=self.do_export_xlsx)
        btn_xlsx.pack(side="left", padx=4)
        self._register_button(btn_xlsx, "Export Excel")

        btn_folder = ttk.Button(adv_row, text="Open Output Folder", command=self.open_output_folder)
        btn_folder.pack(side="left", padx=4)
        self._register_button(btn_folder, "Open Output Folder")

        ttk.Checkbutton(
            adv_row, text="Printer B/W",
            variable=self.printer_bw_var, command=self._save_config,
        ).pack(side="left", padx=6)

        btn_filters = ttk.Button(adv_row, text="Filters…", command=self.open_filters_window)
        btn_filters.pack(side="left", padx=4)
        self._register_button(btn_filters, "Filters…")

        btn_kawaii_settings_main = ttk.Button(
            adv_row, text="Kawaii PDF Settings…", command=self.open_kawaii_settings,
        )
        btn_kawaii_settings_main.pack(side="left", padx=4)
        self._register_button(btn_kawaii_settings_main, "Kawaii PDF Settings…")

        # Items per page (ANCHOR)
        self.frm_page = ttk.Frame(frm_controls)
        self.frm_page.pack(fill="x", pady=(4, 2))
        ttk.Label(self.frm_page, text="Items per page").pack(side="left")
        ttk.Spinbox(
            self.frm_page, from_=10, to=200,
            textvariable=self.page_items_var, width=6, command=self._save_config,
        ).pack(side="left", padx=6)

        # Status labels
        self.status = StringVar(value="Ready.")
        ttk.Label(frm_controls, textvariable=self.status, anchor="w").pack(fill="x")

        self.rowcount_var = StringVar(value="Items loaded: 0")
        ttk.Label(frm_controls, textvariable=self.rowcount_var, anchor="w").pack(fill="x")

        self.moveupcount_var = StringVar(value="Move-Up items: 0")
        ttk.Label(frm_controls, textvariable=self.moveupcount_var, anchor="w").pack(fill="x")

        self.kuntalcount_var = StringVar(value="Priority! items: 0")
        ttk.Label(frm_controls, textvariable=self.kuntalcount_var, anchor="w").pack(fill="x")

        self.filters_summary_var = StringVar(value="Filters: default")
        ttk.Label(frm_controls, textvariable=self.filters_summary_var, anchor="w", wraplength=480).pack(fill="x")

        # ── Right: Bisa — stretch wide, top-aligned ──
        self.dog_widget = AsciiDogWidget(
            frm_top_row,
            name=self._bisa_name,
            on_rename=self._on_bisa_renamed,
        )
        self.dog_widget.frame.pack(side="left", fill="x", expand=True, anchor="n")

        # ==============================
        # MIDDLE: Notebook — expands to fill all remaining space
        # ==============================
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True, padx=10, pady=(4, 0))

        self.tab_moveup = ttk.Frame(self.nb)
        self.tab_kuntal = ttk.Frame(self.nb)
        self.tab_excluded = ttk.Frame(self.nb)
        self.tab_all = ttk.Frame(self.nb)

        self.nb.add(self.tab_moveup, text="Move-Up List")
        self.nb.add(self.tab_kuntal, text="Priority!")
        self.nb.add(self.tab_excluded, text="Excluded / Removed")
        self.nb.add(self.tab_all, text="All Items")

        # Colored dot indicators on each tab (PhotoImage is the only reliable
        # cross-platform way to get per-tab color accents in ttk.Notebook)
        def _make_tab_dot(color: str) -> tk.PhotoImage:
            img = tk.PhotoImage(width=10, height=10)
            img.put(color, to=(1, 1, 9, 9))
            return img

        self._tab_dot_moveup   = _make_tab_dot("#4a9fd4")   # steel blue
        self._tab_dot_kuntal   = _make_tab_dot("#e8a020")   # amber
        self._tab_dot_excluded = _make_tab_dot("#d46060")   # soft red
        self._tab_dot_all      = _make_tab_dot("#9988cc")   # muted purple

        self.nb.tab(self.tab_moveup,   image=self._tab_dot_moveup,   compound="left")
        self.nb.tab(self.tab_kuntal,   image=self._tab_dot_kuntal,   compound="left")
        self.nb.tab(self.tab_excluded, image=self._tab_dot_excluded, compound="left")
        self.nb.tab(self.tab_all,      image=self._tab_dot_all,      compound="left")

        self.tree = ttk.Treeview(self.tab_moveup, columns=tuple(self.active_columns), show="headings", height=18)
        self._configure_tree_columns(self.tree, self.active_columns)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self._on_moveup_double_click)
        self.tree.bind("<ButtonRelease-1>", self._on_moveup_single_click)

        self.k_tree = ttk.Treeview(self.tab_kuntal, columns=tuple(COLUMNS_TO_USE), show="headings", height=18)
        self._configure_tree_columns(self.k_tree, COLUMNS_TO_USE)
        self.k_tree.pack(fill="both", expand=True)
        self.k_tree.bind("<Double-Button-1>", self._kuntal_tree_double_click)

        self.x_tree = ttk.Treeview(self.tab_excluded, columns=tuple(COLUMNS_TO_USE), show="headings", height=18)
        self._configure_tree_columns(self.x_tree, COLUMNS_TO_USE)
        self.x_tree.pack(fill="both", expand=True)
        self.x_tree.bind("<Double-1>", self._on_excluded_double_click)

        frm_all_top = ttk.Frame(self.tab_all)
        frm_all_top.pack(fill="x", padx=6, pady=(6, 2))
        ttk.Label(frm_all_top, text="Search:").pack(side="left")
        self.all_search_var = StringVar(value="")
        ttk.Entry(frm_all_top, textvariable=self.all_search_var, width=40).pack(side="left", padx=6)
        ttk.Button(frm_all_top, text="Clear", command=lambda: self.all_search_var.set("")).pack(side="left")
        self.all_items_count_var = StringVar(value="")
        ttk.Label(frm_all_top, textvariable=self.all_items_count_var, foreground="#555").pack(side="left", padx=12)

        all_frm = ttk.Frame(self.tab_all)
        all_frm.pack(fill="both", expand=True)
        self.all_tree = ttk.Treeview(all_frm, columns=tuple(COLUMNS_TO_USE), show="headings", height=18)
        self._configure_tree_columns(self.all_tree, COLUMNS_TO_USE)
        all_sb = ttk.Scrollbar(all_frm, orient="vertical", command=self.all_tree.yview)
        self.all_tree.configure(yscrollcommand=all_sb.set)
        self.all_tree.pack(side="left", fill="both", expand=True)
        all_sb.pack(side="right", fill="y")
        self.all_tree.bind("<Double-Button-1>", self._all_tree_double_click)

        self.all_search_var.trace_add("write", lambda *_: self._render_all_tree(self.current_df))

        # ==============================
        # BOTTOM: action buttons + diag
        # ==============================
        frm_remove = ttk.Frame(self.root, padding=(10, 4, 10, 4))
        frm_remove.pack(fill="x")

        btn_toggle = ttk.Button(frm_remove, text="Toggle Remove", command=self._toggle_remove_selected)
        btn_toggle.pack(side="left", padx=4)
        self._register_button(btn_toggle, "Toggle Remove")

        btn_clear = ttk.Button(frm_remove, text="Clear Removed", command=self._clear_removed)
        btn_clear.pack(side="left", padx=4)
        self._register_button(btn_clear, "Clear Removed")

        ttk.Separator(frm_remove, orient="vertical").pack(side="left", fill="y", padx=8)

        btn_kuntal = ttk.Button(frm_remove, text="Toggle Priority!", command=self._toggle_kuntal_selected)
        btn_kuntal.pack(side="left", padx=4)
        self._register_button(btn_kuntal, "Toggle Priority!")

        btn_manual = ttk.Button(frm_remove, text="Manual Add…", command=self._manual_add_dialog)
        btn_manual.pack(side="left", padx=4)
        self._register_button(btn_manual, "Manual Add…")

        btn_clear_k = ttk.Button(frm_remove, text="Clear Priority! List", command=self._clear_kuntal_list)
        btn_clear_k.pack(side="left", padx=4)
        self._register_button(btn_clear_k, "Clear Priority! List")

        self.diag_var = StringVar(value="")
        ttk.Label(
            self.root, textvariable=self.diag_var,
            anchor="w", foreground="#555",
        ).pack(fill="x", padx=10, pady=(0, 6))



    def _configure_tree_columns(self, tree: ttk.Treeview, cols: List[str]):
        configure_tree_columns(tree, cols, self._sort_state, self._sort_tree)

    def _sort_tree(self, tree: ttk.Treeview, tree_id: str, col: str):
        sort_tree(tree, tree_id, col, self._sort_state)

    def _toggle_advanced(self):
        show = not self.show_advanced_var.get()
        self.show_advanced_var.set(show)

        if show:
            self.frm_advanced.pack(fill="x", before=self.frm_page)
            self._adv_button.config(text="▼ Advanced")
        else:
            self.frm_advanced.pack_forget()
            self._adv_button.config(text="▶ Advanced")

    # ------------------------------
    # ------------------------------
    # Status counters
    # ------------------------------
    def _update_rowcount(self, df: Optional[pd.DataFrame]):
        n = 0 if df is None else len(df)
        self.rowcount_var.set(f"Items loaded: {n}")

    def _update_moveupcount(self, df: Optional[pd.DataFrame]):
        n = 0 if df is None else len(df)
        self.moveupcount_var.set(f"Move-Up items: {n}")
        try:
            self.nb.tab(self.tab_moveup, text=f"Move-Up ({n})")
        except Exception:
            pass

    def _update_kuntalcount(self):
        n = len(self.kuntal_priority_barcodes)
        self.kuntalcount_var.set(f"Priority! items: {n}")
        try:
            self.nb.tab(self.tab_kuntal, text=f"Priority! ({n})")
        except Exception:
            pass

    # ------------------------------
    # Window-wide treat throwing
    # ------------------------------
    def _bind_window_treat(self):
        """Any click on blank space throws Bisa a treat."""
        # Widget types that should NOT trigger a treat (they have their own click behaviour)
        _SKIP_TYPES = (
            "Button", "TButton", "Treeview", "Entry", "TEntry",
            "Combobox", "TCombobox", "Scrollbar", "TScrollbar",
            "Checkbutton", "TCheckbutton", "Radiobutton", "TRadiobutton",
            "Scale", "TScale", "Spinbox", "TSpinbox", "Text",
            "Notebook", "TNotebook",
        )

        def _on_click(event):
            if not hasattr(self, "dog_widget"):
                return
            # Skip if clicking on an interactive widget
            w = event.widget
            wtype = w.winfo_class()
            if any(wtype == t or wtype.endswith(t) for t in _SKIP_TYPES):
                return
            # Also skip if the widget is inside Bisa's own frame
            try:
                parent = w
                while parent:
                    if parent == self.dog_widget.frame:
                        return
                    parent = parent.master
            except Exception:
                pass

            # Convert absolute screen coords to window-relative x
            win_x = event.x_root - self.root.winfo_rootx()
            win_w = self.root.winfo_width()
            self.dog_widget.throw_treat_at_window_x(win_x, win_w)

        self.root.bind_all("<Button-1>", _on_click, add="+")

    # ------------------------------
    # Display columns (core + optional extras if present in data)
    # ------------------------------
    DISPLAY_EXTRA_COLUMNS = ["Received Date"]

    def _display_cols_for(self, df: "Optional[pd.DataFrame]" = None) -> List[str]:
        return get_display_cols(
            self.active_columns,
            df if df is not None else self.current_df,
            self.DISPLAY_EXTRA_COLUMNS,
        )

    # ------------------------------
    # Open folder
    # ------------------------------
    def open_output_folder(self):
        if self._export_run_dir is None:
            messagebox.showinfo("Open Folder", "No exports yet this session.")
            return
        path = self._export_run_dir
        try:
            if os.name == "nt":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.run(["open", path], check=False)
            else:
                subprocess.run(["xdg-open", path], check=False)
        except Exception as e:
            messagebox.showerror("Open Folder", f"Could not open folder:\n{path}\n\n{e}")

    # ------------------------------
    # Import / mapping
    # ------------------------------
    def import_file(self):
        initialdir = self.last_import_dir if (self.last_import_dir and os.path.isdir(self.last_import_dir)) else None
        path = filedialog.askopenfilename(
            title="Select Inventory File",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")],
            initialdir=initialdir
        )
        if not path:
            return

        try:
            self.last_import_dir = os.path.dirname(path)
            self.current_file_path = path
            self._save_config()

            self.status.set(f"Loading {os.path.basename(path)}…")
            raw = load_raw_df(path)
            self.raw_df = raw

            try:
                mapped, _used = automap_columns(raw)
                self.current_df = mapped

                present = set(self.current_df["Package Barcode"].astype(str).fillna("").str.strip().tolist())
                self.excluded_barcodes = {bc for bc in self.excluded_barcodes if bc in present}
                self.kuntal_priority_barcodes = {bc for bc in self.kuntal_priority_barcodes if bc in present}
                self._update_kuntalcount()

                self.status.set(f"Loaded {len(mapped)} rows. Auto-mapped columns.")

                # --- Bisa move-up detection ---
                # Build a normalized room snapshot for the new inventory
                _snap_df = normalize_rooms(mapped.copy(), self.room_alias_map)
                _sf_set = {"sales floor"} | SALES_FLOOR_ALIASES
                new_snap: Dict[str, str] = {}
                if "Package Barcode" in _snap_df.columns and "Room" in _snap_df.columns:
                    new_snap = {
                        str(bc).strip(): str(rm).strip().lower()
                        for bc, rm in zip(
                            _snap_df["Package Barcode"].astype(str),
                            _snap_df["Room"].astype(str),
                        )
                        if str(bc).strip() and str(bc).strip().lower() != "nan"
                    }
                # Count SKUs that moved from a non-SF room → SF room
                # Compare against most recent snapshot only
                _moved = 0
                if self._prev_inventory_snapshot:
                    for bc, new_room in new_snap.items():
                        old_room = self._prev_inventory_snapshot.get(bc)
                        if old_room is not None:
                            if old_room not in _sf_set and new_room in _sf_set:
                                _moved += 1
                # Replace snapshot with current import (single instance only)
                self._prev_inventory_snapshot = new_snap
                self._save_config()

                if _moved > 0 and hasattr(self, "dog_widget"):
                    self.dog_widget.react_moveups(_moved)
                elif hasattr(self, "dog_widget"):
                    self.dog_widget.react_data_loaded(len(mapped))
                # ------------------------------------

                self._update_rowcount(mapped)
                self._recompute_from_current()
                return

            except Exception:
                self.current_df = None
                self._update_rowcount(raw)
                self.status.set(f"Loaded raw file ({len(raw)} rows). Needs manual column mapping.")

                go = messagebox.askyesno(
                    "Manual Mapping Needed",
                    "This file doesn't match the expected columns.\n\n"
                    "Do you want to manually map columns now?"
                )
                if go:
                    self.map_columns_dialog(force=True)
                return

        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status.set(f"Error: {e}")

    def map_columns_dialog(self, force: bool = False):
        if self.raw_df is None or self.raw_df.empty:
            messagebox.showinfo("Map Columns", "Import a file first.")
            return

        def _on_apply(mapped_df, mapping):
            self.col_mapping_override = mapping
            self.current_df = mapped_df
            present = set(mapped_df["Package Barcode"].astype(str).fillna("").str.strip().tolist())
            self.excluded_barcodes = {bc for bc in self.excluded_barcodes if bc in present}
            self.kuntal_priority_barcodes = {bc for bc in self.kuntal_priority_barcodes if bc in present}
            self._update_kuntalcount()
            self._update_rowcount(mapped_df)
            self._recompute_from_current()
            self.status.set("Column mapping applied (METRC forced to Package Barcode).")

        open_map_columns_dialog(self.root, self.raw_df, on_apply=_on_apply, force=force)

    # ------------------------------
    # Filters window
    # ------------------------------
    def open_filters_window(self):
        if self.current_df is None or self.current_df.empty:
            messagebox.showinfo("Filters", "Import a file first.")
            return

        if self.filters_window is not None and self.filters_window.winfo_exists():
            try:
                self.filters_window.lift()
                self.filters_window.focus_force()
            except Exception:
                pass
            return

        def _on_apply(rooms, brands, types):
            self.selected_rooms = rooms
            self.selected_brands = brands
            self.selected_types = types
            self._save_config()
            self._recompute_from_current()
            self.filters_window = None

        self.filters_window = _dlg_filters(
            parent=self.root,
            current_df=self.current_df,
            room_alias_map=self.room_alias_map,
            selected_rooms=self.selected_rooms,
            selected_brands=self.selected_brands,
            selected_types=self.selected_types,
            on_apply=_on_apply,
            on_alias_changed=self._save_config,
            on_close=lambda: setattr(self, "filters_window", None),
        )

    # ------------------------------
    # Audit window
    # ------------------------------
    def open_audit_window(self):
        if self.current_df is None or self.current_df.empty:
            messagebox.showinfo("Audit PDFs", "Import a file first.")
            return

        def _on_success(msg):
            if hasattr(self, "dog_widget"):
                self.dog_widget.react_success(msg)

        def _on_error(msg):
            if hasattr(self, "dog_widget"):
                self.dog_widget.react_error(msg)

        _dlg_audit(
            parent=self.root,
            current_df=self.current_df,
            room_alias_map=self.room_alias_map,
            export_run_dir=self.export_run_dir,
            printer_bw=bool(self.printer_bw_var.get()),
            auto_open=bool(self.auto_open_var.get()),
            on_status=self.status.set,
            on_success=_on_success,
            on_error=_on_error,
        )

    # ------------------------------
    # Tree rendering
    # ------------------------------
    def _refresh_treeview_columns(self, df: "Optional[pd.DataFrame]" = None):
        refresh_treeview_columns(
            self.tree, self.k_tree, self.x_tree, self.all_tree,
            self.active_columns, self.DISPLAY_EXTRA_COLUMNS,
            df if df is not None else self.current_df,
            self._sort_state, self._sort_tree,
        )

    def _render_tree(self, df: pd.DataFrame):
        render_moveup_tree(
            self.tree, df, self._display_cols_for(df),
            self.kuntal_priority_barcodes, self.excluded_barcodes,
            self.hide_removed_var.get(),
        )

    def _render_kuntal_tree(self, df: pd.DataFrame):
        render_kuntal_tree(self.k_tree, df, self._display_cols_for(df))

    def _render_excluded_tree(self, df: pd.DataFrame):
        render_excluded_tree(self.x_tree, df, self._display_cols_for(df))

    def _render_all_tree(self, df: Optional[pd.DataFrame]):
        render_all_tree(
            self.all_tree, df, self._display_cols_for(df),
            self.all_search_var.get(),
            self.excluded_barcodes, self.kuntal_priority_barcodes,
            self.all_items_count_var,
        )


    # ------------------------------
    # ── NEW: Double-click to exclude ──
    # ------------------------------
    def _on_moveup_single_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region not in ("cell", "tree"):
            return
        if not self.tree.identify_row(event.y):
            return
        if hasattr(self, "dog_widget"):
            self.dog_widget.react_row_selected()

    def _on_moveup_double_click(self, event):
        """Double-clicking a row in the Move-Up tree immediately excludes it
        and switches to the Excluded tab so the user sees where it went."""
        region = self.tree.identify("region", event.x, event.y)
        if region not in ("cell", "tree"):
            return

        iid = self.tree.identify_row(event.y)
        if not iid:
            return

        # Package Barcode is always in active_columns (enforced by column editor)
        idx_bar = self.active_columns.index("Package Barcode")
        vals = self.tree.item(iid, "values")
        if not vals or len(vals) <= idx_bar:
            return

        bc = str(vals[idx_bar]).strip()
        if not bc:
            return

        already_excluded = bc in self.excluded_barcodes
        if already_excluded:
            self.excluded_barcodes.discard(bc)
            self.status.set(f"Restored from excluded: …{bc[-6:]}")
            if hasattr(self, "dog_widget"):
                self.dog_widget.react_restored(1)
        else:
            self.excluded_barcodes.add(bc)
            self.status.set(f"Excluded (double-click): …{bc[-6:]}")
            if hasattr(self, "dog_widget"):
                self.dog_widget.react_excluded(1)

        self._recompute_from_current()
        self._save_config()

    # ------------------------------
    # Remove / Kuntal
    # ------------------------------
    def _remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Remove", "Select row(s) first.")
            return
        idx_bar = self.active_columns.index("Package Barcode")
        removed = 0
        for iid in sel:
            vals = self.tree.item(iid, "values")
            if not vals or len(vals) <= idx_bar:
                continue
            bc = str(vals[idx_bar]).strip()
            if bc:
                self.excluded_barcodes.add(bc)
                removed += 1
        self._recompute_from_current()
        self.status.set(f"Removed {removed} item(s) this session.")
        self._save_config()

    def _toggle_remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx_bar = self.active_columns.index("Package Barcode")
        toggled = 0
        for iid in sel:
            vals = self.tree.item(iid, "values")
            if not vals or len(vals) <= idx_bar:
                continue
            bc = str(vals[idx_bar]).strip()
            if not bc:
                continue
            if bc in self.excluded_barcodes:
                self.excluded_barcodes.remove(bc)
            else:
                self.excluded_barcodes.add(bc)
            toggled += 1
        self._recompute_from_current()
        self.status.set(f"Toggled remove on {toggled} item(s).")
        self._save_config()
        if toggled and hasattr(self, "dog_widget"):
            self.dog_widget.react_excluded(toggled)

    def _clear_removed(self):
        self.excluded_barcodes.clear()
        self._recompute_from_current()
        self.status.set("Cleared manually removed items.")
        self._save_config()
        if hasattr(self, "dog_widget"):
            self.dog_widget.react_cleared()

    def _toggle_kuntal_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Priority!", "Select row(s) first.")
            return
        idx_bar = self.active_columns.index("Package Barcode")
        toggled = 0
        for iid in sel:
            vals = self.tree.item(iid, "values")
            if not vals or len(vals) <= idx_bar:
                continue
            bc = str(vals[idx_bar]).strip()
            if not bc:
                continue
            if bc in self.kuntal_priority_barcodes:
                self.kuntal_priority_barcodes.remove(bc)
            else:
                self.kuntal_priority_barcodes.add(bc)
            toggled += 1
        self._update_kuntalcount()
        self._recompute_from_current()
        self.status.set(f"Toggled Priority! on {toggled} item(s).")
        self._save_config()
        if toggled and hasattr(self, "dog_widget"):
            self.dog_widget.react_kuntal(toggled)

    def _clear_kuntal_list(self):
        self.kuntal_priority_barcodes.clear()
        self._update_kuntalcount()
        self._recompute_from_current()
        self.status.set("Cleared Priority! list.")
        self._save_config()
        if hasattr(self, "dog_widget"):
            self.dog_widget.react_cleared()

    def _all_tree_double_click(self, event):
        iid = self.all_tree.identify_row(event.y)
        if not iid:
            return
        vals = self.all_tree.item(iid, "values")
        if not vals:
            return
        try:
            bc_idx = list(COLUMNS_TO_USE).index("Package Barcode")
            bc = str(vals[bc_idx]).strip()
        except (ValueError, IndexError):
            return
        if not bc:
            return
        self.kuntal_priority_barcodes.add(bc)
        self._update_kuntalcount()
        self._recompute_from_current()
        self.status.set(f"Added to Priority!: …{bc[-6:]}")
        self._save_config()
        if hasattr(self, "dog_widget"):
            self.dog_widget.react_kuntal(1)

    def _kuntal_tree_double_click(self, event):
        iid = self.k_tree.identify_row(event.y)
        if not iid:
            return
        vals = self.k_tree.item(iid, "values")
        if not vals:
            return
        try:
            bc_idx = list(COLUMNS_TO_USE).index("Package Barcode")
            bc = str(vals[bc_idx]).strip()
        except (ValueError, IndexError):
            return
        if not bc or bc not in self.kuntal_priority_barcodes:
            return
        self.kuntal_priority_barcodes.discard(bc)
        self._update_kuntalcount()
        self._recompute_from_current()
        self.status.set(f"Removed from Priority!: …{bc[-6:]}")
        self._save_config()

    # ------------------------------
    # Excluded single-click restore
    # ------------------------------
    def _on_excluded_double_click(self, event):
        region = self.x_tree.identify("region", event.x, event.y)
        if region not in ("cell", "tree"):
            return

        iid = self.x_tree.identify_row(event.y)
        if not iid:
            return

        try:
            self.x_tree.selection_set(iid)
        except Exception:
            pass

        self._restore_excluded_selected(go_to_moveup=False, quiet=True)

    def _restore_excluded_selected(self, go_to_moveup: bool = True, quiet: bool = False):
        sel = self.x_tree.selection()
        if not sel:
            if not quiet:
                messagebox.showinfo("Restore", "Select an excluded item first.")
            return

        idx_bar = COLUMNS_TO_USE.index("Package Barcode")
        restored = 0
        restored_bcs = []

        for iid in sel:
            vals = self.x_tree.item(iid, "values")
            if not vals or len(vals) <= idx_bar:
                continue
            bc = str(vals[idx_bar]).strip()
            if not bc:
                continue
            if bc in self.excluded_barcodes:
                self.excluded_barcodes.remove(bc)
                restored += 1
                restored_bcs.append(bc)

        if restored == 0:
            return

        self._recompute_from_current()

        if go_to_moveup:
            try:
                self.nb.select(self.tab_moveup)
                self.tree.selection_remove(self.tree.selection())
                idx_bar_main = self.active_columns.index("Package Barcode")

                to_select = []
                for iid2 in self.tree.get_children():
                    v = self.tree.item(iid2, "values")
                    if not v or len(v) <= idx_bar_main:
                        continue
                    bc2 = str(v[idx_bar_main]).strip()
                    if bc2 in restored_bcs:
                        to_select.append(iid2)

                if to_select:
                    self.tree.selection_set(to_select)
                    self.tree.focus(to_select[0])
                    self.tree.see(to_select[0])
            except Exception:
                pass

        self.status.set(f"Restored {restored} item(s) from Excluded.")
        self._save_config()
        if hasattr(self, "dog_widget"):
            self.dog_widget.react_restored(restored)

    # ------------------------------
    # Manual Add
    # ------------------------------
    def _manual_add_dialog(self):
        if self.current_df is None or self.current_df.empty:
            messagebox.showinfo("Manual Add", "Import a file first.")
            return

        def _on_apply(added_barcodes):
            self.kuntal_priority_barcodes |= added_barcodes
            self._update_kuntalcount()
            self._recompute_from_current()
            self.status.set(f"Manual added {len(added_barcodes)} item(s) to Priority!")
            self._save_config()

        open_manual_add_dialog(
            parent=self.root,
            current_df=self.current_df,
            kuntal_priority_barcodes=self.kuntal_priority_barcodes,
            on_apply=_on_apply,
        )

    # ------------------------------
    # Data getters
    # ------------------------------
    def _get_kuntal_priority_df(self) -> pd.DataFrame:
        if self.current_df is None or self.current_df.empty:
            return pd.DataFrame(columns=COLUMNS_TO_USE)
        if not self.kuntal_priority_barcodes:
            return pd.DataFrame(columns=COLUMNS_TO_USE)

        df = self.current_df.copy()
        df["Package Barcode"] = df["Package Barcode"].astype(str).fillna("").str.strip()
        keep = df["Package Barcode"].isin({str(x).strip() for x in self.kuntal_priority_barcodes})
        out = df.loc[keep, COLUMNS_TO_USE].copy()
        if not out.empty:
            out = out.sort_values(by=["Room", "Brand", "Product Name"], kind="stable")
        return out

    def _get_excluded_df(self) -> pd.DataFrame:
        if self.current_df is None or self.current_df.empty:
            return pd.DataFrame(columns=COLUMNS_TO_USE)
        if not self.excluded_barcodes:
            return pd.DataFrame(columns=COLUMNS_TO_USE)

        df = self.current_df.copy()
        df["Package Barcode"] = df["Package Barcode"].astype(str).fillna("").str.strip()
        keep = df["Package Barcode"].isin({str(x).strip() for x in self.excluded_barcodes})
        out = df.loc[keep, COLUMNS_TO_USE].copy()
        if not out.empty:
            out = out.sort_values(by=["Room", "Brand", "Product Name"], kind="stable")
        return out

    # ------------------------------
    # Effective filters
    # ------------------------------
    def _effective_rooms(self, df: pd.DataFrame) -> List[str]:
        all_rooms = set(get_all_rooms_normalized(df, self.room_alias_map))
        if self.selected_rooms:
            cleaned = [r for r in self.selected_rooms if r in all_rooms]
            if cleaned:
                return cleaned
        return default_candidate_rooms(df, self.room_alias_map)

    def _effective_brands(self, df: pd.DataFrame) -> List[str]:
        all_brands = set(get_all_brands(df))
        if self.selected_brands:
            cleaned = [b for b in self.selected_brands if b in all_brands]
            return cleaned
        return []

    def _effective_types(self, df: pd.DataFrame) -> List[str]:
        all_types = set(get_all_types(df))
        if self.selected_types:
            cleaned = [t for t in self.selected_types if t in all_types]
            return cleaned
        return []

    def _effective_brand_filter(self, df: pd.DataFrame) -> List[str]:
        cleaned = self._effective_brands(df)
        return cleaned if cleaned else ["ALL"]

    def _effective_type_filter(self, df: pd.DataFrame) -> List[str]:
        cleaned = self._effective_types(df)
        return cleaned if cleaned else ["ALL"]

    # ------------------------------
    # Recompute
    # ------------------------------
    def _recompute_from_current(self):
        df = self.current_df
        if df is None or df.empty:
            self._render_tree(pd.DataFrame(columns=COLUMNS_TO_USE))
            self._render_kuntal_tree(pd.DataFrame(columns=COLUMNS_TO_USE))
            self._render_excluded_tree(pd.DataFrame(columns=COLUMNS_TO_USE))
            self._render_all_tree(None)
            self._update_rowcount(None)
            self._update_moveupcount(None)
            self._update_kuntalcount()
            self.status.set("No data loaded.")
            self.diag_var.set("")
            self.filters_summary_var.set("Filters: none (no data)")
            return

        rooms = self._effective_rooms(df)

        move_up_df, diag = compute_moveup_from_df(
            df,
            rooms,
            self.room_alias_map,
            brand_filter=self._effective_brand_filter(df),
            type_filter=self._effective_type_filter(df),
            skip_sales_floor=self.skip_sales_floor_var.get()
        )

        move_up_df = aggregate_split_packages_by_room(move_up_df)

        if self.excluded_barcodes and self.hide_removed_var.get():
            move_up_df = move_up_df[~move_up_df["Package Barcode"].astype(str).fillna("").isin(self.excluded_barcodes)].copy()

        move_up_df = sort_with_backstock_priority(move_up_df)
        self.moveup_df = move_up_df

        # Rebuild all treeview column sets in case Received Date appeared/disappeared
        self._refresh_treeview_columns(df)

        self._render_tree(move_up_df)
        self._update_moveupcount(move_up_df)

        prio_df = self._get_kuntal_priority_df()
        self._render_kuntal_tree(prio_df)
        self._update_kuntalcount()

        excl_df = self._get_excluded_df()
        self._render_excluded_tree(excl_df)
        self._render_all_tree(df)

        self.status.set(f"Loaded {len(df)} rows; Move-Up {len(move_up_df)}")

        self.diag_var.set(
            f"Diagnostics — after dropna: {diag.get('after_dropna')}, "
            f"after brand: {diag.get('after_brand')}, "
            f"after category filter: {diag.get('after_type_filter')}, "
            f"after type(accessories removed): {diag.get('after_type')}, "
            f"candidate pool: {diag.get('candidate_pool')}, "
            f"removed as on Sales Floor: {diag.get('removed_as_on_sf')}."
        )

        b = len(self._effective_brands(df))
        t = len(self._effective_types(df))
        self.filters_summary_var.set(
            f"Filters — Rooms: {len(rooms)} | Brands: {'ALL' if b == 0 else b} | Types: {'ALL' if t == 0 else t} | "
            f"Skip SF: {'Yes' if self.skip_sales_floor_var.get() else 'No'}"
        )

    # ------------------------------
    # Exports
    # ------------------------------
    def do_export_pdf(self):
        if self.moveup_df is None:
            messagebox.showwarning("No data", "Import first.")
            return

        # When hide_removed=True, excluded items were already stripped from moveup_df in
        # _recompute_from_current. We only need to filter here when hide_removed=False
        # (items are visible in the list but should still be omitted from the export).
        if self.excluded_barcodes and not self.hide_removed_var.get():
            mu_use = self.moveup_df[~self.moveup_df["Package Barcode"].astype(str).fillna("").isin(self.excluded_barcodes)].copy()
        else:
            mu_use = self.moveup_df.copy()

        prio_df = self._get_kuntal_priority_df()

        try:
            p = export_moveup_pdf_paginated(
                move_up_df=mu_use,
                priority_df=prio_df,
                base_dir=self.export_run_dir,
                timestamp=self.timestamp_var.get(),
                prefix=self.prefix_var.get() or None,
                auto_open=self.auto_open_var.get(),
                items_per_page=int(self.page_items_var.get() or 30),
                kawaii_pdf=True,
                printer_bw=bool(self.printer_bw_var.get()),
            )
            self.status.set(f"PDF saved: {os.path.basename(p)}")
            if hasattr(self, "dog_widget"):
                self.dog_widget.react_success("PDF exported ✅")
        except Exception as e:
            messagebox.showerror("Export PDF", str(e))
            self.status.set(f"Export error: {e}")
            if hasattr(self, "dog_widget"):
                self.dog_widget.react_error("PDF failed 💥")

    def do_export_xlsx(self):
        if self.moveup_df is None:
            messagebox.showwarning("No data", "Import first.")
            return

        # Same logic as do_export_pdf — only filter here when hide_removed=False.
        if self.excluded_barcodes and not self.hide_removed_var.get():
            mu_use = self.moveup_df[~self.moveup_df["Package Barcode"].astype(str).fillna("").isin(self.excluded_barcodes)].copy()
        else:
            mu_use = self.moveup_df.copy()

        prio_df = self._get_kuntal_priority_df()

        try:
            p = export_excel(
                move_up_df=mu_use,
                priority_df=prio_df,
                base_dir=self.export_run_dir,
                timestamp=self.timestamp_var.get(),
                prefix=self.prefix_var.get() or None,
            )
            self.status.set(f"Excel saved: {os.path.basename(p)}")
            if hasattr(self, "dog_widget"):
                self.dog_widget.react_success("Excel exported ✅")
        except Exception as e:
            messagebox.showerror("Export Excel", str(e))
            self.status.set(f"Export error: {e}")
            if hasattr(self, "dog_widget"):
                self.dog_widget.react_error("Excel failed 💥")


# ------------------------------
# Main
# ------------------------------
def main():
    root = Tk()
    _gui = MoveUpGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()