"""
Standalone config manager for MoveUp.

Owns all file I/O, path resolution, defaults, validation, and atomic-write
logic for moveup_config.json. No Tk dependency for basic usage; accepts an
optional *tk_root* for the messagebox backup-restore prompt.

HOW CONFIG PERSISTENCE WORKS:
=============================
All app settings (filters, excluded barcodes, Bisa stats, etc.) are stored
in a single JSON file: moveup_config.json. This module handles:

  1. DEFAULTS: All 22 config keys have defaults (see DEFAULT_CONFIG below).
     If a key is missing from the file, the default is used silently.

  2. LOADING: On startup, load() reads the JSON and merges it into self.data.
     Type coercion ensures booleans stay booleans, lists stay lists, etc.
     If the config file is missing but a backup exists in ~/.moveup/,
     the user is prompted to restore from backup.

  3. SAVING: save() writes self.data to JSON atomically:
     - First writes to moveup_config.json.tmp
     - Then atomically renames .tmp → .json (via os.replace)
     - This prevents corruption if the app crashes mid-write
     - Also writes a backup copy to ~/.moveup/moveup_config_backup.json

  4. DICT-LIKE ACCESS: cfg["key"] and cfg["key"] = value are shortcuts
     for cfg.data["key"]. Keeps calling code clean.

CONFIG KEYS (see DEFAULT_CONFIG for full list):
  - room_alias_map: user-defined room name aliases
  - selected_rooms/brands/types: filter selections
  - excluded_barcodes: items hidden from the move-up list
  - kuntal_priority_barcodes: items starred as priority
  - lifetime_pets/treats/moveups: Bisa companion stats
  - And more... see DEFAULT_CONFIG below
"""

import json
import os
import sys
from typing import Any, Dict, List, Optional


CONFIG_FILENAME = "moveup_config.json"
BACKUP_DIR_NAME = ".moveup"
BACKUP_FILENAME = "moveup_config_backup.json"


# All 22 config keys with their default values.
# Defaults here match what MoveUpGUI.__init__ historically set before
# _load_config() was called.
DEFAULT_CONFIG: Dict[str, Any] = {
    "room_alias_map":           {},
    "selected_rooms":           [],
    "selected_brands":          [],
    "selected_types":           [],
    "printer_bw":               False,
    "skip_sales_floor":         False,
    "hide_removed":             True,
    "auto_open_pdf":            os.name == "nt",
    "timestamp":                True,
    "items_per_page":           35,
    "prefix":                   "",
    "last_import_dir":          "",
    "current_file_path":        "",
    "excluded_barcodes":        [],
    "kuntal_priority_barcodes": [],
    "active_columns":           [],       # empty → fall back to COLUMNS_TO_USE
    "lifetime_pets":            0,
    "lifetime_treats":          0,
    "lifetime_moveups":         0,
    "bisa_name":                "Bisa",
    "catnip_redeemed":          0,
    "prev_inventory_snapshot":  {},
}


def determine_app_dir() -> str:
    """Return the directory containing the running script / frozen exe."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


class ConfigManager:
    """
    Manages load / save of moveup_config.json.

    Usage::

        cfg = ConfigManager()
        cfg.load()                     # reads JSON into cfg.data
        val = cfg["some_key"]          # shorthand for cfg.data["some_key"]
        cfg["some_key"] = new_val      # shorthand for cfg.data[...] = ...
        cfg.save()                     # atomic write + backup
    """

    def __init__(
        self,
        app_dir: Optional[str] = None,
        tk_root=None,
    ):
        self.app_dir = app_dir or determine_app_dir()
        self.config_path = os.path.join(self.app_dir, CONFIG_FILENAME)

        self._backup_dir = os.path.join(
            os.path.expanduser("~"), BACKUP_DIR_NAME,
        )
        self._backup_config_path = os.path.join(
            self._backup_dir, BACKUP_FILENAME,
        )

        # Optional Tk root for messagebox prompts (backup restore).
        self._tk_root = tk_root

        # The config data dict — initialised to defaults.
        self.data: Dict[str, Any] = dict(DEFAULT_CONFIG)

    # ---- dict-like access shortcuts ----

    def __getitem__(self, key: str) -> Any:
        """
        Return ``self.data[key]``.

        Raises ``KeyError`` if *key* is not in ``self.data``.  All keys in
        ``DEFAULT_CONFIG`` are always present after construction, so a
        ``KeyError`` only occurs if you request a key that was never defined.

        Parameters
        ----------
        key : str
            A config key (e.g. ``"hide_removed"``, ``"excluded_barcodes"``).

        Returns
        -------
        Any
            The stored value for that key.
        """
        return self.data[key]

    def __setitem__(self, key: str, value: Any) -> None:
        """
        Set ``self.data[key] = value`` in memory only.

        This does **not** call ``save()`` — callers must persist explicitly.
        Useful for batching multiple updates before a single ``save()`` call.

        Parameters
        ----------
        key : str
            Config key to update.
        value : Any
            New value to store.  No type checking is performed here; type
            coercion only happens during ``load()`` / ``_apply_loaded()``.
        """
        self.data[key] = value

    def get(self, key: str, default: Any = None) -> Any:
        """
        Return ``self.data.get(key, default)``.

        Safe alternative to ``__getitem__`` when the key might not be present
        (e.g. accessing a key added in a newer version of the app that an
        older config file may not contain).

        Parameters
        ----------
        key : str
            Config key to look up.
        default : Any
            Value to return if *key* is absent.  Defaults to ``None``.

        Returns
        -------
        Any
            The stored value, or *default* if the key is absent.
        """
        return self.data.get(key, default)

    # ---- Load ----

    def load(self, valid_columns: Optional[List[str]] = None) -> None:
        """
        Load config from disk into ``self.data``.

        Reads ``moveup_config.json`` from ``self.app_dir``.  If the primary
        file is missing but a backup exists at ``~/.moveup/moveup_config_backup.json``,
        the user is prompted via ``_prompt_restore()``; if they accept, the
        backup is copied back to the primary location and then loaded.

        All values are merged into ``self.data`` via ``_apply_loaded()``, which
        enforces type coercion (booleans, lists, integers, validated paths) so
        that a corrupt or partial JSON file cannot leave the app in a broken
        state.  Missing keys fall back to ``DEFAULT_CONFIG`` values set in
        ``__init__``.

        Any ``Exception`` during the entire load is caught and printed; the app
        continues with defaults rather than crashing on startup.

        Parameters
        ----------
        valid_columns : list[str] | None
            If provided (pass ``data_core.COLUMNS_TO_USE``), the saved
            ``active_columns`` list is validated against this set.  If any
            saved column name is not in *valid_columns*, the saved value is
            replaced with the full *valid_columns* list.  This prevents stale
            column selections from a previous app version causing KeyErrors
            in the treeview.
        """
        try:
            if not os.path.exists(self.config_path):
                if os.path.exists(self._backup_config_path):
                    if self._prompt_restore():
                        import shutil
                        shutil.copy2(self._backup_config_path, self.config_path)
                        print("[moveup] Restored config from backup.")
                    else:
                        return
                else:
                    return

            with open(self.config_path, "r", encoding="utf-8") as f:
                raw = json.load(f)

            self._apply_loaded(raw, valid_columns)

        except Exception as e:
            print(
                f"[moveup] Warning: could not load config "
                f"({self.config_path}): {e}"
            )

    # ---- Save ----

    def save(self) -> None:
        """
        Persist ``self.data`` to ``moveup_config.json`` atomically.

        Write strategy (crash-safe):
        1. Serialise ``self.data`` as JSON and write to ``moveup_config.json.tmp``.
        2. Atomically rename ``.tmp`` → ``.json`` via ``os.replace()``.
           On Windows and POSIX this is an atomic file-system operation — if
           the app crashes mid-write the ``.tmp`` is left behind but the
           previous ``.json`` is never corrupted.
        3. Attempt to write the same data to
           ``~/.moveup/moveup_config_backup.json`` using the same two-step
           pattern.  Backup failure is non-critical and only prints a warning.

        Any ``Exception`` during the primary write is caught and printed; the
        backup failure is caught separately so a permission error on the home
        directory never prevents the primary save.
        """
        try:
            tmp = self.config_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=2)
            os.replace(tmp, self.config_path)

            # Backup to ~/.moveup/
            try:
                os.makedirs(self._backup_dir, exist_ok=True)
                tmp_bk = self._backup_config_path + ".tmp"
                with open(tmp_bk, "w", encoding="utf-8") as f:
                    json.dump(self.data, f, indent=2)
                os.replace(tmp_bk, self._backup_config_path)
            except OSError as e:
                print(f"[moveup] Backup write failed (non-critical): {e}")

        except Exception as e:
            print(
                f"[moveup] Warning: could not save config "
                f"({self.config_path}): {e}"
            )

    # ---- Internal helpers ----

    def _prompt_restore(self) -> bool:
        """
        Show a Tk messagebox asking whether to restore config from backup.

        Only called from ``load()`` when the primary config file is missing but
        a backup file exists.  Requires ``self._tk_root`` to be set (the
        MoveUpGUI root window); if it is ``None`` (headless or pre-Tk context)
        the method returns ``False`` silently so the app starts with defaults.

        Returns
        -------
        bool
            ``True`` if the user clicked "Yes" (restore), ``False`` if they
            clicked "No" (start fresh) or no Tk root is available.
        """
        if self._tk_root is not None:
            from tkinter import messagebox
            return messagebox.askyesno(
                "Config Not Found",
                "Your main config file is missing, but a backup was found.\n\n"
                "Would you like to restore from the backup?\n\n"
                "(Choose 'No' to start fresh.)",
            )
        # No UI available — start fresh silently.
        return False

    def _apply_loaded(
        self, raw: dict, valid_columns: Optional[List[str]]
    ) -> None:
        """
        Merge *raw* JSON data into ``self.data`` with strict type coercion.

        Called only from ``load()``.  Every value is coerced to its expected
        Python type so that a hand-edited or partially corrupt JSON file cannot
        put the app in an inconsistent state:

        - **Dicts** (``room_alias_map``, ``prev_inventory_snapshot``): wrapped in
          ``dict()`` so a ``None`` or missing value becomes an empty dict.
        - **Lists of strings** (``selected_rooms/brands/types``,
          ``excluded_barcodes``, ``kuntal_priority_barcodes``): wrapped in
          ``list()``; ``None`` or missing becomes ``[]``.
        - **Booleans** (``printer_bw``, ``skip_sales_floor``, etc.): explicitly
          cast with ``bool()`` so JSON ``0``/``1`` or string ``"true"`` are
          handled consistently.
        - **``active_columns``**: validated against *valid_columns* if supplied.
          Any column name that no longer exists in the app schema causes the
          entire saved list to be replaced with the current *valid_columns*.
        - **Integers** (``items_per_page``, Bisa stats): wrapped in a
          ``try/except (TypeError, ValueError)`` so a string like ``"35"``
          is coerced but a completely invalid value leaves the default intact.
        - **Strings** (``prefix``, ``bisa_name``): cast with ``str()``.
        - **Validated paths** (``last_import_dir``, ``current_file_path``):
          checked with ``os.path.isdir`` / ``os.path.isfile`` so stale paths
          from a previous machine or renamed folder are silently discarded.

        Parameters
        ----------
        raw : dict
            The raw dict decoded from the JSON config file.
        valid_columns : list[str] | None
            Current app column names for ``active_columns`` validation.
        """
        d = self.data

        # Dicts
        d["room_alias_map"] = dict(raw.get("room_alias_map", {}) or {})

        # Lists of strings
        d["selected_rooms"]  = list(raw.get("selected_rooms", []) or [])
        d["selected_brands"] = list(raw.get("selected_brands", []) or [])
        d["selected_types"]  = list(raw.get("selected_types", []) or [])

        # Booleans
        for key in (
            "printer_bw", "skip_sales_floor", "hide_removed",
            "auto_open_pdf", "timestamp",
        ):
            if key in raw:
                d[key] = bool(raw[key])

        # Sets stored as sorted lists
        d["excluded_barcodes"] = list(
            raw.get("excluded_barcodes", []) or []
        )
        d["kuntal_priority_barcodes"] = list(
            raw.get("kuntal_priority_barcodes", []) or []
        )

        # active_columns with validation
        saved_cols = raw.get("active_columns", [])
        if (
            valid_columns
            and saved_cols
            and all(c in valid_columns for c in saved_cols)
        ):
            d["active_columns"] = list(saved_cols)
        elif valid_columns:
            d["active_columns"] = list(valid_columns)

        # Integer
        try:
            d["items_per_page"] = int(
                raw.get("items_per_page", d["items_per_page"])
            )
        except (TypeError, ValueError):
            pass

        # String
        d["prefix"] = str(raw.get("prefix", d["prefix"]) or "")

        # Validated paths
        last_dir = raw.get("last_import_dir")
        if isinstance(last_dir, str) and last_dir.strip() and os.path.isdir(last_dir):
            d["last_import_dir"] = last_dir.strip()

        last_file = raw.get("current_file_path")
        if (
            isinstance(last_file, str)
            and last_file.strip()
            and os.path.isfile(last_file)
        ):
            d["current_file_path"] = last_file.strip()

        # Bisa stats (integers)
        for _key in ("lifetime_pets", "lifetime_treats", "lifetime_moveups", "catnip_redeemed"):
            try:
                d[_key] = int(raw.get(_key, 0))
            except (TypeError, ValueError):
                pass
        d["bisa_name"] = str(raw.get("bisa_name", "Bisa")) or "Bisa"

        # Inventory snapshot
        snap = raw.get("prev_inventory_snapshot", {})
        if isinstance(snap, dict):
            d["prev_inventory_snapshot"] = snap
