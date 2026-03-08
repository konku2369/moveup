"""
Standalone config manager for MoveUp.

Owns all file I/O, path resolution, defaults, validation, and atomic-write
logic for moveup_config.json.  No Tk dependency for basic usage; accepts an
optional *tk_root* for the messagebox backup-restore prompt.
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
        return self.data[key]

    def __setitem__(self, key: str, value: Any) -> None:
        self.data[key] = value

    def get(self, key: str, default: Any = None) -> Any:
        return self.data.get(key, default)

    # ---- Load ----

    def load(self, valid_columns: Optional[List[str]] = None) -> None:
        """
        Load config from disk into *self.data*.

        Parameters
        ----------
        valid_columns : list[str] | None
            If provided, used to validate the saved ``active_columns``.
            Pass ``data_core.COLUMNS_TO_USE`` here.
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
        """Atomic write to primary config file + non-critical backup."""
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
            except Exception:
                pass  # backup failure is non-critical

        except Exception as e:
            print(
                f"[moveup] Warning: could not save config "
                f"({self.config_path}): {e}"
            )

    # ---- Internal helpers ----

    def _prompt_restore(self) -> bool:
        """Ask the user whether to restore from backup.  Returns True → yes."""
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
        """Merge a raw JSON dict into *self.data* with type coercion."""
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
        d["lifetime_pets"]    = int(raw.get("lifetime_pets", 0))
        d["lifetime_treats"]  = int(raw.get("lifetime_treats", 0))
        d["lifetime_moveups"] = int(raw.get("lifetime_moveups", 0))
        d["bisa_name"]        = str(raw.get("bisa_name", "Bisa")) or "Bisa"
        d["catnip_redeemed"]  = int(raw.get("catnip_redeemed", 0))

        # Inventory snapshot
        snap = raw.get("prev_inventory_snapshot", {})
        if isinstance(snap, dict):
            d["prev_inventory_snapshot"] = snap
