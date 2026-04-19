"""
Import history storage for MoveUp.

Manages import_history.json — stores aggregate statistics for each
METRC file import (row counts, move-up counts, brand/type/room counts).
No Tk dependency.

STORAGE FORMAT (import_history.json):
=====================================
{
  "version": 1,
  "entries": [
    {
      "timestamp": "2026-03-20T14:30:00.123456",
      "file_name": "Inventory_Export_March_20.xlsx",
      "total_rows": 1250,
      "mapped_rows": 1230,
      "moveup_count": 87,
      "diag": { ... },
      "unique_brands": 45,
      "unique_types": 12,
      "unique_rooms": 6,
      "total_qty": 5840,
      "velocity_entries_count": 1230,
      "bisa_moveups": 3
    },
    ...
  ]
}

Each entry is appended when the user imports a file.  The diag dict is
stored verbatim from compute_moveup_from_df() so future diagnostic keys
are automatically captured.

PERSISTENCE:
  - Primary: import_history.json in the app directory
  - Backup: ~/.moveup/import_history_backup.json
  - Writes are atomic (.tmp -> os.replace) to prevent corruption on crash
"""

import json
import os
import sys
from typing import Any, Dict, List, Optional


IMPORT_HISTORY_FILENAME = "import_history.json"
IMPORT_HISTORY_BACKUP_FILENAME = "import_history_backup.json"
BACKUP_DIR_NAME = ".moveup"


def _determine_app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


class ImportHistoryManager:
    """
    Manages load / save of import_history.json.

    Each entry captures aggregate stats for one import:
    {timestamp, file_name, total_rows, mapped_rows, moveup_count, diag, ...}
    """

    def __init__(self, app_dir: Optional[str] = None):
        """
        Initialise the manager but do not read from disk yet.

        Call ``load()`` explicitly after construction to populate
        ``self.entries`` from the persisted JSON file.

        Parameters
        ----------
        app_dir : str | None
            Directory that contains (or will contain) ``import_history.json``.
            If ``None``, defaults to the directory of the running script or
            executable (same convention as ``ConfigManager``).
        """
        self.app_dir = app_dir or _determine_app_dir()
        self.history_path = os.path.join(self.app_dir, IMPORT_HISTORY_FILENAME)

        self._backup_dir = os.path.join(
            os.path.expanduser("~"), BACKUP_DIR_NAME,
        )
        self._backup_path = os.path.join(
            self._backup_dir, IMPORT_HISTORY_BACKUP_FILENAME,
        )

        self.entries: List[Dict[str, Any]] = []

    def load(self) -> None:
        """
        Load import history from ``import_history.json`` into ``self.entries``.

        Handles two legacy storage formats gracefully:
        - **Dict format** (current): ``{"version": 1, "entries": [...]}`` — the
          ``entries`` list is extracted directly.
        - **List format** (legacy): the raw JSON array is used as-is.

        If the file does not exist, ``self.entries`` remains an empty list and
        no error is raised — this is normal on first launch.

        Any ``json.JSONDecodeError``, ``OSError``, ``KeyError``, or
        ``TypeError`` is caught and printed; the app continues with an empty
        history rather than crashing.
        """
        try:
            if not os.path.exists(self.history_path):
                return
            with open(self.history_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if isinstance(raw, dict) and isinstance(raw.get("entries"), list):
                self.entries = raw["entries"]
            elif isinstance(raw, list):
                self.entries = raw
        except (json.JSONDecodeError, OSError, KeyError, TypeError) as e:
            print(f"[moveup] Warning: could not load import history: {e}")

    def save(self) -> None:
        """
        Persist ``self.entries`` to ``import_history.json`` atomically.

        Always writes in the current dict format:
        ``{"version": 1, "entries": [...]}``

        Write strategy (crash-safe):
        1. Write to ``import_history.json.tmp``.
        2. Atomically rename ``.tmp`` → ``.json`` via ``os.replace()``.
        3. Attempt a second write to ``~/.moveup/import_history_backup.json``
           using the same pattern.  Backup failure is non-critical (``OSError``
           only) and does not affect the primary save.

        Any unexpected ``Exception`` during the primary write is caught and
        printed; the in-memory ``self.entries`` list is never modified here.
        """
        try:
            data = {"version": 1, "entries": self.entries}
            tmp = self.history_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=1)
            os.replace(tmp, self.history_path)

            try:
                os.makedirs(self._backup_dir, exist_ok=True)
                tmp_bk = self._backup_path + ".tmp"
                with open(tmp_bk, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=1)
                os.replace(tmp_bk, self._backup_path)
            except OSError as e_bk:
                print(f"[moveup] Import history backup write failed (non-critical): {e_bk}")

        except Exception as e:
            print(f"[moveup] Warning: could not save import history: {e}")

    def add_entry(
        self,
        timestamp: str,
        file_name: str,
        total_rows: int,
        mapped_rows: int,
        moveup_count: int,
        diag: Dict[str, int],
        unique_brands: int,
        unique_types: int,
        unique_rooms: int,
        total_qty: int,
        velocity_entries_count: int,
        bisa_moveups: int,
    ) -> None:
        """
        Append one import summary entry to ``self.entries`` and persist to disk.

        Called by ``main.py`` after each successful file import.  The entry
        captures aggregate statistics for the import so the Import History
        window (``mainImportHistory.py``) can show a per-file trend over time.

        Parameters
        ----------
        timestamp : str
            ISO-8601 datetime string (``datetime.now().isoformat()``), used as
            the entry's primary sort key and for ``purge_before()`` comparisons.
        file_name : str
            Base name of the imported file (e.g. ``"Inventory_Export.xlsx"``).
        total_rows : int
            Row count of the raw DataFrame before column mapping and filtering.
        mapped_rows : int
            Row count after column mapping succeeded (junk/footer rows removed).
        moveup_count : int
            Number of items in the final move-up result list.
        diag : dict[str, int]
            The diagnostic dict returned by ``compute_moveup_from_df()``.
            Stored verbatim so future diagnostic keys are automatically captured
            without schema changes to this module.
        unique_brands : int
            Number of distinct brands seen in the mapped data.
        unique_types : int
            Number of distinct product types seen in the mapped data.
        unique_rooms : int
            Number of distinct (alias-normalised) rooms seen in the mapped data.
        total_qty : int
            Sum of the ``Qty On Hand`` column across all mapped rows.
        velocity_entries_count : int
            Number of rows written to the velocity snapshot for this import.
        bisa_moveups : int
            The value of ``lifetime_moveups`` from Bisa's stats at import time,
            used to show Bisa's cumulative milestone counter in history.
        """
        self.entries.append({
            "timestamp": timestamp,
            "file_name": file_name,
            "total_rows": total_rows,
            "mapped_rows": mapped_rows,
            "moveup_count": moveup_count,
            "diag": dict(diag),
            "unique_brands": unique_brands,
            "unique_types": unique_types,
            "unique_rooms": unique_rooms,
            "total_qty": total_qty,
            "velocity_entries_count": velocity_entries_count,
            "bisa_moveups": bisa_moveups,
        })
        self.save()

    def get_entries(self) -> List[Dict[str, Any]]:
        """
        Return a shallow copy of all import history entries.

        Returns a new list so callers cannot accidentally mutate
        ``self.entries`` by modifying the returned list.  The dicts inside
        the list are not deep-copied — do not mutate individual entry dicts
        in place unless you intend to affect the in-memory state.

        Returns
        -------
        list[dict]
            All stored import history entries, oldest first (insertion order).
        """
        return list(self.entries)

    def entry_count(self) -> int:
        """
        Return the number of stored import history entries.

        Returns
        -------
        int
            ``len(self.entries)``; 0 when no history has been recorded or
            after a ``purge_all()`` call.
        """
        return len(self.entries)

    def purge_before(self, cutoff_timestamp: str) -> int:
        """
        Remove all entries with a timestamp earlier than *cutoff_timestamp*.

        Comparison is performed as a plain string comparison, which is correct
        because timestamps are stored as ISO-8601 strings
        (``"2026-03-20T14:30:00.123456"``) — lexicographic order matches
        chronological order for this format.

        If any entries are removed the updated list is immediately persisted
        via ``save()``.

        Parameters
        ----------
        cutoff_timestamp : str
            ISO-8601 datetime string.  Entries whose ``"timestamp"`` field is
            strictly less than this value are removed.  Entries at exactly
            the cutoff are kept.

        Returns
        -------
        int
            Number of entries that were removed.
        """
        before = len(self.entries)
        self.entries = [
            e for e in self.entries
            if e.get("timestamp", "") >= cutoff_timestamp
        ]
        removed = before - len(self.entries)
        if removed:
            self.save()
        return removed

    def purge_all(self) -> int:
        """
        Remove all import history entries and persist the empty list.

        Immediately calls ``save()`` if there was anything to remove, so the
        cleared state is written to disk even if the app exits abnormally
        afterward.

        Returns
        -------
        int
            Number of entries that were removed (0 if history was already empty).
        """
        removed = len(self.entries)
        self.entries = []
        if removed:
            self.save()
        return removed
