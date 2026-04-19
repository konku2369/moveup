"""
Velocity history storage for MoveUp.

Manages velocity_history.json — stores snapshots of inventory state
across successive imports for velocity/movement tracking.
No Tk dependency.

STORAGE FORMAT (velocity_history.json):
=======================================
{
  "version": 1,
  "snapshots": [
    {
      "timestamp": "2026-03-20T14:30:00.123456",
      "file_name": "Inventory_Export_March_20.xlsx",
      "entries": [
        {"barcode": "1A406030...", "room": "backstock", "qty": 5, "received_date": "2026-01-15"},
        ...
      ]
    },
    ...
  ]
}

Each snapshot is appended when the user imports a file. The history is used
by data_core.compute_velocity_metrics() to detect movement patterns.

PERSISTENCE:
  - Primary: velocity_history.json in the app directory
  - Backup: ~/.moveup/velocity_history_backup.json
  - Writes are atomic (.tmp → os.replace) to prevent corruption on crash
"""

import json
import os
import sys
from typing import Any, Dict, List, Optional


VELOCITY_FILENAME = "velocity_history.json"
VELOCITY_BACKUP_FILENAME = "velocity_history_backup.json"
BACKUP_DIR_NAME = ".moveup"


def _determine_app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


class VelocityHistoryManager:
    """
    Manages load / save of velocity_history.json.

    Each snapshot captures the full inventory state at import time:
    {timestamp, file_name, entries: [{barcode, room, qty, received_date}]}
    """

    def __init__(self, app_dir: Optional[str] = None):
        """
        Initialise the manager without reading from disk.

        Call ``load()`` explicitly to populate ``self.snapshots`` from the
        persisted JSON file.

        Parameters
        ----------
        app_dir : str | None
            Directory containing (or to contain) ``velocity_history.json``.
            Defaults to the directory of the running script or executable.
        """
        self.app_dir = app_dir or _determine_app_dir()
        self.history_path = os.path.join(self.app_dir, VELOCITY_FILENAME)

        self._backup_dir = os.path.join(
            os.path.expanduser("~"), BACKUP_DIR_NAME,
        )
        self._backup_path = os.path.join(
            self._backup_dir, VELOCITY_BACKUP_FILENAME,
        )

        self.snapshots: List[Dict[str, Any]] = []

    def load(self) -> None:
        """
        Load velocity history from ``velocity_history.json`` into ``self.snapshots``.

        Handles two storage formats:
        - **Dict format** (current): ``{"version": 1, "snapshots": [...]}``
        - **List format** (legacy): bare JSON array used in early development

        If the file is absent, ``self.snapshots`` stays empty — normal on first
        launch.  Any ``json.JSONDecodeError``, ``OSError``, ``KeyError``, or
        ``TypeError`` is caught and printed; the app continues with empty
        history rather than crashing.
        """
        try:
            if not os.path.exists(self.history_path):
                return
            with open(self.history_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            # Support both formats for backward compatibility:
            # v1: {"version": 1, "snapshots": [...]}  (current)
            # legacy: bare list [...]                   (early development)
            if isinstance(raw, dict) and isinstance(raw.get("snapshots"), list):
                self.snapshots = raw["snapshots"]
            elif isinstance(raw, list):
                self.snapshots = raw
        except (json.JSONDecodeError, OSError, KeyError, TypeError) as e:
            print(f"[moveup] Warning: could not load velocity history: {e}")

    def save(self) -> None:
        """
        Persist ``self.snapshots`` to ``velocity_history.json`` atomically.

        Always writes the current dict format: ``{"version": 1, "snapshots": [...]}``.

        Write strategy (crash-safe):
        1. Write to ``velocity_history.json.tmp``.
        2. Atomically rename ``.tmp`` → ``.json`` via ``os.replace()``.
        3. Attempt a backup write to ``~/.moveup/velocity_history_backup.json``
           using the same pattern.  Backup ``OSError`` is non-critical.

        Any unexpected ``Exception`` from the primary write is caught and
        printed; ``self.snapshots`` is never modified by this method.
        """
        try:
            data = {"version": 1, "snapshots": self.snapshots}
            tmp = self.history_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=1)
            os.replace(tmp, self.history_path)  # atomic: if write fails, old file is intact

            # Backup to ~/.moveup/
            try:
                os.makedirs(self._backup_dir, exist_ok=True)
                tmp_bk = self._backup_path + ".tmp"
                with open(tmp_bk, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=1)
                os.replace(tmp_bk, self._backup_path)
            except OSError as e_bk:
                print(f"[moveup] Velocity backup write failed (non-critical): {e_bk}")

        except Exception as e:
            print(f"[moveup] Warning: could not save velocity history: {e}")

    def add_snapshot(
        self, timestamp: str, file_name: str, entries: List[Dict[str, Any]]
    ) -> None:
        """
        Append one inventory snapshot and immediately persist to disk.

        Called by ``main.py`` after each successful import.  The snapshot
        records the full per-barcode inventory state at that point in time so
        ``data_core.compute_velocity_metrics()`` can compare successive imports
        and produce movement scores (room changes, qty delta, sell rate, etc.).

        Parameters
        ----------
        timestamp : str
            ISO-8601 datetime string (``datetime.now().isoformat()``), used as
            the snapshot's primary sort key for ``purge_before()`` comparisons.
        file_name : str
            Base name of the imported file (e.g. ``"Inventory_Export.xlsx"``).
        entries : list[dict]
            Per-barcode inventory state.  Each dict should contain at minimum:
            ``{"barcode": str, "room": str, "qty": int, "received_date": str}``.
            The list is stored verbatim; extra keys are preserved and ignored.
        """
        self.snapshots.append({
            "timestamp": timestamp,
            "file_name": file_name,
            "entries": entries,
        })
        self.save()

    def purge_before(self, cutoff_timestamp: str) -> int:
        """
        Remove all snapshots with a timestamp earlier than *cutoff_timestamp*.

        Comparison is lexicographic string comparison, which is correct for
        ISO-8601 timestamps (``"2026-03-20T14:30:00.123456"``).  Snapshots at
        exactly the cutoff are kept.

        If any snapshots are removed the updated list is immediately persisted.

        Parameters
        ----------
        cutoff_timestamp : str
            ISO-8601 datetime string.  Snapshots with ``"timestamp" < cutoff``
            are deleted.

        Returns
        -------
        int
            Number of snapshots removed.
        """
        before = len(self.snapshots)
        self.snapshots = [
            s for s in self.snapshots
            if s.get("timestamp", "") >= cutoff_timestamp
        ]
        removed = before - len(self.snapshots)
        if removed:
            self.save()
        return removed

    def purge_all(self) -> int:
        """
        Remove all velocity snapshots and persist the empty state.

        Calls ``save()`` immediately if there was anything to remove.

        Returns
        -------
        int
            Number of snapshots removed; 0 if the history was already empty.
        """
        removed = len(self.snapshots)
        self.snapshots = []
        if removed:
            self.save()
        return removed

    def get_snapshots(self) -> List[Dict[str, Any]]:
        """
        Return a shallow copy of all velocity snapshots.

        Returns a new list so callers cannot accidentally mutate
        ``self.snapshots``.  The dicts inside are not deep-copied.

        Returns
        -------
        list[dict]
            All stored snapshots, oldest first (insertion order).
        """
        return list(self.snapshots)

    def snapshot_count(self) -> int:
        """
        Return the number of stored velocity snapshots.

        Returns
        -------
        int
            ``len(self.snapshots)``; 0 before ``load()`` or after ``purge_all()``.
        """
        return len(self.snapshots)
