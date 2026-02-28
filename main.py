import os
import sys
import json
import subprocess
from datetime import datetime
from typing import Dict, List, Optional

import pandas as pd

# GUI
from tkinter import (
    Tk, Toplevel, StringVar, IntVar, BooleanVar, filedialog, messagebox,
    ttk
)
from tkinter import Listbox, MULTIPLE, END
import tkinter as tk

# PDF exports live in pdf_export.py
from pdf_export import export_moveup_pdf_paginated, export_audit_pdfs

# Core logic
from data_core import (
    APP_VERSION,
    COLUMNS_TO_USE,
    AUDIT_OPTIONAL_FIELDS,
    TYPE_TRUNC_LEN,
    SALES_FLOOR_ALIASES,
    load_raw_df,
    automap_columns,
    compute_moveup_from_df,
    normalize_rooms,
    detect_metrc_source_column,
    sort_with_backstock_priority,
    ellipses,
    sanitize_prefix,
    aggregate_split_packages_by_room,
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
        prio_bcs = set(prio["Package Barcode"].astype(str).str.strip().tolist())
        mu = mu[~mu["Package Barcode"].astype(str).str.strip().isin(prio_bcs)].copy()

    prio = sort_with_backstock_priority(prio) if not prio.empty else prio
    mu = sort_with_backstock_priority(mu) if not mu.empty else mu

    with pd.ExcelWriter(out, engine="openpyxl") as w:
        if not prio.empty:
            prio.to_excel(w, sheet_name="Kuntal_Priority", index=False)
        mu.to_excel(w, sheet_name="Move_Up_Items", index=False)

    return out


# ------------------------------
# Filters UI helper
# ------------------------------
class _FilterList:
    def __init__(self, parent, title: str):
        self.title = title
        self.all_items: List[str] = []
        self.filtered_idx: List[int] = []

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

        self.search_var.trace_add("write", lambda *_: self.refresh())

    def set_items(self, items: List[str]):
        self.all_items = list(items or [])
        self.refresh()

    def refresh(self):
        q = (self.search_var.get() or "").strip().lower()
        selected_vals = set(self.get_selected_values())
        self.lb.delete(0, END)

        if not q:
            self.filtered_idx = list(range(len(self.all_items)))
        else:
            tokens = q.split()

            def match(i: int) -> bool:
                s = self.all_items[i].lower()
                return all(t in s for t in tokens)

            self.filtered_idx = [i for i in range(len(self.all_items)) if match(i)]

        for i in self.filtered_idx:
            self.lb.insert(END, self.all_items[i])

        for pos, i in enumerate(self.filtered_idx):
            if self.all_items[i] in selected_vals:
                self.lb.selection_set(pos)

    def select_all(self):
        self.lb.selection_set(0, END)

    def clear_selection(self):
        self.lb.selection_clear(0, END)

    def set_selected_values(self, values: List[str]):
        values_set = set(values or [])
        self.lb.selection_clear(0, END)
        self.search_var.set("")
        self.refresh()
        for pos, i in enumerate(self.filtered_idx):
            if self.all_items[i] in values_set:
                self.lb.selection_set(pos)

    def get_selected_values(self) -> List[str]:
        sel = list(self.lb.curselection())
        out = []
        for pos in sel:
            if pos < 0 or pos >= len(self.filtered_idx):
                continue
            i = self.filtered_idx[pos]
            out.append(self.all_items[i])
        return out



# ------------------------------
# ASCII Dog Widget
# Bisa — earthmed-style husky
# ------------------------------
class AsciiDogWidget:
    """Animated ASCII dog widget (Bisa) with expanded behaviors.

    ✅ Keeps all previous features/APIs:
      - click dog to pet (receive_pet)
      - click box/blank space to throw treat (throw_treat_at_window_x / frame click)
      - stats counter (pets/treats)
      - react_* methods used by the app

    ➕ Adds:
      - more idle micro-animations (wag, blink, sleep, zoomies)
      - contextual reactions (success/warning/error)
      - seasonal theme accents (Oct/Dec)
      - rare "legendary" easter egg (1% chance)
    """

    # ------------------------------
    # Existing frames
    # ------------------------------
    IDLE_FRAMES = [
        "  /\\_/\\  \n ( o.o ) \n  > ^ <",
        "  /\\_/\\  \n ( o.o ) \n  > ^ <",
        "  /\\_/\\  \n ( ^.^ ) \n  > ^ <",
        "  /\\_/\\  \n ( O.O ) \n  > W <",
    ]
    PET_FRAMES = [
        "  /\\_/\\  \n ( u.u )♥\n  > ^ <",
        "  /\\_/\\  \n ( ^w^ )♥\n  > ^ <",
        "  /\\_/\\  \n ( ^.^ )♥\n  > ^~<",
        "  /\\_/\\  \n ( u.u ) \n  > ^ <",
    ]
    TREAT_SHORT = [
        "  /\\_/\\    🦴\n ( o.o )  \n  > ^ <",
        "    /\\_/\\ 🦴\n   ( ^.^) \n    > ^ <",
    ]
    TREAT_MEDIUM = [
        "  /\\_/\\      🦴\n ( o.o )    \n  > ^ <",
        "    /\\_/\\  🦴\n   ( ^o^ ) \n    > ^ <",
        "      /\\_/\\🦴\n     ( ^.^)\n      > ^ <",
    ]
    TREAT_FAR = [
        "  /\\_/\\        🦴\n ( o.o )      \n  > ^ <",
        "    /\\_/\\    🦴\n   ( ^o^ )   \n    > ^ <",
        "      /\\_/\\ 🦴\n     ( ^.^) \n      > ^ <",
        "        /\\_/\\🦴\n       ( ^O^)\n        > ^ <",
    ]
    RUN_BACK = [
        "      /\\_/\\  \n     🦴(^.^) \n      > ^ <",
        "    /\\_/\\    \n   🦴( ^w^)  \n    > ^ <",
        "  /\\_/\\      \n 🦴( ^.^)   \n  > ^ <",
    ]
    HAPPY_FRAMES = [
        "    /\\_/\\   \n   ( ^o^)o  \n    > ^ <",
        "      /\\_/\\ \n     ( ^o^)o\n      > ^ <",
        "    /\\_/\\   \n   o(^w^)   \n    > ^ <",
        "    /\\_/\\   \n   ( ^.^)o  \n    >w^ <",
        "      /\\_/\\ \n     o(^v^) \n      > ^ <",
        "    /\\_/\\   \n   \\(^o^)/o \n    > ^ <",
    ]
    LOAD_FRAMES = [
        "  /\\_/\\   \n ( O.O )! \n  > W <",
        "  /\\_/\\   \n ( ^o^)!! \n  >w^ <",
        "   /\\_/\\  \n  (\\^o^/) \n   >W< ",
        "  /\\_/\\   \n  (^o^)/  \n  > ^ <",
        "  /\\_/\\   \n \\(^w^)/  \n  >w^ <",
        "  /\\_/\\   \n  ( ^.^)~ \n  > ^ <",
    ]
    EXCLUDED_FRAMES = [
        "  /\\_/\\  \n ( o.o ) \n  > ^ <",
        "  /\\_/\\  \n ( ;.; ) \n  > ^ <",
        "  /\\_/\\  \n ( T.T ) \n  > ^ <",
        "  /\\_/\\  \n ( u.u ) \n  > ^ <",
    ]
    ALERT_FRAMES = [
        "  /|_|\\  \n ( o.o ) \n  > ^ <",
        "  /|_|\\  \n ( O.O ) \n  > ! <",
        "  /|_|\\  \n ( ^.^ ) \n  > ^ <",
        "  /\\_/\\  \n ( ^.^ ) \n  > ^ <",
    ]
    SNIFF_FRAMES = [
        "  /\\_/\\  \n ( o.o ) \n  > ^ <",
        "  /\\_/\\  \n ( o.~ ) \n  >sniff",
        "  /\\_/\\  \n ( ^.o ) \n  >sniff",
        "  /\\_/\\  \n ( o.o ) \n  > ^ <",
    ]
    KUNTAL_FRAMES = [
        "  /\\_/\\  \n ( ^o^)★\n  > ^ <",
        "   /\\_/\\ \n  (★^o^)\n   > ^ <",
        "  /\\_/\\  \n  (^w^)★\n  >w^ <",
        "  /\\_/\\  \n \\(^o^)/ \n  > ^ <",
        "  /\\_/\\  \n  ( ^.^)★\n  > ^ <",
    ]
    STRETCH_FRAMES = [
        "  /\\_/\\  \n ( -.- ) \n  > ^ <",
        "  /\\_/\\  \n ( o.o ) \n  >str<",
        "  /\\_/~  \n ( o.o ) \n  > ^ <",
        "  /\\_/\\  \n ( ^.^ ) \n  > ^ <",
    ]
    CLEARED_FRAMES = [
        "  /\\_/\\  \n ( o.o ) \n  > ^ <",
        "  /\\_/\\  \n ( -.- ) \n  > ^ <",
        "  /\\_/\\  \n ( u.u ) \n  > ^ <",
        "  /\\_/\\  \n ( ^.^ ) \n  > ^ <",
    ]

    # ------------------------------
    # New frames (added)
    # ------------------------------
    WAG_FRAMES = [
        "  /\\_/\\  ~\n ( ^.^ )  \n  > ^ <",
        "~  /\\_/\\  \n  ( ^.^ ) \n   > ^ <",
        "  /\\_/\\  ~\n ( ^w^ )  \n  > ^ <",
        "~  /\\_/\\  \n  ( ^w^ ) \n   > ^ <",
    ]
    BLINK_FRAMES = [
        "  /\\_/\\  \n ( o.o ) \n  > ^ <",
        "  /\\_/\\  \n ( -.- ) \n  > ^ <",
        "  /\\_/\\  \n ( o.o ) \n  > ^ <",
    ]
    SLEEP_FRAMES = [
        "  /\\_/\\  \n ( -.- ) zZ\n  > ^ <",
        "  /\\_/\\  \n ( -.- ) Zz\n  > ^ <",
        "  /\\_/\\  \n ( -.- ) zz\n  > ^ <",
        "  /\\_/\\  \n ( u.u ) zZ\n  > ^ <",
    ]
    ZOOMIES_FRAMES = [
        "  /\\_/\\      \n ( ^o^ )  ⚡\n  > ^ <      ",
        "      /\\_/\\  \n ⚡ ( ^o^ ) \n      > ^ <  ",
        "  /\\_/\\      \n ( ^w^ )  ⚡\n  > ^ <      ",
        "    /\\_/\\    \n ⚡ ( ^w^ ) \n    > ^ <    ",
    ]
    CONFUSED_FRAMES = [
        "  /\\_/\\  \n ( o.o ) ?\n  > ^ <",
        "  /\\_/\\  \n ( O.o ) ?\n  > ^ <",
        "  /\\_/\\  \n ( o.O ) ?\n  > ^ <",
    ]
    SUCCESS_FRAMES = [
        "    /\\_/\\   \n   ( ^o^)✨ \n    > ^ <",
        "      /\\_/\\ \n     ( ^w^)✨\n      > ^ <",
        "    /\\_/\\   \n   \\(^o^)/✨\n    > ^ <",
    ]
    WARNING_FRAMES = [
        "  /|_|\\  \n ( O.O ) !\n  > ! <",
        "  /|_|\\  \n ( o.o ) !\n  > ! <",
        "  /\\_/\\  \n ( o.o ) !\n  > ! <",
    ]
    LEGENDARY_FRAMES = [
        "  /\\_/\\   ★★★\n ( ✧o✧ )  ★\n  > W <   ★",
        "  /\\_/\\   ★★★\n ( ✧w✧ )  ★\n  > W <   ★",
        "  /\\_/\\   ★★★\n ( ✧.^✧ ) ★\n  > W <   ★",
        "  /\\_/\\   ★★★\n ( ✧o✧ )  ★\n  > W <   ★",
    ]

    HALLOWEEN_FRAMES = [
        "  /\\_/\\   🎃\n ( o.o )  \n  > ^ <",
        "  /\\_/\\   🎃\n ( O.O )  \n  > W <",
        "  /\\_/\\   👻\n ( ^.^ )  \n  > ^ <",
    ]
    WINTER_FRAMES = [
        "  /\\_/\\   ❄️\n ( o.o )  \n  > ^ <",
        "  /\\_/\\   ❄️\n ( ^.^ )  \n  > ^ <",
        "  /\\_/\\   ☃️\n ( u.u )  \n  > ^ <",
    ]

    MESSAGES = {
        "idle":     "...",
        "pet":      "so nice~ ♥",
        "treat":    "treat?? 🦴",
        "running":  "nom nom! 🦴",
        "happy":    "yay!!!! ✨",
        "loaded":   "new data!! 📋",
        "excluded": "oh no... 😢",
        "sniff":    "sniff sniff...",
        "alert":    "! what's that?",
        "kuntal":   "ooh priority! ★",
        "stretch":  "zzz... yawn~",
        "cleared":  "phew~ clean!",
        "restored": "yay, back!! ✅",
        "wag":      "tail wag!!",
        "blink":    "blink~",
        "sleep":    "zzz…",
        "zoomies":  "ZOOMIES!! ⚡",
        "confused": "huh?",
        "success":  "nice!! ✅",
        "warning":  "uh oh… ⚠️",
        "error":    "nope… 💥",
        "legendary": "LEGENDARY BISAAAA ★★★",
        "halloween": "spooky Bisa 🎃",
        "winter":   "brr… ❄️",
    }

    def __init__(self, parent: tk.Widget):
        import random
        from datetime import datetime

        self.parent = parent
        self._state = "idle"
        self._after_id = None
        self._idle_idx = 0
        self._anim_idx = 0
        self._anim_frames = []
        self._total_pets = 0
        self._total_treats = 0

        # Animation tuning
        self._speed_scale = 1.0
        self._legendary_chance = 0.01

        # Theme
        self._apply_seasonal_theme(datetime.now())

        self.frame = tk.Frame(
            parent,
            relief="ridge",
            bd=2,
            bg=self._theme_bg,
            highlightbackground=self._theme_border,
            highlightthickness=1,
            padx=8,
            pady=6,
        )

        tk.Label(
            self.frame,
            text="✦ Bisa ✦",
            font=("Segoe UI", 10, "bold"),
            bg=self._theme_bg,
            foreground=self._theme_accent,
        ).pack()

        self.dog_var = tk.StringVar()
        self.dog_label = tk.Label(
            self.frame,
            textvariable=self.dog_var,
            font=("Courier", 11, "bold"),
            justify="center",
            cursor="hand2",
            bg=self._theme_bg,
            fg=self._theme_accent,
        )
        self.dog_label.pack(pady=(2, 0), fill="x", expand=True)
        self.dog_label.bind("<Button-1>", lambda _e: self.receive_pet())

        self.msg_var = tk.StringVar(value="...")
        tk.Label(
            self.frame,
            textvariable=self.msg_var,
            font=("Segoe UI", 9),
            bg=self._theme_bg,
            fg=self._theme_msg,
        ).pack()

        tk.Frame(self.frame, bg=self._theme_border, height=1).pack(fill="x", pady=4)

        tk.Label(
            self.frame,
            text="click box → throw treat  |  click Bisa → pet",
            font=("Segoe UI", 8),
            bg=self._theme_bg,
            fg=self._theme_hint,
        ).pack(pady=(0, 2))

        self.stats_var = tk.StringVar(value="pets:0  treats:0")
        tk.Label(
            self.frame,
            textvariable=self.stats_var,
            font=("Segoe UI", 8),
            bg=self._theme_bg,
            fg=self._theme_stats,
        ).pack()

        self.frame.bind("<Button-1>", self._on_frame_click)

        self._render_frame(self.IDLE_FRAMES[0])
        self._idle_loop()

    # ------------------------------
    # Theme
    # ------------------------------
    def _apply_seasonal_theme(self, now):
        # Defaults: your purple
        self._theme_bg = "#f0eaf4"
        self._theme_border = "#c9a8d4"
        self._theme_accent = "#7a4a9a"
        self._theme_msg = "#9c6dbf"
        self._theme_hint = "#c9a8d4"
        self._theme_stats = "#c9a8d4"

        self._seasonal_idle_frames = None

        # October: Halloween
        if now.month == 10:
            self._theme_bg = "#1f1326"
            self._theme_border = "#7a4a9a"
            self._theme_accent = "#ff7a18"
            self._theme_msg = "#c9a8d4"
            self._theme_hint = "#7f6a86"
            self._theme_stats = "#6a4b73"
            self._seasonal_idle_frames = self.HALLOWEEN_FRAMES

        # December: Winter
        elif now.month == 12:
            self._theme_bg = "#eef6ff"
            self._theme_border = "#b8d7ff"
            self._theme_accent = "#2a5aa5"
            self._theme_msg = "#3b76c9"
            self._theme_hint = "#8aa9d6"
            self._theme_stats = "#c7d9f2"
            self._seasonal_idle_frames = self.WINTER_FRAMES

    # ------------------------------
    # Core rendering
    # ------------------------------
    def _render_frame(self, text: str, msg: str = ""):
        self.dog_var.set(text)
        if msg:
            self.msg_var.set(msg)

    def _cancel(self):
        if self._after_id is not None:
            try:
                self.parent.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _update_stats(self):
        self.stats_var.set(f"pets:{self._total_pets}  treats:{self._total_treats}")

    # ------------------------------
    # Animation engine
    # ------------------------------
    def _run_anim(self, frames, msg, speed_ms, on_done):
        import random
        self._anim_frames = list(frames or [])
        self._anim_idx = 0

        def _step():
            if self._anim_idx < len(self._anim_frames):
                self._render_frame(self._anim_frames[self._anim_idx], msg)
                self._anim_idx += 1

                # Tiny timing variance (feels less robotic)
                jitter = int(speed_ms * 0.08)
                delay = max(40, speed_ms + random.randint(-jitter, jitter))
                self._after_id = self.parent.after(delay, _step)
            else:
                on_done()

        _step()

    def _return_idle(self):
        self._state = "idle"
        self._idle_idx = 0
        self._render_frame(self.IDLE_FRAMES[0], "...")
        self._idle_loop()

    def _maybe_play_legendary(self) -> bool:
        import random
        if random.random() < self._legendary_chance:
            self._cancel()
            self._state = "legendary"
            self._run_anim(
                self.LEGENDARY_FRAMES,
                self.MESSAGES["legendary"],
                int(200 * self._speed_scale),
                lambda: self._return_idle(),
            )
            return True
        return False

    # ------------------------------
    # Idle loop (expanded)
    # ------------------------------
    def _idle_loop(self):
        import random
        self._cancel()
        self._after_id = self.parent.after(random.randint(650, 1500), self._idle_tick)

    def _idle_tick(self):
        import random
        from datetime import datetime

        if self._state != "idle":
            return

        # Rare legendary pop
        if self._maybe_play_legendary():
            return

        # Seasonal cameo (occasional)
        if self._seasonal_idle_frames and random.random() < 0.12:
            msg_key = "halloween" if datetime.now().month == 10 else "winter"
            self._cancel()
            self._state = "idle"
            self._run_anim(self._seasonal_idle_frames, self.MESSAGES.get(msg_key, "..."), int(420 * self._speed_scale),
                           lambda: self._return_idle())
            return

        r = random.random()
        if r < 0.06:
            self._cancel(); self._state = "blink"
            self._run_anim(self.BLINK_FRAMES, self.MESSAGES["blink"], int(220 * self._speed_scale),
                           lambda: self._return_idle())
            return
        if r < 0.10:
            self._cancel(); self._state = "wag"
            self._run_anim(self.WAG_FRAMES, self.MESSAGES["wag"], int(160 * self._speed_scale),
                           lambda: self._return_idle())
            return
        if r < 0.12:
            self._cancel(); self._state = "sleep"
            self._run_anim(self.SLEEP_FRAMES, self.MESSAGES["sleep"], int(520 * self._speed_scale),
                           lambda: self._return_idle())
            return
        if r < 0.14:
            self._cancel(); self._state = "zoomies"
            self._run_anim(self.ZOOMIES_FRAMES, self.MESSAGES["zoomies"], int(140 * self._speed_scale),
                           lambda: self._return_idle())
            return
        if r < 0.24:
            self._cancel(); self._state = "stretch"
            self._run_anim(self.STRETCH_FRAMES, self.MESSAGES["stretch"], int(550 * self._speed_scale),
                           lambda: self._return_idle())
            return

        # Original idle cycle
        self._idle_idx = (self._idle_idx + 1) % len(self.IDLE_FRAMES)
        self._render_frame(self.IDLE_FRAMES[self._idle_idx], "...")
        self._idle_loop()

    # ------------------------------
    # User interactions (kept)
    # ------------------------------
    def _on_frame_click(self, event):
        if self._state != "idle":
            return

        try:
            frame_w = max(self.frame.winfo_width(), 1)
            rel = max(0.0, min(1.0, (event.x_root - self.frame.winfo_rootx()) / frame_w))
            go_frames = self.TREAT_SHORT if rel < 0.30 else (self.TREAT_MEDIUM if rel < 0.65 else self.TREAT_FAR)
            self._cancel()
            self._state = "treat"
            self._total_treats += 1
            self._update_stats()

            if self._maybe_play_legendary():
                return

            self._run_anim(go_frames, self.MESSAGES["treat"], int(200 * self._speed_scale),
                           lambda: self._run_anim(self.RUN_BACK, self.MESSAGES["running"], int(200 * self._speed_scale),
                                                  lambda: self._return_idle()))
        except Exception:
            pass

    def receive_pet(self):
        if self._state != "idle":
            return

        self._cancel()
        self._state = "pet"
        self._total_pets += 1
        self._update_stats()

        if self._maybe_play_legendary():
            return

        self._run_anim(self.PET_FRAMES, self.MESSAGES["pet"], int(480 * self._speed_scale),
                       lambda: self._run_anim(self.HAPPY_FRAMES[:3], self.MESSAGES["pet"], int(480 * self._speed_scale),
                                              lambda: self._return_idle()))

    def throw_treat_at_window_x(self, window_x: int, window_width: int):
        if self._state != "idle":
            return

        try:
            rel = max(0.0, min(1.0, (window_x - (self.frame.winfo_rootx() - self.frame.winfo_toplevel().winfo_rootx())) /
                                    max(self.frame.winfo_width(), 1)))
        except Exception:
            rel = 0.5

        go_frames = self.TREAT_SHORT if rel < 0.30 else (self.TREAT_MEDIUM if rel < 0.65 else self.TREAT_FAR)
        self._cancel()
        self._state = "treat"
        self._total_treats += 1
        self._update_stats()

        if self._maybe_play_legendary():
            return

        self._run_anim(go_frames, self.MESSAGES["treat"], int(200 * self._speed_scale),
                       lambda: self._run_anim(self.RUN_BACK, self.MESSAGES["running"], int(200 * self._speed_scale),
                                              lambda: self._return_idle()))

    # ------------------------------
    # Reactions (kept)
    # ------------------------------
    def celebrate(self):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "happy"

        if self._maybe_play_legendary():
            return

        self._run_anim(self.HAPPY_FRAMES, self.MESSAGES["happy"], int(380 * self._speed_scale),
                       lambda: self._return_idle())

    def react_data_loaded(self, row_count: int = 0):
        if self._state not in ("idle", "happy"):
            return
        self._cancel()
        self._state = "loaded"
        self._run_anim(self.LOAD_FRAMES, self.MESSAGES["loaded"], int(320 * self._speed_scale),
                       lambda: self._return_idle())

    def react_excluded(self, count: int = 1):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "excluded"
        self._run_anim(self.EXCLUDED_FRAMES, self.MESSAGES["excluded"], int(420 * self._speed_scale),
                       lambda: self._return_idle())

    def react_restored(self, count: int = 1):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "happy"
        self._run_anim(self.HAPPY_FRAMES[:4], self.MESSAGES["restored"], int(400 * self._speed_scale),
                       lambda: self._return_idle())

    def react_row_selected(self):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "alert"
        self._run_anim(self.ALERT_FRAMES, self.MESSAGES["alert"], int(280 * self._speed_scale),
                       lambda: self._run_anim(self.SNIFF_FRAMES, self.MESSAGES["sniff"], int(380 * self._speed_scale),
                                              lambda: self._return_idle()))

    def react_kuntal(self, count: int = 1):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "kuntal"
        self._run_anim(self.KUNTAL_FRAMES, self.MESSAGES["kuntal"], int(340 * self._speed_scale),
                       lambda: self._return_idle())

    def react_cleared(self):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "cleared"
        self._run_anim(self.CLEARED_FRAMES, self.MESSAGES["cleared"], int(400 * self._speed_scale),
                       lambda: self._return_idle())

    # ------------------------------
    # New contextual reactions
    # ------------------------------
    def react_success(self, msg: str = "nice!! ✅"):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "success"
        frames = self.SUCCESS_FRAMES + self.WAG_FRAMES
        self._run_anim(frames, msg, int(170 * self._speed_scale), lambda: self._return_idle())

    def react_warning(self, msg: str = "uh oh… ⚠️"):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "warning"
        self._run_anim(self.WARNING_FRAMES, msg, int(220 * self._speed_scale), lambda: self._return_idle())

    def react_error(self, msg: str = "nope… 💥"):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "error"
        self._run_anim(self.CONFUSED_FRAMES, msg, int(210 * self._speed_scale), lambda: self._return_idle())
# ------------------------------
# GUI
# ------------------------------
class MoveUpGUI:

    def __init__(self, root: Tk):
        self.root = root
        self.base_title = f"Move-Up v{APP_VERSION} — KK"
        self.root.title(self.base_title)
        self.root.geometry("1240x920")

        self.style = ttk.Style(self.root)
        self.base_theme = self.style.theme_use()

        # Persisted vars
        self.kawaii_var = BooleanVar(value=False)
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

        self.app_dir = self._determine_app_dir()
        self.config_path = os.path.join(self.app_dir, "moveup_config.json")

        self.export_root = os.path.join(self.app_dir, "generated")
        os.makedirs(self.export_root, exist_ok=True)
        self.export_run_dir = os.path.join(self.export_root, datetime.now().strftime("%Y-%m-%d_%H-%M"))
        os.makedirs(self.export_run_dir, exist_ok=True)

        # Persistent filters + aliases
        self.room_alias_map: Dict[str, str] = {}
        self.selected_rooms: List[str] = []
        self.selected_brands: List[str] = []
        self.selected_types: List[str] = []

        self.last_import_dir: Optional[str] = None

        # Runtime state
        self.raw_df: Optional[pd.DataFrame] = None
        self.current_df: Optional[pd.DataFrame] = None
        self.col_mapping_override: Dict[str, str] = {}
        self.moveup_df: Optional[pd.DataFrame] = None
        self.excluded_barcodes: set = set()
        self.kuntal_priority_barcodes: set = set()

        self.filters_window: Optional[Toplevel] = None

        self._load_config()
        self._build_ui()
        self._bind_window_treat()
        self._toggle_theme(initial=True)
        self._refresh_button_labels()
        self._update_kuntalcount()

        self.root.protocol("WM_DELETE_WINDOW", self._on_app_close)

    # ------------------------------
    # Config persistence
    # ------------------------------
    def _load_config(self):
        try:
            if not os.path.exists(self.config_path):
                return
            with open(self.config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)

            self.room_alias_map = dict(cfg.get("room_alias_map", {}) or {})
            self.selected_rooms = list(cfg.get("selected_rooms", []) or [])
            self.selected_brands = list(cfg.get("selected_brands", []) or [])
            self.selected_types = list(cfg.get("selected_types", []) or [])

            self.kawaii_var.set(bool(cfg.get("kawaii_mode", self.kawaii_var.get())))
            self.printer_bw_var.set(bool(cfg.get("printer_bw", self.printer_bw_var.get())))
            self.skip_sales_floor_var.set(bool(cfg.get("skip_sales_floor", self.skip_sales_floor_var.get())))
            self.hide_removed_var.set(bool(cfg.get("hide_removed", self.hide_removed_var.get())))
            self.auto_open_var.set(bool(cfg.get("auto_open_pdf", self.auto_open_var.get())))
            self.timestamp_var.set(bool(cfg.get("timestamp", self.timestamp_var.get())))

            self.excluded_barcodes = set(cfg.get("excluded_barcodes", []) or [])
            self.kuntal_priority_barcodes = set(cfg.get("kuntal_priority_barcodes", []) or [])

            # Restore active_columns, but validate against COLUMNS_TO_USE
            saved_cols = cfg.get("active_columns", [])
            if saved_cols and all(c in COLUMNS_TO_USE for c in saved_cols):
                self.active_columns = list(saved_cols)
            else:
                self.active_columns = list(COLUMNS_TO_USE)

            try:
                self.page_items_var.set(int(cfg.get("items_per_page", self.page_items_var.get())))
            except Exception:
                pass

            self.prefix_var.set(str(cfg.get("prefix", self.prefix_var.get()) or ""))

            last_dir = cfg.get("last_import_dir")
            if isinstance(last_dir, str) and last_dir.strip() and os.path.isdir(last_dir):
                self.last_import_dir = last_dir.strip()

        except Exception:
            return

    def _save_config(self):
        try:
            cfg = {
                "room_alias_map": self.room_alias_map,
                "selected_rooms": self.selected_rooms,
                "selected_brands": self.selected_brands,
                "selected_types": self.selected_types,
                "kawaii_mode": bool(self.kawaii_var.get()),
                "printer_bw": bool(self.printer_bw_var.get()),
                "skip_sales_floor": bool(self.skip_sales_floor_var.get()),
                "hide_removed": bool(self.hide_removed_var.get()),
                "auto_open_pdf": bool(self.auto_open_var.get()),
                "timestamp": bool(self.timestamp_var.get()),
                "items_per_page": int(self.page_items_var.get() or 30),
                "prefix": str(self.prefix_var.get() or ""),
                "last_import_dir": self.last_import_dir or "",
                "excluded_barcodes": sorted(list(self.excluded_barcodes)),
                "kuntal_priority_barcodes": sorted(list(self.kuntal_priority_barcodes)),
                "active_columns": self.active_columns,
            }
            tmp = self.config_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
            os.replace(tmp, self.config_path)
        except Exception:
            pass

    def _on_app_close(self):
        self._save_config()
        try:
            self.root.destroy()
        except Exception:
            pass

    # ------------------------------
    # Base helpers
    # ------------------------------
    def _determine_app_dir(self) -> str:
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))

    def _create_kawaii_theme(self):
        if "kawaii_daisy" in self.style.theme_names():
            return
        self.style.theme_create(
            "kawaii_daisy",
            parent=self.base_theme,
            settings={
                "TFrame": {"configure": {"background": "#ffe6f2"}},
                "TLabel": {"configure": {"background": "#ffe6f2", "foreground": "#4a154b"}},
                "TButton": {
                    "configure": {"padding": 6, "relief": "raised", "background": "#ffcce5", "foreground": "#4a154b"},
                    "map": {"background": [("active", "#ffb6d9"), ("pressed", "#ff9fcf")]}
                },
                "Treeview": {
                    "configure": {"background": "#fff7fb", "fieldbackground": "#fff7fb", "foreground": "#333333",
                                  "rowheight": 20},
                    "map": {"background": [("selected", "#ffb6d9")], "foreground": [("selected", "#000000")]}
                },
                "TCheckbutton": {"configure": {"background": "#ffe6f2"}},
            }
        )

    def _register_button(self, btn, base_text: str):
        self._button_registry.append((btn, base_text))

    def _refresh_button_labels(self):
        for btn, base in self._button_registry:
            btn.config(text=(f"🌼 {base} 🌼" if self.kawaii_var.get() else base))

    def _toggle_theme(self, initial: bool = False):
        if self.kawaii_var.get():
            self.style.theme_use("kawaii_daisy")
            self.root.title(self.base_title + " 🌼🌼🌼")
        else:
            self.style.theme_use(self.base_theme)
            self.root.title(self.base_title)
        self._refresh_button_labels()
        if not initial:
            self._save_config()

    def open_kawaii_settings(self):
        try:
            from kawaii_preview import open_kawaii_settings_window
            open_kawaii_settings_window(self.root)
        except Exception as e:
            messagebox.showerror("Kawaii PDF Settings", f"Could not open settings window:\n\n{e}")

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

        ttk.Checkbutton(
            btn_row, text="Kawaii mode",
            variable=self.kawaii_var, command=self._toggle_theme,
        ).pack(side="left", padx=8)

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

        btn_kawaii_settings = ttk.Button(
            adv_row, text="Kawaii PDF Settings…", command=self.open_kawaii_settings,
        )
        btn_kawaii_settings.pack(side="left", padx=4)
        self._register_button(btn_kawaii_settings, "Kawaii PDF Settings…")

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

        self.kuntalcount_var = StringVar(value="Kuntal's priority items: 0")
        ttk.Label(frm_controls, textvariable=self.kuntalcount_var, anchor="w").pack(fill="x")

        self.filters_summary_var = StringVar(value="Filters: default")
        ttk.Label(frm_controls, textvariable=self.filters_summary_var, anchor="w", wraplength=480).pack(fill="x")

        # ── Right: Bisa — stretch wide, top-aligned ──
        self.dog_widget = AsciiDogWidget(frm_top_row)
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
        self.nb.add(self.tab_kuntal, text="Kuntal's Priority")
        self.nb.add(self.tab_excluded, text="Excluded / Removed")
        self.nb.add(self.tab_all, text="All Items")

        self.tree = ttk.Treeview(self.tab_moveup, columns=tuple(self.active_columns), show="headings", height=18)
        self._configure_tree_columns(self.tree, self.active_columns)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self._on_moveup_double_click)
        self.tree.bind("<ButtonRelease-1>", self._on_moveup_single_click)

        self.k_tree = ttk.Treeview(self.tab_kuntal, columns=tuple(COLUMNS_TO_USE), show="headings", height=18)
        self._configure_tree_columns(self.k_tree, COLUMNS_TO_USE)
        self.k_tree.pack(fill="both", expand=True)

        self.x_tree = ttk.Treeview(self.tab_excluded, columns=tuple(COLUMNS_TO_USE), show="headings", height=18)
        self._configure_tree_columns(self.x_tree, COLUMNS_TO_USE)
        self.x_tree.pack(fill="both", expand=True)
        self.x_tree.bind("<ButtonRelease-1>", self._on_excluded_single_click)

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

        btn_kuntal = ttk.Button(frm_remove, text="Toggle Kuntal's Priority", command=self._toggle_kuntal_selected)
        btn_kuntal.pack(side="left", padx=4)
        self._register_button(btn_kuntal, "Toggle Kuntal's Priority")

        btn_manual = ttk.Button(frm_remove, text="Manual Add…", command=self._manual_add_dialog)
        btn_manual.pack(side="left", padx=4)
        self._register_button(btn_manual, "Manual Add…")

        btn_clear_k = ttk.Button(frm_remove, text="Clear Kuntal's List", command=self._clear_kuntal_list)
        btn_clear_k.pack(side="left", padx=4)
        self._register_button(btn_clear_k, "Clear Kuntal's List")

        self.diag_var = StringVar(value="")
        ttk.Label(
            self.root, textvariable=self.diag_var,
            anchor="w", foreground="#555",
        ).pack(fill="x", padx=10, pady=(0, 6))



    def _configure_tree_columns(self, tree: ttk.Treeview, cols: List[str]):
        """Apply standard column widths and wire up click-to-sort on every heading."""
        tree_id = str(id(tree))
        if tree_id not in self._sort_state:
            self._sort_state[tree_id] = {}

        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=150 if col != "Product Name" else 440, anchor="w")
            # Bind after setting heading so the command captures col correctly
            tree.heading(col, command=lambda c=col, t=tree, tid=tree_id: self._sort_tree(t, tid, c))

    def _sort_tree(self, tree: ttk.Treeview, tree_id: str, col: str):
        """Sort a Treeview in-place by col, toggling asc/desc. Updates heading arrows."""
        state = self._sort_state.setdefault(tree_id, {})
        ascending = not state.get(col, True)   # first click → ascending
        state[col] = ascending

        # Collect all rows
        rows = [(tree.set(iid, col), iid) for iid in tree.get_children("")]

        # Try numeric sort first, fall back to case-insensitive string
        try:
            rows.sort(key=lambda x: float(x[0]) if x[0] != "" else float("-inf"),
                      reverse=not ascending)
        except (ValueError, TypeError):
            rows.sort(key=lambda x: str(x[0]).lower(), reverse=not ascending)

        for pos, (_, iid) in enumerate(rows):
            tree.move(iid, "", pos)

        # Update all headings: clear arrows on others, set on sorted col
        for c in tree["columns"]:
            current = tree.heading(c, "text")
            # Strip any existing arrow
            clean = current.rstrip(" ▲▼")
            if c == col:
                arrow = " ▲" if ascending else " ▼"
                tree.heading(c, text=clean + arrow)
            else:
                tree.heading(c, text=clean)

    def _rebuild_main_tree_columns(self):
        """
        Destroy and recreate the Move-Up Treeview so it reflects
        self.active_columns. Called after the column editor applies changes.
        """
        # Unbind before destroying
        try:
            self.tree.unbind("<Double-1>")
        except Exception:
            pass

        # Clear stale sort state for the old tree
        old_id = str(id(self.tree))
        self._sort_state.pop(old_id, None)

        self.tree.destroy()

        self.tree = ttk.Treeview(
            self.tab_moveup,
            columns=tuple(self.active_columns),
            show="headings",
            height=18
        )
        self._configure_tree_columns(self.tree, self.active_columns)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self._on_moveup_double_click)

        # Re-render with current data
        if self.moveup_df is not None:
            self._render_tree(self.moveup_df)

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
    # ── NEW: Column Editor ──
    # ------------------------------
    def open_column_editor(self):
        """
        Opens a window that lets the user choose which columns to show
        in the Move-Up tree, and in what order.
        """
        win = Toplevel(self.root)
        win.title("Edit Display Columns")
        win.geometry("520x480")
        win.transient(self.root)
        win.grab_set()

        ttk.Label(
            win,
            text="Select which columns to display in the Move-Up list,\nand reorder them using the buttons.",
            justify="left"
        ).pack(anchor="w", padx=12, pady=(12, 6))

        # ── listbox showing current active order ──
        frm_list = ttk.Frame(win)
        frm_list.pack(fill="both", expand=True, padx=12, pady=6)

        lb = Listbox(frm_list, selectmode=tk.SINGLE, exportselection=False, height=16)
        lb.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(frm_list, orient="vertical", command=lb.yview)
        sb.pack(side="right", fill="y")
        lb.config(yscrollcommand=sb.set)

        # Track which columns are enabled via checkmarks in the label.
        # We store the full ordered list (all COLUMNS_TO_USE) and which are active.
        ordered: List[str] = list(COLUMNS_TO_USE)  # canonical order pool
        enabled: Dict[str, bool] = {c: (c in self.active_columns) for c in ordered}

        # Reorder so active_columns come first in the displayed order,
        # followed by any hidden ones at the bottom.
        ordered = list(self.active_columns) + [c for c in COLUMNS_TO_USE if c not in self.active_columns]

        def refresh_lb():
            sel = lb.curselection()
            lb.delete(0, END)
            for c in ordered:
                mark = "✓" if enabled[c] else "✗"
                lb.insert(END, f"  {mark}  {c}")
            if sel:
                lb.selection_set(sel[0])

        refresh_lb()

        # ── side buttons: move up / down / toggle ──
        frm_btns = ttk.Frame(win)
        frm_btns.pack(fill="x", padx=12, pady=4)

        def move_up():
            sel = lb.curselection()
            if not sel:
                return
            i = sel[0]
            if i == 0:
                return
            ordered[i], ordered[i - 1] = ordered[i - 1], ordered[i]
            refresh_lb()
            lb.selection_set(i - 1)

        def move_down():
            sel = lb.curselection()
            if not sel:
                return
            i = sel[0]
            if i >= len(ordered) - 1:
                return
            ordered[i], ordered[i + 1] = ordered[i + 1], ordered[i]
            refresh_lb()
            lb.selection_set(i + 1)

        def toggle_enabled():
            sel = lb.curselection()
            if not sel:
                return
            col = ordered[sel[0]]
            # Always keep Package Barcode enabled — it's used as the key everywhere
            if col == "Package Barcode" and enabled[col]:
                messagebox.showwarning(
                    "Edit Columns",
                    "'Package Barcode' must remain visible — it's used as the row key."
                )
                return
            enabled[col] = not enabled[col]
            refresh_lb()
            lb.selection_set(sel[0])

        def reset_defaults():
            nonlocal ordered
            ordered[:] = list(COLUMNS_TO_USE)
            for c in ordered:
                enabled[c] = True
            refresh_lb()

        ttk.Button(frm_btns, text="▲ Move Up", command=move_up).pack(side="left", padx=4)
        ttk.Button(frm_btns, text="▼ Move Down", command=move_down).pack(side="left", padx=4)
        ttk.Button(frm_btns, text="Toggle Show/Hide", command=toggle_enabled).pack(side="left", padx=4)
        ttk.Button(frm_btns, text="Reset Defaults", command=reset_defaults).pack(side="left", padx=12)

        ttk.Label(win, text="Double-click a row to toggle it on/off.", foreground="#666").pack(
            anchor="w", padx=12, pady=(0, 2)
        )

        lb.bind("<Double-1>", lambda _e: toggle_enabled())

        # ── Apply / Cancel ──
        frm_bot = ttk.Frame(win)
        frm_bot.pack(fill="x", padx=12, pady=(6, 12))

        def apply():
            new_active = [c for c in ordered if enabled[c]]
            if not new_active:
                messagebox.showerror("Edit Columns", "At least one column must be visible.")
                return
            if "Package Barcode" not in new_active:
                messagebox.showerror("Edit Columns", "'Package Barcode' must remain visible.")
                return
            self.active_columns = new_active
            self._save_config()
            self._rebuild_main_tree_columns()
            self.status.set(f"Display columns updated: {', '.join(self.active_columns)}")
            win.destroy()

        ttk.Button(frm_bot, text="Apply", command=apply).pack(side="left")
        ttk.Button(frm_bot, text="Cancel", command=win.destroy).pack(side="left", padx=8)

    # ------------------------------
    # Simple callbacks
    # ------------------------------
    def _on_hide_removed_changed(self):
        self._save_config()
        self._recompute_from_current()

    # ------------------------------
    # Status counters
    # ------------------------------
    def _update_rowcount(self, df: Optional[pd.DataFrame]):
        n = 0 if df is None else len(df)
        self.rowcount_var.set(f"Items loaded: {n}")

    def _update_moveupcount(self, df: Optional[pd.DataFrame]):
        n = 0 if df is None else len(df)
        self.moveupcount_var.set(f"Move-Up items: {n}")

    def _update_kuntalcount(self):
        self.kuntalcount_var.set(f"Kuntal's priority items: {len(self.kuntal_priority_barcodes)}")

    # ------------------------------
    # Window-wide treat throwing
    # ------------------------------
    def _bind_window_treat(self):
        """Any click on blank space throws a treat for Bisa."""
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
            # Also skip if the widget is inside the dog's own frame
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
        """
        Returns the columns to show in treeviews: active_columns +
        any DISPLAY_EXTRA_COLUMNS that are present in df (or current_df).
        """
        base = list(self.active_columns)
        src = df if df is not None else self.current_df
        if src is not None and not src.empty:
            for col in self.DISPLAY_EXTRA_COLUMNS:
                if col in src.columns and col not in base:
                    base.append(col)
        return base

    # ------------------------------
    # Open folder
    # ------------------------------
    def open_output_folder(self):
        path = self.export_run_dir
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
                if hasattr(self, "dog_widget"):
                    self.dog_widget.react_data_loaded(len(mapped))
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

        src_cols = list(self.raw_df.columns)
        metrc_src_detected = detect_metrc_source_column(self.raw_df)

        auto_map = {}
        try:
            _auto_df, auto_map = automap_columns(self.raw_df)
        except Exception:
            auto_map = {}

        win = Toplevel(self.root)
        win.title("Map Columns (Manual Override)")
        win.geometry("760x560")
        ttk.Label(win, text="Choose which source column maps to each required field.").pack(anchor="w", padx=10, pady=10)

        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=10)

        combos = {}

        ttk.Label(frame, text="METRC Source Column (required):").grid(row=0, column=0, sticky="e", pady=6)
        metrc_var = StringVar(value=metrc_src_detected or "")
        metrc_cb = ttk.Combobox(frame, textvariable=metrc_var, values=src_cols, width=52, state="readonly")
        metrc_cb.grid(row=0, column=1, sticky="w", pady=6)
        ttk.Label(frame, text="⚠ Changing this resets the barcode key", foreground="#aa6600").grid(
            row=0, column=2, sticky="w", padx=(8, 0)
        )

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
                    "excluded / Kuntal priority lists if the barcodes no longer match.\n\n"
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
            ttk.Label(frame, text=target + ":").grid(row=i + row_offset, column=0, sticky="e", pady=4)
            var = StringVar(value="")

            if target == "Package Barcode":
                var.set(metrc_var.get().strip())
                cb = ttk.Combobox(
                    frame,
                    textvariable=var,
                    values=[metrc_var.get().strip()] if metrc_var.get().strip() else src_cols,
                    width=52,
                    state="disabled"
                )
            else:
                cb = ttk.Combobox(frame, textvariable=var, values=src_cols, width=52, state="readonly")
                pre = next((src for src, dst in auto_map.items() if dst == target), None)
                if pre:
                    var.set(pre)

            cb.grid(row=i + row_offset, column=1, sticky="w", pady=4)
            combos[target] = {"_var": var, "_cb": cb}

        rebuild_non_metrc_dropdown_values()

        ttk.Separator(frame, orient="horizontal").grid(
            row=len(COLUMNS_TO_USE) + row_offset, column=0, columnspan=2, sticky="ew", pady=10
        )

        opt_start = len(COLUMNS_TO_USE) + row_offset + 1
        ttk.Label(frame, text="Optional (used by Audit PDFs):", font=("Helvetica", 9, "bold")).grid(
            row=opt_start, column=0, sticky="e", pady=4
        )

        opt_vars = {}
        for j, opt in enumerate(AUDIT_OPTIONAL_FIELDS):
            ttk.Label(frame, text=f"{opt} (optional):").grid(row=opt_start + 1 + j, column=0, sticky="e", pady=4)
            v = StringVar(value="")
            pre = next((src for src, dst in auto_map.items() if dst == opt), None)
            if pre:
                v.set(pre)
            cb = ttk.Combobox(frame, textvariable=v, values=[""] + src_cols, width=52, state="readonly")
            cb.grid(row=opt_start + 1 + j, column=1, sticky="w", pady=4)
            opt_vars[opt] = v

        btns = ttk.Frame(win)
        btns.pack(fill="x", pady=10)

        def _apply_mapping():
            try:
                chosen_metrc = metrc_var.get().strip()
                if not chosen_metrc:
                    messagebox.showerror("Missing", "Please choose the METRC source column (required).")
                    return

                mapping = {}
                used_sources = set()

                mapping[chosen_metrc] = "Package Barcode"
                used_sources.add(chosen_metrc)

                for target in COLUMNS_TO_USE:
                    if target == "Package Barcode":
                        continue
                    src = combos[target]["_var"].get().strip()
                    if not src:
                        messagebox.showerror("Missing", f"Please choose a source for '{target}'.")
                        return
                    if src in used_sources:
                        messagebox.showerror("Duplicate Source", f"The source column '{src}' is used more than once.")
                        return
                    used_sources.add(src)
                    mapping[src] = target

                for opt in AUDIT_OPTIONAL_FIELDS:
                    src_opt = opt_vars.get(opt).get().strip()
                    if src_opt:
                        if src_opt in used_sources:
                            messagebox.showerror("Duplicate Source", f"The source column '{src_opt}' is already used.")
                            return
                        used_sources.add(src_opt)
                        mapping[src_opt] = opt

                df = self.raw_df.rename(columns=mapping)

                missing = [c for c in COLUMNS_TO_USE if c not in df.columns]
                if missing:
                    raise ValueError("After mapping, still missing: " + ", ".join(missing))

                df["Package Barcode"] = df["Package Barcode"].astype("string").fillna("")
                df["Qty On Hand"] = pd.to_numeric(df["Qty On Hand"], errors="coerce").fillna(0).astype(int)
                for col in ["Product Name", "Brand", "Type", "Room"]:
                    df[col] = df[col].astype(str)

                if "Distributor" in df.columns:
                    df["Distributor"] = df["Distributor"].astype(str).fillna("").str.strip()
                if "Store" in df.columns:
                    df["Store"] = df["Store"].astype(str).fillna("").str.strip()
                if "Size" in df.columns:
                    df["Size"] = df["Size"].astype(str).fillna("").str.strip()

                self.col_mapping_override = mapping
                self.current_df = df

                present = set(self.current_df["Package Barcode"].astype(str).fillna("").str.strip().tolist())
                self.excluded_barcodes = {bc for bc in self.excluded_barcodes if bc in present}
                self.kuntal_priority_barcodes = {bc for bc in self.kuntal_priority_barcodes if bc in present}
                self._update_kuntalcount()

                self._update_rowcount(df)
                self._recompute_from_current()
                win.destroy()
                self.status.set("Column mapping applied (METRC forced to Package Barcode).")
            except Exception as e:
                messagebox.showerror("Mapping Error", str(e))

        ttk.Button(btns, text="Apply", command=_apply_mapping).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="left", padx=6)

    # ------------------------------
    # Filters helpers
    # ------------------------------
    def _get_all_rooms_normalized(self, df: pd.DataFrame) -> List[str]:
        if df is None or df.empty or "Room" not in df.columns:
            return []
        df_norm = normalize_rooms(df, self.room_alias_map)
        return sorted(set(str(x).strip() for x in df_norm["Room"].dropna().astype(str).tolist()))

    def _get_all_brands(self, df: pd.DataFrame) -> List[str]:
        if df is None or df.empty or "Brand" not in df.columns:
            return []
        vals = sorted(set(str(x).strip() for x in df["Brand"].dropna().astype(str).tolist()))
        return [v for v in vals if v]

    def _get_all_types(self, df: pd.DataFrame) -> List[str]:
        if df is None or df.empty or "Type" not in df.columns:
            return []
        vals = sorted(set(str(x).strip() for x in df["Type"].dropna().astype(str).tolist()))
        return [v for v in vals if v]

    def _default_candidate_rooms(self, df: pd.DataFrame) -> List[str]:
        rooms = self._get_all_rooms_normalized(df)
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

        win = Toplevel(self.root)
        self.filters_window = win
        win.title("Filters")
        win.geometry("1180x860")
        win.transient(self.root)
        win.grab_set()

        df = self.current_df

        top = ttk.Frame(win, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="Room Aliases (optional) — normalize messy room names into a clean canonical name.").pack(anchor="w")

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
            for k, v in sorted(self.room_alias_map.items(), key=lambda kv: kv[0].lower()):
                alias_tree.insert("", "end", values=(k, v))

        def add_alias():
            f = (alias_from.get() or "").strip()
            t = (alias_to.get() or "").strip()
            if not f or not t:
                messagebox.showinfo("Alias", "Enter both From and To.")
                return
            self.room_alias_map[f] = t
            alias_from.set("")
            alias_to.set("")
            refresh_alias_tree()
            rooms_list.set_items(self._get_all_rooms_normalized(df))
            self._save_config()

        def remove_alias():
            sel = alias_tree.selection()
            if not sel:
                return
            for iid in sel:
                vals = alias_tree.item(iid, "values")
                if vals and vals[0] in self.room_alias_map:
                    del self.room_alias_map[vals[0]]
            refresh_alias_tree()
            rooms_list.set_items(self._get_all_rooms_normalized(df))
            self._save_config()

        ttk.Button(alias_row, text="Add/Update", command=add_alias).pack(side="left", padx=6)
        ttk.Button(alias_row, text="Remove Selected", command=remove_alias).pack(side="left", padx=6)
        refresh_alias_tree()

        mid = ttk.Frame(win, padding=10)
        mid.pack(fill="both", expand=True)

        rooms_list = _FilterList(mid, "Rooms (Move-Up source rooms)")
        brands_list = _FilterList(mid, "Brands (empty = ALL)")
        types_list = _FilterList(mid, "Types (empty = ALL)")

        rooms = self._get_all_rooms_normalized(df)
        brands = self._get_all_brands(df)
        types_ = self._get_all_types(df)

        rooms_list.set_items(rooms)
        brands_list.set_items(brands)
        types_list.set_items(types_)

        if self.selected_rooms:
            rooms_list.set_selected_values(self.selected_rooms)
        else:
            rooms_list.set_selected_values(self._default_candidate_rooms(df))

        brands_list.set_selected_values(self.selected_brands)
        types_list.set_selected_values(self.selected_types)

        bot = ttk.Frame(win, padding=10)
        bot.pack(fill="x")

        def apply_filters():
            sel_rooms = rooms_list.get_selected_values()
            if not sel_rooms:
                sel_rooms = self._default_candidate_rooms(df)

            self.selected_rooms = sel_rooms
            self.selected_brands = brands_list.get_selected_values()
            self.selected_types = types_list.get_selected_values()

            self._save_config()
            self._recompute_from_current()
            win.destroy()
            self.filters_window = None

        def reset_defaults():
            self.selected_rooms = []
            self.selected_brands = []
            self.selected_types = []
            rooms_list.set_selected_values(self._default_candidate_rooms(df))
            brands_list.clear_selection()
            types_list.clear_selection()
            self._save_config()

        ttk.Button(bot, text="Apply", command=apply_filters).pack(side="left")
        ttk.Button(bot, text="Reset Defaults", command=reset_defaults).pack(side="left", padx=8)
        ttk.Button(bot, text="Close", command=win.destroy).pack(side="left", padx=8)

        def _on_close():
            self.filters_window = None
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", _on_close)

    # ------------------------------
    # Audit window
    # ------------------------------
    def open_audit_window(self):
        if self.current_df is None or self.current_df.empty:
            messagebox.showinfo("Audit PDFs", "Import a file first.")
            return

        df = self.current_df.copy()
        has_dist = "Distributor" in df.columns
        if has_dist:
            blanks = (df["Distributor"].astype(str).fillna("").str.strip() == "").sum()
            self.status.set(f"Audit: Distributor column present. Blank rows: {blanks}/{len(df)}")
        else:
            self.status.set("Audit: Distributor column NOT present (will show Unknown Distributor).")

        if "Distributor" not in df.columns:
            df["Distributor"] = ""
        if "Store" not in df.columns:
            df["Store"] = ""
        if "Size" not in df.columns:
            df["Size"] = ""

        df["Distributor"] = df["Distributor"].astype(str).fillna("").str.strip()
        df.loc[df["Distributor"] == "", "Distributor"] = "Unknown Distributor"

        df_norm = normalize_rooms(df, self.room_alias_map)

        types_ = sorted(set(df_norm["Type"].dropna().astype(str).str.strip().tolist()))
        types_ = [t for t in types_ if t]

        brands = sorted(set(df_norm["Brand"].dropna().astype(str).str.strip().tolist()))
        brands = [b for b in brands if b]

        rooms = self._get_all_rooms_normalized(df_norm)

        dists = sorted(set(df_norm["Distributor"].dropna().astype(str).str.strip().tolist()))
        dists = [d for d in dists if d]

        dist_to_brands: Dict[str, set] = {}
        try:
            sub = df_norm[["Distributor", "Brand"]].copy()
            sub["Distributor"] = sub["Distributor"].astype(str).str.strip().replace({"": "Unknown Distributor"})
            sub["Brand"] = sub["Brand"].astype(str).str.strip()
            sub = sub[(sub["Distributor"] != "") & (sub["Brand"] != "")]
            for _, r in sub.drop_duplicates().iterrows():
                dist_to_brands.setdefault(r["Distributor"], set()).add(r["Brand"])
        except Exception:
            dist_to_brands = {}

        win = Toplevel(self.root)
        win.title("Audit PDF Export (Distributor Groups)")
        win.geometry("1320x820")
        win.transient(self.root)
        win.grab_set()

        pad = 10
        top = ttk.Frame(win, padding=pad)
        top.pack(fill="x")
        ttk.Label(
            top,
            text="Select filters for the Audit PDFs (Master + Blank). Page breaks follow the Sort Mode.",
            font=("Helvetica", 10, "bold"),
        ).pack(anchor="w")

        defaults = ttk.LabelFrame(win, text="Defaults (used if Store/Room missing)", padding=pad)
        defaults.pack(fill="x", padx=pad, pady=(0, pad))

        default_store_var = StringVar(value="Store")
        default_room_var = StringVar(value="Sales Floor")

        ttk.Label(defaults, text="Default Store:").pack(side="left")
        ttk.Entry(defaults, textvariable=default_store_var, width=26).pack(side="left", padx=(6, 18))
        ttk.Label(defaults, text="Default Room:").pack(side="left")
        ttk.Entry(defaults, textvariable=default_room_var, width=26).pack(side="left", padx=6)

        title_row = ttk.Frame(win, padding=(pad, 0))
        title_row.pack(fill="x")
        title_var = StringVar(value=f"Inventory Audit — {datetime.now().strftime('%m-%d-%Y')}")
        ttk.Label(title_row, text="Title").pack(side="left")
        ttk.Entry(title_row, textvariable=title_var).pack(side="left", fill="x", expand=True, padx=8)

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

        frm_sort = ttk.LabelFrame(win, text="Sort Mode (controls page breaks)", padding=pad)
        frm_sort.pack(fill="x", padx=pad, pady=(0, pad))

        ttk.Radiobutton(
            frm_sort,
            text="Distributor → Type → Size → Product (page break by Distributor)",
            variable=sort_mode_var,
            value="distributor_type_size_product"
        ).pack(anchor="w")

        ttk.Radiobutton(
            frm_sort,
            text="Brand → Type → Product (page break by Brand)",
            variable=sort_mode_var,
            value="brand_type_product"
        ).pack(anchor="w")

        ttk.Radiobutton(
            frm_sort,
            text="Type → Brand → Product (page break by Type)",
            variable=sort_mode_var,
            value="type_brand_product"
        ).pack(anchor="w")

        bot = ttk.Frame(win, padding=pad)
        bot.pack(fill="x")

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
                master_path, blank_path = export_audit_pdfs(
                    df=use,
                    base_dir=self.export_run_dir,
                    title_text=title_var.get().strip() or "Inventory Audit",
                    sort_mode=sort_mode_var.get(),
                    kawaii_pdf=bool(self.kawaii_var.get()),
                    printer_bw=bool(self.printer_bw_var.get()),
                    auto_open=bool(self.auto_open_var.get()),
                    default_store=default_store_var.get().strip() or "Store",
                    default_room=default_room_var.get().strip() or "Sales Floor",
                    type_trunc_len=TYPE_TRUNC_LEN,
                )

                self.status.set(f"Audit PDFs saved: {os.path.basename(master_path)} + {os.path.basename(blank_path)}")
                if hasattr(self, "dog_widget"):
                    self.dog_widget.react_success("Audit PDFs ✅")
                win.destroy()
            except Exception as e:
                messagebox.showerror("Audit PDFs", str(e))
                if hasattr(self, "dog_widget"):
                    self.dog_widget.react_error("Audit failed 💥")

        ttk.Button(bot, text="Export Audit PDFs", command=export_now).pack(side="left")
        ttk.Button(bot, text="Close", command=win.destroy).pack(side="right")

    # ------------------------------
    # Tree rendering
    # ------------------------------
    def _refresh_treeview_columns(self, df: "Optional[pd.DataFrame]" = None):
        """
        Reconfigure all four treeviews to reflect current display columns.
        Called when data is loaded or recomputed so Received Date appears/disappears cleanly.
        """
        moveup_cols = self._display_cols_for(df)
        extra_cols = self._display_cols_for(df)  # kuntal/excluded/all use full set + extras

        # Only rebuild if columns actually changed to avoid flicker
        current_moveup = list(self.tree["columns"])
        if current_moveup != moveup_cols:
            self.tree.config(columns=tuple(moveup_cols))
            self._configure_tree_columns(self.tree, moveup_cols)

        current_extra = list(self.k_tree["columns"])
        if current_extra != extra_cols:
            for t in (self.k_tree, self.x_tree, self.all_tree):
                t.config(columns=tuple(extra_cols))
                self._configure_tree_columns(t, extra_cols)

    def _render_tree(self, df: pd.DataFrame):
        for i in self.tree.get_children():
            self.tree.delete(i)

        if df is None or df.empty:
            return

        # Use active_columns + any present extra cols for display
        display_cols = self._display_cols_for(df)
        core_cols = COLUMNS_TO_USE

        idx_bar  = core_cols.index("Package Barcode")
        idx_room = core_cols.index("Room")
        idx_type = core_cols.index("Type")

        disp_idx_bar  = display_cols.index("Package Barcode") if "Package Barcode" in display_cols else None
        disp_idx_room = display_cols.index("Room")            if "Room"            in display_cols else None
        disp_idx_type = display_cols.index("Type")            if "Type"            in display_cols else None

        # Build a lookup for extra (non-core) columns → series for fast access
        extra_cols = [c for c in display_cols if c not in core_cols]
        extra_series = {c: df[c].reset_index(drop=True) for c in extra_cols if c in df.columns}

        core_missing = [c for c in core_cols if c not in df.columns]
        if core_missing:
            return

        for row_idx, full_row in enumerate(df[core_cols].itertuples(index=False, name=None)):
            bc = str(full_row[idx_bar]).strip()
            room_lower = str(full_row[idx_room]).strip().lower()
            is_backstock = (room_lower == "backstock")
            is_kuntal = (bc in self.kuntal_priority_barcodes)

            vals = []
            for c in display_cols:
                if c in core_cols:
                    vals.append(full_row[core_cols.index(c)])
                else:
                    vals.append(extra_series[c].iloc[row_idx] if c in extra_series else "")

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
            if bc and (bc in self.excluded_barcodes) and not self.hide_removed_var.get():
                tags.append("excluded")
            if is_backstock:
                tags.append("backstock")
            if is_kuntal:
                tags.append("kuntal")

            self.tree.insert("", "end", values=vals, tags=tuple(tags))

    def _render_kuntal_tree(self, df: pd.DataFrame):
        for i in self.k_tree.get_children():
            self.k_tree.delete(i)
        if df is None or df.empty:
            return
        display_cols = self._display_cols_for(df)
        core_cols = COLUMNS_TO_USE
        idx_type = core_cols.index("Type")
        extra_cols = [c for c in display_cols if c not in core_cols]
        extra_series = {c: df[c].reset_index(drop=True) for c in extra_cols if c in df.columns}
        for row_idx, row in enumerate(df[core_cols].itertuples(index=False, name=None)):
            vals = list(row)
            vals[idx_type] = ellipses(str(vals[idx_type]), TYPE_TRUNC_LEN)
            for c in extra_cols:
                vals.append(extra_series[c].iloc[row_idx] if c in extra_series else "")
            self.k_tree.insert("", "end", values=vals)

    def _render_excluded_tree(self, df: pd.DataFrame):
        for i in self.x_tree.get_children():
            self.x_tree.delete(i)
        if df is None or df.empty:
            return
        display_cols = self._display_cols_for(df)
        core_cols = COLUMNS_TO_USE
        idx_type = core_cols.index("Type")
        extra_cols = [c for c in display_cols if c not in core_cols]
        extra_series = {c: df[c].reset_index(drop=True) for c in extra_cols if c in df.columns}
        for seq, row in enumerate(df[core_cols].itertuples(index=False, name=None)):
            vals = list(row)
            vals[idx_type] = ellipses(str(vals[idx_type]), TYPE_TRUNC_LEN)
            for c in extra_cols:
                vals.append(extra_series[c].iloc[seq] if c in extra_series else "")
            self.x_tree.insert("", "end", iid=f"x_{seq}", values=vals)

    def _render_all_tree(self, df: Optional[pd.DataFrame]):
        for i in self.all_tree.get_children():
            self.all_tree.delete(i)

        if df is None or df.empty:
            self.all_items_count_var.set("")
            return

        display_cols = self._display_cols_for(df)
        core_cols = COLUMNS_TO_USE
        idx_type = core_cols.index("Type")
        idx_bar  = core_cols.index("Package Barcode")
        extra_cols = [c for c in display_cols if c not in core_cols]
        extra_series = {c: df[c].reset_index(drop=True) for c in extra_cols if c in df.columns}

        q = (self.all_search_var.get() or "").strip().lower()
        tokens = q.split() if q else []

        shown = 0
        for row_idx, row in enumerate(df[core_cols].itertuples(index=False, name=None)):
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
            if bc in self.excluded_barcodes:
                tags.append("excluded_all")
            if bc in self.kuntal_priority_barcodes:
                tags.append("kuntal_all")

            self.all_tree.insert("", "end", values=vals, tags=tuple(tags))
            shown += 1

        total = len(df)
        self.all_items_count_var.set(
            f"Showing {shown} of {total}" if tokens else f"{total} items"
        )
        self.all_tree.tag_configure("excluded_all", foreground="#999999")
        self.all_tree.tag_configure("kuntal_all", foreground="#c0007a")


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
            messagebox.showinfo("Kuntal's Priority", "Select row(s) first.")
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
        self.status.set(f"Toggled Kuntal's Priority on {toggled} item(s).")
        self._save_config()
        if toggled and hasattr(self, "dog_widget"):
            self.dog_widget.react_kuntal(toggled)

    def _clear_kuntal_list(self):
        self.kuntal_priority_barcodes.clear()
        self._update_kuntalcount()
        self._recompute_from_current()
        self.status.set("Cleared Kuntal's Priority list.")
        self._save_config()
        if hasattr(self, "dog_widget"):
            self.dog_widget.react_cleared()

    # ------------------------------
    # Excluded single-click restore
    # ------------------------------
    def _on_excluded_single_click(self, event):
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

        self._restore_excluded_selected(go_to_moveup=True, quiet=True)

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

        df = self.current_df.copy()
        missing = [c for c in COLUMNS_TO_USE if c not in df.columns]
        if missing:
            messagebox.showerror("Manual Add", "Missing required columns: " + ", ".join(missing))
            return

        win = Toplevel(self.root)
        win.title("Manual Add to Kuntal's Priority")
        win.geometry("920x620")
        win.transient(self.root)
        win.grab_set()

        ttk.Label(win, text="Search inventory. Select one or more, then Add.").pack(anchor="w", padx=10, pady=(10, 6))
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
            lambda r: f"{r['Brand']} | {r['Product Name']} | {r['Room']} | Qty:{r['Qty On Hand']} | …{str(r['Package Barcode'])[-6:]}",
            axis=1
        )
        rows = rows.sort_values(by=["Brand", "Product Name"], kind="stable").reset_index(drop=True)

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
            added = 0
            for pos in sel:
                i = filtered_idx[pos]
                bc = str(rows.loc[i, "__bc"]).strip()
                if bc and bc not in self.kuntal_priority_barcodes:
                    self.kuntal_priority_barcodes.add(bc)
                    added += 1
            self._update_kuntalcount()
            self._recompute_from_current()
            self.status.set(f"Manual added {added} item(s) to Kuntal's Priority.")
            self._save_config()
            win.destroy()

        btns = ttk.Frame(win)
        btns.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(btns, text="Add Selected", command=do_add).pack(side="left")
        ttk.Button(btns, text="Close", command=win.destroy).pack(side="left", padx=8)

        search_var.trace_add("write", refresh_list)
        refresh_list()
        ent.focus_set()

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
        all_rooms = set(self._get_all_rooms_normalized(df))
        if self.selected_rooms:
            cleaned = [r for r in self.selected_rooms if r in all_rooms]
            if cleaned:
                return cleaned
        return self._default_candidate_rooms(df)

    def _effective_brands(self, df: pd.DataFrame) -> List[str]:
        all_brands = set(self._get_all_brands(df))
        if self.selected_brands:
            cleaned = [b for b in self.selected_brands if b in all_brands]
            return cleaned
        return []

    def _effective_types(self, df: pd.DataFrame) -> List[str]:
        all_types = set(self._get_all_types(df))
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
            move_up_df = move_up_df[~move_up_df["Package Barcode"].astype(str).isin(self.excluded_barcodes)].copy()

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

        if self.excluded_barcodes:
            mu_use = self.moveup_df[~self.moveup_df["Package Barcode"].astype(str).isin(self.excluded_barcodes)].copy()
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
                kawaii_pdf=bool(self.kawaii_var.get()),
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

        if self.excluded_barcodes:
            mu_use = self.moveup_df[~self.moveup_df["Package Barcode"].astype(str).isin(self.excluded_barcodes)].copy()
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