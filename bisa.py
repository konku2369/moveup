"""
Bisa — Animated ASCII companion for the Move-Up Utility.

She's an earthmed-style husky who reacts to user interactions
and app events with various animations, tricks, and seasonal themes.
"""

import random
from datetime import datetime
import tkinter as tk


class AsciiDogWidget:
    """Animated ASCII companion widget — Bisa the husky.

    - click her to pet (receive_pet)
    - click box/blank space to throw her a treat (throw_treat_at_window_x / frame click)
    - stats counter (pets/treats)
    - react_* methods used by the app
    - idle micro-animations (wag, blink, sleep, zoomies)
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
        "  /\\_/\\  \n ( u.u )\u2665\n  > ^ <",
        "  /\\_/\\  \n ( ^w^ )\u2665\n  > ^ <",
        "  /\\_/\\  \n ( ^.^ )\u2665\n  > ^~<",
        "  /\\_/\\  \n ( u.u ) \n  > ^ <",
    ]
    TREAT_SHORT = [
        "  /\\_/\\    \U0001f9b4\n ( o.o )  \n  > ^ <",
        "    /\\_/\\ \U0001f9b4\n   ( ^.^) \n    > ^ <",
    ]
    TREAT_MEDIUM = [
        "  /\\_/\\      \U0001f9b4\n ( o.o )    \n  > ^ <",
        "    /\\_/\\  \U0001f9b4\n   ( ^o^ ) \n    > ^ <",
        "      /\\_/\\\U0001f9b4\n     ( ^.^)\n      > ^ <",
    ]
    TREAT_FAR = [
        "  /\\_/\\        \U0001f9b4\n ( o.o )      \n  > ^ <",
        "    /\\_/\\    \U0001f9b4\n   ( ^o^ )   \n    > ^ <",
        "      /\\_/\\ \U0001f9b4\n     ( ^.^) \n      > ^ <",
        "        /\\_/\\\U0001f9b4\n       ( ^O^)\n        > ^ <",
    ]
    RUN_BACK = [
        "      /\\_/\\  \n     \U0001f9b4(^.^) \n      > ^ <",
        "    /\\_/\\    \n   \U0001f9b4( ^w^)  \n    > ^ <",
        "  /\\_/\\      \n \U0001f9b4( ^.^)   \n  > ^ <",
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
        "  /\\_/\\  \n ( ^o^)\u2605\n  > ^ <",
        "   /\\_/\\ \n  (\u2605^o^)\n   > ^ <",
        "  /\\_/\\  \n  (^w^)\u2605\n  >w^ <",
        "  /\\_/\\  \n \\(^o^)/ \n  > ^ <",
        "  /\\_/\\  \n  ( ^.^)\u2605\n  > ^ <",
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
        "  /\\_/\\      \n ( ^o^ )  \u26a1\n  > ^ <      ",
        "      /\\_/\\  \n \u26a1 ( ^o^ ) \n      > ^ <  ",
        "  /\\_/\\      \n ( ^w^ )  \u26a1\n  > ^ <      ",
        "    /\\_/\\    \n \u26a1 ( ^w^ ) \n    > ^ <    ",
    ]
    CONFUSED_FRAMES = [
        "  /\\_/\\  \n ( o.o ) ?\n  > ^ <",
        "  /\\_/\\  \n ( O.o ) ?\n  > ^ <",
        "  /\\_/\\  \n ( o.O ) ?\n  > ^ <",
    ]
    BELLY_FRAMES = [
        "  /\\_/\\  \n  ( ^o^ ) \n  ~> w <~",
        "  /\\_/\\  \n  ( ^w^ )\u2665\n  ~> w <~",
        "  /\\_/\\  \n  ( ^.^ )\u2665\n  ~> w <~",
        "  /\\_/\\  \n  ( u.u )\u2665\n  ~> ^ <~",
    ]
    SUCCESS_FRAMES = [
        "    /\\_/\\   \n   ( ^o^)\u2728 \n    > ^ <",
        "      /\\_/\\ \n     ( ^w^)\u2728\n      > ^ <",
        "    /\\_/\\   \n   \\(^o^)/\u2728\n    > ^ <",
    ]
    WARNING_FRAMES = [
        "  /|_|\\  \n ( O.O ) !\n  > ! <",
        "  /|_|\\  \n ( o.o ) !\n  > ! <",
        "  /\\_/\\  \n ( o.o ) !\n  > ! <",
    ]
    LEGENDARY_FRAMES = [
        "  /\\_/\\   \u2605\u2605\u2605\n ( \u2727o\u2727 )  \u2605\n  > W <   \u2605",
        "  /\\_/\\   \u2605\u2605\u2605\n ( \u2727w\u2727 )  \u2605\n  > W <   \u2605",
        "  /\\_/\\   \u2605\u2605\u2605\n ( \u2727.^\u2727 ) \u2605\n  > W <   \u2605",
        "  /\\_/\\   \u2605\u2605\u2605\n ( \u2727o\u2727 )  \u2605\n  > W <   \u2605",
    ]

    # Trick frames
    SIT_FRAMES = [
        "  /\\_/\\  \n ( o.o ) \n  > ^ <",
        "  /\\_/\\  \n ( ^.^ ) \n  |   |",
        "  /\\_/\\  \n ( ^.^ ) \n  | W |",
        "  /\\_/\\  \n ( u.u ) \n  | W |",
    ]
    SHAKE_FRAMES = [
        "  /\\_/\\  \n ( o.o )/ \n  > ^ <",
        "  /\\_/\\  \n ( ^.^ )\U0001f91d\n  > ^ <",
        "  /\\_/\\  \n ( ^w^ )\U0001f91d\n  > ^ <",
        "  /\\_/\\  \n ( ^.^ )/ \n  > ^ <",
    ]
    SPIN_FRAMES = [
        "  /\\_/\\  \n ( ^.^ ) \n  > ^ <",
        "   (\\  \n    )  \n   /  ",
        "  \\_/\\_ \n   ( \u25cf ) \n  > ^ <",
        "      /) \n     (   \n      \\  ",
        "  /\\_/\\  \n ( ^o^ )~\n  > ^ <",
    ]
    PLAY_DEAD_FRAMES = [
        "  /\\_/\\  \n ( O.O )!\n  > ^ <",
        "  /\\_/\\  \n ( x.x ) \n  > ^ <",
        "         \n  /\\_/\\_ \n  ( x.x )",
        "         \n  /\\_/\\_ \n  ( x.x ) ~",
    ]
    SNEEZE_FRAMES = [
        "  /\\_/\\  \n ( o.o ) \n  > ^ <",
        "  /\\_/\\  \n ( O.O ) \n  > o <",
        "  /\\_/\\  \n (>w< )!! \n  > ^ < !!",
        "  /\\_/\\  \n ( ^.^ ) \n  > ^ < ~",
    ]

    HALLOWEEN_FRAMES = [
        "  /\\_/\\   \U0001f383\n ( o.o )  \n  > ^ <",
        "  /\\_/\\   \U0001f383\n ( O.O )  \n  > W <",
        "  /\\_/\\   \U0001f47b\n ( ^.^ )  \n  > ^ <",
    ]
    WINTER_FRAMES = [
        "  /\\_/\\   \u2744\ufe0f\n ( o.o )  \n  > ^ <",
        "  /\\_/\\   \u2744\ufe0f\n ( ^.^ )  \n  > ^ <",
        "  /\\_/\\   \u2603\ufe0f\n ( u.u )  \n  > ^ <",
    ]

    MESSAGES = {
        "idle":     "...",
        "pet":      "so nice~ \u2665",
        "treat":    "treat?? \U0001f9b4",
        "running":  "nom nom! \U0001f9b4",
        "happy":    "yay!!!! \u2728",
        "loaded":   "new data!! \U0001f4cb",
        "excluded": "oh no... \U0001f622",
        "sniff":    "sniff sniff...",
        "alert":    "! what's that?",
        "kuntal":   "ooh priority! \u2605",
        "stretch":  "zzz... yawn~",
        "cleared":  "phew~ clean!",
        "restored": "yay, back!! \u2705",
        "wag":      "tail wag!!",
        "blink":    "blink~",
        "sleep":    "zzz\u2026",
        "zoomies":  "ZOOMIES!! \u26a1",
        "confused": "huh?",
        "success":  "nice!! \u2705",
        "warning":  "uh oh\u2026 \u26a0\ufe0f",
        "error":    "nope\u2026 \U0001f4a5",
        "legendary": "LEGENDARY BISAAAA \u2605\u2605\u2605",
        "halloween": "spooky Bisa \U0001f383",
        "winter":   "brr\u2026 \u2744\ufe0f",
        "belly":    "belly rubs!! \u2665",
        "milestone": "milestone!!  \u2b50",
        "moveup":    "they moved!! \U0001f4e6",
        "sit":       "good sit!! \U0001f43e",
        "shake":     "nice to meet u! \U0001f91d",
        "spin":      "wheee~! \U0001f300",
        "play_dead": "... \U0001f480 (jk!!)",
        "sneeze":    "ACHOO!! \U0001f927",
    }

    THOUGHTS = [
        "thinking about treats...",
        "I wonder what's in Backstock...",
        "is it lunch yet? \U0001f355",
        "so many barcodes...",
        "~dreaming of zoomies~",
        "*stares at spreadsheet*",
        "who's a good dog? me??",
        "need... more... pets...",
        "what does METRC even mean",
        "tail wag loading... 10%",
        "\u2728 sparkle sparkle \u2728",
        "hmm... sus barcode \U0001f50d",
        "inventory is my passion",
        "*pretends to help*",
        "one more export plz \U0001f4cb",
        "bork? bork.",
        "cannabis... the good stuff \U0001f33f",
    ]

    def __init__(self, parent: tk.Widget):
        self.parent = parent
        self._state = "idle"
        self._after_id = None
        self._idle_idx = 0
        self._anim_idx = 0
        self._anim_frames = []
        self._total_pets = 0
        self._total_treats = 0
        self._total_moveups = 0
        self._interactions_since_milestone = 0
        self._next_milestone_interval = random.randint(60, 100)

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
            text="\u2726 Bisa \u2726",
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
        self.dog_label.bind("<Double-Button-1>", lambda _e: self._sneeze())
        self.dog_label.bind("<Button-3>", lambda _e: self._belly_rub())
        self.dog_label.bind("<Enter>", self._on_hover)

        # Secret trick input buffer
        self._trick_buffer = ""
        self.frame.bind("<Key>", self._on_key)
        self.frame.configure(takefocus=True)
        # Also let clicking the frame give it focus for key events
        self.frame.bind("<Button-1>", lambda e: (self.frame.focus_set(), self._on_frame_click(e)))

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
            text="click \u2192 pet  |  dbl-click \u2192 boop  |  right-click \u2192 belly rub",
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

    TITLES = [
        (0,    "New Pup \U0001f423"),
        (10,   "Good Boy \U0001f415"),
        (50,   "Loyal Friend \U0001f43e"),
        (100,  "Treat Fiend \U0001f9b4"),
        (200,  "Inventory Hound \U0001f4e6"),
        (500,  "Store Guardian \U0001f6e1\ufe0f"),
        (1000, "LEGENDARY BISA \u2605"),
    ]

    def _get_title(self) -> str:
        total = self._total_pets + self._total_treats
        title = self.TITLES[0][1]
        for threshold, t in self.TITLES:
            if total >= threshold:
                title = t
        return title

    def _update_stats(self):
        title = self._get_title()
        self.stats_var.set(
            f"{title}  |  pets:{self._total_pets}  treats:{self._total_treats}  moved:{self._total_moveups}"
        )

    # ------------------------------
    # Animation engine
    # ------------------------------
    def _run_anim(self, frames, msg, speed_ms, on_done):
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
        self._cancel()
        self._after_id = self.parent.after(random.randint(650, 1500), self._idle_tick)

    def _idle_tick(self):
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

        # Original idle cycle — occasionally show a random thought
        self._idle_idx = (self._idle_idx + 1) % len(self.IDLE_FRAMES)
        thought = random.choice(self.THOUGHTS) if random.random() < 0.15 else "..."
        self._render_frame(self.IDLE_FRAMES[self._idle_idx], thought)
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
            if self._maybe_milestone(self._total_treats):
                return

            self._run_anim(go_frames, self.MESSAGES["treat"], int(110 * self._speed_scale),
                           lambda: self._run_anim(self.RUN_BACK, self.MESSAGES["running"], int(110 * self._speed_scale),
                                                  lambda: self._return_idle()))
        except Exception as e:
            print(f"[moveup] Bisa click error: {e}")

    def receive_pet(self):
        if self._state != "idle":
            return

        self._cancel()
        self._state = "pet"
        self._total_pets += 1
        self._update_stats()

        if self._maybe_play_legendary():
            return
        if self._maybe_milestone(self._total_pets):
            return

        self._run_anim(self.PET_FRAMES, self.MESSAGES["pet"], int(180 * self._speed_scale),
                       lambda: self._run_anim(self.HAPPY_FRAMES[:3], self.MESSAGES["pet"], int(180 * self._speed_scale),
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
        if self._maybe_milestone(self._total_treats):
            return

        self._run_anim(go_frames, self.MESSAGES["treat"], int(200 * self._speed_scale),
                       lambda: self._run_anim(self.RUN_BACK, self.MESSAGES["running"], int(200 * self._speed_scale),
                                              lambda: self._return_idle()))

    # ------------------------------
    # New interactions
    # ------------------------------
    def _belly_rub(self):
        if self._state not in ("idle", "pet"):
            return
        self._cancel()
        self._state = "pet"
        self._total_pets += 1
        self._update_stats()
        if self._maybe_play_legendary():
            return
        if self._maybe_milestone(self._total_pets):
            return
        self._run_anim(self.BELLY_FRAMES, self.MESSAGES["belly"],
                       int(150 * self._speed_scale),
                       lambda: self._return_idle())

    def _sneeze(self):
        """Double-click boop -> Bisa sneezes (interrupts pet, doesn't double-count)."""
        if self._state not in ("idle", "pet"):
            return
        was_pet = (self._state == "pet")  # already counted by receive_pet
        self._cancel()
        self._state = "sneeze"
        if not was_pet:
            self._total_pets += 1
            self._update_stats()
        self._run_anim(self.SNEEZE_FRAMES, self.MESSAGES["sneeze"],
                       int(160 * self._speed_scale),
                       lambda: self._return_idle())

    # --- Secret tricks (type while Bisa panel is focused) ---
    TRICKS = {
        "sit":       ("sit",       "SIT_FRAMES"),
        "shake":     ("shake",     "SHAKE_FRAMES"),
        "spin":      ("spin",      "SPIN_FRAMES"),
        "roll":      ("spin",      "SPIN_FRAMES"),      # alias
        "play dead": ("play_dead", "PLAY_DEAD_FRAMES"),
        "dead":      ("play_dead", "PLAY_DEAD_FRAMES"),  # alias
    }

    def _on_key(self, event):
        """Buffer keypresses on Bisa's frame. If the buffer ends with a trick name, she performs it."""
        if not event.char or not event.char.isprintable():
            return
        self._trick_buffer = (self._trick_buffer + event.char.lower())[-12:]  # keep last 12 chars
        for trigger, (msg_key, frames_attr) in self.TRICKS.items():
            if self._trick_buffer.endswith(trigger):
                self._trick_buffer = ""
                self._do_trick(msg_key, getattr(self, frames_attr))
                return

    def _do_trick(self, msg_key: str, frames: list):
        if self._state not in ("idle", "pet", "happy"):
            return
        self._cancel()
        self._state = "trick"
        self._total_pets += 1
        self._update_stats()
        self._run_anim(frames, self.MESSAGES[msg_key],
                       int(170 * self._speed_scale),
                       lambda: self._run_anim(self.HAPPY_FRAMES[:2],
                                              "good dog!! \u2728",
                                              int(150 * self._speed_scale),
                                              lambda: self._return_idle()))

    def _on_hover(self, event=None):
        if self._state != "idle":
            return
        if random.random() < 0.28:
            self._cancel()
            self._state = "sniff"
            self._run_anim(self.SNIFF_FRAMES, self.MESSAGES["sniff"],
                           int(180 * self._speed_scale),
                           lambda: self._return_idle())

    def _maybe_milestone(self, count: int) -> bool:
        """Bisa celebrates at fixed milestones OR every ~80 interactions (+/-20)."""
        self._interactions_since_milestone += 1
        fixed_milestones = {10, 25, 50, 100, 200, 500, 1000}
        interval_hit = self._interactions_since_milestone >= self._next_milestone_interval
        if count not in fixed_milestones and not interval_hit:
            return False
        # Reset interval counter and pick next random threshold
        self._interactions_since_milestone = 0
        self._next_milestone_interval = random.randint(60, 100)
        stars = "\u2b50" * min(5, (count // 100) + 1)
        msg = f"{stars} {count} total!!"
        self._run_anim(
            self.KUNTAL_FRAMES,
            msg,
            int(140 * self._speed_scale),
            lambda: self._run_anim(self.HAPPY_FRAMES[:2], msg,
                                   int(130 * self._speed_scale),
                                   lambda: self._return_idle()),
        )
        return True

    def greet_startup(self):
        """Bisa greets the user based on time of day when the app launches."""
        hour = datetime.now().hour
        if hour < 6:
            msg, frames = "up late?? \U0001f319", self.SLEEP_FRAMES
        elif hour < 12:
            msg, frames = "good morning!! \u2600\ufe0f", self.LOAD_FRAMES[:3]
        elif hour < 17:
            msg, frames = "good afternoon~ \U0001f324\ufe0f", self.HAPPY_FRAMES[:3]
        elif hour < 21:
            msg, frames = "good evening! \U0001f306", self.WAG_FRAMES
        else:
            msg, frames = "working late? \U0001f319", self.BLINK_FRAMES
        self._cancel()
        self._state = "idle"
        self._run_anim(frames, msg, int(200 * self._speed_scale),
                       lambda: self._return_idle())

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

    def react_moveups(self, count: int):
        """Bisa celebrates when SKUs are detected as moved to Sales Floor since last load."""
        self._cancel()
        self._state = "moveup"
        msg = f"{count} SKU{'s' if count != 1 else ''} moved!! \U0001f4e6"
        self._run_anim(
            self.ZOOMIES_FRAMES, msg, int(130 * self._speed_scale),
            lambda: self._run_anim(self.HAPPY_FRAMES[:3], msg, int(160 * self._speed_scale),
                                   lambda: self._return_idle()),
        )

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
    def react_success(self, msg: str = "nice!! \u2705"):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "success"
        frames = self.SUCCESS_FRAMES + self.WAG_FRAMES
        self._run_anim(frames, msg, int(170 * self._speed_scale), lambda: self._return_idle())

    def react_warning(self, msg: str = "uh oh\u2026 \u26a0\ufe0f"):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "warning"
        self._run_anim(self.WARNING_FRAMES, msg, int(220 * self._speed_scale), lambda: self._return_idle())

    def react_error(self, msg: str = "nope\u2026 \U0001f4a5"):
        if self._state != "idle":
            return
        self._cancel()
        self._state = "error"
        self._run_anim(self.CONFUSED_FRAMES, msg, int(210 * self._speed_scale), lambda: self._return_idle())
