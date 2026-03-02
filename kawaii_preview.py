# kawaii_preview.py
# Popup: Kawaii PDF Settings + Live Preview
# Uses kawaii_settings.py for persistence and math.

from __future__ import annotations

import math
import os
import random
import tkinter as tk
from tkinter import ttk

from kawaii_settings import (
    KawaiiSettings,
    PRESETS_COLOR,
    load_settings,
    save_settings,
    reset_defaults,
    compute_effective_profile,
)

def _rgb01_to_hex(rgb):
    r = int(max(0, min(1, rgb[0])) * 255)
    g = int(max(0, min(1, rgb[1])) * 255)
    b = int(max(0, min(1, rgb[2])) * 255)
    return f"#{r:02x}{g:02x}{b:02x}"

def _mix_hex(a_hex: str, b_hex: str, t: float) -> str:
    t = max(0.0, min(1.0, t))
    a_hex = a_hex.lstrip("#")
    b_hex = b_hex.lstrip("#")
    ar, ag, ab = int(a_hex[0:2], 16), int(a_hex[2:4], 16), int(a_hex[4:6], 16)
    br, bg, bb = int(b_hex[0:2], 16), int(b_hex[2:4], 16), int(b_hex[4:6], 16)
    r = int(ar + (br - ar) * t)
    g = int(ag + (bg - ag) * t)
    b = int(ab + (bb - ab) * t)
    return f"#{r:02x}{g:02x}{b:02x}"

def _blend_over_white(hex_color: str, alpha: float) -> str:
    alpha = max(0.0, min(1.0, alpha))
    s = hex_color.lstrip("#")
    r, g, b = int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)
    out_r = int(alpha * r + (1 - alpha) * 255)
    out_g = int(alpha * g + (1 - alpha) * 255)
    out_b = int(alpha * b + (1 - alpha) * 255)
    return f"#{out_r:02x}{out_g:02x}{out_b:02x}"


class KawaiiPreviewDialog:
    def __init__(self, parent: tk.Tk | tk.Toplevel):
        self.win = tk.Toplevel(parent)
        self.win.title("Kawaii PDF Settings (Live Preview)")
        self.win.geometry("1100x720")
        self.win.transient(parent)
        self.win.grab_set()

        self.settings = load_settings()

        # variables
        self.printer_bw = tk.BooleanVar(value=self.settings.printer_bw)
        self.preset_var = tk.StringVar(value=self.settings.preset)

        self.bg_hue_pct = tk.IntVar(value=self.settings.bg_hue_pct)                 # 0..100
        self.bg_intensity_pct = tk.IntVar(value=self.settings.bg_intensity_pct)     # 0..300
        self.elem_intensity_pct = tk.IntVar(value=self.settings.elem_intensity_pct) # 0..300

        self.readout = tk.StringVar(value="")

        self._build_ui()
        self.redraw()

        self.win.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        outer = ttk.Frame(self.win, padding=10)
        outer.pack(fill="both", expand=True)

        left = ttk.Frame(outer)
        left.pack(side="left", fill="y", padx=(0, 10))

        right = ttk.Frame(outer)
        right.pack(side="left", fill="both", expand=True)

        self.canvas = tk.Canvas(right, width=780, height=640, bg="white", highlightthickness=1)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda _e: self.redraw())

        ttk.Label(left, text="Controls", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 10))

        ttk.Checkbutton(
            left,
            text="Printer B/W mode",
            variable=self.printer_bw,
            command=self._save_now
        ).pack(anchor="w", pady=(0, 10))

        row = ttk.Frame(left)
        row.pack(fill="x", pady=(0, 10))
        ttk.Label(row, text="Preset", width=8).pack(side="left")
        preset_cb = ttk.Combobox(row, textvariable=self.preset_var, values=list(PRESETS_COLOR.keys()),
                                 state="readonly", width=12)
        preset_cb.pack(side="left", padx=6)
        preset_cb.bind("<<ComboboxSelected>>", lambda _e: self._save_now())

        ttk.Label(left, text="BG Hue (Pink → Purple)").pack(anchor="w")
        ttk.Scale(left, from_=0, to=100, orient="horizontal",
                  variable=self.bg_hue_pct, command=lambda _v: self._save_now()).pack(fill="x", pady=(2, 10))

        ttk.Label(left, text="BG Intensity (tint alpha)").pack(anchor="w")
        ttk.Scale(left, from_=0, to=300, orient="horizontal",
                  variable=self.bg_intensity_pct, command=lambda _v: self._save_now()).pack(fill="x", pady=(2, 10))

        ttk.Label(left, text="Element Intensity (everything pretty)").pack(anchor="w")
        ttk.Scale(left, from_=0, to=300, orient="horizontal",
                  variable=self.elem_intensity_pct, command=lambda _v: self._save_now()).pack(fill="x", pady=(2, 12))

        btn_row = ttk.Frame(left)
        btn_row.pack(fill="x", pady=(8, 6))
        ttk.Button(btn_row, text="Reset to Defaults", command=self._reset).pack(side="left")
        ttk.Button(btn_row, text="Save", command=self._save_now).pack(side="left", padx=8)
        ttk.Button(btn_row, text="Close", command=self._on_close).pack(side="left", padx=8)

        ttk.Label(left, textvariable=self.readout, foreground="#444").pack(anchor="w", pady=(12, 0))

    def _to_settings(self) -> KawaiiSettings:
        s = KawaiiSettings(
            preset=str(self.preset_var.get()),
            printer_bw=bool(self.printer_bw.get()),
            bg_hue_pct=int(self.bg_hue_pct.get()),
            bg_intensity_pct=int(self.bg_intensity_pct.get()),
            elem_intensity_pct=int(self.elem_intensity_pct.get()),
            # Preserve stars tuning from loaded settings (no UI slider for these)
            stars_base=self.settings.stars_base,
            stars_max_extra=self.settings.stars_max_extra,
        )
        s.clamp_self()
        return s

    def _save_now(self):
        if getattr(self, "_resetting", False):
            return
        s = self._to_settings()
        save_settings(s)
        self.settings = s
        self.redraw()

    def _reset(self):
        self._resetting = True
        s = reset_defaults()
        self.printer_bw.set(s.printer_bw)
        self.preset_var.set(s.preset)
        self.bg_hue_pct.set(s.bg_hue_pct)
        self.bg_intensity_pct.set(s.bg_intensity_pct)
        self.elem_intensity_pct.set(s.elem_intensity_pct)
        self._resetting = False
        save_settings(self._to_settings())
        self.redraw()

    def redraw(self):
        c = self.canvas
        c.delete("all")

        w = max(10, int(c.winfo_width()))
        h = max(10, int(c.winfo_height()))

        prof = compute_effective_profile(self._to_settings())

        tint_hex = _rgb01_to_hex(prof["tint_rgb"])
        stroke_hex = _rgb01_to_hex(prof["stroke_rgb"])

        tint_a = float(prof["tint_alpha"])
        stroke_a = float(prof["stroke_alpha"])
        sparkle_a = float(prof["sparkle_alpha"])
        border_a = float(prof["border_alpha"])

        # background wash
        bg = _blend_over_white(tint_hex, tint_a)
        c.create_rectangle(0, 0, w, h, fill=bg, outline="")

        # pseudo-alpha for lines (mix toward bg)
        line_col = _mix_hex(bg, stroke_hex, max(0.0, min(1.0, stroke_a * 5.5)))
        sparkle_col = _mix_hex(bg, stroke_hex, max(0.0, min(1.0, sparkle_a * 5.5)))
        border_col = _mix_hex(bg, stroke_hex, max(0.0, min(1.0, border_a * 5.5)))

        # border
        margin = 24
        top_pad = 24
        c.create_rectangle(margin, margin + top_pad, w - margin, h - margin, outline=border_col, width=3)

        # daisy watermark
        cx, cy = int(w * 0.64), int(h * 0.52)
        petals = 12
        petal_r = int(min(w, h) * 0.08)
        petal_dist = int(min(w, h) * 0.135)

        for i in range(petals):
            ang = (i / petals) * (2 * math.pi)
            px = cx + int(petal_dist * math.cos(ang))
            py = cy + int(petal_dist * math.sin(ang))
            c.create_oval(px - petal_r, py - petal_r, px + petal_r, py + petal_r, outline=line_col, width=3)

        center_r = int(petal_r * 1.12)
        c.create_oval(cx - center_r, cy - center_r, cx + center_r, cy + center_r, outline=line_col, width=3)

        # stars: random spread, count increases linearly with element intensity
        stars = int(prof.get("stars_count", 10))
        rng = random.Random(1337)  # fixed seed so it doesn’t flicker while sliding
        for _ in range(stars):
            x = rng.randint(int(w * 0.08), int(w * 0.92))
            y = rng.randint(int(h * 0.10), int(h * 0.90))
            r = rng.randint(6, 14)
            self._star(c, x, y, r, sparkle_col)

        # corner daisies: count scales with element intensity (same pool order as PDF)
        daisy_count = int(prof.get("daisy_count", 6))
        daisy_positions = [
            (w * 0.11, h * 0.86, 0.95),
            (w * 0.89, h * 0.86, 0.95),
            (w * 0.12, h * 0.18, 0.85),
            (w * 0.88, h * 0.18, 0.85),
            (w * 0.16, h * 0.83, 0.70),
            (w * 0.84, h * 0.83, 0.70),
            (w * 0.05, h * 0.52, 0.75),
            (w * 0.95, h * 0.52, 0.75),
            (w * 0.22, h * 0.50, 0.55),
            (w * 0.78, h * 0.50, 0.55),
        ]
        d_r = int(min(w, h) * 0.025)  # small fixed radius for corner daisies
        for dx, dy, _ in daisy_positions[:daisy_count]:
            self._daisy(c, int(dx), int(dy), d_r, line_col)

        mode = "B/W" if prof["printer_bw"] else "Color"
        paw_count = int(prof.get("paw_count", 4))
        preset_name = prof.get("preset", "")
        bg_hue = prof.get("bg_hue_pct", 0)
        bg_int = prof.get("bg_intensity_pct", 100)
        el_int = prof.get("elem_intensity_pct", 100)
        self.readout.set(
            f"Mode: {mode} | Preset: {preset_name} | Hue: {bg_hue}% Purple | "
            f"BG: {bg_int}% | Elem: {el_int}%\n"
            f"Effective a tint/stroke/sparkle/border: "
            f"{tint_a:.3f}/{stroke_a:.3f}/{sparkle_a:.3f}/{border_a:.3f} | "
            f"Stars: {stars} | Daisies: {daisy_count} | Paws: {paw_count}"
        )

    @staticmethod
    def _star(c: tk.Canvas, x: int, y: int, r: int, color: str):
        c.create_line(x - r, y, x + r, y, fill=color, width=3)
        c.create_line(x, y - r, x, y + r, fill=color, width=3)
        c.create_line(x - int(r * 0.7), y - int(r * 0.7), x + int(r * 0.7), y + int(r * 0.7), fill=color, width=3)
        c.create_line(x - int(r * 0.7), y + int(r * 0.7), x + int(r * 0.7), y - int(r * 0.7), fill=color, width=3)

    @staticmethod
    def _daisy(c: tk.Canvas, x: int, y: int, r: int, color: str):
        petals = 10
        petal_r = r
        petal_dist = int(r * 1.7)
        for i in range(petals):
            ang = (i / petals) * (2 * math.pi)
            px = x + int(petal_dist * math.cos(ang))
            py = y + int(petal_dist * math.sin(ang))
            c.create_oval(px - petal_r, py - petal_r, px + petal_r, py + petal_r, outline=color, width=2)
        center_r = int(r * 0.7)
        c.create_oval(x - center_r, y - center_r, x + center_r, y + center_r, outline=color, width=2)

    def _on_close(self):
        # always save current state on close
        save_settings(self._to_settings())
        try:
            self.win.grab_release()
        except Exception:
            pass
        self.win.destroy()


def open_kawaii_settings_window(parent):
    KawaiiPreviewDialog(parent)



# standalone test
def main():
    root = tk.Tk()
    root.title("Kawaii Preview Test Host")
    ttk.Button(root, text="Open Kawaii Settings…", command=lambda: open_kawaii_settings_window(root)).pack(padx=20, pady=20)
    root.geometry("360x120")
    root.mainloop()

if __name__ == "__main__":
    main()