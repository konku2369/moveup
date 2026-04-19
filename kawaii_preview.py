"""
Live preview dialog for kawaii PDF settings.

Tk Canvas-based preview that mirrors pdf_export.py decoration rendering
in real time as the user adjusts sliders. Persists via kawaii_settings.py.
"""

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
    """Convert a 0–1 float RGB tuple to a Tk-compatible '#rrggbb' hex string.

    Parameters
    ----------
    rgb : tuple[float, float, float]
        RGB components in [0.0, 1.0]. Values are clamped before conversion.

    Returns
    -------
    str
        Hex color string, e.g. ``'#ff80a0'``.
    """
    r = int(max(0, min(1, rgb[0])) * 255)
    g = int(max(0, min(1, rgb[1])) * 255)
    b = int(max(0, min(1, rgb[2])) * 255)
    return f"#{r:02x}{g:02x}{b:02x}"

def _mix_hex(a_hex: str, b_hex: str, t: float) -> str:
    """Linearly interpolate between two hex colors.

    Parameters
    ----------
    a_hex : str
        Start color (t=0), with or without leading '#'.
    b_hex : str
        End color (t=1), with or without leading '#'.
    t : float
        Blend factor in [0.0, 1.0]; clamped before use. 0 → a, 1 → b.

    Returns
    -------
    str
        Interpolated '#rrggbb' hex string.
    """
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
    """Composite *hex_color* at *alpha* over a pure white background.

    Implements the standard Porter-Duff 'over white' formula:
      out = alpha * color + (1 − alpha) * 255

    Used by the preview canvas to approximate the semi-transparent PDF tint wash,
    since Tk Canvas doesn't support real alpha fills.

    Parameters
    ----------
    hex_color : str
        Foreground color, with or without leading '#'.
    alpha : float
        Opacity in [0.0, 1.0]; 0 → white, 1 → full color.

    Returns
    -------
    str
        '#rrggbb' hex string of the blended result.
    """
    alpha = max(0.0, min(1.0, alpha))
    s = hex_color.lstrip("#")
    r, g, b = int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)
    out_r = int(alpha * r + (1 - alpha) * 255)
    out_g = int(alpha * g + (1 - alpha) * 255)
    out_b = int(alpha * b + (1 - alpha) * 255)
    return f"#{out_r:02x}{out_g:02x}{out_b:02x}"


class KawaiiPreviewDialog:
    """Modal settings dialog with real-time kawaii decoration preview.

    Displays a Tk Canvas that mirrors the decoration rendering from pdf_export.py —
    background tint, border, daisy watermark, sparkle stars, corner daisies, and
    cat faces — updated live as the user moves sliders.

    Controls exposed:
    - Printer B/W mode checkbox
    - Preset selector (Cute / Elegant / etc.)
    - BG Hue slider (0 = pink, 100 = purple)
    - Element Intensity slider (0–300%)
    - Reset to Defaults / Save / Close buttons
    - "Generate Test PDF" button — exports a 10-item sample PDF and auto-opens it

    Settings are persisted to kawaii_pdf_settings.json via kawaii_settings.save_settings()
    on every slider change and on close, so the preview state and the live app state
    are always in sync.

    Parameters
    ----------
    parent : tk.Tk or tk.Toplevel
        Owner window; dialog is modal (grab_set) relative to *parent*.
    """

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
        self.elem_intensity_pct = tk.IntVar(value=self.settings.elem_intensity_pct) # 0..300

        self.readout = tk.StringVar(value="")

        self._build_ui()
        self.redraw()

        self.win.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        """Construct the two-panel layout: left=controls, right=canvas preview."""
        outer = ttk.Frame(self.win, padding=10)
        outer.pack(fill="both", expand=True)

        left = ttk.Frame(outer)
        left.pack(side="left", fill="y", padx=(0, 10))

        right = ttk.Frame(outer)
        right.pack(side="left", fill="both", expand=True)

        self.canvas = tk.Canvas(right, width=780, height=640, bg="white", highlightthickness=1)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda _e: self.redraw())  # redraws on window resize so preview fills the panel

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

        ttk.Label(left, text="Element Intensity (everything pretty)").pack(anchor="w")
        ttk.Scale(left, from_=0, to=300, orient="horizontal",
                  variable=self.elem_intensity_pct, command=lambda _v: self._save_now()).pack(fill="x", pady=(2, 12))

        btn_row = ttk.Frame(left)
        btn_row.pack(fill="x", pady=(8, 6))
        ttk.Button(btn_row, text="Reset to Defaults", command=self._reset).pack(side="left")
        ttk.Button(btn_row, text="Save", command=self._save_now).pack(side="left", padx=8)
        ttk.Button(btn_row, text="Close", command=self._on_close).pack(side="left", padx=8)

        test_row = ttk.Frame(left)
        test_row.pack(fill="x", pady=(0, 4))
        ttk.Button(test_row, text="📄 Generate Test PDF", command=self._generate_test_pdf).pack(side="left")

        ttk.Label(left, textvariable=self.readout, foreground="#444").pack(anchor="w", pady=(12, 0))

    def _to_settings(self) -> KawaiiSettings:
        """Snapshot current control state into a clamped KawaiiSettings object.

        stars_base and stars_max_extra are preserved from the last loaded settings because
        there is no UI slider for those fields — they can only be changed by editing
        kawaii_pdf_settings.json directly.
        """
        s = KawaiiSettings(
            preset=str(self.preset_var.get()),
            printer_bw=bool(self.printer_bw.get()),
            bg_hue_pct=int(self.bg_hue_pct.get()),
            bg_intensity_pct=100,
            elem_intensity_pct=int(self.elem_intensity_pct.get()),
            # Preserve stars tuning from loaded settings (no UI slider for these)
            stars_base=self.settings.stars_base,
            stars_max_extra=self.settings.stars_max_extra,
        )
        s.clamp_self()
        return s

    def _save_now(self):
        """Persist current settings and refresh the canvas preview.

        Called on every slider change and checkbox toggle. The _resetting guard prevents
        recursive saves while _reset() is programmatically updating the control variables.
        """
        if getattr(self, "_resetting", False):
            return
        s = self._to_settings()
        save_settings(s)
        self.settings = s
        self.redraw()

    def _reset(self):
        """Reset all controls to factory defaults and persist the reset settings.

        Sets _resetting=True while updating Tk variables so _save_now() no-ops on each
        individual variable change; a single save is done after all variables are set.
        """
        self._resetting = True
        s = reset_defaults()
        self.printer_bw.set(s.printer_bw)
        self.preset_var.set(s.preset)
        self.bg_hue_pct.set(s.bg_hue_pct)
        self.elem_intensity_pct.set(s.elem_intensity_pct)
        self._resetting = False
        save_settings(self._to_settings())
        self.redraw()

    def redraw(self):
        """Repaint the full preview canvas from the current settings.

        Rendering steps mirror pdf_export._draw_kawaii_background() as closely as possible
        using Tk Canvas primitives instead of ReportLab canvas calls:
          1. Solid background fill (tint blended over white via _blend_over_white)
          2. Border rectangle
          3. Large watermark daisy (center-right)
          4. Random sparkle stars (count from profile stars_count)
          5. Corner daisies (count from profile daisy_count, same pool-order as PDF)
          6. Cat faces scattered in margin bands (count from profile cat_count)

        Also updates the readout label with current effective alpha values and element counts.

        Note: Tk Canvas doesn't support real alpha compositing, so line colors are blended
        toward the background using _mix_hex() to approximate transparency.
        """
        c = self.canvas
        c.delete("all")

        w = max(10, int(c.winfo_width()))
        h = max(10, int(c.winfo_height()))

        prof = compute_effective_profile(self._to_settings())

        tint_hex = _rgb01_to_hex(prof.get("tint_rgb", (0.95, 0.85, 0.90)))
        stroke_hex = _rgb01_to_hex(prof.get("stroke_rgb", (0.55, 0.40, 0.50)))

        tint_a = float(prof.get("tint_alpha", 0.055))
        stroke_a = float(prof.get("stroke_alpha", 0.08))
        sparkle_a = float(prof.get("sparkle_alpha", 0.07))
        border_a = float(prof.get("border_alpha", 0.10))

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

        cat_count = int(prof.get("cat_count", 18))
        cat_rng = random.Random(1338)
        for _ in range(cat_count):
            band = cat_rng.randint(0, 3)
            if band == 0:
                cx2 = cat_rng.randint(int(w * 0.08), int(w * 0.92))
                cy2 = cat_rng.randint(int(h * 0.02), int(h * 0.12))
            elif band == 1:
                cx2 = cat_rng.randint(int(w * 0.08), int(w * 0.92))
                cy2 = cat_rng.randint(int(h * 0.88), int(h * 0.98))
            elif band == 2:
                cx2 = cat_rng.randint(int(w * 0.01), int(w * 0.09))
                cy2 = cat_rng.randint(int(h * 0.12), int(h * 0.88))
            else:
                cx2 = cat_rng.randint(int(w * 0.91), int(w * 0.99))
                cy2 = cat_rng.randint(int(h * 0.12), int(h * 0.88))
            cat_r = cat_rng.randint(int(min(w, h) * 0.018), int(min(w, h) * 0.038))
            self._cat_face(c, cx2, cy2, cat_r, line_col)

        mode = "B/W" if prof.get("printer_bw", False) else "Color"
        paw_count = int(prof.get("paw_count", 4))
        preset_name = prof.get("preset", "")
        bg_hue = prof.get("bg_hue_pct", 0)
        el_int = prof.get("elem_intensity_pct", 100)
        self.readout.set(
            f"Mode: {mode} | Preset: {preset_name} | Hue: {bg_hue}% Purple | Elem: {el_int}%\n"
            f"Effective a tint/stroke/sparkle/border: "
            f"{tint_a:.3f}/{stroke_a:.3f}/{sparkle_a:.3f}/{border_a:.3f} | "
            f"Stars: {stars} | Daisies: {daisy_count} | Paws: {paw_count} | Cats: {cat_count}"
        )

    @staticmethod
    def _star(c: tk.Canvas, x: int, y: int, r: int, color: str):
        """Draw a 4-line sparkle star on the Tk Canvas at (x, y) with radius r."""
        c.create_line(x - r, y, x + r, y, fill=color, width=3)
        c.create_line(x, y - r, x, y + r, fill=color, width=3)
        c.create_line(x - int(r * 0.7), y - int(r * 0.7), x + int(r * 0.7), y + int(r * 0.7), fill=color, width=3)
        c.create_line(x - int(r * 0.7), y + int(r * 0.7), x + int(r * 0.7), y - int(r * 0.7), fill=color, width=3)

    @staticmethod
    def _daisy(c: tk.Canvas, x: int, y: int, r: int, color: str):
        """Draw a 10-petal daisy on the Tk Canvas. Petal center offset is 1.7×r."""
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

    @staticmethod
    def _cat_face(c: tk.Canvas, x: int, y: int, r: int, color: str):
        """Draw a minimal Tk Canvas cat face: head circle, ear triangles, eyes, nose, whiskers."""
        # Head
        c.create_oval(x - r, y - r, x + r, y + r, outline=color, width=2)
        # Ears — two triangles
        for side in (-1, 1):
            ex = x + side * int(r * 0.72)
            ey_base = y - int(r * 0.6)
            tip_x = ex + side * int(r * 0.28)
            tip_y = ey_base - int(r * 0.85)
            c.create_polygon(
                ex - int(r * 0.44), ey_base,
                ex + int(r * 0.44), ey_base,
                tip_x, tip_y,
                outline=color, fill="", width=2
            )
        # Eyes
        er = max(2, int(r * 0.20))
        c.create_oval(x - int(r * 0.35) - er, y - int(r * 0.15) - er,
                      x - int(r * 0.35) + er, y - int(r * 0.15) + er, outline=color, width=2)
        c.create_oval(x + int(r * 0.35) - er, y - int(r * 0.15) - er,
                      x + int(r * 0.35) + er, y - int(r * 0.15) + er, outline=color, width=2)
        # Nose
        nr = max(1, int(r * 0.11))
        c.create_oval(x - nr, y + int(r * 0.17) - nr, x + nr, y + int(r * 0.17) + nr,
                      outline=color, width=2)
        # Whiskers — 3 per side
        wlen = int(r * 0.78)
        wy = y + int(r * 0.17)
        for side in (-1, 1):
            x0 = x + side * int(r * 0.20)
            for dy in (-int(wlen * 0.35), 0, int(wlen * 0.35)):
                c.create_line(x0, wy, x0 + side * wlen, wy + dy, fill=color, width=1)

    def _generate_test_pdf(self):
        """Generate a small sample move-up PDF with the current kawaii settings and open it."""
        try:
            import tempfile
            import pandas as pd
            from pdf_export import export_moveup_pdf_paginated

            self._save_now()  # make sure current settings are persisted before generating

            dummy = pd.DataFrame([
                ["1A4050300012345", "Flower",        "Green Thumb",  "Blue Dream 3.5g",            "Backstock", 4],
                ["1A4050300012346", "Flower",        "Harvest Moon", "OG Kush 1g Pre-Pack",        "Backstock", 12],
                ["1A4050300012347", "Vapes",         "CloudCo",      "Pineapple Express Cart 1g",  "Backstock", 7],
                ["1A4050300012348", "Edibles",       "Sweet Leaf",   "Gummies Watermelon 100mg",   "Backstock", 3],
                ["1A4050300012349", "Concentrates",  "Wax Works",    "Live Resin 1g",              "Backstock", 5],
                ["1A4050300012350", "Pre-Rolls",     "RollEasy",     "Pre-Roll 5 Pack",            "Backstock", 8],
                ["1A4050300012351", "Flower",        "Green Thumb",  "Gelato #41 7g",              "Backstock", 2],
                ["1A4050300012352", "Tinctures",     "HerbCraft",    "CBD:THC 1:1 Tincture 30ml",  "Backstock", 6],
                ["1A4050300012353", "Flower",        "Harvest Moon", "Sunset Sherbet 3.5g",        "Backstock", 9],
                ["1A4050300012354", "Vapes",         "CloudCo",      "Watermelon Zkittlez Cart",   "Backstock", 4],
            ], columns=["Package Barcode", "Type", "Brand", "Product Name", "Room", "Qty On Hand"])

            tmp_dir = tempfile.mkdtemp(prefix="moveup_kawaii_test_")
            printer_bw = bool(self.printer_bw.get())
            export_moveup_pdf_paginated(
                move_up_df=dummy,
                priority_df=None,
                base_dir=tmp_dir,
                timestamp=False,
                prefix="KawaiiTest",
                auto_open=True,
                kawaii_pdf=True,
                printer_bw=printer_bw,
            )
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror("Test PDF Error", str(e), parent=self.win)

    def _on_close(self):
        """Save current settings, release the modal grab, and destroy the window."""
        # always save current state on close
        save_settings(self._to_settings())
        try:
            self.win.grab_release()
        except Exception:
            pass
        self.win.destroy()


def open_kawaii_settings_window(parent):
    """Open the kawaii PDF settings dialog as a modal child of *parent*.

    The dialog is grab-set (modal) so the user must close it before interacting
    with the main window. Settings are auto-saved to kawaii_pdf_settings.json on
    every change and again on close — there is no 'Cancel' that discards changes.

    Parameters
    ----------
    parent : tk.Tk or tk.Toplevel
        Owner window for the modal dialog.
    """
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