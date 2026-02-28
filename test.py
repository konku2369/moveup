# kawaii_preview.py
# Kawaii PDF Background Preview tool (4 sliders + feral starfield)
# - Slider 1: Element intensity (stroke + sparkle + border alphas together)
# - Slider 2: Background intensity (tint alpha only)
# - Slider 3: Background hue (Pink <-> Purple blend)
# - Slider 4: Hot Magenta Inject (adds feral magenta bias)
# - Presets: Minimal / Cute / Extra (single preset table)
# - Stars increase linearly with element intensity, scattered randomly across page (stable seed)
# - Persists to kawaii_preview_config.json

import json
import os
import math
import random
import tkinter as tk
from tkinter import ttk

CONFIG_FILENAME = "kawaii_preview_config.json"

# Endpoints for hue blending (established vibe)
PINK_TINT_HEX = "#ffd6ea"
PURPLE_TINT_HEX = "#e6d9ff"

PINK_STROKE_HEX = "#8c667f"    # close to your Color(0.55,0.40,0.50)
PURPLE_STROKE_HEX = "#6f6292"  # lavender-ish stroke

# Feral injection color (pure chaos)
MAGENTA_INJECT_HEX = "#ff2bd6"  # hot magenta

# -----------------------------
# One preset table (base alphas)
# -----------------------------
PRESET_BASES = {
    "Minimal": {
        "tint_alpha":    0.050,
        "stroke_alpha":  0.090,
        "sparkle_alpha": 0.085,
        "border_alpha":  0.110,
    },
    "Cute": {
        "tint_alpha":    0.140,
        "stroke_alpha":  0.170,
        "sparkle_alpha": 0.160,
        "border_alpha":  0.210,
    },
    # ☣️ FERAL VARIANT MODE ☣️
    "Extra": {
        "tint_alpha":    0.55,
        "stroke_alpha":  0.42,
        "sparkle_alpha": 0.40,
        "border_alpha":  0.50,
    },
}

# B/W mode preset table (kept conservative)
PRESETS_BW = {
    "Minimal": {"tint_alpha": 0.004, "stroke_alpha": 0.040, "sparkle_alpha": 0.035, "border_alpha": 0.050},
    "Cute":    {"tint_alpha": 0.010, "stroke_alpha": 0.070, "sparkle_alpha": 0.060, "border_alpha": 0.080},
    "Extra":   {"tint_alpha": 0.020, "stroke_alpha": 0.110, "sparkle_alpha": 0.100, "border_alpha": 0.130},
}

# Raised ceilings (you asked for feral)
LIMITS = {
    "tint_alpha":    (0.0, 1.00),
    "stroke_alpha":  (0.0, 0.95),
    "sparkle_alpha": (0.0, 0.95),
    "border_alpha":  (0.0, 0.95),
}

DEFAULTS = {
    "printer_bw": False,
    "preset": "Cute",
    # sliders in percent:
    # - bg_hue_pct: 0=Pink, 100=Purple
    # - bg_intensity_pct: multiplies only tint alpha
    # - elem_intensity_pct: multiplies stroke/sparkle/border
    # - magenta_pct: injects hot magenta into tint+stroke colors
    "bg_hue_pct": 0,           # default full kawaii pink hue
    "bg_intensity_pct": 100,
    "elem_intensity_pct": 100,
    "magenta_pct": 0,
}

# Slider ranges (tune here)
BG_INTENSITY_MAX = 600
ELEM_INTENSITY_MAX = 600


def clamp(v: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, v))


def hex_to_rgb(hex_color: str):
    s = (hex_color or "").strip().lstrip("#")
    if len(s) != 6:
        raise ValueError("Expected a 6-digit hex color like #ffd6ea")
    r = int(s[0:2], 16)
    g = int(s[2:4], 16)
    b = int(s[4:6], 16)
    return r, g, b


def rgb_to_hex(r: int, g: int, b: int) -> str:
    r = int(clamp(r, 0, 255))
    g = int(clamp(g, 0, 255))
    b = int(clamp(b, 0, 255))
    return f"#{r:02x}{g:02x}{b:02x}"


def mix(hex_a: str, hex_b: str, t: float) -> str:
    """Linear mix: t=0 => a, t=1 => b"""
    t = clamp(t, 0.0, 1.0)
    ar, ag, ab = hex_to_rgb(hex_a)
    br, bg, bb = hex_to_rgb(hex_b)
    r = int(ar + (br - ar) * t)
    g = int(ag + (bg - ag) * t)
    b = int(ab + (bb - ab) * t)
    return rgb_to_hex(r, g, b)


def blend_over_white(hex_color: str, alpha: float) -> str:
    alpha = clamp(alpha, 0.0, 1.0)
    r, g, b = hex_to_rgb(hex_color)
    out_r = int(alpha * r + (1 - alpha) * 255)
    out_g = int(alpha * g + (1 - alpha) * 255)
    out_b = int(alpha * b + (1 - alpha) * 255)
    return rgb_to_hex(out_r, out_g, out_b)


def inject_magenta(base_hex: str, magenta_hex: str, amt: float) -> str:
    """
    Adds a hot-magenta bias. amt: 0..1.
    This is NOT just a mix. It "pushes" toward magenta but keeps some base character.
    """
    amt = clamp(amt, 0.0, 1.0)
    br, bg, bb = hex_to_rgb(base_hex)
    mr, mg, mb = hex_to_rgb(magenta_hex)

    # nonlinear ramp so the first half is subtle, second half gets wild
    a = amt ** 1.6

    r = int(br + (mr - br) * a)
    g = int(bg + (mg - bg) * a)
    b = int(bb + (mb - bb) * a)
    return rgb_to_hex(r, g, b)


class KawaiiPreviewApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Kawaii PDF Background Preview (Feral Mode)")
        self.root.geometry("1080x720")

        self.app_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_path = os.path.join(self.app_dir, CONFIG_FILENAME)

        self.printer_bw = tk.BooleanVar(value=DEFAULTS["printer_bw"])
        self.preset_var = tk.StringVar(value=DEFAULTS["preset"])

        self.bg_hue_pct = tk.IntVar(value=DEFAULTS["bg_hue_pct"])                  # 0..100
        self.bg_intensity_pct = tk.IntVar(value=DEFAULTS["bg_intensity_pct"])      # 0..BG_INTENSITY_MAX
        self.elem_intensity_pct = tk.IntVar(value=DEFAULTS["elem_intensity_pct"])  # 0..ELEM_INTENSITY_MAX
        self.magenta_pct = tk.IntVar(value=DEFAULTS["magenta_pct"])                # 0..100

        self.readout = tk.StringVar(value="")
        self._build_ui()
        self._load_config()
        self.redraw()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # --------------------------
    # Config
    # --------------------------
    def _load_config(self):
        if not os.path.exists(self.config_path):
            self._reset_defaults(save=False)
            return
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)

            self.printer_bw.set(bool(cfg.get("printer_bw", DEFAULTS["printer_bw"])))

            preset = str(cfg.get("preset", DEFAULTS["preset"]))
            if preset not in PRESET_BASES:
                preset = DEFAULTS["preset"]
            self.preset_var.set(preset)

            def load_int(key, default, lo, hi):
                v = cfg.get(key, default)
                try:
                    v = int(v)
                except Exception:
                    v = default
                return int(clamp(v, lo, hi))

            self.bg_hue_pct.set(load_int("bg_hue_pct", DEFAULTS["bg_hue_pct"], 0, 100))
            self.bg_intensity_pct.set(load_int("bg_intensity_pct", DEFAULTS["bg_intensity_pct"], 0, BG_INTENSITY_MAX))
            self.elem_intensity_pct.set(load_int("elem_intensity_pct", DEFAULTS["elem_intensity_pct"], 0, ELEM_INTENSITY_MAX))
            self.magenta_pct.set(load_int("magenta_pct", DEFAULTS["magenta_pct"], 0, 100))

        except Exception:
            self._reset_defaults(save=False)

    def _save_config(self):
        try:
            cfg = {
                "printer_bw": bool(self.printer_bw.get()),
                "preset": str(self.preset_var.get()),
                "bg_hue_pct": int(self.bg_hue_pct.get()),
                "bg_intensity_pct": int(self.bg_intensity_pct.get()),
                "elem_intensity_pct": int(self.elem_intensity_pct.get()),
                "magenta_pct": int(self.magenta_pct.get()),
            }
            tmp = self.config_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
            os.replace(tmp, self.config_path)
        except Exception:
            pass

    def _reset_defaults(self, save: bool = True):
        self.printer_bw.set(DEFAULTS["printer_bw"])
        self.preset_var.set(DEFAULTS["preset"])
        self.bg_hue_pct.set(DEFAULTS["bg_hue_pct"])
        self.bg_intensity_pct.set(DEFAULTS["bg_intensity_pct"])
        self.elem_intensity_pct.set(DEFAULTS["elem_intensity_pct"])
        self.magenta_pct.set(DEFAULTS["magenta_pct"])
        if save:
            self._save_config()
        self.redraw()

    # --------------------------
    # UI
    # --------------------------
    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=10)
        outer.pack(fill="both", expand=True)

        left = ttk.Frame(outer)
        left.pack(side="left", fill="y", padx=(0, 10))

        right = ttk.Frame(outer)
        right.pack(side="left", fill="both", expand=True)

        self.canvas = tk.Canvas(right, width=760, height=640, bg="white", highlightthickness=1)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda _e: self.redraw())

        ttk.Label(left, text="Controls", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 8))

        ttk.Checkbutton(
            left,
            text="Printer B/W mode (preview)",
            variable=self.printer_bw,
            command=self._changed
        ).pack(anchor="w", pady=(0, 10))

        row = ttk.Frame(left)
        row.pack(fill="x", pady=(0, 10))
        ttk.Label(row, text="Preset", width=8).pack(side="left")
        preset_cb = ttk.Combobox(row, textvariable=self.preset_var, values=list(PRESET_BASES.keys()),
                                 state="readonly", width=12)
        preset_cb.pack(side="left", padx=6)
        preset_cb.bind("<<ComboboxSelected>>", lambda _e: self._changed())

        # BG Hue slider
        ttk.Label(left, text="BG Hue (Pink → Purple)").pack(anchor="w")
        ttk.Scale(
            left,
            from_=0,
            to=100,
            orient="horizontal",
            variable=self.bg_hue_pct,
            command=lambda _v: self._changed(save=False)
        ).pack(fill="x", pady=(2, 10))

        # Magenta injection slider
        ttk.Label(left, text="Hot Magenta Inject (Feral)").pack(anchor="w")
        ttk.Scale(
            left,
            from_=0,
            to=100,
            orient="horizontal",
            variable=self.magenta_pct,
            command=lambda _v: self._changed(save=False)
        ).pack(fill="x", pady=(2, 10))

        # BG intensity slider
        ttk.Label(left, text="BG Intensity (tint alpha only)").pack(anchor="w")
        ttk.Scale(
            left,
            from_=0,
            to=BG_INTENSITY_MAX,
            orient="horizontal",
            variable=self.bg_intensity_pct,
            command=lambda _v: self._changed(save=False)
        ).pack(fill="x", pady=(2, 10))

        # Element intensity slider
        ttk.Label(left, text="Element Intensity (stroke/sparkle/border + star count)").pack(anchor="w")
        ttk.Scale(
            left,
            from_=0,
            to=ELEM_INTENSITY_MAX,
            orient="horizontal",
            variable=self.elem_intensity_pct,
            command=lambda _v: self._changed(save=False)
        ).pack(fill="x", pady=(2, 12))

        btn_row = ttk.Frame(left)
        btn_row.pack(fill="x", pady=(6, 6))
        ttk.Button(btn_row, text="Reset to Defaults", command=lambda: self._reset_defaults(save=True)).pack(side="left")
        ttk.Button(btn_row, text="Save Now", command=self._save_config).pack(side="left", padx=8)

        ttk.Label(left, textvariable=self.readout, foreground="#444", wraplength=260).pack(anchor="w", pady=(10, 0))

    def _changed(self, save: bool = True):
        if save:
            self._save_config()
        self.redraw()

    # --------------------------
    # Derived profile (sliders)
    # --------------------------
    def _effective_alphas(self):
        preset = self.preset_var.get()
        bg_mult = float(self.bg_intensity_pct.get()) / 100.0
        el_mult = float(self.elem_intensity_pct.get()) / 100.0

        base = PRESETS_BW[preset] if self.printer_bw.get() else PRESET_BASES[preset]

        # BG slider affects only tint alpha
        tint = clamp(base["tint_alpha"] * bg_mult, *LIMITS["tint_alpha"])

        # Element slider affects everything decorative
        stroke = clamp(base["stroke_alpha"] * el_mult, *LIMITS["stroke_alpha"])
        sparkle = clamp(base["sparkle_alpha"] * el_mult, *LIMITS["sparkle_alpha"])
        border = clamp(base["border_alpha"] * el_mult, *LIMITS["border_alpha"])

        return tint, stroke, sparkle, border

    def _effective_colors(self):
        hue_t = float(self.bg_hue_pct.get()) / 100.0  # 0 pink, 1 purple
        mag_t = float(self.magenta_pct.get()) / 100.0

        if self.printer_bw.get():
            # BW stays neutral
            return "#f3f3f6", "#777777"

        tint_hex = mix(PINK_TINT_HEX, PURPLE_TINT_HEX, hue_t)
        stroke_hex = mix(PINK_STROKE_HEX, PURPLE_STROKE_HEX, hue_t)

        # Inject feral magenta
        tint_hex = inject_magenta(tint_hex, MAGENTA_INJECT_HEX, mag_t * 0.85)
        stroke_hex = inject_magenta(stroke_hex, MAGENTA_INJECT_HEX, mag_t * 1.00)

        return tint_hex, stroke_hex

    # --------------------------
    # Starfield count + stable RNG
    # --------------------------
    def _star_count(self):
        """
        Linear growth with element intensity.
        At 0% => ~0
        At 100% => base moderate
        At 600% => feral
        """
        el = int(self.elem_intensity_pct.get())
        # linear scale; tweak the slope if you want more/less
        # 100% => 18 stars, 600% => 108 stars (plus a few depending on canvas size)
        return int(round(0.18 * el))

    def _rng_seed(self, w: int, h: int) -> int:
        """
        Deterministic seed so random placement is stable for a given state.
        Include canvas size so it redistributes when resized (but still stable at that size).
        """
        return hash((
            int(w), int(h),
            bool(self.printer_bw.get()),
            str(self.preset_var.get()),
            int(self.bg_hue_pct.get()),
            int(self.magenta_pct.get()),
            int(self.bg_intensity_pct.get()),
            int(self.elem_intensity_pct.get()),
        )) & 0xFFFFFFFF

    # --------------------------
    # Drawing
    # --------------------------
    def redraw(self):
        c = self.canvas
        c.delete("all")
        w = max(10, int(c.winfo_width()))
        h = max(10, int(c.winfo_height()))

        tint_hex, stroke_hex = self._effective_colors()
        tint_a, stroke_a, sparkle_a, border_a = self._effective_alphas()

        # Background wash
        bg = blend_over_white(tint_hex, tint_a)
        c.create_rectangle(0, 0, w, h, fill=bg, outline="")

        # Simulate alpha by mixing toward background.
        # Higher boost => lines pop harder (Tk has no true alpha strokes).
        boost = 18.0
        line_col = mix(bg, stroke_hex, clamp(stroke_a * boost, 0.0, 1.0))
        sparkle_col = mix(bg, stroke_hex, clamp(sparkle_a * boost, 0.0, 1.0))
        border_col = mix(bg, stroke_hex, clamp(border_a * boost, 0.0, 1.0))

        # Border
        margin = 22
        top_pad = 22
        c.create_rectangle(margin, margin + top_pad, w - margin, h - margin, outline=border_col, width=3)

        # Daisy watermark
        cx, cy = int(w * 0.65), int(h * 0.52)
        petals = 12
        petal_r = int(min(w, h) * 0.07)
        petal_dist = int(min(w, h) * 0.12)

        for i in range(petals):
            angle = (i / petals) * (2 * math.pi)
            px = cx + int(petal_dist * math.cos(angle))
            py = cy + int(petal_dist * math.sin(angle))
            c.create_oval(px - petal_r, py - petal_r, px + petal_r, py + petal_r, outline=line_col, width=3)

        center_r = int(petal_r * 1.15)
        c.create_oval(cx - center_r, cy - center_r, cx + center_r, cy + center_r, outline=line_col, width=3)

        # Star drawing helper
        def star(x, y, r, col, width=2):
            c.create_line(x - r, y, x + r, y, fill=col, width=width)
            c.create_line(x, y - r, x, y + r, fill=col, width=width)
            c.create_line(x - int(r * 0.7), y - int(r * 0.7),
                          x + int(r * 0.7), y + int(r * 0.7), fill=col, width=width)
            c.create_line(x - int(r * 0.7), y + int(r * 0.7),
                          x + int(r * 0.7), y - int(r * 0.7), fill=col, width=width)

        # Deterministic “random” starfield that scales with intensity
        n_stars = self._star_count()
        rng = random.Random(self._rng_seed(w, h))

        # Keep stars within content area (respect margin)
        x0 = margin + 6
        x1 = w - margin - 6
        y0 = margin + top_pad + 6
        y1 = h - margin - 6

        # Bigger stars at higher intensity
        el_norm = clamp(float(self.elem_intensity_pct.get()) / 100.0, 0.0, 8.0)  # up to 800%
        r_min = 4
        r_max = int(8 + 3 * el_norm)  # grows with intensity

        # Scatter
        for _ in range(n_stars):
            x = rng.randint(x0, x1) if x1 > x0 else x0
            y = rng.randint(y0, y1) if y1 > y0 else y0
            r = rng.randint(r_min, max(r_min, r_max))
            wline = 2 if r < 10 else 3
            star(x, y, r, sparkle_col, width=wline)

        # A few anchor sparkles so it never looks empty at low intensity
        if n_stars < 8:
            star(int(w * 0.16), int(h * 0.18), 10, sparkle_col, width=2)
            star(int(w * 0.84), int(h * 0.82), 10, sparkle_col, width=2)

        # Readout
        mode = "B/W" if self.printer_bw.get() else "Color"
        self.readout.set(
            f"Mode: {mode}\n"
            f"Preset: {self.preset_var.get()}\n"
            f"BG Hue: {int(self.bg_hue_pct.get())}% Purple\n"
            f"Magenta Inject: {int(self.magenta_pct.get())}%\n"
            f"BG Intensity: {int(self.bg_intensity_pct.get())}% (max {BG_INTENSITY_MAX})\n"
            f"Elements: {int(self.elem_intensity_pct.get())}% (max {ELEM_INTENSITY_MAX})\n"
            f"Stars: {n_stars}\n"
            f"Effective α (tint/stroke/sparkle/border): "
            f"{tint_a:.3f}/{stroke_a:.3f}/{sparkle_a:.3f}/{border_a:.3f}"
        )

    def _on_close(self):
        self._save_config()
        self.root.destroy()


def main():
    root = tk.Tk()
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names() and os.name == "nt":
            style.theme_use("vista")
    except Exception:
        pass
    _app = KawaiiPreviewApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
