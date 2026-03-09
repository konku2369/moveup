# kawaii_settings.py
# Shared Kawaii PDF background settings (single source of truth)
# - Presets editable at top
# - Persists to kawaii_pdf_settings.json
# - Provides effective profile for pdf_export.py (tint + stroke + border + sparkles)

from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass
from typing import Dict, Tuple


CONFIG_FILENAME = "kawaii_pdf_settings.json"

# -----------------------------------------
# Presets (EDIT THESE ONLY, if you want)
# These are BASE values before sliders multiply them.
# -----------------------------------------
PRESETS_COLOR: Dict[str, Dict[str, float]] = {
    "Minimal": {"tint_alpha": 0.020, "stroke_alpha": 0.050, "sparkle_alpha": 0.045, "border_alpha": 0.060},
    "Cute":    {"tint_alpha": 0.055, "stroke_alpha": 0.080, "sparkle_alpha": 0.070, "border_alpha": 0.100},
    "Extra":   {"tint_alpha": 0.110, "stroke_alpha": 0.140, "sparkle_alpha": 0.120, "border_alpha": 0.180},
}

PRESETS_BW: Dict[str, Dict[str, float]] = {
    "Minimal": {"tint_alpha": 0.004, "stroke_alpha": 0.030, "sparkle_alpha": 0.028, "border_alpha": 0.040},
    "Cute":    {"tint_alpha": 0.007, "stroke_alpha": 0.045, "sparkle_alpha": 0.040, "border_alpha": 0.055},
    "Extra":   {"tint_alpha": 0.012, "stroke_alpha": 0.065, "sparkle_alpha": 0.055, "border_alpha": 0.080},
}

# -----------------------------------------
# Limits (how intense it’s allowed to get)
# You asked for more feral: here you go.
# -----------------------------------------
LIMITS = {
    "tint_alpha":    (0.0, 0.30),  # capped at 30% — above that washes out table text
    "stroke_alpha":  (0.0, 0.55),
    "sparkle_alpha": (0.0, 0.55),
    "border_alpha":  (0.0, 0.55),
}

# Hue endpoints
PINK_TINT_RGB = (1.00, 0.86, 0.92)     # Color(1, 0.86, 0.92)
PURPLE_TINT_RGB = (0.90, 0.85, 1.00)   # soft lavender tint

PINK_STROKE_RGB = (0.55, 0.40, 0.50)   # your original
PURPLE_STROKE_RGB = (0.44, 0.36, 0.62) # purple-ish outline


@dataclass
class KawaiiSettings:
    preset: str = "Cute"
    printer_bw: bool = False

    # sliders in percent
    bg_hue_pct: int = 100          # 0 = pink, 100 = purple
    bg_intensity_pct: int = 100    # multiplies tint_alpha only
    elem_intensity_pct: int = 100  # multiplies stroke/sparkle/border

    # stars intensity: number of stars scales with element intensity
    stars_base: int = 18
    stars_max_extra: int = 140     # added at 200% intensity

    def clamp_self(self) -> None:
        if self.preset not in PRESETS_COLOR:
            self.preset = "Cute"
        self.bg_hue_pct = max(0, min(100, int(self.bg_hue_pct)))
        self.bg_intensity_pct = max(0, min(300, int(self.bg_intensity_pct)))
        self.elem_intensity_pct = max(0, min(300, int(self.elem_intensity_pct)))
        self.stars_base = max(0, int(self.stars_base))
        self.stars_max_extra = max(0, int(self.stars_max_extra))


def _app_dir() -> str:
    # When running as a PyInstaller exe, __file__ points to the temp _MEIPASS
    # folder which is cleaned up on exit. Use sys.executable instead so the
    # settings file lives next to the .exe (same as config_manager.py does).
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def config_path() -> str:
    return os.path.join(_app_dir(), CONFIG_FILENAME)


def load_settings() -> KawaiiSettings:
    path = config_path()
    s = KawaiiSettings()
    if not os.path.exists(path):
        return s
    try:
        with open(path, "r", encoding="utf-8") as f:
            cfg = json.load(f) or {}
        s.preset = str(cfg.get("preset", s.preset))
        s.printer_bw = bool(cfg.get("printer_bw", s.printer_bw))
        s.bg_hue_pct = int(cfg.get("bg_hue_pct", s.bg_hue_pct))
        s.bg_intensity_pct = int(cfg.get("bg_intensity_pct", s.bg_intensity_pct))
        s.elem_intensity_pct = int(cfg.get("elem_intensity_pct", s.elem_intensity_pct))
        s.stars_base = int(cfg.get("stars_base", s.stars_base))
        s.stars_max_extra = int(cfg.get("stars_max_extra", s.stars_max_extra))
        s.clamp_self()
        return s
    except Exception as e:
        print(f"[moveup] Warning: could not load kawaii settings ({path}): {e}")
        return KawaiiSettings()


def save_settings(s: KawaiiSettings) -> None:
    s.clamp_self()
    cfg = {
        "preset": s.preset,
        "printer_bw": bool(s.printer_bw),
        "bg_hue_pct": int(s.bg_hue_pct),
        "bg_intensity_pct": int(s.bg_intensity_pct),
        "elem_intensity_pct": int(s.elem_intensity_pct),
        "stars_base": int(s.stars_base),
        "stars_max_extra": int(s.stars_max_extra),
    }
    path = config_path()
    tmp = path + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
        os.replace(tmp, path)
    except Exception as e:
        print(f"[moveup] Warning: could not save kawaii settings ({path}): {e}")


def reset_defaults() -> KawaiiSettings:
    return KawaiiSettings()


def _clamp(v: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, v))


def _mix_rgb(a: Tuple[float, float, float], b: Tuple[float, float, float], t: float) -> Tuple[float, float, float]:
    t = _clamp(t, 0.0, 1.0)
    return (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t, a[2] + (b[2] - a[2]) * t)


def compute_effective_profile(s: KawaiiSettings) -> Dict[str, object]:
    """
    Returns a dict consumable by pdf_export.py:
      - tint_alpha, stroke_alpha, sparkle_alpha, border_alpha
      - tint_rgb (0..1), stroke_rgb (0..1)
      - stars_count
    """
    s.clamp_self()

    base = PRESETS_BW[s.preset] if s.printer_bw else PRESETS_COLOR[s.preset]

    bg_mult = float(s.bg_intensity_pct) / 100.0
    el_mult = float(s.elem_intensity_pct) / 100.0

    tint_alpha = _clamp(base["tint_alpha"] * bg_mult, *LIMITS["tint_alpha"])
    stroke_alpha = _clamp(base["stroke_alpha"] * el_mult, *LIMITS["stroke_alpha"])
    sparkle_alpha = _clamp(base["sparkle_alpha"] * el_mult, *LIMITS["sparkle_alpha"])
    border_alpha = _clamp(base["border_alpha"] * el_mult, *LIMITS["border_alpha"])

    # hue: 0..1
    t = float(s.bg_hue_pct) / 100.0

    if s.printer_bw:
        tint_rgb = (0.95, 0.95, 0.97)
        stroke_rgb = (0.45, 0.45, 0.48)
    else:
        tint_rgb = _mix_rgb(PINK_TINT_RGB, PURPLE_TINT_RGB, t)
        stroke_rgb = _mix_rgb(PINK_STROKE_RGB, PURPLE_STROKE_RGB, t)

    # All element counts scale linearly with elem_intensity (0–200%).
    # At 100%: stars=base+half_extra, daisies=9, paws=6.
    el_pct = _clamp(float(s.elem_intensity_pct), 0.0, 200.0)
    t_el = el_pct / 200.0

    extra = int(round(t_el * float(s.stars_max_extra)))
    stars_count = int(max(0, s.stars_base + extra))

    # daisies: 3 at 0% → 9 at 100% → 15 at 200%
    daisy_count = int(max(0, round(3 + t_el * 12)))
    # paws:    2 at 0% → 6 at 100% → 10 at 200%
    paw_count = int(max(0, round(2 + t_el * 8)))
    # cats: same scale as stars — a fuck-ton at high intensity
    cat_count = int(max(0, s.stars_base + int(round(t_el * float(s.stars_max_extra)))))

    return {
        "tint_alpha": float(tint_alpha),
        "stroke_alpha": float(stroke_alpha),
        "sparkle_alpha": float(sparkle_alpha),
        "border_alpha": float(border_alpha),
        "tint_rgb": tint_rgb,
        "stroke_rgb": stroke_rgb,
        "stars_count": stars_count,
        "daisy_count": daisy_count,
        "paw_count": paw_count,
        "cat_count": cat_count,
        "printer_bw": bool(s.printer_bw),
        "preset": s.preset,
        "bg_hue_pct": int(s.bg_hue_pct),
        "bg_intensity_pct": int(s.bg_intensity_pct),
        "elem_intensity_pct": int(s.elem_intensity_pct),
    }
