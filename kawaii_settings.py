"""
Kawaii PDF decoration settings.

Manages KawaiiSettings dataclass, preset system, color math (pink-to-purple
hue interpolation), and persistence to kawaii_pdf_settings.json.
compute_effective_profile() converts user-facing slider values into the
exact alpha/count/color values consumed by pdf_export.py decorations.

HOW THE KAWAII SYSTEM WORKS:
============================
1. USER SETTINGS: KawaiiSettings dataclass holds slider values:
   - preset: "Minimal", "Cute", or "Extra" (base intensity level)
   - bg_hue_pct: 0-100 slider (0=pink, 100=lavender/purple)
   - elem_intensity: 0-100 slider (controls how many daisies/paws/cats)
   - printer_bw: True for B/W printers (disables color tint)

2. PRESETS: Define base alpha values for tint, stroke, sparkle, border.
   "Minimal" = barely visible decorations, "Extra" = maximum kawaii.

3. compute_effective_profile(): Takes settings → produces a profile dict
   with exact RGB colors, alpha values, element counts, and jitter seed.
   This dict is consumed directly by pdf_export.py's drawing functions.

4. PERSISTENCE: Saved to kawaii_pdf_settings.json. Loaded on startup.
   Preview dialog (kawaii_preview.py) lets users see changes live.
"""

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
    """
    User-facing kawaii PDF decoration settings.

    All sliders are stored as integer percentages so they can be serialised
    cleanly to JSON and bound directly to Tk ``IntVar`` widgets in the preview
    dialog.  ``compute_effective_profile()`` converts these percent values into
    the exact float alpha/RGB/count values consumed by ``pdf_export.py``.

    Attributes
    ----------
    preset : str
        Decoration intensity tier: ``"Minimal"``, ``"Cute"``, or ``"Extra"``.
        Selects the base alpha values from ``PRESETS_COLOR`` or ``PRESETS_BW``
        before the intensity sliders multiply them.
    printer_bw : bool
        When ``True``, switch to greyscale-safe ``PRESETS_BW`` alphas and use
        neutral grey RGB values instead of pink/purple tints.
    bg_hue_pct : int
        Background tint hue slider, 0–100.  0 = warm pink
        (``PINK_TINT_RGB``), 100 = soft lavender purple (``PURPLE_TINT_RGB``).
        Linearly interpolated in ``compute_effective_profile()``.
    bg_intensity_pct : int
        Background tint opacity slider, 0–300.  Multiplied against the preset's
        ``tint_alpha`` base value.  Only affects background fill color, not
        element strokes.
    elem_intensity_pct : int
        Element intensity slider, 0–300.  Multiplied against ``stroke_alpha``,
        ``sparkle_alpha``, and ``border_alpha`` from the preset; also controls
        star/daisy/paw/cat counts.
    stars_base : int
        Minimum star count when ``elem_intensity_pct`` is 0.
    stars_max_extra : int
        Additional stars added at 200% intensity.  At 100% half this value is
        added; the final count is ``stars_base + round(t_el * stars_max_extra)``
        where ``t_el = elem_intensity_pct / 200`` (clamped 0–1).
    """

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
        """
        Coerce all fields to their valid ranges in place.

        Called before save and before ``compute_effective_profile()`` so that
        hand-edited JSON or out-of-range slider values are silently corrected
        rather than causing downstream exceptions:

        - ``preset`` → reset to ``"Cute"`` if not in ``PRESETS_COLOR``.
        - ``bg_hue_pct`` → clamped to [0, 100].
        - ``bg_intensity_pct`` / ``elem_intensity_pct`` → clamped to [0, 300].
        - ``stars_base`` / ``stars_max_extra`` → clamped to ≥ 0.
        """
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
    """
    Load kawaii PDF settings from ``kawaii_pdf_settings.json``.

    If the file does not exist, a default ``KawaiiSettings()`` instance is
    returned silently.  Any ``Exception`` during read or parse is caught and
    printed; defaults are returned rather than propagating the error.

    All loaded values are coerced to their correct types (``str``, ``bool``,
    ``int``) and ``clamp_self()`` is called before returning so that stale or
    hand-edited files cannot produce out-of-range settings.

    Returns
    -------
    KawaiiSettings
        Populated from disk, or default values if the file is absent or invalid.
    """
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
    """
    Persist kawaii PDF settings to ``kawaii_pdf_settings.json``.

    Calls ``s.clamp_self()`` before writing so out-of-range in-memory values
    are corrected prior to serialisation.  Uses the atomic write pattern
    (``.tmp`` → ``os.replace()``) to prevent file corruption on crash.

    Any ``Exception`` during write is caught and printed; the in-memory
    settings object is not modified (aside from the clamp call).

    Parameters
    ----------
    s : KawaiiSettings
        Settings to persist.  Fields are coerced to their expected Python
        types (``int``, ``bool``, ``str``) before being written to JSON.
    """
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
    """Return a new ``KawaiiSettings`` instance with all default values."""
    return KawaiiSettings()


def _clamp(v: float, lo: float, hi: float) -> float:
    """Clamp *v* to the range [*lo*, *hi*]."""
    return max(lo, min(hi, v))


def _mix_rgb(
    a: Tuple[float, float, float],
    b: Tuple[float, float, float],
    t: float,
) -> Tuple[float, float, float]:
    """
    Linear interpolation between two RGB tuples.

    Parameters
    ----------
    a, b : tuple[float, float, float]
        Source and destination RGB colors with components in [0, 1].
    t : float
        Blend factor: 0.0 returns *a*, 1.0 returns *b*, clamped to [0, 1].

    Returns
    -------
    tuple[float, float, float]
        Interpolated RGB color.
    """
    t = _clamp(t, 0.0, 1.0)
    return (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t, a[2] + (b[2] - a[2]) * t)


def compute_effective_profile(s: KawaiiSettings) -> Dict[str, object]:
    """
    Convert a ``KawaiiSettings`` instance into a concrete decoration profile dict.

    This is the central computation that bridges user-facing sliders and the
    per-element drawing parameters in ``pdf_export.py``.  The output dict is
    passed directly to decoration functions like ``_draw_kawaii_background()``.

    Computation steps:
    1. **Clamp** all settings via ``s.clamp_self()``.
    2. **Pick base alphas** from ``PRESETS_BW`` (if ``printer_bw``) or
       ``PRESETS_COLOR``, keyed by ``s.preset``.
    3. **Apply sliders**: multiply ``tint_alpha`` by ``bg_intensity_pct / 100``;
       multiply ``stroke_alpha``, ``sparkle_alpha``, ``border_alpha`` by
       ``elem_intensity_pct / 100``.  Each result is clamped to ``LIMITS``.
    4. **Interpolate hue**: blend ``PINK_TINT_RGB`` → ``PURPLE_TINT_RGB`` and
       ``PINK_STROKE_RGB`` → ``PURPLE_STROKE_RGB`` using ``t = bg_hue_pct / 100``.
       B/W mode uses neutral grey values instead.
    5. **Compute element counts**:
       - ``t_el = clamp(elem_intensity_pct, 0, 200) / 200``  (normalised 0–1)
       - stars:  ``stars_base + round(t_el × stars_max_extra)``
       - daisies: ``round(3 + t_el × 12)``  → 3 at 0%, 15 at 200%
       - paws:    ``round(2 + t_el × 8)``   → 2 at 0%, 10 at 200%
       - cats:    same formula as stars

    Parameters
    ----------
    s : KawaiiSettings
        User settings.  Modified in place by ``clamp_self()`` before use.

    Returns
    -------
    dict
        Keys: ``tint_alpha``, ``stroke_alpha``, ``sparkle_alpha``,
        ``border_alpha`` (floats), ``tint_rgb``, ``stroke_rgb``
        (3-tuples of floats in [0, 1]), ``stars_count``, ``daisy_count``,
        ``paw_count``, ``cat_count`` (ints), plus echo-back of the input
        slider values (``preset``, ``bg_hue_pct``, ``bg_intensity_pct``,
        ``elem_intensity_pct``, ``printer_bw``).
    """
    s.clamp_self()

    base = PRESETS_BW[s.preset] if s.printer_bw else PRESETS_COLOR[s.preset]

    bg_mult = float(s.bg_intensity_pct) / 100.0
    el_mult = float(s.elem_intensity_pct) / 100.0

    tint_alpha = _clamp(base["tint_alpha"] * bg_mult, *LIMITS["tint_alpha"])
    stroke_alpha = _clamp(base["stroke_alpha"] * el_mult, *LIMITS["stroke_alpha"])
    sparkle_alpha = _clamp(base["sparkle_alpha"] * el_mult, *LIMITS["sparkle_alpha"])
    border_alpha = _clamp(base["border_alpha"] * el_mult, *LIMITS["border_alpha"])

    # Hue interpolation: t=0 → pink, t=1 → lavender purple.
    # bg_hue_pct slider (0-100) controls the tint color of the PDF background.
    t = float(s.bg_hue_pct) / 100.0

    if s.printer_bw:
        # B/W mode uses neutral greys instead of pink/purple
        tint_rgb = (0.95, 0.95, 0.97)
        stroke_rgb = (0.45, 0.45, 0.48)
    else:
        # Linearly interpolate between hand-tuned pink and purple RGB values
        tint_rgb = _mix_rgb(PINK_TINT_RGB, PURPLE_TINT_RGB, t)
        stroke_rgb = _mix_rgb(PINK_STROKE_RGB, PURPLE_STROKE_RGB, t)

    # Element counts (stars, daisies, paws, cats) scale linearly with intensity.
    # 0% → minimal decorations, 100% → normal cute, 200% → maximum feral.
    el_pct = _clamp(float(s.elem_intensity_pct), 0.0, 200.0)
    t_el = el_pct / 200.0  # normalize to 0..1 range

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
