"""Shared visual theme for the generated deck (colors, fonts, spacing).

Derived from the neutral brand palette of the source template: a violet
primary accent, a dark ink color for headings/backgrounds, and a status
traffic-light scheme for on-track / cautious / at-risk.
"""
from pptx.util import Emu

# --- Palette (hex, no '#') ---
PRIMARY = "6B47F5"       # violet accent
INK = "302E52"           # dark heading / footer text
INK_SOFT = "5B5B6B"
MUTED = "8A889E"         # secondary / caption text
SURFACE = "F7F7FB"       # page background tint
SURFACE_ALT = "ECEBF5"   # card / tile background
LINE = "D6D1ED"          # hairlines, table borders
WHITE = "FFFFFF"

STATUS = {
    "on_track": "2E9E7B",   # green
    "cautious": "E0A52E",   # amber
    "at_risk": "C0392B",    # red
}
STATUS_LABEL = {
    "on_track": "ON TRACK",
    "cautious": "WATCH",
    "at_risk": "AT RISK",
}

FONT = "Calibri"

# --- Slide geometry (16:9, matches source deck) ---
SLIDE_W = Emu(12192000)
SLIDE_H = Emu(6858000)

MARGIN = Emu(457200)  # 0.5in


def status_color(key: str) -> str:
    return STATUS.get(key, MUTED)


def status_label(key: str) -> str:
    return STATUS_LABEL.get(key, key.upper() if key else "")
