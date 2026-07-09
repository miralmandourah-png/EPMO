"""In-place branded PPTX export.

Opens the user's own Tawuniya-branded template and substitutes ONLY text-run
values, leaving every other piece of slide XML (backgrounds, theme colors,
fonts, logo, wave shapes, layout) untouched — so brand fidelity is exact.

Anchoring never uses the template's confidential figures. Boxes are located
by (a) their position on the slide (a coordinate map calibrated to this
36-slide deck, per the user's choice) and (b) non-confidential structural
labels (sector names, section numbers). The duplicate CEO/appendix slides are
driven from one model copy via the *_SLIDES lists below.

Coverage note: this module currently maps the cover, section dividers,
closing cards, and the IPI-by-sector values (the clean, high-confidence
anchors). The dense per-initiative grid tables are the next calibration step;
see CHANGES.md. Unmapped slides are left exactly as they are in the template.
"""
from __future__ import annotations

import io
import re
from typing import Any, Dict, List, Optional

from pptx import Presentation
from pptx.util import Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE

IN = 914400  # EMU per inch


# --------------------------------------------------------------------------
# Low-level text mutation (format-preserving)
# --------------------------------------------------------------------------

def _set_paragraph_text(paragraph, text: str) -> None:
    """Replace a paragraph's text, keeping the first run's formatting."""
    runs = paragraph.runs
    if not runs:
        run = paragraph.add_run()
        run.text = text
        return
    runs[0].text = text
    for extra in runs[1:]:
        extra._r.getparent().remove(extra._r)


def _set_box_text(shape, text: str) -> None:
    if not shape.has_text_frame:
        return
    _set_paragraph_text(shape.text_frame.paragraphs[0], str(text))


def _set_run_text(shape, run_index: int, text: str, para_index: int = 0) -> None:
    """Replace exactly one run's text, leaving every other run/format untouched.

    Used for boxes where a value is one isolated run inside a longer label
    (e.g. "95  Strategic projects   773  milestones") so we never disturb
    surrounding tabs, line breaks, or label formatting.
    """
    if not shape.has_text_frame:
        return
    paras = shape.text_frame.paragraphs
    if para_index >= len(paras):
        return
    runs = paras[para_index].runs
    if 0 <= run_index < len(runs):
        runs[run_index].text = str(text)


def _left_top_in(shape):
    try:
        return shape.left / IN, shape.top / IN
    except (TypeError, ZeroDivisionError):
        return None, None


def _text_boxes(slide):
    return [sh for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text.strip()]


def _find_near(slide, left_in: float, top_in: float, tol: float = 0.3):
    """Text box whose position is closest to (left_in, top_in) within tol."""
    best, best_d = None, tol
    for sh in _text_boxes(slide):
        L, T = _left_top_in(sh)
        if L is None:
            continue
        d = abs(L - left_in) + abs(T - top_in)
        if d <= best_d:
            best, best_d = sh, d
    return best


def _find_exact(slide, left_in: float, top_in: float, tol: float = 0.05):
    """Tight-tolerance lookup for dense slides (e.g. Executive Summary) where
    _find_near's default tolerance would risk matching a neighboring box."""
    return _find_near(slide, left_in, top_in, tol=tol)


def _find_by_label(slide, label: str):
    """First text box whose text starts with a structural label (case-insensitive)."""
    lab = label.strip().lower()
    for sh in _text_boxes(slide):
        if sh.text_frame.text.strip().lower().startswith(lab):
            return sh
    return None


def _value_right_of(slide, anchor, min_dx: float = 0.5, tol_top: float = 0.15):
    """Nearest text box to the right of `anchor` on the same visual row."""
    aL, aT = _left_top_in(anchor)
    if aL is None:
        return None
    best, best_dx = None, 99.0
    for sh in _text_boxes(slide):
        if sh._element is anchor._element:
            continue
        L, T = _left_top_in(sh)
        if L is None:
            continue
        if abs(T - aT) <= tol_top and L - aL >= min_dx and (L - aL) < best_dx:
            best, best_dx = sh, L - aL
    return best


# --------------------------------------------------------------------------
# Dynamic status coloring
#
# The threshold this deck states in its own text is "on-track ≥ 3.0" (slides
# 4, 6, 12, 17). The amber/red split below that line (>=2.6 cautious, <2.6
# at-risk) is the rule from the dashboard's own IPI bars, applied here too so
# the exported deck and the live dashboard never disagree about what a given
# number means. Colors match the template's own status palette (sampled from
# its existing status bubbles), not an invented one.
# --------------------------------------------------------------------------

STATUS_COLORS = {"on_track": "2E9E7B", "cautious": "E0A52E", "at_risk": "C0392B"}


def _ipi_status(value: float) -> str:
    if value >= 3.0:
        return "on_track"
    if value >= 2.6:
        return "cautious"
    return "at_risk"


def _set_shape_fill(shape, hex_color: str) -> None:
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(hex_color)


def _set_run_font_color(shape, hex_color: str, run_index: int = 0, para_index: int = 0) -> None:
    if not shape.has_text_frame:
        return
    paras = shape.text_frame.paragraphs
    if para_index >= len(paras):
        return
    runs = paras[para_index].runs
    if 0 <= run_index < len(runs):
        runs[run_index].font.color.rgb = RGBColor.from_string(hex_color)


def _find_status_bubble(slide, left_in: float, top_in: float, tol: float = 0.15):
    """An AUTO_SHAPE near (left_in, top_in) whose current fill is one of the
    deck's own status colors — i.e. the colored dot/bubble itself, not its
    lighter background track shape."""
    known = {c.upper() for c in STATUS_COLORS.values()}
    best, best_d = None, tol
    for sh in slide.shapes:
        if sh.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
            continue
        try:
            if sh.fill.type != 1 or str(sh.fill.fore_color.rgb).upper() not in known:
                continue
        except Exception:
            continue
        L, T = _left_top_in(sh)
        if L is None:
            continue
        d = abs(L - left_in) + abs(T - top_in)
        if d <= best_d:
            best, best_d = sh, d
    return best


# --------------------------------------------------------------------------
# Coordinate map (calibrated to the 36-slide Strategic Health Check deck)
# --------------------------------------------------------------------------

COVER_SLIDES = [1]
CLOSING_SLIDES = [13, 36]
# slide index -> section-number suffix used in div.sec.<NN>
DIVIDER_SLIDES = {3: "01", 14: "02", 16: "03", 18: "04", 32: "05"}
IPI_SECTOR_SLIDES = [6, 17]

# cover box positions (inches) on this deck
COVER_POS = {
    "eyebrow": (0.8, 1.85),
    "title": (0.8, 2.35),   # paragraph 0 = title, paragraph 1 = date
    "footer": (0.8, 5.2),
}
CLOSING_POS = {"message": (0.8, 2.9), "footer": (0.8, 5.2)}
DIVIDER_TITLE_POS = (1.0, 4.4)


def _get(fields: Dict[str, Any], key: str) -> Optional[str]:
    v = fields.get(key)
    if v is None or v == "":
        return None
    return str(v)


def _sector_rows(fields: Dict[str, Any], rows: Dict[str, int]) -> List[Dict[str, str]]:
    n = int(rows.get("ipi.sector", 0) or 0)
    out = []
    for i in range(n):
        out.append({
            "sector": _get(fields, f"ipi.sector.{i}.sector") or "",
            "ipi": _get(fields, f"ipi.sector.{i}.ipi"),
        })
    return out


def _apply_cover(slide, fields):
    box = _find_near(slide, *COVER_POS["eyebrow"])
    v = _get(fields, "div.cover.eyebrow")
    if box and v is not None:
        _set_box_text(box, v)
    title_box = _find_near(slide, *COVER_POS["title"])
    if title_box:
        paras = title_box.text_frame.paragraphs
        tv = _get(fields, "div.cover.title")
        if tv is not None and paras:
            _set_paragraph_text(paras[0], tv)
        dv = _get(fields, "div.cover.date")
        if dv is not None and len(paras) > 1:
            _set_paragraph_text(paras[1], dv)
    fbox = _find_near(slide, *COVER_POS["footer"])
    fv = _get(fields, "div.cover.footer")
    if fbox and fv is not None:
        _set_box_text(fbox, fv)


def _apply_divider(slide, fields, suffix):
    box = _find_near(slide, *DIVIDER_TITLE_POS, tol=0.6)
    v = _get(fields, f"div.sec.{suffix}")
    if box and v is not None:
        _set_box_text(box, v)


def _apply_closing(slide, fields):
    box = _find_near(slide, *CLOSING_POS["message"], tol=0.5)
    v = _get(fields, "div.closing.message")
    if box and v is not None:
        _set_box_text(box, v)
    fbox = _find_near(slide, *CLOSING_POS["footer"])
    fv = _get(fields, "div.closing.footer")
    if fbox and fv is not None:
        _set_box_text(fbox, fv)


def _apply_ipi_sector(slide, sector_rows):
    for sh in _text_boxes(slide):
        label = sh.text_frame.text.strip()
        L, T = _left_top_in(sh)
        if L is None or L > 2.0:  # sector labels sit in the left column
            continue
        match = None
        low = label.lower()
        for row in sector_rows:
            s = row["sector"].lower()
            if not s:
                continue
            # tolerant match: "Mobility" vs "Mobility/Motor"
            token = low.split("/")[0].strip()
            stoken = s.split("/")[0].strip()
            if token and (token == stoken or token in s or stoken in low):
                match = row
                break
        if match and match["ipi"] is not None:
            val = _value_right_of(slide, sh, min_dx=1.0)
            if val is not None:
                _set_box_text(val, match["ipi"])
            # dynamic status bubble: on-track >=3.0 (this deck's own stated
            # rule), cautious >=2.6, at-risk below -- recomputed from the
            # live number rather than left at whatever the template shipped.
            try:
                ipi_val = float(match["ipi"])
            except (TypeError, ValueError):
                ipi_val = None
            if ipi_val is not None:
                bubble = _find_status_bubble(slide, 2.10, T, tol=0.15)
                if bubble is not None:
                    _set_shape_fill(bubble, STATUS_COLORS[_ipi_status(ipi_val)])


# Executive Summary slide (not duplicated elsewhere in the deck).
EXEC_SUMMARY_SLIDES = [4]

# (left_in, top_in) of each target box on slide 4, calibrated to this deck.
EXEC_POS = {
    "committed_bri": (9.182, 1.820),          # single run, e.g. "SAR 2,961m"
    "enterprise_ipi": (4.703, 2.369),         # single run, e.g. "2.79"
    "projects_milestones_line": (4.703, 3.186),  # run0=strategic projects, run6=ms total
    "ms_total_line": (4.703, 4.013),          # last run ends "<N> Total"
    "ms_complete": (4.655, 4.265),            # run0=count, run2="(NN%)"
    "ms_not_yet_due": (6.708, 4.265),         # single run
    "ms_delayed": (8.253, 4.295),             # single run
}


def _fmt_int(v: Any) -> Optional[str]:
    try:
        return f"{int(round(float(v))):,}"
    except (TypeError, ValueError):
        return None


def _fmt_sar_m(v: Any) -> Optional[str]:
    try:
        return f"SAR {int(round(float(v))):,}m"
    except (TypeError, ValueError):
        return None


def _apply_exec_summary(slide, fields: Dict[str, Any]) -> None:
    bri = _get(fields, "exec.bri.total")
    if bri is not None:
        box = _find_exact(slide, *EXEC_POS["committed_bri"])
        v = _fmt_sar_m(bri)
        if box and v:
            _set_run_text(box, 0, v)

    ipi = _get(fields, "exec.health.enterprise_ipi")
    if ipi is not None:
        box = _find_exact(slide, *EXEC_POS["enterprise_ipi"])
        try:
            ipi_val = float(ipi)
        except (TypeError, ValueError):
            ipi_val = None
        if box and ipi_val is not None:
            _set_run_text(box, 0, f"{ipi_val:.2f}")
            _set_run_font_color(box, STATUS_COLORS[_ipi_status(ipi_val)])

    sp = _get(fields, "exec.health.strategic_projects")
    mt = _get(fields, "exec.health.ms_total")
    if sp is not None or mt is not None:
        box = _find_exact(slide, *EXEC_POS["projects_milestones_line"])
        if box:
            if sp is not None and (v := _fmt_int(sp)):
                _set_run_text(box, 0, v)
            if mt is not None and (v := _fmt_int(mt)):
                _set_run_text(box, 6, v)

    if mt is not None:
        box = _find_exact(slide, *EXEC_POS["ms_total_line"])
        v = _fmt_int(mt)
        if box and v:
            runs = box.text_frame.paragraphs[0].runs
            if runs:
                last = runs[-1]
                last.text = re.sub(r"\d+", v, last.text, count=1) if re.search(r"\d+", last.text) else f"{v} Total"

    mc = _get(fields, "exec.health.ms_complete")
    if mc is not None:
        box = _find_exact(slide, *EXEC_POS["ms_complete"])
        v = _fmt_int(mc)
        if box and v:
            _set_run_text(box, 0, v)
            if mt not in (None,) and _fmt_int(mt):
                try:
                    pct = round(float(mc) / float(mt) * 100)
                    _set_run_text(box, 2, f"({pct}%)")
                except (TypeError, ValueError, ZeroDivisionError):
                    pass

    mnyd = _get(fields, "exec.health.ms_not_yet_due")
    if mnyd is not None:
        box = _find_exact(slide, *EXEC_POS["ms_not_yet_due"])
        v = _fmt_int(mnyd)
        if box and v:
            _set_run_text(box, 0, v)

    md = _get(fields, "exec.health.ms_delayed")
    if md is not None:
        box = _find_exact(slide, *EXEC_POS["ms_delayed"])
        v = _fmt_int(md)
        if box and v:
            _set_run_text(box, 0, v)


# Recovery tracker table (identical row grid on the CEO slide and its
# appendix duplicate). Calibrated to this deck: 10 physical rows starting at
# T=1.590in, spaced exactly 0.503in apart.
RECOVERY_SLIDES = [12, 35]
RECOVERY_ROW0_T = 1.590
RECOVERY_ROW_H = 0.503
RECOVERY_MAX_ROWS = 10
RECOVERY_COLS = {
    "rank": 0.620,
    "initiative": 0.940,       # 2 paragraphs: name / "{lob}  {type}"
    "bri": 3.350,
    "status": 4.220,           # text box + a status-colored bubble at the same spot
    "ipi_ti": 5.160,           # 2 paragraphs: "IPI x.xx" / "TI x.xx"
    "root_cause": 6.200,
    "corrective_action": 8.740,
    "owner": 11.500,           # 2 paragraphs: split on comma/newline
}
RECOVERY_STATUS_LABEL = {"Watch": "WATCH", "At-risk": "AT RISK"}
RECOVERY_STATUS_KEY = {"Watch": "cautious", "At-risk": "at_risk"}


def _set_two_paragraph(shape, line1: Optional[str], line2: Optional[str]) -> None:
    if not shape.has_text_frame:
        return
    paras = shape.text_frame.paragraphs
    if len(paras) >= 1 and line1 is not None:
        _set_paragraph_text(paras[0], line1)
    if len(paras) >= 2 and line2 is not None:
        _set_paragraph_text(paras[1], line2)


def _recovery_items(fields: Dict[str, Any], rows: Dict[str, int]) -> List[Dict[str, Any]]:
    n = int(rows.get("recovery.row", 0) or 0)
    items = []
    for i in range(n):
        p = f"recovery.row.{i}"
        name = fields.get(f"{p}.initiative")
        if not name:
            continue
        items.append({
            "rank": fields.get(f"{p}.rank") or (i + 1),
            "initiative": str(name),
            "lob": fields.get(f"{p}.lob") or "",
            "item_type": fields.get(f"{p}.item_type") or "",
            "bri": fields.get(f"{p}.bri"),
            "status": fields.get(f"{p}.status") or "Watch",
            "ipi": fields.get(f"{p}.ipi"),
            "ti": fields.get(f"{p}.ti"),
            "root_cause": fields.get(f"{p}.root_cause"),
            "corrective_action": fields.get(f"{p}.corrective_action"),
            "owner": fields.get(f"{p}.owner"),
        })
    items.sort(key=lambda it: _num_or(it["rank"], 9999))
    return items


def _num_or(v: Any, default: float) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _apply_recovery_tracker(slide, items: List[Dict[str, Any]]) -> None:
    for idx, item in enumerate(items[:RECOVERY_MAX_ROWS]):
        row_t = RECOVERY_ROW0_T + idx * RECOVERY_ROW_H

        box = _find_near(slide, RECOVERY_COLS["rank"], row_t, tol=0.2)
        if box:
            _set_box_text(box, str(int(_num_or(item["rank"], idx + 1))))

        box = _find_near(slide, RECOVERY_COLS["initiative"], row_t, tol=0.2)
        if box:
            _set_two_paragraph(box, item["initiative"], f"{item['lob']}  {item['item_type']}".strip())

        box = _find_near(slide, RECOVERY_COLS["bri"], row_t, tol=0.2)
        if box and item["bri"] is not None and (v := _fmt_int(item["bri"])):
            _set_box_text(box, v)

        box = _find_near(slide, RECOVERY_COLS["ipi_ti"], row_t, tol=0.2)
        if box:
            ipi_v = _num_or(item["ipi"], None) if item["ipi"] not in (None, "") else None
            ti_v = _num_or(item["ti"], None) if item["ti"] not in (None, "") else None
            ipi_s = f"IPI {ipi_v:.2f}" if ipi_v is not None else "IPI n/a"
            ti_s = f"TI {ti_v:.2f}" if ti_v is not None else "TI n/a"
            _set_two_paragraph(box, ipi_s, ti_s)

        box = _find_near(slide, RECOVERY_COLS["root_cause"], row_t, tol=0.2)
        if box and item["root_cause"]:
            _set_box_text(box, item["root_cause"])

        box = _find_near(slide, RECOVERY_COLS["corrective_action"], row_t, tol=0.2)
        if box and item["corrective_action"]:
            _set_box_text(box, item["corrective_action"])

        box = _find_near(slide, RECOVERY_COLS["owner"], row_t, tol=0.2)
        if box and item["owner"]:
            parts = re.split(r",|\n", str(item["owner"]), maxsplit=1)
            l1 = parts[0].strip()
            l2 = parts[1].strip() if len(parts) > 1 else ""
            _set_two_paragraph(box, l1, l2)

        status = item["status"] if item["status"] in RECOVERY_STATUS_LABEL else "Watch"
        box = _find_near(slide, RECOVERY_COLS["status"], row_t, tol=0.25)
        if box:
            _set_box_text(box, RECOVERY_STATUS_LABEL[status])
        bubble = _find_status_bubble(slide, RECOVERY_COLS["status"], row_t, tol=0.25)
        if bubble:
            _set_shape_fill(bubble, STATUS_COLORS[RECOVERY_STATUS_KEY[status]])


# --------------------------------------------------------------------------
# Line-of-business deep-dive tables (Health/Motor/General/Life initiatives +
# their "projects under each initiative" tables, plus CX and HR).
#
# Column X-positions are identical across every LoB (same template layout);
# only row Y-positions differ per LoB (initiative/project counts vary), so
# each LoB gets its own calibrated row list/grid, calibrated to this deck.
# --------------------------------------------------------------------------

STATUS_6_COLORS = {
    "On-track": "2E9E7B", "Cautious": "E0A52E", "Critical": "C0392B",
    "At-risk": "E8833A", "Not scored": "8A889E", "Overachieved": "2D9CDB",
}

# (left_min, left_max) column bands shared by every LoB's initiative table.
INIT_COLS = {
    "ipi": (4.0, 4.7),
    "ti": (5.7, 6.3),
    "bri_committed": (6.4, 8.6),
    "realization_date": (8.65, 9.9),
    "bri_actual": (10.85, 11.45),
    "status_dot": (12.25, 12.7),
}
NAME_COL_L = 0.620

# Project-table columns are identical across every LoB.
PROJ_NAME_L = 0.860
PROJ_IPI_L = 4.840
PROJ_TI_L = 6.440
PROJ_COMMENT_L = 7.100

LOB_INIT_CONFIG = {
    "health": {"slides": [8, 20], "row0_t": 1.760, "row_h": 0.337, "n_rows": 6},
    "motor": {"slides": [22], "row0_t": 1.760, "row_h": 0.345, "n_rows": 6},
    "general": {"slides": [24], "row0_t": 1.760, "row_h": 0.255, "n_rows": 8},
    "life": {"slides": [26], "row0_t": 1.760, "row_h": 0.345, "n_rows": 5},
}

# Project row Y-positions: irregular (depends on how many projects sit under
# each initiative), so calibrated as an explicit ordered list per LoB rather
# than a formula. Order matches the deck's own visual top-to-bottom order.
LOB_PROJ_CONFIG = {
    "health": {"slides": [9, 21], "row_ts": [1.842, 2.497, 2.830, 3.485, 3.819, 4.153, 4.808, 5.142, 5.797, 6.452]},
    "motor": {"slides": [23], "row_ts": [1.863, 2.225, 2.935, 3.297, 4.007, 4.717, 5.428, 5.789]},
    "general": {"slides": [25], "row_ts": [1.813, 2.396, 2.694, 2.991, 3.574, 4.158, 4.741, 5.325, 5.908, 6.492]},
    "life": {"slides": [27], "row_ts": [1.863, 2.574, 3.284, 3.994, 4.704]},
}


def _find_in_band(slide, row_t: float, l_min: float, l_max: float, tol_t: float = 0.15,
                    require_shape: bool = False):
    """First shape whose row is near row_t and whose left falls in [l_min, l_max].
    require_shape=True restricts to AUTO_SHAPE (used for status dots)."""
    for sh in slide.shapes:
        if require_shape and sh.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
            continue
        if not require_shape and not (sh.has_text_frame and sh.text_frame.text.strip()):
            continue
        L, T = _left_top_in(sh)
        if L is None or abs(T - row_t) > tol_t:
            continue
        if l_min <= L <= l_max:
            return sh
    return None


def _apply_lob_initiatives(slide, lob_id: str, fields: Dict[str, Any], rows: Dict[str, int]) -> None:
    cfg = LOB_INIT_CONFIG[lob_id]
    table_id = f"{lob_id}.init"
    n = min(int(rows.get(table_id, 0) or 0), cfg["n_rows"])
    for i in range(n):
        row_t = cfg["row0_t"] + i * cfg["row_h"]
        p = f"{table_id}.{i}"

        name = fields.get(f"{p}.name")
        if name:
            box = _find_near(slide, NAME_COL_L, row_t, tol=0.05)
            if box:
                _set_two_paragraph(box, str(name), fields.get(f"{p}.note") or "")

        ipi = fields.get(f"{p}.ipi")
        if ipi not in (None, ""):
            box = _find_in_band(slide, row_t, *INIT_COLS["ipi"])
            if box:
                try:
                    _set_box_text(box, f"{float(ipi):.2f}")
                except (TypeError, ValueError):
                    pass

        ti = fields.get(f"{p}.ti")
        if ti not in (None, ""):
            box = _find_in_band(slide, row_t, *INIT_COLS["ti"])
            if box:
                try:
                    _set_box_text(box, f"{float(ti):.2f}")
                except (TypeError, ValueError):
                    pass

        bri_c = fields.get(f"{p}.bri_committed")
        if bri_c not in (None, ""):
            box = _find_in_band(slide, row_t, *INIT_COLS["bri_committed"])
            if box and (v := _fmt_int(bri_c)):
                _set_box_text(box, v)

        realization = fields.get(f"{p}.realization_date")
        if realization:
            box = _find_in_band(slide, row_t, *INIT_COLS["realization_date"])
            if box:
                _set_box_text(box, str(realization))

        bri_a = fields.get(f"{p}.bri_actual")
        if bri_a not in (None, ""):
            box = _find_in_band(slide, row_t, *INIT_COLS["bri_actual"])
            if box and (v := _fmt_int(bri_a)):
                _set_box_text(box, v)

        status = fields.get(f"{p}.status")
        if status in STATUS_6_COLORS:
            dot = _find_in_band(slide, row_t, *INIT_COLS["status_dot"], require_shape=True)
            if dot:
                _set_shape_fill(dot, STATUS_6_COLORS[status])


def _apply_lob_projects(slide, lob_id: str, fields: Dict[str, Any], rows: Dict[str, int]) -> None:
    cfg = LOB_PROJ_CONFIG[lob_id]
    table_id = f"{lob_id}.proj"
    row_ts = cfg["row_ts"]
    n = min(int(rows.get(table_id, 0) or 0), len(row_ts))
    for i in range(n):
        row_t = row_ts[i]
        p = f"{table_id}.{i}"

        project = fields.get(f"{p}.project")
        if project:
            box = _find_near(slide, PROJ_NAME_L, row_t, tol=0.05)
            if box:
                _set_box_text(box, str(project))

        ipi = fields.get(f"{p}.ipi")
        if ipi not in (None, ""):
            box = _find_near(slide, PROJ_IPI_L, row_t, tol=0.1)
            if box:
                try:
                    _set_box_text(box, f"{float(ipi):.2f}")
                except (TypeError, ValueError):
                    pass

        ti = fields.get(f"{p}.ti")
        if ti not in (None, ""):
            box = _find_near(slide, PROJ_TI_L, row_t, tol=0.1)
            if box:
                try:
                    _set_box_text(box, f"{float(ti):.2f}")
                except (TypeError, ValueError):
                    pass

        comment = fields.get(f"{p}.comment")
        if comment:
            box = _find_near(slide, PROJ_COMMENT_L, row_t, tol=0.1)
            if box:
                _set_box_text(box, str(comment))


# CX projects table (identical row grid on the CEO slide and its appendix
# duplicate) and the single-initiative HR slide.
CX_PROJ_SLIDES = [10, 28]
CX_PROJ_ROW_TS = [3.963, 4.285]

HR_SLIDE = 31
HR_INIT_POS = (0.620, 2.260)   # 2-paragraph: note only (name is fixed "Operating-model redesign")
HR_INIT_IPI_L = 4.350
HR_INIT_TI_L = 6.050
HR_PROJ_ROW_T = 4.046
HR_PROJ_NAME_L = 0.896
HR_PROJ_COMMENT_L = 7.220


def _apply_cx_projects(slide, fields: Dict[str, Any], rows: Dict[str, int]) -> None:
    n = min(int(rows.get("cx.proj", 0) or 0), len(CX_PROJ_ROW_TS))
    for i in range(n):
        row_t = CX_PROJ_ROW_TS[i]
        p = f"cx.proj.{i}"
        name = fields.get(f"{p}.name")
        if name:
            box = _find_near(slide, PROJ_NAME_L, row_t, tol=0.05)
            if box:
                _set_box_text(box, str(name))
        ipi = fields.get(f"{p}.ipi")
        if ipi not in (None, ""):
            box = _find_near(slide, PROJ_IPI_L, row_t, tol=0.1)
            if box:
                try:
                    _set_box_text(box, f"{float(ipi):.2f}")
                except (TypeError, ValueError):
                    pass
        ti = fields.get(f"{p}.ti")
        if ti not in (None, ""):
            box = _find_near(slide, PROJ_TI_L, row_t, tol=0.1)
            if box:
                try:
                    _set_box_text(box, f"{float(ti):.2f}")
                except (TypeError, ValueError):
                    pass


def _apply_hr(slide, fields: Dict[str, Any]) -> None:
    note = fields.get("hr.init.note")
    if note:
        box = _find_near(slide, *HR_INIT_POS, tol=0.05)
        if box:
            paras = box.text_frame.paragraphs
            if len(paras) >= 2:
                _set_paragraph_text(paras[1], str(note))
    ipi = fields.get("hr.init.ipi")
    if ipi not in (None, ""):
        box = _find_near(slide, HR_INIT_IPI_L, HR_INIT_POS[1], tol=0.1)
        if box:
            try:
                _set_box_text(box, f"{float(ipi):.2f}")
            except (TypeError, ValueError):
                pass
    ti = fields.get("hr.init.ti")
    if ti not in (None, ""):
        box = _find_near(slide, HR_INIT_TI_L, HR_INIT_POS[1], tol=0.1)
        if box:
            try:
                _set_box_text(box, f"{float(ti):.2f}")
            except (TypeError, ValueError):
                pass
    project = fields.get("hr.proj.0.project")
    if project:
        box = _find_near(slide, HR_PROJ_NAME_L, HR_PROJ_ROW_T, tol=0.05)
        if box:
            _set_box_text(box, str(project))
    comment = fields.get("hr.proj.0.comment")
    if comment:
        box = _find_near(slide, HR_PROJ_COMMENT_L, HR_PROJ_ROW_T, tol=0.15)
        if box:
            _set_box_text(box, str(comment))


def export_pptx(template_bytes: bytes, state: Dict[str, Any]) -> bytes:
    """Return a new .pptx (bytes) = template with mapped values substituted."""
    fields = state.get("fields", {})
    rows = state.get("rows", {})
    prs = Presentation(io.BytesIO(template_bytes))
    n = len(prs.slides)
    sector_rows = _sector_rows(fields, rows)

    for one_based, suffix in DIVIDER_SLIDES.items():
        if one_based <= n:
            _apply_divider(prs.slides[one_based - 1], fields, suffix)
    for idx in COVER_SLIDES:
        if idx <= n:
            _apply_cover(prs.slides[idx - 1], fields)
    for idx in CLOSING_SLIDES:
        if idx <= n:
            _apply_closing(prs.slides[idx - 1], fields)
    for idx in IPI_SECTOR_SLIDES:
        if idx <= n:
            _apply_ipi_sector(prs.slides[idx - 1], sector_rows)
    for idx in EXEC_SUMMARY_SLIDES:
        if idx <= n:
            _apply_exec_summary(prs.slides[idx - 1], fields)
    recovery_items = _recovery_items(fields, rows)
    if recovery_items:
        for idx in RECOVERY_SLIDES:
            if idx <= n:
                _apply_recovery_tracker(prs.slides[idx - 1], recovery_items)

    for lob_id, cfg in LOB_INIT_CONFIG.items():
        for idx in cfg["slides"]:
            if idx <= n:
                _apply_lob_initiatives(prs.slides[idx - 1], lob_id, fields, rows)
    for lob_id, cfg in LOB_PROJ_CONFIG.items():
        for idx in cfg["slides"]:
            if idx <= n:
                _apply_lob_projects(prs.slides[idx - 1], lob_id, fields, rows)
    for idx in CX_PROJ_SLIDES:
        if idx <= n:
            _apply_cx_projects(prs.slides[idx - 1], fields, rows)
    if HR_SLIDE <= n:
        _apply_hr(prs.slides[HR_SLIDE - 1], fields)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def template_info(template_bytes: bytes) -> Dict[str, Any]:
    prs = Presentation(io.BytesIO(template_bytes))
    slides = []
    for i, s in enumerate(prs.slides, 1):
        title = ""
        for sh in _text_boxes(s):
            title = sh.text_frame.text.strip().split("\n")[0][:60]
            break
        slides.append({"index": i, "title": title, "text_boxes": len(_text_boxes(s))})
    return {"slide_count": len(prs.slides), "slides": slides}
