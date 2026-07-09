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
        if sh is anchor:
            continue
        L, T = _left_top_in(sh)
        if L is None:
            continue
        if abs(T - aT) <= tol_top and L - aL >= min_dx and (L - aL) < best_dx:
            best, best_dx = sh, L - aL
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
        L, _ = _left_top_in(sh)
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
        return str(int(round(float(v))))
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
            if box:
                _set_run_text(box, 0, f"{float(ipi):.2f}")
        except (TypeError, ValueError):
            pass

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
