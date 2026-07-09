# Changes

## What already existed (commit "Add Deck Builder: Excel-driven PPTX generator")

A first-pass "Excel-driven PPTX generator": a FastAPI app that generated a
deck **from scratch** with python-pptx, round-tripped a multi-sheet Excel
workbook, and served a small generic tab UI. It was **stateless** (in-memory
only, no persistence, no autosave), had a **shallow data model** that didn't
match the real "Strategic Health Check" deck, built slides from a blank
presentation (so no brand fidelity), and had no knowledge of the real
operational Excel exports.

## Audit finding that shaped this work

The real Tawuniya template is **36 slides with zero tables** — every data
slide is 60–200 individually-positioned text boxes arranged as a visual grid.
So "edit table cells" doesn't apply; in-place export must be **coordinate-
mapped** box-by-box to this specific deck (Phase 2). Duplicate CEO/appendix
slide pairs were confirmed (7≡19, 8≡20, 9≡21, 10≡28, 11≡29, 12≡35, 5≡15,
6≡17), so the model stores one copy and populates both on export.

## Phase 1 — Dashboard, storage, ingestion (this release)

**Added**
- `app/schema.py` — single source of truth for all **14 sections** with a
  stable dotted-key scheme (`exec.bri.total`, `health.init.0.ipi`, …) that
  every layer shares.
- `app/store.py` — **local JSON persistence** at `data/state.json` with atomic
  writes, `default_state` (seeds section labels only, never numbers), autosave
  merge, and reset.
- Flat **`Key | Label | Value` Excel template** generation + parse in
  `app/excel_io.py`, replacing the old multi-sheet round-trip.
- **IPI Accountability importer** — finds the summary label row by exact
  string match (no hardcoded row index), writes Enterprise IPI and
  `SBP + SEP`, returns a read-only reference summary, errors cleanly if labels
  are missing (verified: IPI 2.78, 81 strategic projects, 291 total).
- **Milestones importer** — column lookup by header name, skips *Used filters*,
  classifies complete/delayed/not-yet-due with a **configurable grace period**
  (verified: 128 → 65 / 17 / 46; grace 30 → 65 / 14 / 49).
- **New 14-tab single-page dashboard** (`app/static/`) — inline click-to-edit
  fields, add/remove table rows, **debounced autosave** with a status
  indicator, status dropdowns rendered as colored **badges**, **IPI bars**
  (green ≥3.0 / amber ≥2.6 / red <2.6), computed columns (totals, gaps),
  import modals with reference summaries, and Reset-to-blank confirm.
- New backend routes: `GET /api/schema`, `GET/PUT /api/state`, `POST
  /api/reset`, `GET /api/template`, `POST /api/import/template`,
  `POST /api/import/ipi`, `POST /api/import/milestones`.
- Tawuniya brand palette (primary `#5B4FE0`), README, this file, `.gitignore`
  hardening (`/data/`, uploaded files never committed).

**Changed** — `app/main.py` rewritten to the stateful model; `app/theme.py`
primary color aligned to `#5B4FE0`.

**Removed** — `app/ppt_builder.py`, `app/ppt_helpers.py`, `app/models.py` (the
old build-from-scratch export and shallow model). Phase 2 replaces these with
an in-place template editor.

**Verified** — every backend route via HTTP (TestClient) against the three
real uploaded files, and the UI in a real headless-Chromium smoke test
(tab switching, autosave to disk, add-row, badges, IPI bars, computed cells,
menus, reset modal). A CSS bug found by that test (an invisible modal overlay
intercepting all clicks) was fixed.

## Phase 2 — In-place branded PPTX export (this release)

**Added** — `app/pptx_export.py` plus template routes (`POST
/api/template/pptx`, `GET /api/template/pptx/info`, real `POST
/api/export/pptx`) and a **Upload PPTX template** UI action.

- Opens the user's uploaded template and mutates **only text runs** — every
  other piece of slide XML is untouched, so brand fidelity is exact.
  **Verified**: exported deck keeps all 36 slides and an identical shape count
  on every slide; it reopens cleanly.
- Anchoring never uses the template's confidential figures — boxes are found
  by position (coordinate map calibrated to this deck, per the user's choice)
  and non-confidential structural labels (sector names, section numbers).
- **Duplicate-slide rule** honored: the IPI-by-sector values are written from
  one model copy to both slide 6 and its appendix duplicate slide 17 (verified).
- Template file is stored locally at `data/template.pptx` (gitignored); its
  name is remembered in `data/state.json`.

**Mapped so far**: cover, the five section dividers, both closing cards, the
six IPI-by-sector values (× their duplicates), and — as of this update — the
**Executive Summary slide's committed BRI, Enterprise IPI, strategic project
count, and Strategic Milestones total/complete/not-yet-due/delayed** (the
exact fields the IPI Accountability and Milestones importers populate). This
closes the gap where those two importers wrote correctly to the dashboard but
the export didn't reflect them yet: several of that slide's numbers live as
one isolated run inside a longer label (e.g. "95  Strategic projects …773
milestones"), so a new `_set_run_text` helper replaces exactly one run,
leaving every neighboring tab/line-break/format untouched. Verified against
the real uploaded files end-to-end, including through a live browser click-
through (upload template → import IPI → import Milestones → Export),
confirming the numbers land correctly and per-slide shape counts are
unchanged.

**Dynamic status coloring (this update)**: the deck states its own rule in
plain text ("on-track ≥ 3.0", slides 4/6/12/17). Export now recomputes color
from the live number instead of leaving whatever color the template shipped
with — the sector status bubbles on the IPI-by-sector slide (and its
duplicate) and the Enterprise IPI number both recolor green/amber/red as
values change (≥3.0 on-track, ≥2.6 cautious, <2.6 at-risk — the same rule
already used for the dashboard's IPI bars, so dashboard and export never
disagree). Verified: boundary values (3.0, 2.6, 2.59) land in the correct
bucket, and forcing a value into "at-risk" turns the actual bubble shape and
number red end-to-end through the export route, with brand shape counts
unchanged.

**Not yet mapped**: the dense per-initiative deep-dive grids (Health/Motor/
General/Life), the Strategy & Ambition 2026→2030 table, benefit-status-by-LoB
and savings-breakdown tables, recovery tracker, scenarios — these are left
exactly as in the template until their per-cell coordinates are calibrated.
The engine (`_find_near`/`_find_exact`, `_find_by_label`, `_value_right_of`,
`_set_box_text`, `_set_run_text`) and the `*_SLIDES` duplicate lists are in
place to extend to them.
