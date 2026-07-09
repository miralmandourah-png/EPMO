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

## Phase 2 — In-place branded PPTX export (next)

Coordinate-mapped, in-place editing of the uploaded Tawuniya template:
open with python-pptx, locate each value's text box by position on the mapped
slide, replace only the text (preserving all brand XML), honoring the
duplicate-slide rule. Template file path and the tab→slide mapping persist in
`data/state.json`.
