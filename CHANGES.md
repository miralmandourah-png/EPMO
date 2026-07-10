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

## Phase 7 — Blank NPS content from the exported PPTX (this release)

Phase 6 removed NPS from the dashboard, but the exported deck still showed
whatever NPS content was baked into the user's own template (since export
only overwrites what the data model maps). User asked for the export itself
to have NPS removed too. Chose (per the user's decision) to **blank values
only** — clear NPS text/tables wherever they appear, but leave every slide,
its layout, and page numbers exactly as-is; no slide deletion or
renumbering.

**Added to `pptx_export.py`**
- Blanks the "NPS (sector)" row (label + actual + target-delta) on every
  LoB's Financials block (Health incl. its duplicate slide, Motor, General,
  Life).
- Blanks the CX NPS-by-LoB slide (and its appendix duplicate): the 5-row
  actual/target table, both NPS section headers, and the "Read:" note.
  Surgically strips just the NPS clause from its headline via regex
  (`"...strong execution (IPI 3.82), but NPS lags..."` → `"...strong
  execution (IPI 3.82)"`) rather than blanking the whole headline, since the
  rest of the sentence isn't NPS content.
- Blanks the CX NPS-by-segment slide's headline/subheadline/headers/note,
  and **removes the nested GROUP shape** holding the per-LoB segment tables
  (that content lives inside a PowerPoint group, invisible to a flat
  top-level shape scan -- found this while investigating why the slide's
  text dump looked emptier than its visible content).
- Found two more headlines that weave NPS into otherwise-unrelated sentences
  (General's "best NPS, but execution is just below the line (IPI 2.82) and
  GWP is 22% short"; Life's "execution and NPS lag (IPI 2.66)") plus their
  subheadlines and one EPMO-feedback note. Used the same surgical-regex
  approach so the non-NPS content (IPI, GWP, loss ratio, committed benefit)
  survives -- full-sentence blanking would have thrown away real information
  that happened to share a sentence with an NPS mention.

**Verified**: scanned the full 36-slide exported deck case-insensitively for
"nps" -- zero remaining mentions. Confirmed shape-count diff is exactly one
slide (30, where the group was removed) and every other slide is untouched
structurally. Re-ran the full existing test suite (constants extraction, IPI/
Milestones import, recovery sync, LoB initiative export) on top of this to
confirm no interference with anything shipped earlier.

## Phase 6 — Remove NPS entirely

Per user request. Removed from `schema.py`:
- Each LoB's "May-26 YTD Financials" NPS actual/target field.
- The Customer Experience tab's NPS-by-line-of-business table and the four
  NPS-by-segment tables (Health/Motor/General/Life).
- The now-unused `cx.nps_lob` default-row seed.

Nothing else in the codebase referenced NPS (no export mapping had been
built for those slides yet, so there was nothing to unwind there). Verified
the app still boots, the schema loads, and the full import/export round
trip still works after removal.

## Phase 5 — Map the LoB deep-dive tables, CX, and HR

Per the user's request to map everything, plus "the initiative names and
project names ... populate them from the presentation into the platform,
they will be the same for now" (i.e. treat names as constant identifiers,
same principle as Phase 4).

**Added**
- Export mapping for all four LoB initiative tables (Health/Motor/General/
  Life) and their "projects under each initiative" tables, plus CX projects
  and the HR slide. Column X-positions are identical across every LoB;
  initiative-table row heights are calibrated per LoB (Health 0.337in,
  Motor/Life 0.345in, General 0.255in for its 8 rows) and project-table rows
  use an explicit per-LoB Y-position list (project counts per initiative are
  irregular, so no formula fits).
- Discovered the deck actually uses **two different status vocabularies**:
  the recovery tracker's Watch(amber)/At-risk(red), already shipped
  correctly, and a separate **6-value initiative status** (On-track/
  Cautious/Critical/At-risk/Not scored/Overachieved) with its own distinct
  colors read directly from the deck's own legend swatches (On-track green,
  Cautious amber, Critical red, **At-risk orange** `#E8833A` -- distinct from
  Critical's red, Not scored grey, Overachieved blue). Initiative status dots
  now recolor dynamically from this palette.
- Extended `pptx_constants.py` (same principle as Phase 4) to also read
  initiative name/note/committed-BRI and project names straight from the
  template, reusing the exact same coordinate config as the export so read
  and write can never drift apart.
- Fixed a formatting bug found while verifying the round-trip: BRI values
  were losing their thousands separator on export (e.g. "1,066" became
  "1066"); `_fmt_int` now matches the deck's own number style everywhere
  it's used.

**Verified**: synthetic non-confidential test data written to every LoB's
initiative/project rows and read back correctly on both the Health slide and
its appendix duplicate, status-dot recoloring across all six values,
brand shape counts unchanged across the full 36-slide deck. Then the real
round-trip: Extract targets from template → dashboard's Health tab shows
"Corporate KAM" / committed BRI 576 with IPI/TI/actual left blank → Export →
slide 8 shows "SME acquisition" and "1,066" -- all confirmed through a live
browser run with zero console errors.

**Not yet mapped**: CX's NPS-by-LoB and NPS-by-segment slides, and the
Executive Summary's execution-tier bars / benefit-status bubbles (those are
live status breakdowns, not names, so lower priority per the user's own
constant-vs-status distinction).

## Phase 4 — Extract committed/target constants from the template

Per user instruction: "any number that mentions benefits or BRI meaning if
it's a target is a constant ... write them for me in the platform (source
them from the slides)". Committed/target figures are stable for the year and
unlikely to change monthly, unlike status numbers (actuals, IPI/TI, % on
track) which stay driven by manual entry or the other importers.

**Added**
- `pptx_constants.py` — the read-direction mirror of the export mapping:
  extracts committed BRI, committed savings, the Strategy & Ambition
  2026→2030 table (GWP/profit/ROE/initiative-share), committed BRI by line
  of business, the savings breakdown, and GWP by year/by-LoB directly from
  the uploaded template's own slides (Executive Summary + Context).
- New **Extract targets from template** action (Data ▾), reusing whichever
  PPTX template is already configured -- no separate upload needed.
- Coordinate-anchoring bug fixed along the way: python-pptx returns a new
  wrapper object each time you iterate `slide.shapes`, so comparing shapes
  with `is` across two separate iterations never actually excludes anything.
  Switched to comparing the underlying XML element (`sh._element is
  other._element`), and added horizontal region bounds so a same-row lookup
  can't wander into an unrelated panel on the other side of the slide at a
  coincidentally similar height.

**Verified**: every extracted value checked field-by-field against the
source deck (committed BRI 2,961; savings 177; all 8 ambition figures;
benefit-by-LoB 1,726/599/503/133; GWP by year and by LoB, cross-checked
arithmetically against each row's own displayed total) -- all exact matches,
through the full HTTP API and a live browser run (upload template → Extract
targets → dashboard fields populated correctly, zero console errors).

**Known simplification**: the savings-breakdown "AI 8-12" range is stored as
its first number (8); a true range isn't representable in a single numeric
field.

## Phase 3 — EPMO Recovery Tracker auto-flagging

Per user instruction: "projects with low TI ... should be shown in the EPMO
recovery tracker" so they can chase root cause / corrective action with the
responsible department. Threshold agreed with the user: **TI < 3.0**,
matching the deck's own "TI (vs 3.0)" labeling.

**Added**
- `automation.py` — `sync_recovery_tracker`: scans every LoB's initiatives
  table for TI below the threshold and upserts a matching row into the
  recovery tracker, matched by (LoB, initiative). New rows get root cause /
  corrective action / owner left **blank**; existing rows get BRI/IPI/TI
  refreshed without touching anything already typed into those blank fields.
  Nothing is ever auto-removed (if an item recovers, its row is left for the
  user to delete once resolved). Rows are re-ranked by BRI size descending on
  every sync, matching the deck's own "largest ... by size" ordering.
- Runs only when the user clicks **Sync Recovery Tracker** (Data ▾) — not on
  every autosave — so deleting a flagged row doesn't have it silently
  reappear on the next keystroke.
- Schema change: added a `Type` (growth/savings) column to each LoB's
  initiatives table, and split the recovery tracker's `lob_type`/`ipi_ti`
  combined text columns into structured `lob`/`item_type`/`ipi`/`ti` fields
  so the sync engine can match and update rows reliably (the combined
  "Health\ngrowth" / "IPI 2.21\nTI 1.72" display is now composed only at
  export time).
- Export mapping for the recovery tracker slide (and its appendix duplicate):
  10-row grid calibrated to this deck, including the colored status bubble
  (Watch=amber, At-risk=red) recolored dynamically from the row's own status.

**Verified**: end-to-end through the real HTTP API and a live browser click-
through (seed an initiative → Sync Recovery Tracker → summary modal → row
appears on the Recovery Tracker tab with correct dropdowns); re-sync
preserves manually-entered root cause; BRI-descending re-ranking confirmed;
export writes correct text/colors to both the CEO slide and its appendix
duplicate with brand shape counts unchanged.

## Phase 2 — In-place branded PPTX export

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
