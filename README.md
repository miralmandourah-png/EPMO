# Strategic Health Check Platform

A local, single-user web app that replaces the manual monthly *"Strategic
Health Check by LoB"* CEO PowerPoint process. You edit an executive reporting
dashboard inline across **14 sections**, populate it from Excel exports, it
**autosaves to a local file**, and (Phase 2) it exports a branded PowerPoint
deck by editing your existing template in place.

## Your data stays on your machine

The app runs entirely on your computer. It makes **no calls to any cloud
service, third-party API, analytics, or telemetry**. All dashboard data is
stored in a single local file, `data/state.json`. Excel and PowerPoint files
you upload are read in memory only for that one request and are never sent
anywhere. The `data/` folder and any uploaded source files are gitignored so
they are never committed.

---

## Setup & run (for a non-technical user)

You need **Python 3.9 or newer**. Check with `python3 --version`; if you don't
have it, install from [python.org](https://www.python.org/downloads/).

Open a terminal **in this project folder** and run these once:

```bash
python3 -m venv .venv
source .venv/bin/activate          # on Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

Then start the app (this is the command you'll use every time):

```bash
uvicorn app.main:app --reload
```

Open **http://localhost:8000** in your browser. To stop it, press `Ctrl+C`.

Next time, you only need:

```bash
source .venv/bin/activate
uvicorn app.main:app --reload
```

---

## Using it

The top bar has a **tab per section** (Executive Summary, Context, Execution
Signal, LoB Scorecard, the four LoB deep-dives, Customer Experience, Human
Resources, Recovery Tracker, Scenarios, Summary & Actions, Dividers). Every
number and text box is click-to-edit and **autosaves** — the top bar shows
"Saving…" then "All changes saved".

### Three ways to get data in

Use the **Data ▾** menu (top right):

1. **Manual editing** — just type into any field.
2. **Full Excel template** — *Download full Excel template* gives you an
   `.xlsx` with a `Key | Label | Value` row for every field, pre-filled with
   your current values. Fill in the **Value** column and *Import filled
   template* to overwrite everything by key.
3. **Source-system exports** (targeted fields only — everything else is left
   untouched):
   - **Import IPI Accountability** — reads the *Data Table* sheet's summary
     block, writes **Enterprise IPI** and **Strategic project count**
     (`SBP + SEP`) to the Executive Summary, and shows a read-only reference
     summary (total projects, CP vs strategic split, full status breakdown)
     so you can sanity-check. If the expected labels aren't found it shows an
     error and changes nothing.
   - **Import Milestones** — classifies each milestone as **complete**
     (progress ≥ 100%), **delayed** (past due date and not complete), or
     **not-yet-due**, and writes the counts to the Executive Summary. The
     delayed cutoff has a configurable **grace period** (default 0 days).

**Reset to blank** (in the Data menu) clears everything behind a confirm
dialog.

### Branded PowerPoint export

1. In the **Data ▾** menu, choose **Upload PPTX template** and select your
   Tawuniya-branded `.pptx` once (it's remembered locally).
2. Click **Export PPTX** (top right). The app opens *your* template and
   substitutes only the text values — all backgrounds, theme colors, fonts,
   logo, and layout are left exactly as they are, so brand fidelity is exact.
   Duplicate CEO/appendix slides are filled from the same data automatically.

The export button is dimmed until a template is uploaded. Currently mapped:
the cover, section dividers, closing cards, and the IPI-by-sector values;
the dense per-initiative tables are the next mapping step (see CHANGES.md).

---

## Project layout

```
app/
  schema.py     # single source of truth: 14 sections, every field + dotted key
  store.py      # local JSON persistence (data/state.json), autosave, reset
  excel_io.py   # flat Key|Label|Value template + IPI & Milestones importers
  theme.py      # Tawuniya brand colors / fonts / geometry (used by export)
  main.py       # FastAPI routes
  static/       # single-page dashboard (plain HTML/CSS/JS, no build step)
data/
  state.json    # your data (created on first run; gitignored)
```
