# Deck Builder

A local web app for turning an Excel workbook of strategy-review numbers into a
branded PowerPoint deck: upload a workbook, review/edit the numbers in the
browser, then export a `.pptx` (and, if you want, a cleaned-up `.xlsx` back).

It reproduces the recurring slide types from a strategic-review deck
(cover, contents, executive summary KPIs, GWP context charts, IPI-by-sector,
line-of-business deep-dive tables, NPS, an EPMO recovery tracker, scenarios,
and recommended actions) as reusable templates driven entirely by your data —
no company-specific numbers are built into the tool itself.

## Confidentiality

- Nothing you upload or edit is written to disk or logged on the server.
  Uploaded workbooks and edited data live only in server request memory
  (parsed and discarded) and in your browser tab's memory (lost on reload).
- This is meant to run locally / on infrastructure you control. Don't expose
  it on the open internet without adding authentication.

## Running it

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Then open http://127.0.0.1:8000 in your browser.

## Workflow

1. Click **Blank Template** (or **Sample Template** to see the shape of the
   data with placeholder numbers) to download an `.xlsx`.
2. Fill in your real numbers in Excel — one sheet per section (see below).
3. Click **Upload Excel** and select your filled-in workbook. The app parses
   it and populates every editable tab.
4. Tweak numbers, headlines, or table rows directly in the browser.
5. Click **Export PPTX** to download the generated deck, or **Save as Excel**
   to download your edits back into the same workbook shape.

## Workbook structure

- **Meta** — headlines, subheadlines, narratives and single numbers (one
  `Section / Field / Value` row per item).
- One sheet per repeating table: `TOC`, `IPI_Tiers`, `LoB_Status`,
  `Savings_Breakdown`, `GWP_Years`, `GWP_By_LoB`, `IPI_By_Sector`,
  `LoB_Overview`, `Initiatives`, `Projects`, `NPS_By_LoB`, `NPS_Detail`,
  `Recovery_Tracker`, `Scenarios`, `Recommended_Actions`.

`Initiatives` and `Projects` are long-format tables with a `lob` column —
each distinct value in that column becomes its own line-of-business
deep-dive slide (and, for `Projects`, a matching "projects under each
initiative" slide), so you don't need to add sheets per line of business.

`status` and `item_type` columns are constrained to fixed values
(`on_track` / `cautious` / `at_risk`, and `growth` / `savings`) via dropdown
validation in the template, since slide coloring is keyed off those exact
strings.

A slide type is only generated if its underlying sheet(s) have data — leave a
sheet empty and that slide/section is skipped.

## Project layout

```
app/
  models.py       # DeckData schema (pydantic)
  theme.py        # colors, fonts, slide geometry
  excel_io.py     # xlsx <-> DeckData (template generation + parsing)
  ppt_helpers.py  # low-level python-pptx drawing helpers
  ppt_builder.py  # one build_* function per slide type + orchestration
  main.py         # FastAPI routes
  static/         # vanilla HTML/CSS/JS frontend (no build step)
```

## Known limitations

- Headlines/subheadlines for the per-LoB deep-dive slides are generated
  automatically from the LoB name (e.g. "Health — Initiative Execution")
  rather than freely editable per line of business.
- The GWP trajectory chart uses a native PowerPoint chart object; every other
  chart-like visual (IPI bars, NPS bars, progress bars) is drawn as plain
  shapes so an on-track reference line and custom coloring could be added
  precisely.
