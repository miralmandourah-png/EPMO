"""Single source of truth for the Strategic Health Check data model.

Everything in the app is driven by SECTIONS below: the 14 dashboard tabs,
the inline-edit UI, the flat `Key | Label | Value` Excel template, the two
smart importers, autosave, and (phase 2) the PPTX coordinate map all read
these same dotted keys.

Key scheme
----------
* Scalar field:      "<section>.<group>.<name>"     e.g. "exec.amb.gwp_2026"
* Repeating table:   "<tableId>.<rowIndex>.<col>"    e.g. "health.init.0.ipi"

A "block" is one visual group inside a tab and is either:
    {"type": "fields", ...}  -> a grid of scalar fields
    {"type": "table",  ...}  -> an add/removable grid of rows

Field types: text | number | pct | textarea | select (needs "options").
Table columns reuse the same type vocabulary; a column may set
"computed": "<expr>" to render a read-only derived cell (evaluated client
side from sibling columns, e.g. total = bau + init).
"""
from __future__ import annotations

from typing import Any, Dict, List

STATUS_6 = ["On-track", "Cautious", "Critical", "At-risk", "Not scored", "Overachieved"]
WATCH_2 = ["Watch", "At-risk"]
TYPE_2 = ["growth", "savings"]

LOBS = [
    ("health", "Health"),
    ("motor", "Motor"),
    ("general", "General"),
    ("life", "Life"),
]


def _f(key: str, label: str, type: str = "text", **kw) -> Dict[str, Any]:
    d = {"key": key, "label": label, "type": type}
    d.update(kw)
    return d


def _col(name: str, label: str, type: str = "text", **kw) -> Dict[str, Any]:
    d = {"name": name, "label": label, "type": type}
    d.update(kw)
    return d


def _actual_target(prefix: str, label: str) -> List[Dict[str, Any]]:
    """A metric with an actual value and a free-text target/budget comparison."""
    return [
        _f(f"{prefix}.actual", f"{label} — actual", "text"),
        _f(f"{prefix}.target", f"{label} — target/budget", "text"),
    ]


def _lob_deepdive(lob_id: str, lob_label: str) -> Dict[str, Any]:
    """Sections 5-8: Health / Motor / General / Life share this structure."""
    return {
        "id": lob_id,
        "label": lob_label,
        "blocks": [
            {"type": "fields", "title": "Slide heading", "fields": [
                _f(f"{lob_id}.head.title", "Headline", "text"),
                _f(f"{lob_id}.head.subtitle", "Subtitle", "textarea"),
            ]},
            {"type": "table", "title": "Initiatives", "id": f"{lob_id}.init", "maxRows": 12, "columns": [
                _col("name", "Initiative"),
                _col("note", "Note"),
                _col("item_type", "Type", "select", options=TYPE_2),
                _col("ipi", "IPI", "number"),
                _col("ti", "TI", "number"),
                _col("bri_committed", "BRI committed (SAR m)", "number"),
                _col("bri_actual", "BRI actual (SAR m)", "number"),
                _col("status", "Status", "select", options=STATUS_6),
                _col("realization_date", "BRI realization date"),
            ]},
            {"type": "fields", "title": "May-26 YTD Financials", "fields": [
                *_actual_target(f"{lob_id}.fin.gwp_ytd", "GWP YTD"),
                *_actual_target(f"{lob_id}.fin.loss_ratio", "Loss ratio"),
                *_actual_target(f"{lob_id}.fin.combined_ratio", "Combined ratio"),
                *_actual_target(f"{lob_id}.fin.profit_margin", "Profit margin"),
            ]},
            {"type": "fields", "title": "Stakeholder View", "fields": [
                _f(f"{lob_id}.stake.strategy", "Strategy View", "textarea"),
                _f(f"{lob_id}.stake.epmo", "EPMO Feedback", "textarea"),
                _f(f"{lob_id}.stake.finance", "Finance Feedback", "textarea"),
            ]},
            {"type": "table", "title": "Projects under each initiative", "id": f"{lob_id}.proj", "maxRows": 30, "columns": [
                _col("initiative", "Initiative (group)"),
                _col("project", "Project"),
                _col("ipi", "IPI", "number"),
                _col("ti", "TI", "number"),
                _col("comment", "Comment"),
            ]},
        ],
    }


def _build_sections() -> List[Dict[str, Any]]:
    sections: List[Dict[str, Any]] = []

    # 1. Executive Summary --------------------------------------------------
    sections.append({"id": "exec", "label": "Executive Summary", "blocks": [
        {"type": "fields", "title": "Headline", "fields": [
            _f("exec.headline.committed_bri", "Headline committed BRI (SAR)", "text"),
            _f("exec.headline.strapline", "Strapline", "text"),
        ]},
        {"type": "fields", "title": "Strategy & Ambition 2026 → 2030", "fields": [
            _f("exec.amb.gwp_2026", "GWP 2026", "text"), _f("exec.amb.gwp_2030", "GWP 2030", "text"),
            _f("exec.amb.profit_2026", "Net profit 2026", "text"), _f("exec.amb.profit_2030", "Net profit 2030", "text"),
            _f("exec.amb.roe_2026", "ROE 2026", "text"), _f("exec.amb.roe_2030", "ROE 2030", "text"),
            _f("exec.amb.initshare_2026", "Initiative share of GWP 2026", "text"),
            _f("exec.amb.initshare_2030", "Initiative share of GWP 2030", "text"),
        ]},
        {"type": "fields", "title": "Execution Health", "fields": [
            _f("exec.health.enterprise_ipi", "Enterprise IPI", "number"),
            _f("exec.health.strategic_projects", "Strategic project count", "number"),
            _f("exec.health.ms_total", "Strategic Milestones — total", "number"),
            _f("exec.health.ms_complete", "Milestones — complete", "number"),
            _f("exec.health.ms_not_yet_due", "Milestones — not-yet-due", "number"),
            _f("exec.health.ms_delayed", "Milestones — delayed", "number"),
        ]},
        {"type": "fields", "title": "Committed Benefit BRI", "fields": [
            _f("exec.bri.total", "Committed benefit total (SAR m)", "number"),
            _f("exec.bri.savings_total", "Savings total (SAR m)", "number"),
            _f("exec.bri.tier_on_track", "By tier — on-track (SAR m)", "number"),
            _f("exec.bri.tier_cautious", "By tier — cautious (SAR m)", "number"),
            _f("exec.bri.tier_at_risk", "By tier — at-risk (SAR m)", "number"),
        ]},
        {"type": "table", "title": "Benefit status by LoB", "id": "exec.benefit", "maxRows": 4, "columns": [
            _col("lob", "Line of business"),
            _col("pct", "% on track", "pct"),
            _col("sar", "SAR m", "number"),
        ]},
        {"type": "table", "title": "Savings BRI breakdown", "id": "exec.savings", "maxRows": 4, "columns": [
            _col("name", "Name"),
            _col("amount", "Amount (SAR m)", "number"),
            _col("note", "Note"),
        ]},
    ]})

    # 2. Context ------------------------------------------------------------
    sections.append({"id": "context", "label": "Context", "blocks": [
        {"type": "table", "title": "GWP by year 2026–2030 (SAR B)", "id": "ctx.gwp_year", "maxRows": 5, "columns": [
            _col("year", "Year"),
            _col("bau", "BAU base", "number"),
            _col("init", "Initiative layer", "number"),
            _col("total", "Total", "number", computed="bau+init"),
        ]},
        {"type": "table", "title": "2026 GWP by line of business (SAR m)", "id": "ctx.gwp_lob", "maxRows": 6, "columns": [
            _col("lob", "Line of business"),
            _col("bau", "BAU", "number"),
            _col("init", "Initiative", "number"),
            _col("total", "Total", "number", computed="bau+init"),
        ]},
    ]})

    # 3. Execution Signal (IPI) --------------------------------------------
    sections.append({"id": "ipi", "label": "Execution Signal (IPI)", "blocks": [
        {"type": "table", "title": "IPI by sector (0–5)", "id": "ipi.sector", "maxRows": 8, "columns": [
            _col("sector", "Sector"),
            _col("ipi", "IPI", "number"),
        ]},
        {"type": "fields", "title": "BRI by execution tier (SAR m)", "fields": [
            _f("ipi.tier.on_track", "On-track", "number"),
            _f("ipi.tier.cautious", "Cautious", "number"),
            _f("ipi.tier.at_risk", "At-risk", "number"),
        ]},
        {"type": "table", "title": "Benefit status by LoB", "id": "ipi.benefit", "maxRows": 6, "columns": [
            _col("lob", "Line of business"),
            _col("bri", "BRI (SAR m)", "number"),
            _col("pct", "% on track", "pct"),
            _col("note", "Note"),
        ]},
        {"type": "fields", "title": "Lever note", "fields": [
            _f("ipi.lever_note", "Lever note", "textarea"),
        ]},
    ]})

    # 4. LoB Scorecard ------------------------------------------------------
    sections.append({"id": "scorecard", "label": "LoB Scorecard", "blocks": [
        {"type": "table", "title": "Line-of-business scorecard", "id": "scorecard.row", "maxRows": 6, "columns": [
            _col("lob", "Line of business"),
            _col("note", "Note"),
            _col("ipi", "Strategic IPI", "number"),
            _col("bri_committed", "BRI committed (SAR m)", "number"),
            _col("bri_actual", "BRI actual (SAR m)", "number"),
            _col("status", "Status", "select", options=STATUS_6),
        ]},
    ]})

    # 5-8. Deep dives -------------------------------------------------------
    for lob_id, lob_label in LOBS:
        sections.append(_lob_deepdive(lob_id, lob_label))

    # 9. Customer Experience ------------------------------------------------
    sections.append({"id": "cx", "label": "Customer Experience", "blocks": [
        {"type": "fields", "title": "Slide heading", "fields": [
            _f("cx.head.title", "Headline", "text"),
            _f("cx.head.subtitle", "Subtitle", "textarea"),
        ]},
        {"type": "table", "title": "Projects", "id": "cx.proj", "maxRows": 12, "columns": [
            _col("name", "Project"), _col("ipi", "IPI", "number"), _col("ti", "TI", "number"),
        ]},
        {"type": "fields", "title": "Stakeholder View", "fields": [
            _f("cx.stake.strategy", "Strategy View", "textarea"),
            _f("cx.stake.epmo", "EPMO Feedback", "textarea"),
        ]},
    ]})

    # 10. Human Resources ---------------------------------------------------
    sections.append({"id": "hr", "label": "Human Resources", "blocks": [
        {"type": "fields", "title": "Operating-model redesign", "fields": [
            _f("hr.init.ipi", "IPI", "number"),
            _f("hr.init.ti", "TI", "number"),
            _f("hr.init.note", "Note", "textarea"),
        ]},
        {"type": "fields", "title": "Stakeholder View", "fields": [
            _f("hr.stake.strategy", "Strategy View", "textarea"),
            _f("hr.stake.epmo", "EPMO Feedback", "textarea"),
            _f("hr.stake.finance", "Finance Feedback", "textarea"),
            _f("hr.stake.hrlob", "HR / LoB Feedback", "textarea"),
        ]},
        {"type": "table", "title": "Project", "id": "hr.proj", "maxRows": 6, "columns": [
            _col("project", "Project"), _col("comment", "Comment"),
        ]},
    ]})

    # 11. EPMO Recovery Tracker --------------------------------------------
    sections.append({"id": "recovery", "label": "EPMO Recovery Tracker", "blocks": [
        {"type": "fields", "title": "Auto-flag", "fields": [
            _f("recovery.config.ti_threshold", "Flag initiatives with TI below", "number"),
        ]},
        {"type": "table", "title": "Recovery tracker", "id": "recovery.row", "maxRows": 20, "columns": [
            _col("rank", "#", "number"),
            _col("initiative", "Initiative"),
            _col("lob", "LoB", "select", options=[lbl for _, lbl in LOBS] + ["CX", "HR"]),
            _col("item_type", "Type", "select", options=TYPE_2),
            _col("bri", "BRI (SAR m)", "number"),
            _col("status", "Status", "select", options=WATCH_2),
            _col("ipi", "IPI", "number"),
            _col("ti", "TI", "number"),
            _col("root_cause", "Root cause"),
            _col("corrective_action", "Corrective action"),
            _col("owner", "Owner + check-in"),
        ]},
    ]})

    # 12. BRI Scenarios -----------------------------------------------------
    scen_blocks = []
    for sid, slabel in [("downside", "Downside"), ("base", "Base case"), ("upside", "Upside")]:
        scen_blocks.append({"type": "fields", "title": f"Scenario — {slabel}", "fields": [
            _f(f"scenario.{sid}.sublabel", "Sub-label", "text"),
            _f(f"scenario.{sid}.likelihood", "Likelihood %", "pct"),
            _f(f"scenario.{sid}.sar", "SAR realized value (m)", "number"),
            _f(f"scenario.{sid}.pct_committed", "% of committed BRI realized", "pct"),
            _f(f"scenario.{sid}.converts", "What converts", "textarea"),
            _f(f"scenario.{sid}.requires", "What it requires", "textarea"),
        ]})
    scen_blocks.append({"type": "fields", "title": "Summary line", "fields": [
        _f("scenario.summary.committed", "Committed (SAR m)", "number"),
        _f("scenario.summary.weighted_sar", "Probability-weighted (SAR m)", "number"),
        _f("scenario.summary.weighted_pct", "Probability-weighted %", "pct"),
        _f("scenario.summary.range", "Realisation range", "text"),
    ]})
    sections.append({"id": "scenarios", "label": "BRI Scenarios", "blocks": scen_blocks})

    # 13. Summary & Recommended Actions ------------------------------------
    sections.append({"id": "summary", "label": "Summary & Actions", "blocks": [
        {"type": "table", "title": "2026 Commitment", "id": "summary.commit", "maxRows": 6, "columns": [
            _col("item", "Item"),
            _col("value", "Value"),
            _col("descriptor", "Descriptor"),
        ]},
        {"type": "table", "title": "Five Priority Moves", "id": "summary.moves", "maxRows": 5, "columns": [
            _col("rank", "#", "number"),
            _col("move", "Move"),
            _col("description", "Description"),
            _col("status", "Status", "select", options=STATUS_6),
            _col("bri", "BRI (SAR m)", "number"),
        ]},
    ]})

    # 14. Dividers & Closing -----------------------------------------------
    sections.append({"id": "dividers", "label": "Dividers & Closing", "blocks": [
        {"type": "fields", "title": "Cover", "fields": [
            _f("div.cover.eyebrow", "Cover eyebrow", "text"),
            _f("div.cover.title", "Cover title", "text"),
            _f("div.cover.date", "Cover date", "text"),
            _f("div.cover.footer", "Cover footer", "text"),
        ]},
        {"type": "fields", "title": "Section dividers", "fields": [
            _f("div.sec.01", "01 — title", "text"),
            _f("div.sec.02", "02 — title", "text"),
            _f("div.sec.03", "03 — title", "text"),
            _f("div.sec.04", "04 — title", "text"),
            _f("div.sec.05", "05 — title", "text"),
        ]},
        {"type": "fields", "title": "Closing", "fields": [
            _f("div.closing.message", "Thank-you message", "text"),
            _f("div.closing.footer", "Closing footer", "text"),
        ]},
    ]})

    return sections


SECTIONS: List[Dict[str, Any]] = _build_sections()

# Default row seeds so the dashboard opens with the right labels in place
# (labels only — never confidential numbers).
DEFAULT_ROWS: Dict[str, List[Dict[str, Any]]] = {
    "exec.benefit": [{"lob": n} for n in ("Health", "General", "Life", "Motor")],
    "ctx.gwp_year": [{"year": str(y)} for y in range(2026, 2031)],
    "ctx.gwp_lob": [{"lob": n} for n in ("Health", "General Corp", "Motor", "Life", "General Retail")],
    "ipi.sector": [{"sector": n} for n in ("Mobility/Motor", "General", "Life", "Health", "CX", "HR")],
    "ipi.benefit": [{"lob": n} for n in ("Health", "General", "Life", "Motor")],
    "scorecard.row": [{"lob": n} for n in ("Health", "Motor", "General", "Life", "Customer Experience", "Human Resources")],
    "summary.commit": [{"item": n} for n in (
        "Committed growth BRI", "On track today", "Watch", "At Risk", "Savings BRI", "Govt assets")],
}


# ---- Derived lookups (used by store / excel_io / ui) ----

def iter_scalar_fields():
    for sec in SECTIONS:
        for block in sec["blocks"]:
            if block["type"] == "fields":
                for f in block["fields"]:
                    yield sec, block, f


def iter_tables():
    for sec in SECTIONS:
        for block in sec["blocks"]:
            if block["type"] == "table":
                yield sec, block


def all_scalar_keys() -> List[str]:
    return [f["key"] for _, _, f in iter_scalar_fields()]


def field_label_map() -> Dict[str, str]:
    """Every scalar key -> human label (for the flat Excel template)."""
    out: Dict[str, str] = {}
    for sec, block, f in iter_scalar_fields():
        out[f["key"]] = f"{sec['label']} · {f['label']}"
    return out


def table_by_id() -> Dict[str, Dict[str, Any]]:
    return {block["id"]: block for _, block in iter_tables()}


def section_index() -> Dict[str, Dict[str, Any]]:
    return {s["id"]: s for s in SECTIONS}
