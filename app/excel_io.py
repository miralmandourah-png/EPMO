"""Reads/writes the Excel workbook that drives the deck.

The workbook has one "Meta" sheet (Section / Field / Value key-value rows)
for headlines, narratives and single numbers, plus one sheet per repeating
table (initiatives, projects, recovery tracker, etc). Nothing here ever
persists data to disk on the server --- callers pass bytes in and get
bytes/DeckData back.
"""
from __future__ import annotations

import io
from typing import Any, Dict, List, Optional

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

from .models import (
    DeckData, Cover, TocEntry, ExecSummary, IpiTier, LobStatus, SavingsItem,
    ContextGwp, GwpYear, GwpByLob, IpiSector, IpiSectorItem, LobOverview,
    LobOverviewRow, Initiative, Project, NpsSummary, NpsLob, NpsDetail,
    NpsDetailRow, RecoveryTracker, RecoveryItem, Scenarios, Scenario,
    RecommendedActions, RecommendedAction, Closing,
)

HEADER_FILL = PatternFill(start_color="302E52", end_color="302E52", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(bold=True, size=12, color="6B47F5")
STATUS_OPTIONS = '"on_track,cautious,at_risk"'
TYPE_OPTIONS = '"growth,savings"'

# Meta sheet layout: (section, field, description)
META_FIELDS = [
    ("Cover", "title", "Deck title, e.g. Strategic Health Check"),
    ("Cover", "subtitle", "Kicker line above the title"),
    ("Cover", "date_label", "e.g. June 2026"),
    ("Cover", "footer", "Footer shown on every slide, e.g. team / office name"),
    ("ExecSummary", "headline", "Executive summary slide headline"),
    ("ExecSummary", "subheadline", "Executive summary slide subheadline"),
    ("ExecSummary", "committed_label", "Label for the committed-benefit tile"),
    ("ExecSummary", "committed_bri_m", "Committed benefit, SAR millions"),
    ("ExecSummary", "savings_total_m", "Total savings benefit, SAR millions"),
    ("ContextGwp", "headline", "GWP trajectory slide headline"),
    ("ContextGwp", "subheadline", "GWP trajectory slide subheadline"),
    ("ContextGwp", "by_lob_year_label", "Label for the by-LoB panel, e.g. 2026 GWP by line of business"),
    ("IpiSector", "headline", "IPI-by-sector slide headline"),
    ("IpiSector", "subheadline", "IPI-by-sector slide subheadline"),
    ("IpiSector", "read_note", "Italic 'how to read this' note under the chart"),
    ("LobOverview", "headline", "LoB overview slide headline"),
    ("LobOverview", "narrative", "Narrative paragraph above the LoB table"),
    ("NpsSummary", "headline", "NPS summary slide headline"),
    ("NpsSummary", "subheadline", "NPS summary slide subheadline"),
    ("NpsSummary", "company_actual", "Companywide NPS, actual"),
    ("NpsSummary", "company_target", "Companywide NPS, target"),
    ("NpsDetail", "headline", "NPS-by-segment slide headline"),
    ("NpsDetail", "narrative", "Narrative under the NPS-by-segment headline"),
    ("RecoveryTracker", "headline", "Recovery tracker slide headline"),
    ("RecoveryTracker", "subheadline", "Recovery tracker slide subheadline"),
    ("Scenarios", "headline", "Scenarios slide headline"),
    ("Scenarios", "subheadline", "Scenarios slide subheadline"),
    ("Scenarios", "committed_m", "Committed BRI figure shown above the scenario cards"),
    ("RecommendedActions", "headline", "Recommended actions slide headline"),
    ("RecommendedActions", "subheadline", "Recommended actions slide subheadline"),
    ("Closing", "message", "Closing slide message, e.g. Thank you"),
    ("Closing", "footer", "Closing slide footer (defaults to Cover footer if blank)"),
]

TABLE_SHEETS: Dict[str, List[str]] = {
    "TOC": ["number", "section", "description", "page"],
    "IPI_Tiers": ["label", "bri_m", "color_key"],
    "LoB_Status": ["lob", "percent_on_track", "bri_m"],
    "Savings_Breakdown": ["category", "amount_m", "detail"],
    "GWP_Years": ["year", "bau_m", "initiative_m"],
    "GWP_By_LoB": ["lob", "gwp_m"],
    "IPI_By_Sector": ["lob", "ipi"],
    "LoB_Overview": ["lob", "status_summary", "bri_m", "ipi"],
    "Initiatives": ["lob", "initiative", "subtitle", "ipi", "ti", "bri_committed_m", "bri_actual_m", "status", "comments"],
    "Projects": ["lob", "initiative", "project", "ipi", "ti", "comments"],
    "NPS_By_LoB": ["lob", "actual", "target"],
    "NPS_Detail": ["lob", "segment", "actual", "target"],
    "Recovery_Tracker": ["rank", "initiative", "lob", "item_type", "bri_m", "status", "ipi", "ti", "root_cause", "corrective_action", "owner"],
    "Scenarios": ["name", "description", "value_m", "percent"],
    "Recommended_Actions": ["title", "description"],
}

TABLE_HEADER_LABELS = {
    "lob": "LoB", "bri_m": "BRI (SAR m)", "color_key": "Color Key (on_track/cautious/at_risk)",
    "percent_on_track": "Percent On Track", "category": "Category", "amount_m": "Amount (SAR m)",
    "detail": "Detail", "year": "Year", "bau_m": "BAU (SAR m)", "initiative_m": "Initiative (SAR m)",
    "gwp_m": "GWP (SAR m)", "ipi": "IPI", "status_summary": "Status Summary",
    "initiative": "Initiative", "subtitle": "Subtitle", "ti": "TI",
    "bri_committed_m": "BRI Committed (SAR m)", "bri_actual_m": "BRI Actual (SAR m)",
    "status": "Status (on_track/cautious/at_risk)", "comments": "Comments", "project": "Project",
    "actual": "Actual", "target": "Target", "segment": "Segment", "rank": "Rank",
    "item_type": "Type (growth/savings)", "root_cause": "Root Cause",
    "corrective_action": "Corrective Action", "owner": "Owner", "name": "Name",
    "description": "Description", "value_m": "Value (SAR m)", "percent": "Percent",
    "title": "Title", "number": "Number", "section": "Section", "page": "Page",
}


def _style_header_row(ws, ncols: int):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"


def build_template_workbook(sample: bool = False) -> Workbook:
    """Build a blank (or, if sample=True, fictitiously-populated) template workbook."""
    wb = Workbook()
    meta_ws = wb.active
    meta_ws.title = "Meta"
    meta_ws.append(["Section", "Field", "Value", "Guidance"])
    _style_header_row(meta_ws, 4)
    sample_meta = _sample_meta() if sample else {}
    for section, field, guidance in META_FIELDS:
        value = sample_meta.get((section, field), "")
        meta_ws.append([section, field, value, guidance])
    meta_ws.column_dimensions["A"].width = 20
    meta_ws.column_dimensions["B"].width = 20
    meta_ws.column_dimensions["C"].width = 40
    meta_ws.column_dimensions["D"].width = 55
    for row in meta_ws.iter_rows(min_row=2, max_col=1):
        row[0].font = Font(bold=True)

    sample_tables = _sample_tables() if sample else {}
    for sheet_name, cols in TABLE_SHEETS.items():
        ws = wb.create_sheet(sheet_name)
        ws.append([TABLE_HEADER_LABELS.get(c, c) for c in cols])
        _style_header_row(ws, len(cols))
        for i, col in enumerate(cols):
            ws.column_dimensions[get_column_letter(i + 1)].width = 22
        for row in sample_tables.get(sheet_name, []):
            ws.append(row)
        if "status" in cols:
            dv = DataValidation(type="list", formula1=STATUS_OPTIONS, allow_blank=True)
            ws.add_data_validation(dv)
            col_idx = cols.index("status") + 1
            dv.add(f"{get_column_letter(col_idx)}2:{get_column_letter(col_idx)}500")
        if "item_type" in cols:
            dv = DataValidation(type="list", formula1=TYPE_OPTIONS, allow_blank=True)
            ws.add_data_validation(dv)
            col_idx = cols.index("item_type") + 1
            dv.add(f"{get_column_letter(col_idx)}2:{get_column_letter(col_idx)}500")

    wb.active = 0
    return wb


def workbook_to_bytes(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _num(v: Any, default: float = 0.0) -> float:
    if v is None or v == "":
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _str(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _rows(ws) -> List[Dict[str, Any]]:
    header = [c.value for c in ws[1]]
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None or v == "" for v in row):
            continue
        out.append(dict(zip(header, row)))
    return out


def parse_workbook(file_bytes: bytes) -> DeckData:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    meta: Dict[tuple, str] = {}
    if "Meta" in wb.sheetnames:
        ws = wb["Meta"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            section, field, value = row[0], row[1], row[2]
            meta[(_str(section), _str(field))] = value

    def m(section, field, default=""):
        return meta.get((section, field), default)

    tables: Dict[str, List[Dict[str, Any]]] = {}
    for sheet_name, cols in TABLE_SHEETS.items():
        if sheet_name in wb.sheetnames:
            raw_rows = _rows(wb[sheet_name])
            keyed = []
            for r in raw_rows:
                values = list(r.values())
                keyed.append(dict(zip(cols, values)))
            tables[sheet_name] = keyed
        else:
            tables[sheet_name] = []

    data = DeckData(
        cover=Cover(
            title=_str(m("Cover", "title")),
            subtitle=_str(m("Cover", "subtitle")),
            date_label=_str(m("Cover", "date_label")),
            footer=_str(m("Cover", "footer")),
        ),
        toc=[TocEntry(number=_str(r.get("number")), section=_str(r.get("section")),
                      description=_str(r.get("description")), page=_str(r.get("page")))
             for r in tables["TOC"]],
        exec_summary=ExecSummary(
            headline=_str(m("ExecSummary", "headline")),
            subheadline=_str(m("ExecSummary", "subheadline")),
            committed_label=_str(m("ExecSummary", "committed_label")),
            committed_bri_m=_num(m("ExecSummary", "committed_bri_m")),
            savings_total_m=_num(m("ExecSummary", "savings_total_m")),
            ipi_tiers=[IpiTier(label=_str(r.get("label")), bri_m=_num(r.get("bri_m")),
                                color_key=_str(r.get("color_key")) or "on_track")
                       for r in tables["IPI_Tiers"]],
            benefit_status_by_lob=[LobStatus(lob=_str(r.get("lob")), percent_on_track=_num(r.get("percent_on_track")),
                                              bri_m=_num(r.get("bri_m")))
                                    for r in tables["LoB_Status"]],
            savings_breakdown=[SavingsItem(category=_str(r.get("category")), amount_m=_num(r.get("amount_m")),
                                            detail=_str(r.get("detail")))
                               for r in tables["Savings_Breakdown"]],
        ),
        context_gwp=ContextGwp(
            headline=_str(m("ContextGwp", "headline")),
            subheadline=_str(m("ContextGwp", "subheadline")),
            by_lob_year_label=_str(m("ContextGwp", "by_lob_year_label")),
            years=[GwpYear(year=_str(r.get("year")), bau_m=_num(r.get("bau_m")), initiative_m=_num(r.get("initiative_m")))
                   for r in tables["GWP_Years"]],
            by_lob=[GwpByLob(lob=_str(r.get("lob")), gwp_m=_num(r.get("gwp_m"))) for r in tables["GWP_By_LoB"]],
        ),
        ipi_sector=IpiSector(
            headline=_str(m("IpiSector", "headline")),
            subheadline=_str(m("IpiSector", "subheadline")),
            read_note=_str(m("IpiSector", "read_note")),
            items=[IpiSectorItem(lob=_str(r.get("lob")), ipi=_num(r.get("ipi"))) for r in tables["IPI_By_Sector"]],
        ),
        lob_overview=LobOverview(
            headline=_str(m("LobOverview", "headline")),
            narrative=_str(m("LobOverview", "narrative")),
            rows=[LobOverviewRow(lob=_str(r.get("lob")), status_summary=_str(r.get("status_summary")),
                                  bri_m=_num(r.get("bri_m")), ipi=_num(r.get("ipi")))
                  for r in tables["LoB_Overview"]],
        ),
        initiatives=[Initiative(lob=_str(r.get("lob")), initiative=_str(r.get("initiative")),
                                 subtitle=_str(r.get("subtitle")), ipi=_num(r.get("ipi")), ti=_num(r.get("ti")),
                                 bri_committed_m=_num(r.get("bri_committed_m")), bri_actual_m=_num(r.get("bri_actual_m")),
                                 status=_str(r.get("status")) or "on_track", comments=_str(r.get("comments")))
                     for r in tables["Initiatives"]],
        projects=[Project(lob=_str(r.get("lob")), initiative=_str(r.get("initiative")), project=_str(r.get("project")),
                           ipi=_num(r.get("ipi")), ti=_num(r.get("ti")), comments=_str(r.get("comments")))
                  for r in tables["Projects"]],
        nps_summary=NpsSummary(
            headline=_str(m("NpsSummary", "headline")),
            subheadline=_str(m("NpsSummary", "subheadline")),
            company_actual=_num(m("NpsSummary", "company_actual")),
            company_target=_num(m("NpsSummary", "company_target")),
            by_lob=[NpsLob(lob=_str(r.get("lob")), actual=_num(r.get("actual")), target=_num(r.get("target")))
                    for r in tables["NPS_By_LoB"]],
        ),
        nps_detail=NpsDetail(
            headline=_str(m("NpsDetail", "headline")),
            narrative=_str(m("NpsDetail", "narrative")),
            rows=[NpsDetailRow(lob=_str(r.get("lob")), segment=_str(r.get("segment")),
                                actual=_num(r.get("actual")), target=_num(r.get("target")))
                  for r in tables["NPS_Detail"]],
        ),
        recovery_tracker=RecoveryTracker(
            headline=_str(m("RecoveryTracker", "headline")),
            subheadline=_str(m("RecoveryTracker", "subheadline")),
            items=[RecoveryItem(rank=int(_num(r.get("rank"))), initiative=_str(r.get("initiative")),
                                 lob=_str(r.get("lob")), item_type=_str(r.get("item_type")),
                                 bri_m=_num(r.get("bri_m")), status=_str(r.get("status")) or "cautious",
                                 ipi=_num(r.get("ipi")), ti=_num(r.get("ti")), root_cause=_str(r.get("root_cause")),
                                 corrective_action=_str(r.get("corrective_action")), owner=_str(r.get("owner")))
                   for r in tables["Recovery_Tracker"]],
        ),
        scenarios=Scenarios(
            headline=_str(m("Scenarios", "headline")),
            subheadline=_str(m("Scenarios", "subheadline")),
            committed_m=_num(m("Scenarios", "committed_m")),
            items=[Scenario(name=_str(r.get("name")), description=_str(r.get("description")),
                             value_m=_num(r.get("value_m")), percent=_num(r.get("percent")))
                   for r in tables["Scenarios"]],
        ),
        recommended_actions=RecommendedActions(
            headline=_str(m("RecommendedActions", "headline")),
            subheadline=_str(m("RecommendedActions", "subheadline")),
            items=[RecommendedAction(title=_str(r.get("title")), description=_str(r.get("description")))
                   for r in tables["Recommended_Actions"]],
        ),
        closing=Closing(
            message=_str(m("Closing", "message")) or "Thank you",
            footer=_str(m("Closing", "footer")),
        ),
    )
    return data


def deck_to_workbook(data: DeckData) -> Workbook:
    """Serialize a DeckData back into the same workbook shape (round-trip for editing)."""
    wb = build_template_workbook(sample=False)
    meta_ws = wb["Meta"]
    values = {
        ("Cover", "title"): data.cover.title, ("Cover", "subtitle"): data.cover.subtitle,
        ("Cover", "date_label"): data.cover.date_label, ("Cover", "footer"): data.cover.footer,
        ("ExecSummary", "headline"): data.exec_summary.headline,
        ("ExecSummary", "subheadline"): data.exec_summary.subheadline,
        ("ExecSummary", "committed_label"): data.exec_summary.committed_label,
        ("ExecSummary", "committed_bri_m"): data.exec_summary.committed_bri_m,
        ("ExecSummary", "savings_total_m"): data.exec_summary.savings_total_m,
        ("ContextGwp", "headline"): data.context_gwp.headline,
        ("ContextGwp", "subheadline"): data.context_gwp.subheadline,
        ("ContextGwp", "by_lob_year_label"): data.context_gwp.by_lob_year_label,
        ("IpiSector", "headline"): data.ipi_sector.headline,
        ("IpiSector", "subheadline"): data.ipi_sector.subheadline,
        ("IpiSector", "read_note"): data.ipi_sector.read_note,
        ("LobOverview", "headline"): data.lob_overview.headline,
        ("LobOverview", "narrative"): data.lob_overview.narrative,
        ("NpsSummary", "headline"): data.nps_summary.headline,
        ("NpsSummary", "subheadline"): data.nps_summary.subheadline,
        ("NpsSummary", "company_actual"): data.nps_summary.company_actual,
        ("NpsSummary", "company_target"): data.nps_summary.company_target,
        ("NpsDetail", "headline"): data.nps_detail.headline,
        ("NpsDetail", "narrative"): data.nps_detail.narrative,
        ("RecoveryTracker", "headline"): data.recovery_tracker.headline,
        ("RecoveryTracker", "subheadline"): data.recovery_tracker.subheadline,
        ("Scenarios", "headline"): data.scenarios.headline,
        ("Scenarios", "subheadline"): data.scenarios.subheadline,
        ("Scenarios", "committed_m"): data.scenarios.committed_m,
        ("RecommendedActions", "headline"): data.recommended_actions.headline,
        ("RecommendedActions", "subheadline"): data.recommended_actions.subheadline,
        ("Closing", "message"): data.closing.message,
        ("Closing", "footer"): data.closing.footer,
    }
    for row in meta_ws.iter_rows(min_row=2):
        key = (_str(row[0].value), _str(row[1].value))
        if key in values:
            row[2].value = values[key]

    def clear_and_fill(sheet_name, cols, rows):
        ws = wb[sheet_name]
        ws.delete_rows(2, ws.max_row)
        for r in rows:
            ws.append([r.get(c) for c in cols])

    clear_and_fill("TOC", TABLE_SHEETS["TOC"], [t.model_dump() for t in data.toc])
    clear_and_fill("IPI_Tiers", TABLE_SHEETS["IPI_Tiers"], [t.model_dump() for t in data.exec_summary.ipi_tiers])
    clear_and_fill("LoB_Status", TABLE_SHEETS["LoB_Status"], [t.model_dump() for t in data.exec_summary.benefit_status_by_lob])
    clear_and_fill("Savings_Breakdown", TABLE_SHEETS["Savings_Breakdown"], [t.model_dump() for t in data.exec_summary.savings_breakdown])
    clear_and_fill("GWP_Years", TABLE_SHEETS["GWP_Years"], [t.model_dump() for t in data.context_gwp.years])
    clear_and_fill("GWP_By_LoB", TABLE_SHEETS["GWP_By_LoB"], [t.model_dump() for t in data.context_gwp.by_lob])
    clear_and_fill("IPI_By_Sector", TABLE_SHEETS["IPI_By_Sector"], [t.model_dump() for t in data.ipi_sector.items])
    clear_and_fill("LoB_Overview", TABLE_SHEETS["LoB_Overview"], [t.model_dump() for t in data.lob_overview.rows])
    clear_and_fill("Initiatives", TABLE_SHEETS["Initiatives"], [t.model_dump() for t in data.initiatives])
    clear_and_fill("Projects", TABLE_SHEETS["Projects"], [t.model_dump() for t in data.projects])
    clear_and_fill("NPS_By_LoB", TABLE_SHEETS["NPS_By_LoB"], [t.model_dump() for t in data.nps_summary.by_lob])
    clear_and_fill("NPS_Detail", TABLE_SHEETS["NPS_Detail"], [t.model_dump() for t in data.nps_detail.rows])
    clear_and_fill("Recovery_Tracker", TABLE_SHEETS["Recovery_Tracker"], [t.model_dump() for t in data.recovery_tracker.items])
    clear_and_fill("Scenarios", TABLE_SHEETS["Scenarios"], [t.model_dump() for t in data.scenarios.items])
    clear_and_fill("Recommended_Actions", TABLE_SHEETS["Recommended_Actions"], [t.model_dump() for t in data.recommended_actions.items])
    return wb


def _sample_meta() -> Dict[tuple, Any]:
    return {
        ("Cover", "title"): "Strategic Health Check",
        ("Cover", "subtitle"): "Management Update",
        ("Cover", "date_label"): "Sample Quarter",
        ("Cover", "footer"): "Strategy & Transformation Office",
        ("ExecSummary", "headline"): "Executive summary — sample data for illustration only",
        ("ExecSummary", "subheadline"): "Strategy, committed benefit and execution at a glance.",
        ("ExecSummary", "committed_label"): "Committed growth benefit",
        ("ExecSummary", "committed_bri_m"): 1000,
        ("ExecSummary", "savings_total_m"): 60,
        ("ContextGwp", "headline"): "Context — sample growth trajectory",
        ("ContextGwp", "subheadline"): "Illustrative GWP trajectory for demo purposes.",
        ("ContextGwp", "by_lob_year_label"): "Sample year GWP by line of business",
        ("IpiSector", "headline"): "The execution signal — IPI by sector (sample)",
        ("IpiSector", "subheadline"): "Illustrative execution index by line of business, 0-5 scale.",
        ("IpiSector", "read_note"): "Read: bars at or above the dashed line are on-track.",
        ("LobOverview", "headline"): "Where benefit concentrates by line of business (sample)",
        ("LobOverview", "narrative"): "Sample narrative describing how benefit concentrates across lines of business.",
        ("NpsSummary", "headline"): "Customer experience — NPS (sample)",
        ("NpsSummary", "subheadline"): "Illustrative companywide and by-LoB NPS.",
        ("NpsSummary", "company_actual"): 45,
        ("NpsSummary", "company_target"): 50,
        ("NpsDetail", "headline"): "NPS by segment (sample)",
        ("NpsDetail", "narrative"): "Illustrative segment-level detail behind the sector NPS.",
        ("RecoveryTracker", "headline"): "Recovery tracker (sample)",
        ("RecoveryTracker", "subheadline"): "Illustrative at-risk items ranked by size.",
        ("Scenarios", "headline"): "Probability of achieving the committed benefit (sample)",
        ("Scenarios", "subheadline"): "Illustrative scenario view.",
        ("Scenarios", "committed_m"): 1000,
        ("RecommendedActions", "headline"): "Summary & recommended actions (sample)",
        ("RecommendedActions", "subheadline"): "Illustrative priority moves.",
        ("Closing", "message"): "Thank you",
        ("Closing", "footer"): "Strategy & Transformation Office",
    }


def _sample_tables() -> Dict[str, List[List[Any]]]:
    return {
        "TOC": [
            ["01", "Executive Summary", "Summary of the presentation", "3"],
            ["02", "Line-of-Business Deep-Dives", "Initiatives and projects", "5"],
        ],
        "IPI_Tiers": [
            ["On-track", 400, "on_track"],
            ["Cautious", 350, "cautious"],
            ["At-risk", 250, "at_risk"],
        ],
        "LoB_Status": [
            ["Line A", 40, 500],
            ["Line B", 20, 300],
            ["Line C", 10, 200],
        ],
        "Savings_Breakdown": [
            ["Efficiency programme", 35, "Sample detail"],
            ["Retention", 25, "Sample detail"],
        ],
        "GWP_Years": [
            ["Year 1", 900, 100],
            ["Year 2", 1000, 150],
            ["Year 3", 1050, 220],
        ],
        "GWP_By_LoB": [
            ["Line A", 500],
            ["Line B", 350],
            ["Line C", 270],
        ],
        "IPI_By_Sector": [
            ["Line A", 3.3],
            ["Line B", 2.8],
            ["Line C", 2.6],
        ],
        "LoB_Overview": [
            ["Line A", "On-track core", 500, 3.3],
            ["Line B", "Below target", 350, 2.8],
        ],
        "Initiatives": [
            ["Line A", "Sample initiative 1", "1 project", 3.1, 2.8, 200, 120, "on_track", "Sample comment"],
            ["Line A", "Sample initiative 2", "2 projects", 2.4, 2.1, 150, 40, "at_risk", "Sample comment"],
        ],
        "Projects": [
            ["Line A", "Sample initiative 1", "Sample project A", 3.1, 2.8, "Sample comment"],
        ],
        "NPS_By_LoB": [
            ["Line A", 47, 50],
            ["Line B", 52, 50],
        ],
        "NPS_Detail": [
            ["Line A", "Segment 1", 40, 50],
            ["Line A", "Segment 2", 55, 50],
        ],
        "Recovery_Tracker": [
            [1, "Sample initiative 2", "Line A", "growth", 150, "at_risk", 2.4, 2.1, "Sample root cause",
             "Sample corrective action", "Sample owner"],
        ],
        "Scenarios": [
            ["Conservative", "Illustrative low case", 600, 60],
            ["Base", "Illustrative base case", 800, 80],
            ["Upside", "Illustrative upside case", 950, 95],
        ],
        "Recommended_Actions": [
            ["Sample priority move 1", "Illustrative description of the action."],
            ["Sample priority move 2", "Illustrative description of the action."],
        ],
    }
