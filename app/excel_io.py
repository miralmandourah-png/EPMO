"""Excel I/O — all local, in-memory.

Three responsibilities:
  1. Build a flat ``Key | Label | Value`` template of the whole data model,
     pre-filled with the current values, and parse a filled-in copy back.
  2. Parse the IPI Accountability export (targeted fields only).
  3. Parse the Strategic Projects Milestones export (targeted fields only).

Nothing is written to disk here; callers pass bytes and get bytes / dicts.
"""
from __future__ import annotations

import io
import re
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

from . import schema

HEADER_FILL = PatternFill(start_color="5B4FE0", end_color="5B4FE0", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
SECTION_FILL = PatternFill(start_color="ECEBFB", end_color="ECEBFB", fill_type="solid")
SECTION_FONT = Font(bold=True, color="5B4FE0")


# --------------------------------------------------------------------------
# Flat full-platform template
# --------------------------------------------------------------------------

def _rows_for_table(state_fields: Dict[str, Any], state_rows: Dict[str, int], table_id: str) -> int:
    return max(1, int(state_rows.get(table_id, 1)))


def build_template_workbook(state: Dict[str, Any]) -> Workbook:
    """One row per field (scalars + expanded table cells), grouped by section."""
    fields = state.get("fields", {})
    rows = state.get("rows", {})
    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"
    ws.append(["Key", "Label", "Value"])
    for c in range(1, 4):
        cell = ws.cell(1, c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    ws.freeze_panes = "A2"

    def section_header(text: str):
        r = ws.max_row + 1
        ws.append(["", text, ""])
        for c in range(1, 4):
            ws.cell(r, c).fill = SECTION_FILL
            ws.cell(r, c).font = SECTION_FONT

    for sec in schema.SECTIONS:
        section_header(sec["label"].upper())
        for block in sec["blocks"]:
            if block["type"] == "fields":
                for f in block["fields"]:
                    ws.append([f["key"], f"{block['title']} · {f['label']}", fields.get(f["key"], "")])
            else:  # table
                tid = block["id"]
                nrows = _rows_for_table(fields, rows, tid)
                for i in range(nrows):
                    for col in block["columns"]:
                        if col.get("computed"):
                            continue
                        key = f"{tid}.{i}.{col['name']}"
                        label = f"{block['title']} · row {i + 1} · {col['label']}"
                        ws.append([key, label, fields.get(key, "")])

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 30
    return wb


def workbook_to_bytes(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


_TABLE_KEY_RE = re.compile(r"^(?P<tid>.+)\.(?P<idx>\d+)\.(?P<col>[^.]+)$")


def parse_template_workbook(file_bytes: bytes) -> Tuple[Dict[str, Any], Dict[str, int]]:
    """Return ({key: value}, {tableId: rowCount}) for every keyed row present."""
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb["Dashboard"] if "Dashboard" in wb.sheetnames else wb.worksheets[0]
    valid_scalar = set(schema.all_scalar_keys())
    valid_tables = set(schema.table_by_id().keys())

    updates: Dict[str, Any] = {}
    max_idx: Dict[str, int] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        key = str(row[0]).strip()
        value = row[2] if len(row) > 2 else None
        if not key:
            continue
        if key in valid_scalar:
            updates[key] = value
            continue
        m = _TABLE_KEY_RE.match(key)
        if m and m.group("tid") in valid_tables:
            updates[key] = value
            idx = int(m.group("idx"))
            max_idx[m.group("tid")] = max(max_idx.get(m.group("tid"), -1), idx)
    row_counts = {tid: idx + 1 for tid, idx in max_idx.items()}
    return updates, row_counts


# --------------------------------------------------------------------------
# IPI Accountability importer
# --------------------------------------------------------------------------

IPI_REQUIRED = ["CP", "SBP", "SEP", "IPI", "Active", "Completed this year",
                "On Hold", "Not-Started", "Cancelled", "Closed"]


def _norm(v: Any) -> str:
    return str(v).strip() if v is not None else ""


def parse_ipi_accountability(file_bytes: bytes) -> Dict[str, Any]:
    """Locate the summary label row by exact string match, read the row below."""
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb["Data Table"] if "Data Table" in wb.sheetnames else wb.worksheets[0]

    required = set(IPI_REQUIRED)
    label_row = None
    col_of: Dict[str, int] = {}
    max_scan = min(ws.max_row, 25)
    for r in range(1, max_scan + 1):
        found: Dict[str, int] = {}
        for c in range(1, ws.max_column + 1):
            label = _norm(ws.cell(r, c).value)
            # exact-match; "IPI" must not pick up "IPI with error"
            if label in required and label not in found:
                found[label] = c
        if required.issubset(found.keys()):
            label_row = r
            col_of = found
            break

    if label_row is None:
        missing = ", ".join(sorted(required - set(col_of.keys()))) or "summary labels"
        raise ValueError(
            f"Could not find the IPI summary block (missing: {missing}). "
            "Expected a label row containing CP, SBP, SEP, IPI, Active, Completed this year, "
            "On Hold, Not-Started, Cancelled, Closed with values in the row directly below. "
            "Nothing was changed."
        )

    vrow = label_row + 1

    def num(label: str) -> Optional[float]:
        v = ws.cell(vrow, col_of[label]).value
        try:
            return float(v)
        except (TypeError, ValueError):
            return None

    ipi = num("IPI")
    cp = num("CP") or 0
    sbp = num("SBP") or 0
    sep = num("SEP") or 0
    if ipi is None:
        raise ValueError("Found the IPI label but its value cell is empty or non-numeric. Nothing was changed.")

    strategic = int(round(sbp + sep))
    updates = {
        "exec.health.enterprise_ipi": round(ipi, 2),
        "exec.health.strategic_projects": strategic,
    }
    status_labels = ["Active", "Completed this year", "On Hold", "Not-Started", "Cancelled", "Closed"]
    reference = {
        "total_projects": int(round(cp + sbp + sep)),
        "cp": int(round(cp)),
        "strategic": strategic,
        "sbp": int(round(sbp)),
        "sep": int(round(sep)),
        "status_breakdown": {lbl: (num(lbl) if num(lbl) is None else int(round(num(lbl)))) for lbl in status_labels},
        "enterprise_ipi": round(ipi, 2),
    }
    return {"updates": updates, "reference": reference}


# --------------------------------------------------------------------------
# Strategic Milestones importer
# --------------------------------------------------------------------------

def _find_header(ws, needed_lower: List[str], max_scan: int = 8) -> Optional[Tuple[int, Dict[str, int]]]:
    for r in range(1, min(ws.max_row, max_scan) + 1):
        cols: Dict[str, int] = {}
        for c in range(1, ws.max_column + 1):
            label = _norm(ws.cell(r, c).value).lower()
            if label:
                cols[label] = c
        if all(n in cols for n in needed_lower):
            return r, cols
    return None


def parse_milestones(file_bytes: bytes, grace_days: int = 0) -> Dict[str, Any]:
    """Classify each milestone row: complete / delayed / not-yet-due."""
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)

    target_ws = None
    header_row = None
    cols: Dict[str, int] = {}
    for ws in wb.worksheets:
        if ws.title.strip().lower() == "used filters":
            continue
        found = _find_header(ws, ["progress", "remaining days"])
        if found:
            target_ws, (header_row, cols) = ws, found
            break

    if target_ws is None:
        raise ValueError(
            "Could not find a sheet with both 'Progress' and 'Remaining Days' columns. "
            "Nothing was changed."
        )

    pcol = cols["progress"]
    rcol = cols["remaining days"]
    total = complete = delayed = not_yet = 0
    for r in range(header_row + 1, target_ws.max_row + 1):
        p = target_ws.cell(r, pcol).value
        rd = target_ws.cell(r, rcol).value
        if p is None and rd is None:
            continue
        total += 1
        try:
            p = float(p)
        except (TypeError, ValueError):
            p = 0.0
        try:
            rd = float(rd)
        except (TypeError, ValueError):
            rd = 0.0
        if p >= 99.999:
            complete += 1
        elif rd < -abs(grace_days):
            delayed += 1
        else:
            not_yet += 1

    updates = {
        "exec.health.ms_total": total,
        "exec.health.ms_complete": complete,
        "exec.health.ms_not_yet_due": not_yet,
        "exec.health.ms_delayed": delayed,
    }
    note = (
        f"complete = progress ≥ 100%; delayed = past due date"
        + (f" by more than {abs(grace_days)} day(s) grace" if grace_days else "")
        + " and not complete; not-yet-due = everything else. "
        "Confirm this matches your definition of delayed."
    )
    reference = {"total": total, "complete": complete, "delayed": delayed,
                 "not_yet_due": not_yet, "sheet": target_ws.title, "grace_days": grace_days}
    return {"updates": updates, "reference": reference, "note": note}
