"""FastAPI backend for the Strategic Health Check Platform.

Runs entirely locally. No third-party API, analytics, or telemetry. All
dashboard data lives in <repo>/data/state.json; uploaded Excel/PPTX bytes are
processed in memory for the duration of a request and never sent anywhere.
"""
from __future__ import annotations

import io
from pathlib import Path
from typing import Any, Dict

from fastapi import FastAPI, UploadFile, File, HTTPException, Body
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from . import schema, store, excel_io

app = FastAPI(title="Strategic Health Check Platform", docs_url=None, redoc_url=None)

STATIC_DIR = Path(__file__).parent / "static"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ---- schema + state ------------------------------------------------------

@app.get("/api/schema")
def get_schema():
    return {"sections": schema.SECTIONS}


@app.get("/api/state")
def get_state():
    return store.load_state()


@app.put("/api/state")
def put_state(payload: Dict[str, Any] = Body(...)):
    """Autosave: replace fields + row counts wholesale (frontend debounces)."""
    state = store.load_state()
    if "fields" in payload:
        state["fields"] = payload["fields"] or {}
    if "rows" in payload:
        state["rows"] = payload["rows"] or {}
    if "config" in payload and isinstance(payload["config"], dict):
        state.setdefault("config", {}).update(payload["config"])
    return store.save_state(state)


@app.post("/api/reset")
def reset():
    return store.reset_state()


# ---- full-platform Excel template ---------------------------------------

@app.get("/api/template")
def download_template():
    state = store.load_state()
    wb = excel_io.build_template_workbook(state)
    data = excel_io.workbook_to_bytes(wb)
    return StreamingResponse(
        io.BytesIO(data), media_type=XLSX_MIME,
        headers={"Content-Disposition": 'attachment; filename="strategic_health_check_template.xlsx"'},
    )


@app.post("/api/import/template")
async def import_template(file: UploadFile = File(...)):
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Please upload a .xlsx file.")
    raw = await file.read()
    try:
        updates, row_counts = excel_io.parse_template_workbook(raw)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(400, f"Could not read template: {exc}") from exc
    state = store.load_state()
    state.setdefault("fields", {}).update({k: v for k, v in updates.items()})
    # drop cleared cells
    state["fields"] = {k: v for k, v in state["fields"].items() if v not in (None, "")}
    for tid, count in row_counts.items():
        state.setdefault("rows", {})[tid] = max(state.get("rows", {}).get(tid, 0), count)
    store.save_state(state)
    return {"state": state, "applied": len(updates)}


# ---- smart importers -----------------------------------------------------

@app.post("/api/import/ipi")
async def import_ipi(file: UploadFile = File(...)):
    raw = await file.read()
    try:
        result = excel_io.parse_ipi_accountability(raw)
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from exc
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(400, f"Could not read IPI file: {exc}") from exc
    state = store.apply_field_updates(result["updates"])
    return {"state": state, "updates": result["updates"], "reference": result["reference"]}


@app.post("/api/import/milestones")
async def import_milestones(file: UploadFile = File(...), grace_days: int = 0):
    raw = await file.read()
    try:
        result = excel_io.parse_milestones(raw, grace_days=grace_days)
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from exc
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(400, f"Could not read milestones file: {exc}") from exc
    state = store.apply_field_updates(result["updates"])
    store.set_config({"delayed_grace_days": grace_days})
    return {"state": state, "updates": result["updates"],
            "reference": result["reference"], "note": result["note"]}


# ---- PPTX export (phase 2) ----------------------------------------------

@app.post("/api/export/pptx")
def export_pptx():
    raise HTTPException(
        501,
        "In-place branded PPTX export is being built in phase 2 (coordinate-mapped "
        "to the Tawuniya template). The dashboard, autosave, and all Excel ingestion "
        "are available now.",
    )


app.mount("/", StaticFiles(directory=str(STATIC_DIR), html=True), name="static")
