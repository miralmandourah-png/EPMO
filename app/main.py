"""FastAPI backend for the Strategic Health Check Platform.

Runs entirely locally. No third-party API, analytics, or telemetry. All
dashboard data lives in <repo>/data/state.json; uploaded Excel/PPTX bytes are
processed in memory for the duration of a request and never sent anywhere.
"""
from __future__ import annotations

import io
from pathlib import Path
from typing import Any, Dict, Optional

from fastapi import FastAPI, UploadFile, File, HTTPException, Body
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from . import schema, store, excel_io, pptx_export, automation, pptx_constants

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


# ---- extract committed/target constants from the uploaded template ------

@app.post("/api/import/constants")
def import_constants():
    raw = store.load_template()
    if raw is None:
        raise HTTPException(400, "Upload a PPTX template first (Data ▾ → Upload PPTX template).")
    try:
        result = pptx_constants.extract_constants(raw)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(400, f"Could not read targets from template: {exc}") from exc

    state = store.load_state()
    fields = state.setdefault("fields", {})
    rows = state.setdefault("rows", {})

    for k, v in result["updates"].items():
        if v is not None:
            fields[k] = v
    for table_id, table_rows in result["tables"].items():
        cols = schema.table_by_id()[table_id]["columns"]
        col_names = [c["name"] for c in cols if not c.get("computed")]
        for i, row in enumerate(table_rows):
            for name in col_names:
                v = row.get(name)
                if v is not None:
                    fields[f"{table_id}.{i}.{name}"] = v
        rows[table_id] = max(rows.get(table_id, 0), len(table_rows))

    store.save_state(state)
    return {"state": state, "updates": result["updates"], "tables": result["tables"]}


# ---- recovery tracker auto-sync ------------------------------------------

@app.post("/api/recovery/sync")
def sync_recovery(threshold: Optional[float] = None):
    state = store.load_state()
    summary = automation.sync_recovery_tracker(state, threshold=threshold)
    store.save_state(state)
    return {"state": state, "summary": summary}


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


# ---- PPTX template + in-place export ------------------------------------

@app.post("/api/template/pptx")
async def upload_pptx_template(file: UploadFile = File(...)):
    if not file.filename.lower().endswith((".pptx",)):
        raise HTTPException(400, "Please upload a .pptx file.")
    raw = await file.read()
    try:
        info = pptx_export.template_info(raw)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(400, f"Could not read template: {exc}") from exc
    store.save_template(raw)
    store.set_config({"template_name": file.filename})
    return {"configured": True, "name": file.filename, **info}


@app.get("/api/template/pptx/info")
def pptx_template_info():
    raw = store.load_template()
    if raw is None:
        return {"configured": False}
    name = store.load_state().get("config", {}).get("template_name")
    return {"configured": True, "name": name, **pptx_export.template_info(raw)}


@app.post("/api/export/pptx")
def export_pptx():
    raw = store.load_template()
    if raw is None:
        raise HTTPException(400, "No PowerPoint template configured yet. Upload your "
                                 "branded .pptx template first (Data ▾ → Upload PPTX template).")
    state = store.load_state()
    try:
        out = pptx_export.export_pptx(raw, state)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(500, f"Export failed: {exc}") from exc
    return StreamingResponse(
        io.BytesIO(out),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": 'attachment; filename="strategic_health_check.pptx"'},
    )


app.mount("/", StaticFiles(directory=str(STATIC_DIR), html=True), name="static")
