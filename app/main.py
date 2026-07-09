"""FastAPI app: upload Excel -> edit in browser -> export PPTX.

Everything is handled in-memory per-request. No uploaded workbook, parsed
data, or generated deck is written to disk or logged on the server; each
request's bytes exist only for the duration of that request.
"""
from __future__ import annotations

import io
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from .models import DeckData
from .excel_io import build_template_workbook, workbook_to_bytes, parse_workbook, deck_to_workbook
from .ppt_builder import build_deck

app = FastAPI(title="Deck Builder", docs_url=None, redoc_url=None)

STATIC_DIR = Path(__file__).parent / "static"


@app.get("/api/template")
def download_template(sample: bool = False):
    wb = build_template_workbook(sample=sample)
    data = workbook_to_bytes(wb)
    name = "sample_template.xlsx" if sample else "blank_template.xlsx"
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{name}"'},
    )


@app.post("/api/upload")
async def upload_excel(file: UploadFile = File(...)):
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Please upload a .xlsx file.")
    raw = await file.read()
    try:
        data = parse_workbook(raw)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(400, f"Could not read workbook: {exc}") from exc
    return JSONResponse(data.model_dump())


@app.post("/api/export/pptx")
async def export_pptx(data: DeckData):
    prs = build_deck(data)
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    filename = (data.cover.title or "deck").strip().replace(" ", "_") or "deck"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{filename}.pptx"'},
    )


@app.post("/api/export/xlsx")
async def export_xlsx(data: DeckData):
    wb = deck_to_workbook(data)
    buf = io.BytesIO(workbook_to_bytes(wb))
    filename = (data.cover.title or "deck").strip().replace(" ", "_") or "deck"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}.xlsx"'},
    )


app.mount("/", StaticFiles(directory=str(STATIC_DIR), html=True), name="static")
