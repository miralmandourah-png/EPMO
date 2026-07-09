"""Local, on-disk persistence for the dashboard.

The entire dashboard is one JSON document at <repo>/data/state.json. There
is no database server and nothing leaves the machine. The document is a flat
key/value store so it round-trips cleanly through the Excel template and the
(phase 2) PPTX coordinate map.

Shape:
    {
      "fields": { "<dotted.key>": <value>, ... },   # scalars + table cells
      "rows":   { "<tableId>": <rowCount>, ... },    # how many rows each table has
      "config": { "delayed_grace_days": 0,
                   "template_path": null,
                   "slide_mapping": {} },
      "meta":   { "updated_at": "<iso8601>" }
    }
"""
from __future__ import annotations

import json
import os
import tempfile
import threading
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict

from . import schema

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
STATE_PATH = DATA_DIR / "state.json"
TEMPLATE_PATH = DATA_DIR / "template.pptx"

_lock = threading.Lock()


def save_template(data: bytes) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    TEMPLATE_PATH.write_bytes(data)


def load_template() -> "bytes | None":
    if TEMPLATE_PATH.exists():
        return TEMPLATE_PATH.read_bytes()
    return None


def has_template() -> bool:
    return TEMPLATE_PATH.exists()


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def default_state() -> Dict[str, Any]:
    fields: Dict[str, Any] = {}
    rows: Dict[str, int] = {}
    # Seed default row labels (labels only, never numbers).
    for table_id, seed_rows in schema.DEFAULT_ROWS.items():
        rows[table_id] = len(seed_rows)
        for i, row in enumerate(seed_rows):
            for col, val in row.items():
                fields[f"{table_id}.{i}.{col}"] = val
    # Any table without a seed starts with one empty row.
    for _sec, block in schema.iter_tables():
        rows.setdefault(block["id"], 1)
    return {
        "fields": fields,
        "rows": rows,
        "config": {
            "delayed_grace_days": 0,
            "template_path": None,
            "slide_mapping": {},
        },
        "meta": {"updated_at": _now()},
    }


def _atomic_write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(text)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def load_state() -> Dict[str, Any]:
    with _lock:
        if not STATE_PATH.exists():
            state = default_state()
            _atomic_write(STATE_PATH, json.dumps(state, ensure_ascii=False, indent=2))
            return state
        try:
            return json.loads(STATE_PATH.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            # Corrupt file -> start clean rather than crash.
            state = default_state()
            _atomic_write(STATE_PATH, json.dumps(state, ensure_ascii=False, indent=2))
            return state


def save_state(state: Dict[str, Any]) -> Dict[str, Any]:
    with _lock:
        state.setdefault("meta", {})["updated_at"] = _now()
        _atomic_write(STATE_PATH, json.dumps(state, ensure_ascii=False, indent=2))
        return state


def reset_state() -> Dict[str, Any]:
    return save_state(default_state())


def apply_field_updates(updates: Dict[str, Any]) -> Dict[str, Any]:
    """Merge a partial {key: value} patch into the live state and persist."""
    state = load_state()
    fields = state.setdefault("fields", {})
    for key, val in updates.items():
        if val is None or val == "":
            fields.pop(key, None)
        else:
            fields[key] = val
    return save_state(state)


def set_config(patch: Dict[str, Any]) -> Dict[str, Any]:
    state = load_state()
    state.setdefault("config", {}).update(patch)
    return save_state(state)
