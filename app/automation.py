"""Recovery-tracker auto-flagging.

Scans every line-of-business initiative table for TI below a threshold (the
deck's own convention: "TI (vs 3.0)") and upserts a matching row into the
EPMO Recovery Tracker -- root cause / corrective action / owner are left
blank so the portfolio manager can chase them down with the responsible
department. Existing rows are matched by (lob, initiative) and updated in
place (BRI/IPI/TI refreshed) without touching anything the user already
typed into root cause / corrective action / owner / status. Nothing is ever
auto-removed -- if an item recovers above the threshold, its row is left for
the user to delete once resolved.

This runs only when explicitly requested (a button, not on every autosave)
so a user who deletes a flagged row doesn't see it silently reappear on the
next keystroke.
"""
from __future__ import annotations

from typing import Any, Dict, List, Tuple

from . import schema

DEFAULT_TI_THRESHOLD = 3.0


def _num(v: Any) -> "float | None":
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _lob_initiative_rows(fields: Dict[str, Any], rows: Dict[str, int]) -> List[Dict[str, Any]]:
    out = []
    for lob_id, lob_label in schema.LOBS:
        table_id = f"{lob_id}.init"
        n = int(rows.get(table_id, 0) or 0)
        for i in range(n):
            prefix = f"{table_id}.{i}"
            name = fields.get(f"{prefix}.name")
            if not name:
                continue
            out.append({
                "lob": lob_label,
                "initiative": str(name),
                "item_type": fields.get(f"{prefix}.item_type") or "growth",
                "ipi": _num(fields.get(f"{prefix}.ipi")),
                "ti": _num(fields.get(f"{prefix}.ti")),
                "bri": _num(fields.get(f"{prefix}.bri_committed")),
            })
    return out


def _existing_recovery_keys(fields: Dict[str, Any], rows: Dict[str, int]) -> Dict[Tuple[str, str], int]:
    n = int(rows.get("recovery.row", 0) or 0)
    keys = {}
    for i in range(n):
        lob = fields.get(f"recovery.row.{i}.lob")
        init = fields.get(f"recovery.row.{i}.initiative")
        if lob and init:
            keys[(str(lob), str(init))] = i
    return keys


def sync_recovery_tracker(state: Dict[str, Any], threshold: "float | None" = None) -> Dict[str, Any]:
    """Mutates and returns `state` in place. Returns a summary dict too."""
    fields = state.setdefault("fields", {})
    rows = state.setdefault("rows", {})

    if threshold is None:
        threshold = _num(fields.get("recovery.config.ti_threshold")) or DEFAULT_TI_THRESHOLD
    else:
        fields["recovery.config.ti_threshold"] = threshold

    all_initiatives = _lob_initiative_rows(fields, rows)
    scored = [r for r in all_initiatives if r["ti"] is not None]
    candidates = [r for r in scored if r["ti"] < threshold]
    existing = _existing_recovery_keys(fields, rows)
    n_existing = int(rows.get("recovery.row", 0) or 0)
    used_indices = set(existing.values())
    blank_slots = [i for i in range(n_existing)
                   if i not in used_indices and not fields.get(f"recovery.row.{i}.initiative")]

    added, updated = [], []
    next_idx = n_existing
    for cand in candidates:
        key = (cand["lob"], cand["initiative"])
        if key in existing:
            i = existing[key]
            fields[f"recovery.row.{i}.bri"] = cand["bri"]
            fields[f"recovery.row.{i}.ipi"] = cand["ipi"]
            fields[f"recovery.row.{i}.ti"] = cand["ti"]
            fields[f"recovery.row.{i}.item_type"] = cand["item_type"]
            updated.append(f"{cand['lob']} — {cand['initiative']}")
        else:
            i = blank_slots.pop(0) if blank_slots else next_idx
            if i == next_idx:
                next_idx += 1
            fields[f"recovery.row.{i}.rank"] = i + 1
            fields[f"recovery.row.{i}.initiative"] = cand["initiative"]
            fields[f"recovery.row.{i}.lob"] = cand["lob"]
            fields[f"recovery.row.{i}.item_type"] = cand["item_type"]
            fields[f"recovery.row.{i}.bri"] = cand["bri"]
            fields[f"recovery.row.{i}.status"] = "Watch"
            fields[f"recovery.row.{i}.ipi"] = cand["ipi"]
            fields[f"recovery.row.{i}.ti"] = cand["ti"]
            fields.setdefault(f"recovery.row.{i}.root_cause", "")
            fields.setdefault(f"recovery.row.{i}.corrective_action", "")
            fields.setdefault(f"recovery.row.{i}.owner", "")
            added.append(f"{cand['lob']} — {cand['initiative']}")

    total_rows = max(n_existing, next_idx)
    rows["recovery.row"] = total_rows

    # Re-rank by BRI size descending, matching the deck's own ordering rule
    # ("the largest cautious & at-risk BRIs, by size").
    live = [i for i in range(total_rows) if fields.get(f"recovery.row.{i}.initiative")]
    live.sort(key=lambda i: _num(fields.get(f"recovery.row.{i}.bri")) or -1, reverse=True)
    for rank, i in enumerate(live, start=1):
        fields[f"recovery.row.{i}.rank"] = rank

    return {
        "threshold": threshold,
        "scanned": len(scored),
        "flagged": len(candidates),
        "added": added,
        "updated": updated,
    }
