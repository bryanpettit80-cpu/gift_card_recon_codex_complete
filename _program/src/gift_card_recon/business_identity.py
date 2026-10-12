"""Business identities exclude transport names and retain record multiplicity."""
from __future__ import annotations

import hashlib
import json
from dataclasses import fields, is_dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterable, Mapping
from collections import Counter


def normalized(value: Any) -> Any:
    if is_dataclass(value) and not isinstance(value, type):
        result = {f.name: normalized(getattr(value, f.name)) for f in fields(value)
                  if f.name not in {"source_file", "source_files", "path", "modified_at", "generated_at"}}
        if type(value).__name__ == "ActivityRow":
            result["request_code_listing"] = value.request_label_lower
            for field in ("request", "transaction_no", "authorization_code", "card_no", "promocode"):
                raw = getattr(value, field)
                if isinstance(raw, float) and raw.is_integer():
                    raw = int(raw)
                result[field] = "" if raw is None else str(raw).strip()
        return result
    if isinstance(value, Decimal):
        if not value.is_finite():
            raise ValueError("Business identity requires finite amounts.")
        return "0" if value == 0 else format(value.normalize(), "f")
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, Path):
        raise ValueError("Business identity must not depend on a filesystem path.")
    if isinstance(value, Mapping):
        return {str(k): normalized(v) for k, v in sorted(value.items())}
    if isinstance(value, (list, tuple, set, frozenset)):
        return sorted((normalized(v) for v in value), key=_json)
    return value


def _json(value: Any) -> str:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False)


def identity(value: Any) -> str:
    return hashlib.sha256(_json(normalized(value)).encode("utf-8")).hexdigest()


def activity_identity(activity: Any) -> str:
    # A sorted list, deliberately NOT a set: two identical records are two records.
    return identity(activity)


def activity_changes(previous: Any, current: Any, limit: int = 10) -> dict:
    """Bounded review facts; never print a complete gift-card number."""
    old = Counter(_json(normalized(row)) for row in previous.rows)
    new = Counter(_json(normalized(row)) for row in current.rows)
    def describe(counter):
        records = []
        for raw, count in sorted(counter.items())[:limit]:
            row = json.loads(raw)
            records.append({"date": row["business_date"], "type": row["request_code_listing"],
                            "amount": row["amount"], "transaction": row["transaction_no"],
                            "count": count})
        return {"records": records, "additional_distinct_records": max(0, len(counter) - limit)}
    return {"removed": describe(old - new), "added": describe(new - old)}


def input_set_identity(activities: Iterable[Any], summary: Any = None, extra: Any = None) -> str:
    return identity({"activities": list(activities), "summary": summary, "extra": extra})


def correction_token(previous_id: str, input_set_id: str) -> str:
    return identity({"replace": previous_id, "with": input_set_id})


class CorrectionRequired(ValueError):
    def __init__(self, previous_id: str, input_set_id: str, changes: Mapping[str, Any]):
        self.previous_id = previous_id
        self.input_set_id = input_set_id
        self.token = correction_token(previous_id, input_set_id)
        self.changes = dict(changes)
        super().__init__(
            "Correction requires one deliberate replacement decision. "
            + "; ".join(f"{key}: {value}" for key, value in changes.items())
            + f". Approve correction token {self.token}. Previous result will be retained."
        )


def workbook_values_identity(path: Path, ignored_cells: Mapping[str, Iterable[str]] | None = None) -> str:
    from openpyxl import load_workbook
    ignored = {name: set(cells) for name, cells in (ignored_cells or {}).items()}
    try:
        book = load_workbook(path, data_only=False, keep_links=False)
    except Exception as exc:
        raise ValueError(f"Cannot read report workbook {path}: {exc}") from exc
    try:
        # Sheet order, names, typed values and formulas remain significant; styles,
        # ZIP metadata, cached formula values and empty formatted cells do not.
        contents = []
        for sheet in book:
            cells = [(cell.coordinate, cell.data_type,
                      normalized(Decimal(str(cell.value))) if cell.data_type == "n" and isinstance(cell.value, (int, float, Decimal))
                      else normalized(cell.value))
                     for row in sheet for cell in row if cell.value is not None
                     and cell.coordinate not in ignored.get(sheet.title, set())]
            contents.append([sheet.title, cells])
        names = sorted((name, item.attr_text) for name, item in book.defined_names.items())
        return hashlib.sha256(_json([contents, names]).encode("utf-8")).hexdigest()
    finally:
        book.close()


def verify_workbook_values(current: Path, baseline: Path, ignored_cells=None) -> None:
    if workbook_values_identity(current, ignored_cells) != workbook_values_identity(baseline, ignored_cells):
        raise ValueError(f"Report values or formulas differ from preserved evidence: {current}")
