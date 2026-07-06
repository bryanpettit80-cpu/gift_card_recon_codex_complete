from __future__ import annotations

import hashlib
import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any

MONEY_CENTS = Decimal("0.01")
ZERO = Decimal("0.00")


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").strip().lower()
    return re.sub(r"\s+", " ", text)


def normalize_header(value: Any) -> str:
    text = normalize_text(value)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def clean_code(value: Any) -> str | None:
    if value is None or value == "":
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def to_decimal(value: Any, default: Decimal | None = None) -> Decimal | None:
    if value is None or value == "":
        return default
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool):
        return Decimal(int(value))
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        return Decimal(str(value))
    text = str(value).strip()
    if not text:
        return default
    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]
    text = text.replace("$", "").replace(",", "").replace(" ", "")
    if text.endswith("-"):
        negative = True
        text = text[:-1]
    try:
        result = Decimal(text)
    except InvalidOperation:
        return default
    return -result if negative else result


def money(value: Any) -> Decimal:
    d = to_decimal(value, ZERO)
    assert d is not None
    return d.quantize(MONEY_CENTS, rounding=ROUND_HALF_UP)


def parse_date(value: Any) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if not text:
        return None
    for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d-%b-%Y", "%d-%b-%y", "%d-%B-%Y", "%d-%B-%y"]:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def file_modified_at(path: Path) -> datetime | None:
    try:
        return datetime.fromtimestamp(Path(path).stat().st_mtime)
    except OSError:
        return None


def variance_status(value: Decimal | None, tolerance: Decimal = Decimal("0.01"), review_threshold: Decimal = Decimal("5.00")) -> str:
    if value is None:
        return "N/A"
    av = abs(money(value))
    if av <= tolerance:
        return "OK"
    if av <= review_threshold:
        return "Minor variance"
    return "Review"
