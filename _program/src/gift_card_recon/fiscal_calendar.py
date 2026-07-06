from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from datetime import date, timedelta
from importlib import resources
from pathlib import Path
from typing import Iterable, Sequence

from gift_card_recon.parsers import ParseError
from gift_card_recon.utils import parse_date

FISCAL_CALENDAR_SOURCE = "Darden Fiscal Calendar as of 11_2025.pdf"
FISCAL_CALENDAR_SOURCE_URL = "https://drive.google.com/file/d/1em5tuw4Iz6oH08InQE_-I1P7GtKrggun"


@dataclass(frozen=True)
class FiscalPeriod:
    period_key: str
    folder_name: str
    fiscal_year: int
    month_number: int
    fiscal_month: str
    month_year: int
    start_date: date
    end_date: date

    @property
    def expected_week_endings(self) -> list[date]:
        current = self.start_date
        while current.weekday() != 6:
            current += timedelta(days=1)
        endings: list[date] = []
        while current <= self.end_date:
            endings.append(current)
            current += timedelta(days=7)
        return endings


def load_fiscal_calendar(calendar_file: Path | None = None) -> list[FiscalPeriod]:
    if calendar_file is None:
        with resources.files("gift_card_recon.data").joinpath("darden_fiscal_calendar.csv").open("r", encoding="utf-8", newline="") as f:
            return _read_calendar_rows(f)

    with Path(calendar_file).open("r", encoding="utf-8-sig", newline="") as f:
        return _read_calendar_rows(f)


def fiscal_period_for_date(value: date, periods: Sequence[FiscalPeriod] | None = None) -> FiscalPeriod:
    periods = periods or load_fiscal_calendar()
    for period in periods:
        if period.start_date <= value <= period.end_date:
            return period
    first = min((p.start_date for p in periods), default=None)
    last = max((p.end_date for p in periods), default=None)
    if first and last:
        raise ParseError(
            f"Darden fiscal calendar does not cover {value:%Y-%m-%d}. "
            f"Loaded coverage is {first:%Y-%m-%d} through {last:%Y-%m-%d}."
        )
    raise ParseError("Darden fiscal calendar is empty.")


def fiscal_period_for_label(value: str, periods: Sequence[FiscalPeriod] | None = None) -> FiscalPeriod:
    text = str(value or "").strip()
    if not text:
        raise ParseError("Monthly close period is required.")

    periods = periods or load_fiscal_calendar()
    normalized = _normalize_period_text(text)
    for period in periods:
        aliases = {
            _normalize_period_text(period.period_key),
            _normalize_period_text(period.folder_name),
            _normalize_period_text(f"FY{period.fiscal_year % 100:02d} M{period.month_number:02d}"),
            _normalize_period_text(f"{period.month_year}-{_calendar_month_number(period.fiscal_month):02d}"),
            _normalize_period_text(f"{period.fiscal_month} {period.month_year}"),
            _normalize_period_text(f"Fiscal {period.fiscal_month} {period.month_year}"),
        }
        if normalized in aliases:
            return period

    valid = ", ".join(period.period_key for period in periods[:3])
    raise ParseError(f"Unknown Darden fiscal period {value!r}. Use a key like {valid}, or YYYY-MM for the fiscal month.")


def fiscal_periods_overlapping(start: date, end: date, periods: Sequence[FiscalPeriod] | None = None) -> list[FiscalPeriod]:
    periods = periods or load_fiscal_calendar()
    return [period for period in periods if period.start_date <= end and start <= period.end_date]


def fiscal_month_bounds(period: str) -> tuple[date, date]:
    fiscal_period = fiscal_period_for_label(period)
    return fiscal_period.start_date, fiscal_period.end_date


def fiscal_month_week_endings(period: str) -> list[date]:
    return fiscal_period_for_label(period).expected_week_endings


def _read_calendar_rows(rows: Iterable[str]) -> list[FiscalPeriod]:
    reader = csv.DictReader(rows)
    periods: list[FiscalPeriod] = []
    for row in reader:
        period = FiscalPeriod(
            period_key=str(row["period_key"]).strip(),
            folder_name=str(row["folder_name"]).strip(),
            fiscal_year=int(row["fiscal_year"]),
            month_number=int(row["month_number"]),
            fiscal_month=str(row["fiscal_month"]).strip(),
            month_year=int(row["month_year"]),
            start_date=_required_date(row["start_date"], "start_date"),
            end_date=_required_date(row["end_date"], "end_date"),
        )
        periods.append(period)
    return periods


def _required_date(value: object, field: str) -> date:
    parsed = parse_date(value)
    if parsed is None:
        raise ParseError(f"Darden fiscal calendar has a malformed {field}: {value!r}")
    return parsed


def _normalize_period_text(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def _calendar_month_number(month_name: str) -> int:
    months = {
        "june": 6,
        "july": 7,
        "august": 8,
        "september": 9,
        "october": 10,
        "november": 11,
        "december": 12,
        "january": 1,
        "february": 2,
        "march": 3,
        "april": 4,
        "may": 5,
    }
    key = month_name.strip().lower()
    if key not in months:
        raise ParseError(f"Unsupported fiscal month name in Darden calendar: {month_name!r}")
    return months[key]
