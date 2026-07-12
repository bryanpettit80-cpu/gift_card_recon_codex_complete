from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path
from typing import Iterable, Mapping, Sequence

from gift_card_recon.models import ActivityFileData
from gift_card_recon.parsers import ParseError
from gift_card_recon.utils import money


@dataclass(frozen=True)
class ActivityEvidence:
    """Validated, ordered activity evidence for one fiscal period."""

    store: str
    period_start: date
    period_end: date
    files: tuple[ActivityFileData, ...]

    @property
    def paths(self) -> tuple[Path, ...]:
        return tuple(item.source_file for item in self.files)

    @property
    def by_week_ending(self) -> Mapping[date, ActivityFileData]:
        return {item.report_end: item for item in self.files if item.report_end is not None}

    @property
    def daily_activity_totals(self) -> Mapping[date, Decimal]:
        totals: dict[date, Decimal] = defaultdict(lambda: Decimal("0.00"))
        for item in self.files:
            for row in item.rows:
                if row.business_date is not None:
                    totals[row.business_date] += row.amount
        return {business_date: money(total) for business_date, total in totals.items()}

    @property
    def daily_activity_magnitude(self) -> Mapping[date, Decimal]:
        """Absolute activity by day, so offsetting transactions still count as evidence."""

        totals: dict[date, Decimal] = defaultdict(lambda: Decimal("0.00"))
        for item in self.files:
            for row in item.rows:
                if row.business_date is not None:
                    totals[row.business_date] += abs(row.amount)
        return {business_date: money(total) for business_date, total in totals.items()}


def validate_activity_evidence(
    activities: Sequence[ActivityFileData],
    *,
    store: str | int,
    period_start: date,
    period_end: date,
    expected_week_endings: Iterable[date],
) -> ActivityEvidence:
    """Require one correct-store Monday-Sunday report for every expected week."""

    expected = tuple(expected_week_endings)
    errors: list[str] = []
    store_number = str(store).strip()

    if not activities:
        raise ParseError("No activity evidence was supplied for the monthly close.")

    ends = [item.report_end for item in activities]
    duplicate_ends = sorted(
        week for week, count in Counter(ends).items() if week is not None and count > 1
    )
    if duplicate_ends:
        errors.append(
            "duplicate week-ending report(s): "
            + ", ".join(value.isoformat() for value in duplicate_ends)
        )

    missing = sorted(set(expected) - {value for value in ends if value is not None})
    extra = sorted({value for value in ends if value is not None} - set(expected))
    if missing:
        errors.append(
            "missing expected week-ending report(s): "
            + ", ".join(value.isoformat() for value in missing)
        )
    if extra:
        errors.append(
            "out-of-period week-ending report(s): "
            + ", ".join(value.isoformat() for value in extra)
        )

    ranges: list[tuple[date, date, Path]] = []
    for item in activities:
        if item.store != store_number:
            errors.append(
                f"{item.source_file.name} identifies store {item.store or 'unknown'}; "
                f"expected {store_number}"
            )
        if item.report_begin is None or item.report_end is None:
            errors.append(f"{item.source_file.name} is missing its report date range")
            continue
        expected_begin = item.report_end - timedelta(days=6)
        if item.report_begin != expected_begin:
            errors.append(
                f"{item.source_file.name} covers {item.report_begin.isoformat()} through "
                f"{item.report_end.isoformat()}; expected a Monday-Sunday seven-day report"
            )
        if item.report_begin < period_start or item.report_end > period_end:
            errors.append(
                f"{item.source_file.name} contains an out-of-period report range "
                f"({item.report_begin.isoformat()} through {item.report_end.isoformat()})"
            )
        ranges.append((item.report_begin, item.report_end, item.source_file))
        for row in item.rows:
            if row.business_date is None:
                errors.append(f"{item.source_file.name} contains a row without a business date")
            elif not item.report_begin <= row.business_date <= item.report_end:
                errors.append(
                    f"{item.source_file.name} contains transaction date "
                    f"{row.business_date.isoformat()} outside its report range"
                )
            elif not period_start <= row.business_date <= period_end:
                errors.append(
                    f"{item.source_file.name} contains out-of-period transaction date "
                    f"{row.business_date.isoformat()}"
                )

    ordered_ranges = sorted(ranges, key=lambda item: (item[0], item[1], item[2].name))
    for previous, current in zip(ordered_ranges, ordered_ranges[1:]):
        if current[0] <= previous[1]:
            errors.append(
                f"overlapping activity reports: {previous[2].name} and {current[2].name}"
            )

    if errors:
        raise ParseError("Activity evidence validation failed: " + "; ".join(errors) + ".")

    ordered = tuple(sorted(activities, key=lambda item: (item.report_end, item.source_file.name)))
    return ActivityEvidence(
        store=store_number,
        period_start=period_start,
        period_end=period_end,
        files=ordered,
    )
