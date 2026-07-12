from __future__ import annotations

import csv
import os
import re
import shutil
import subprocess
import tempfile
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path
from typing import Iterable, Mapping, Sequence

from gift_card_recon.models import MicrosDailyPosControl, WeeklyPosVariance
from gift_card_recon.parsers import ParseError
from gift_card_recon.reconcile import rollup_activity_file
from gift_card_recon.source_validation import ActivityEvidence
from gift_card_recon.store_config import StoreConfig
from gift_card_recon.utils import money, parse_date, sha256_file, to_decimal


@dataclass(frozen=True)
class MicrosEvidence:
    source_dir: Path
    system_totals_path: Path
    tender_detail_path: Path
    daily_pos: tuple[MicrosDailyPosControl, ...]
    daily_tender: Mapping[date, Decimal]
    daily_tender_magnitude: Mapping[date, Decimal]
    tender_observed_dates: frozenset[date]
    accepted_closed_dates: tuple[date, ...]

    @property
    def daily_pos_by_date(self) -> Mapping[date, MicrosDailyPosControl]:
        return {row.business_date: row for row in self.daily_pos}


def resolve_micros_export_dir(source_path: Path, work_dir: Path) -> Path:
    """Resolve an extracted Micros folder or safely extract a supported archive."""

    source = Path(source_path)
    if source.is_dir():
        return source
    if not source.exists():
        raise ParseError(f"Micros export path not found: {source}")
    suffix = source.suffix.lower()
    if suffix not in {".zip", ".7z"}:
        raise ParseError(
            f"Unsupported Micros export path {source}. Use an extracted folder, .zip, or .7z."
        )
    work_root = Path(work_dir).resolve()
    work_root.mkdir(parents=True, exist_ok=True)
    digest = sha256_file(source)[:16]
    destination = work_root / f"{source.stem}-{digest}"
    completion_marker = destination / ".extraction-complete"
    if destination.is_dir() and completion_marker.is_file():
        return destination

    with tempfile.TemporaryDirectory(prefix="micros-extract-", dir=work_root) as temp_dir:
        staged = Path(temp_dir) / "payload"
        staged.mkdir()
        if suffix == ".zip":
            import zipfile

            try:
                with zipfile.ZipFile(source) as archive:
                    _safe_extract_zip(archive, staged)
            except (OSError, zipfile.BadZipFile) as exc:
                raise ParseError(f"Could not extract Micros archive {source}: {exc}") from exc
        else:
            executable = _find_7z()
            if executable is None:
                raise ParseError(
                    "Could not find 7-Zip. Install it or provide an extracted Micros folder."
                )
            try:
                subprocess.run(
                    [executable, "x", "-y", f"-o{staged}", str(source)],
                    check=True,
                    capture_output=True,
                    text=True,
                )
            except subprocess.CalledProcessError as exc:
                raise ParseError(f"Could not extract Micros archive {source}: {exc.stderr or exc}") from exc
        (staged / ".extraction-complete").write_text(digest, encoding="ascii")
        try:
            os.replace(staged, destination)
        except FileExistsError:
            if not completion_marker.is_file():
                raise ParseError(f"Incomplete Micros extraction already exists: {destination}")
    return destination


def validate_micros_source(
    source_dir: Path,
    config: StoreConfig,
    *,
    allow_archive_snapshot: bool = True,
    allow_unconfigured_source: bool = False,
) -> None:
    """Prevent one location's live Micros folder from being used for the other."""

    resolved = Path(source_dir).resolve()
    expected = config.micros_default_path.resolve()
    if resolved == expected:
        return
    if allow_archive_snapshot and _looks_like_store_archive(resolved, config.store):
        return
    other_live_sources = {
        item.micros_default_path.resolve()
        for item in _all_store_configs()
        if item.store != config.store
    }
    if resolved in other_live_sources:
        raise ParseError(
            f"Micros source {resolved} is configured for a different location; "
            f"expected {config.micros_source_label} at {expected}."
        )
    if not allow_unconfigured_source:
        raise ParseError(
            f"Micros source {resolved} is not the configured source for "
            f"{config.location_name} store {config.store}; expected {expected}."
        )


def load_micros_evidence(
    source_dir: Path,
    *,
    config: StoreConfig,
    activity_evidence: ActivityEvidence,
    period_start: date,
    period_end: date,
    validate_source: bool = True,
    allow_unconfigured_source: bool = False,
) -> MicrosEvidence:
    source_dir = Path(source_dir)
    if validate_source:
        validate_micros_source(
            source_dir,
            config,
            allow_unconfigured_source=allow_unconfigured_source,
        )
    system_path = _required_case_insensitive_file(source_dir, config.micros_system_totals_file)
    tender_path = _required_case_insensitive_file(source_dir, config.micros_tender_detail_file)
    daily_pos = _parse_system_totals(system_path, config)
    daily_tender, daily_tender_magnitude, tender_observed_dates = _parse_tender_detail(
        tender_path,
        config,
    )

    pos_by_date = {row.business_date: row for row in daily_pos}
    activity_by_date = activity_evidence.daily_activity_magnitude
    accepted_closed_dates: list[date] = []
    coverage_errors: list[str] = []
    for business_date in _date_range(period_start, period_end):
        if business_date in pos_by_date:
            if business_date not in tender_observed_dates:
                pos_row = pos_by_date[business_date]
                activity_total = money(
                    activity_by_date.get(business_date, Decimal("0.00"))
                )
                scheduled_zero_day = (
                    business_date.weekday() in config.scheduled_closed_weekdays
                    and pos_row.pos_gift_card_issue == Decimal("0.00")
                    and pos_row.pos_gift_card_payment == Decimal("0.00")
                    and activity_total == Decimal("0.00")
                )
                if not scheduled_zero_day:
                    coverage_errors.append(
                        f"{business_date.isoformat()} is missing from {tender_path.name}."
                    )
            continue
        activity_total = money(activity_by_date.get(business_date, Decimal("0.00")))
        tender_total = money(daily_tender.get(business_date, Decimal("0.00")))
        tender_magnitude = money(
            daily_tender_magnitude.get(business_date, Decimal("0.00"))
        )
        if (
            business_date.weekday() in config.scheduled_closed_weekdays
            and activity_total == Decimal("0.00")
            and tender_total == Decimal("0.00")
            and tender_magnitude == Decimal("0.00")
        ):
            accepted_closed_dates.append(business_date)
            continue
        coverage_errors.append(
            f"{business_date.isoformat()} is missing from {system_path.name} "
            f"(activity {activity_total:,.2f}; tender {tender_total:,.2f})"
        )
    if coverage_errors:
        raise ParseError("Micros coverage validation failed: " + "; ".join(coverage_errors) + ".")

    return MicrosEvidence(
        source_dir=source_dir,
        system_totals_path=system_path,
        tender_detail_path=tender_path,
        daily_pos=tuple(row for row in daily_pos if period_start <= row.business_date <= period_end),
        daily_tender={
            key: money(value)
            for key, value in daily_tender.items()
            if period_start <= key <= period_end
        },
        daily_tender_magnitude={
            key: money(value)
            for key, value in daily_tender_magnitude.items()
            if period_start <= key <= period_end
        },
        tender_observed_dates=frozenset(
            key for key in tender_observed_dates if period_start <= key <= period_end
        ),
        accepted_closed_dates=tuple(accepted_closed_dates),
    )


def weekly_tender_variances(
    evidence: MicrosEvidence,
    *,
    week_endings: Iterable[date],
) -> dict[str, Decimal]:
    pos_by_date = evidence.daily_pos_by_date
    values: dict[str, Decimal] = {}
    for week_end in week_endings:
        week_start = week_end - timedelta(days=6)
        expected = list(_date_range(week_start, week_end))
        pos_total = sum(
            (pos_by_date[value].pos_gift_card_payment for value in expected if value in pos_by_date),
            Decimal("0.00"),
        )
        tender_total = sum(
            (evidence.daily_tender.get(value, Decimal("0.00")) for value in expected),
            Decimal("0.00"),
        )
        values[f"Week ending {week_end:%m/%d/%Y} tender"] = money(pos_total - tender_total)
    return values


def build_weekly_pos_variances(
    activity_evidence: ActivityEvidence,
    micros_evidence: MicrosEvidence,
    *,
    conversion_promo_codes: set[str],
) -> list[WeeklyPosVariance]:
    """Compare weekly activity with real POS rows; never manufacture missing totals."""

    pos_by_date = micros_evidence.daily_pos_by_date
    accepted_closed = set(micros_evidence.accepted_closed_dates)
    rows: list[WeeklyPosVariance] = []
    for activity in activity_evidence.files:
        if activity.report_begin is None or activity.report_end is None:
            raise ParseError(f"Activity range was not validated for {activity.source_file.name}.")
        report_dates = list(_date_range(activity.report_begin, activity.report_end))
        rollup = rollup_activity_file(activity, conversion_promo_codes)
        activity_issue = money(rollup.net_activations)
        activity_payment = money(abs(rollup.net_redemptions))
        pos_issue = money(
            sum(
                (pos_by_date[value].pos_gift_card_issue for value in report_dates if value in pos_by_date),
                Decimal("0.00"),
            )
        )
        pos_payment = money(
            sum(
                (pos_by_date[value].pos_gift_card_payment for value in report_dates if value in pos_by_date),
                Decimal("0.00"),
            )
        )
        missing = [value for value in report_dates if value not in pos_by_date]
        if not missing:
            coverage_status = "Complete — all business dates present"
        elif set(missing).issubset(accepted_closed):
            dates = ", ".join(value.strftime("%m/%d") for value in missing)
            coverage_status = f"Complete — scheduled closure(s) {dates}"
        else:
            # load_micros_evidence should make this state unreachable.
            raise ParseError(
                f"Unvalidated POS coverage gap for week ending {activity.report_end:%Y-%m-%d}."
            )
        issue_variance = money(pos_issue - activity_issue)
        payment_variance = money(pos_payment - activity_payment)
        rows.append(
            WeeklyPosVariance(
                week_ending=activity.report_end,
                report_begin=activity.report_begin,
                report_end=activity.report_end,
                activity_issue=activity_issue,
                pos_issue=pos_issue,
                issue_variance=issue_variance,
                activity_payment=activity_payment,
                pos_payment=pos_payment,
                payment_variance=payment_variance,
                net_variance=money(issue_variance - payment_variance),
                coverage_status=coverage_status,
            )
        )
    return rows


def period_tender_variance(evidence: MicrosEvidence) -> Decimal:
    pos_total = sum((row.pos_gift_card_payment for row in evidence.daily_pos), Decimal("0.00"))
    tender_total = sum(evidence.daily_tender.values(), Decimal("0.00"))
    return money(pos_total - tender_total)


def _parse_system_totals(path: Path, config: StoreConfig) -> list[MicrosDailyPosControl]:
    rows: list[MicrosDailyPosControl] = []
    seen_dates: Counter[date] = Counter()
    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        for line_no, row in enumerate(csv.reader(stream), start=1):
            if not row:
                continue
            business_date = _strict_micros_date(row[0], path, line_no)
            required_columns = max(
                config.micros_issue_column_index,
                config.micros_payment_column_index,
            ) + 1
            if len(row) < required_columns:
                raise ParseError(
                    f"{path.name} line {line_no} has {len(row)} columns; "
                    f"{config.location_name} requires at least {required_columns}."
                )
            issue = _strict_money(row[config.micros_issue_column_index], path, line_no)
            payment = _strict_money(row[config.micros_payment_column_index], path, line_no)
            rows.append(
                MicrosDailyPosControl(
                    business_date=business_date,
                    pos_gift_card_issue=issue,
                    pos_gift_card_payment=payment,
                )
            )
            seen_dates[business_date] += 1
    duplicates = sorted(value for value, count in seen_dates.items() if count > 1)
    if duplicates:
        raise ParseError(
            f"{path.name} contains duplicate business date(s): "
            + ", ".join(value.isoformat() for value in duplicates)
            + "."
        )
    if not rows:
        raise ParseError(f"No POS rows parsed from {path}.")
    return sorted(rows, key=lambda row: row.business_date)


def _parse_tender_detail(
    path: Path,
    config: StoreConfig,
) -> tuple[dict[date, Decimal], dict[date, Decimal], frozenset[date]]:
    allowed_names = {_normalize_tender_name(value) for value in config.gift_card_payment_tenders}
    totals: dict[date, Decimal] = defaultdict(lambda: Decimal("0.00"))
    magnitudes: dict[date, Decimal] = defaultdict(lambda: Decimal("0.00"))
    observed_dates: set[date] = set()
    parsed_rows = 0
    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        for line_no, row in enumerate(csv.reader(stream), start=1):
            if not row:
                continue
            parsed_rows += 1
            if len(row) < 4:
                raise ParseError(
                    f"{path.name} line {line_no} has {len(row)} columns; expected at least 4."
                )
            business_date = _strict_micros_date(row[0], path, line_no)
            amount = _strict_money(row[1], path, line_no)
            observed_dates.add(business_date)
            if _normalize_tender_name(row[3]) in allowed_names:
                totals[business_date] += amount
                magnitudes[business_date] += abs(amount)
    if parsed_rows == 0:
        raise ParseError(f"Tender evidence is empty: {path}.")
    return (
        {key: money(value) for key, value in totals.items()},
        {key: money(value) for key, value in magnitudes.items()},
        frozenset(observed_dates),
    )


def _required_case_insensitive_file(folder: Path, expected_name: str) -> Path:
    if not folder.is_dir():
        raise ParseError(f"Micros export folder not found: {folder}")
    matches = [item for item in folder.iterdir() if item.is_file() and item.name.casefold() == expected_name.casefold()]
    if len(matches) != 1:
        raise ParseError(
            f"Expected exactly one {expected_name} in {folder}; found {len(matches)}."
        )
    return matches[0]


def _strict_micros_date(value: object, path: Path, line_no: int) -> date:
    text = str(value or "").strip().strip("'").strip()
    parsed = parse_date(text.split()[0] if text else text)
    if parsed is None:
        raise ParseError(f"Could not parse date in {path.name} line {line_no}: {value!r}.")
    return parsed


def _strict_money(value: object, path: Path, line_no: int) -> Decimal:
    parsed = to_decimal(value)
    if parsed is None:
        raise ParseError(
            f"Could not parse monetary value in {path.name} line {line_no}: {value!r}."
        )
    return money(parsed)


def _normalize_tender_name(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().strip("'").strip()).casefold()


def _date_range(start: date, end: date) -> Iterable[date]:
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def _looks_like_store_archive(path: Path, store: str) -> bool:
    return str(store) in {part.strip() for part in path.parts}


def _all_store_configs() -> Sequence[StoreConfig]:
    from gift_card_recon.store_config import STORE_CONFIGS

    return tuple(STORE_CONFIGS.values())


def _find_7z() -> str | None:
    for candidate in (
        shutil.which("7z"),
        shutil.which("7za"),
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
    ):
        if candidate and Path(candidate).is_file():
            return str(candidate)
    return None


def _safe_extract_zip(archive, destination: Path) -> None:
    root = destination.resolve()
    for member in archive.infolist():
        target = (root / member.filename).resolve()
        try:
            target.relative_to(root)
        except ValueError as exc:
            raise ParseError(f"Unsafe path in Micros ZIP archive: {member.filename}") from exc
    archive.extractall(root)
