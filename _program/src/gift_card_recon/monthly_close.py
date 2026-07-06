from __future__ import annotations

import argparse
import csv
import re
import shutil
import subprocess
import tempfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path
from typing import Iterable, Sequence

from gift_card_recon.excel_writer import append_weekly_pos_variance_detail, write_reconciliation_workbook
from gift_card_recon.fiscal_calendar import FiscalPeriod, fiscal_period_for_label
from gift_card_recon.models import ActivityFileData, MicrosDailyPosControl, PosControls, WeeklyPosVariance
from gift_card_recon.parsers import ParseError, discover_input_files, parse_activity_file, parse_summary
from gift_card_recon.reconcile import build_reconciliation, rollup_activity_file
from gift_card_recon.utils import money, parse_date

SYSTEM_TOTALS_FILE = "DLYSYSTT.TXT"
TENDER_DETAIL_FILE = "TENDER_DETAIL.TXT"
GIFT_CARD_PAYMENT_TENDERS = {"G C Payment", "Gift Card Payment"}
CLOSED_BUSINESS_WEEKDAYS = {0}  # Monday; closed days are omitted from Micros POS exports.

# Zero-based indexes in DLYSYSTT.TXT. These are the validated 1-based columns
# 121 and 103 in the Micros 3700 system totals export used for store 9355.
ISSUE_AMOUNT_INDEX = 120
PAYMENT_AMOUNT_INDEX = 102


@dataclass(frozen=True)
class MonthlyClosePreflight:
    store: str
    period: str
    input_dir: Path
    summary_dir: Path
    activity_dir: Path
    expected_week_endings: list[date]
    summary_paths: list[Path]
    activity_by_week_end: dict[date, Path]
    missing_summary_path: Path
    missing_activity_paths: list[Path]
    staged_activity_paths: list[Path]
    micros_path: Path
    micros_ready: bool
    micros_message: str
    micros_missing_paths: list[Path]

    @property
    def ready(self) -> bool:
        return not self.required_missing_paths and self.micros_ready

    @property
    def required_missing_paths(self) -> list[Path]:
        paths: list[Path] = []
        if len(self.summary_paths) != 1:
            paths.append(self.missing_summary_path)
        paths.extend(self.missing_activity_paths)
        paths.extend(self.micros_missing_paths)
        return paths


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    store = str(args.store)
    try:
        fiscal_period = fiscal_period_for_label(str(args.period))
    except ParseError as exc:
        raise SystemExit(str(exc)) from exc
    period = fiscal_period.period_key
    period_start = fiscal_period.start_date
    default_period_end = fiscal_period.end_date
    period_end = parse_date(args.period_end) if args.period_end else default_period_end
    if period_end is None:
        raise SystemExit(f"Could not parse --period-end value: {args.period_end!r}")

    input_root = Path(args.input_root)
    input_dir = Path(args.input_dir) if args.input_dir else input_root / store / fiscal_period.folder_name
    output_dir = Path(args.output_dir)
    output_path = Path(args.output_file) if args.output_file else output_dir / f"Gift_Card_Reconciliation_{store}_{period}.xlsx"
    micros_path = Path(args.micros_path)
    micros_work_dir = Path(args.micros_work_dir)
    archive_root = None if args.no_cleanup else Path(args.archive_root)

    try:
        preflight = prepare_monthly_close_inputs(
            store=store,
            period=period,
            fiscal_period=fiscal_period,
            period_start=period_start,
            period_end=period_end,
            input_root=input_root,
            input_dir=input_dir,
            micros_path=micros_path,
            micros_work_dir=micros_work_dir,
            stage_weekly=not args.no_stage_weekly,
        )
        if args.prepare_only or not preflight.ready:
            print(format_monthly_close_preflight(preflight))
            return 0 if preflight.ready else 1

        saved_path, result, weekly_variances = run_monthly_close(
            store=store,
            period=period,
            period_start=period_start,
            period_end=period_end,
            input_dir=input_dir,
            output_path=output_path,
            micros_path=micros_path,
            micros_work_dir=micros_work_dir,
            cleanup_archive_root=archive_root,
            fiscal_period=fiscal_period,
            adjust_boundary_weeks=not args.no_boundary_adjustment,
        )
    except ParseError as exc:
        raise SystemExit(str(exc)) from exc

    print(f"Created: {saved_path}")
    print(f"POS controls from Micros: issue={result.pos_controls.pos_gift_card_issue:,.2f} payment={result.pos_controls.pos_gift_card_payment:,.2f}")
    print("Weekly POS variance:")
    for row in weekly_variances:
        week = row.week_ending.strftime("%m/%d/%Y") if row.week_ending else "Unknown"
        print(
            f"  - {week}: issue variance={row.issue_variance:+,.2f} "
            f"payment variance={row.payment_variance:+,.2f} net variance={row.net_variance:+,.2f}"
        )
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="gift-card-monthly-close",
        description="Run monthly gift card close using Micros POS exports and append weekly POS variance detail.",
    )
    parser.add_argument("--store", default="9355", help="Store number. Defaults to 9355.")
    parser.add_argument("--period", default="FY27-M01", help="Darden fiscal period, such as FY27-M01 or 2026-06 for Fiscal June 2026.")
    parser.add_argument("--period-end", default=None, help="Optional period end date. Defaults to the Darden fiscal period end.")
    parser.add_argument("--input-root", default="Monthly Close", help="Folder containing monthly close store folders. Defaults to Monthly Close.")
    parser.add_argument("--input-dir", default=None, help="Input folder containing summary/ and activity/. Defaults to Monthly Close/<store>/<fiscal period>.")
    parser.add_argument("--output-dir", default="Output", help="Folder for the generated workbook.")
    parser.add_argument("--output-file", default=None, help="Optional explicit output .xlsx path.")
    parser.add_argument("--micros-path", default="_inspect_micros3700", help="Micros export folder or .7z archive. Defaults to _inspect_micros3700.")
    parser.add_argument("--micros-work-dir", default="_program/tmp/monthly_close_micros", help="Extraction folder used when --micros-path is an archive.")
    parser.add_argument("--archive-root", default="Archive - Old Files", help="Folder where completed monthly close source files are archived.")
    parser.add_argument("--prepare-only", action="store_true", help="Create monthly folders, stage weekly activity files, and report missing inputs without creating a workbook.")
    parser.add_argument("--no-stage-weekly", action="store_true", help="Do not copy available weekly activity files into the monthly close activity folder before preflight.")
    parser.add_argument("--no-cleanup", action="store_true", help="Do not archive monthly close source files after a successful workbook.")
    parser.add_argument(
        "--no-boundary-adjustment",
        action="store_true",
        help="Do not hold boundary weeks to activity totals when Micros dates do not cover dates outside the monthly period.",
    )
    return parser


def prepare_monthly_close_inputs(
    *,
    store: str,
    period: str,
    fiscal_period: FiscalPeriod | None = None,
    period_start: date,
    period_end: date,
    input_root: Path,
    input_dir: Path,
    micros_path: Path,
    micros_work_dir: Path,
    stage_weekly: bool = True,
) -> MonthlyClosePreflight:
    input_root = Path(input_root)
    input_dir = Path(input_dir)
    summary_dir = input_dir / "summary"
    activity_dir = input_dir / "activity"
    summary_dir.mkdir(parents=True, exist_ok=True)
    activity_dir.mkdir(parents=True, exist_ok=True)
    _stage_monthly_summary_files_for_period(
        store=store,
        period_end=period_end,
        store_monthly_dir=input_dir.parent,
        summary_dir=summary_dir,
    )

    staged_activity_paths: list[Path] = []
    if stage_weekly:
        staged_activity_paths = stage_weekly_activity_files_for_month(
            store=store,
            period=period,
            fiscal_period=fiscal_period,
            period_start=period_start,
            period_end=period_end,
            input_root=input_root,
            monthly_activity_dir=activity_dir,
        )

    summary_paths = _monthly_summary_paths(input_dir)
    activity_by_week_end = _monthly_activity_by_week_end(input_dir)
    expected_week_endings = monthly_activity_week_endings(period_start, period_end)
    missing_activity_paths = [
        activity_dir / f"{week_end:%m.%d.%Y} {store} Gift Card Activity.xls"
        for week_end in expected_week_endings
        if week_end not in activity_by_week_end
    ]
    micros_ready, micros_message, micros_missing_paths = _micros_preflight(
        micros_path=Path(micros_path),
        micros_work_dir=Path(micros_work_dir),
        period_end=period_end,
    )

    return MonthlyClosePreflight(
        store=str(store),
        period=str(period),
        input_dir=input_dir,
        summary_dir=summary_dir,
        activity_dir=activity_dir,
        expected_week_endings=expected_week_endings,
        summary_paths=summary_paths,
        activity_by_week_end=activity_by_week_end,
        missing_summary_path=summary_dir / f"{period_end:%m.%d.%Y} {store} Gift Card Summary.xlsx",
        missing_activity_paths=missing_activity_paths,
        staged_activity_paths=staged_activity_paths,
        micros_path=Path(micros_path),
        micros_ready=micros_ready,
        micros_message=micros_message,
        micros_missing_paths=micros_missing_paths,
    )


def stage_activity_files_for_month(
    *,
    store: str,
    period: str,
    monthly_activity_dir: Path,
    activity_paths: Sequence[Path],
    move: bool = False,
) -> list[Path]:
    fiscal_period = fiscal_period_for_label(period)
    period_start, period_end = fiscal_period.start_date, fiscal_period.end_date
    monthly_activity_dir = Path(monthly_activity_dir)
    monthly_activity_dir.mkdir(parents=True, exist_ok=True)
    staged: list[Path] = []
    for source in activity_paths:
        source = Path(source)
        report_end = _activity_report_end(source)
        if report_end is None or not (period_start <= report_end <= period_end):
            continue
        destination = monthly_activity_dir / source.name
        saved_path = _move_if_needed(source, destination) if move else _copy_if_needed(source, destination)
        if saved_path:
            staged.append(saved_path)
    return staged


def stage_weekly_activity_files_for_month(
    *,
    store: str,
    period: str,
    fiscal_period: FiscalPeriod | None = None,
    period_start: date,
    period_end: date,
    input_root: Path,
    monthly_activity_dir: Path,
) -> list[Path]:
    root = Path(input_root)
    search_dirs = [
        root.parent / f"{store} - Weekly" / "activity",
        root.parent / f"{store} - Weekly" / "archive",
        root / str(store) / "weekly" / "activity",
        root / str(store) / "weekly" / "archive",
    ]
    candidates: list[Path] = []
    for folder in search_dirs:
        candidates.extend(_activity_file_candidates(folder))
    staged: list[Path] = []
    monthly_activity_dir = Path(monthly_activity_dir)
    monthly_activity_dir.mkdir(parents=True, exist_ok=True)
    for source in _dedupe_preserving_order(candidates):
        report_end = _activity_report_end(source)
        if report_end is None or not (period_start <= report_end <= period_end):
            continue
        destination = monthly_activity_dir / source.name
        saved_path = _copy_if_needed(source, destination)
        if saved_path:
            staged.append(saved_path)
    return staged


def monthly_activity_week_endings(period_start: date, period_end: date) -> list[date]:
    current = period_start
    while current.weekday() != 6:
        current += timedelta(days=1)
    week_endings: list[date] = []
    while current <= period_end:
        week_endings.append(current)
        current += timedelta(days=7)
    return week_endings


def format_monthly_close_preflight(preflight: MonthlyClosePreflight) -> str:
    status = "READY" if preflight.ready else "NOT READY"
    lines = [
        f"Monthly close preflight for store {preflight.store} {preflight.period}: {status}",
        f"Input folder: {preflight.input_dir}",
    ]
    if preflight.staged_activity_paths:
        lines.append("Staged activity files:")
        lines.extend(f"  - {path}" for path in preflight.staged_activity_paths)
    lines.append(f"Gift Card Summary files found: {len(preflight.summary_paths)}")
    lines.append(f"Weekly activity files found for expected weeks: {len(preflight.activity_by_week_end)} of {len(preflight.expected_week_endings)}")
    lines.append(f"Micros export: {preflight.micros_message}")
    if preflight.required_missing_paths:
        lines.append("Missing required paths:")
        lines.extend(f"  - {path}" for path in preflight.required_missing_paths)
    else:
        lines.append("All required monthly-close inputs are present.")
    return "\n".join(lines)


def run_monthly_close(
    *,
    store: str,
    period: str,
    period_start: date,
    period_end: date,
    input_dir: Path,
    output_path: Path,
    micros_path: Path,
    micros_work_dir: Path,
    cleanup_archive_root: Path | None = None,
    fiscal_period: FiscalPeriod | None = None,
    adjust_boundary_weeks: bool = True,
) -> tuple[Path, object, list[WeeklyPosVariance]]:
    summary_path, activity_paths, _pos_path = discover_input_files(input_dir, mode="monthly")
    if summary_path is None:
        raise ParseError("Monthly close requires a Gift Card Summary file.")

    summary = parse_summary(summary_path, store=store)
    activities = [parse_activity_file(path, summary.conversion_promo_codes) for path in activity_paths]
    micros_dir = resolve_micros_export_dir(micros_path, micros_work_dir)
    daily_controls = parse_micros_daily_pos_controls(micros_dir)
    weekly_variances = build_weekly_pos_variances(
        activities,
        summary.conversion_promo_codes,
        daily_controls,
        period_start=period_start,
        period_end=period_end,
        adjust_boundary_weeks=adjust_boundary_weeks,
    )
    pos_controls = PosControls(
        store=str(store),
        period=str(period),
        pos_gift_card_issue=sum((row.pos_issue for row in weekly_variances), Decimal("0.00")),
        pos_gift_card_payment=sum((row.pos_payment for row in weekly_variances), Decimal("0.00")),
    )

    relevant_dates = _activity_report_dates(activities)
    exceptions = validate_tender_payment_totals(micros_dir, daily_controls, relevant_dates)
    exceptions.extend(_coverage_exceptions(weekly_variances))
    result = build_reconciliation(
        store=store,
        period=period,
        period_end=period_end,
        summary=summary,
        activities=activities,
        pos_controls=pos_controls,
        mode="monthly",
        exceptions=exceptions,
    )

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="monthly-close-", dir=str(output_path.parent)) as tmp_dir:
        tmp_path = Path(tmp_dir) / output_path.name
        write_reconciliation_workbook(result, tmp_path)
        append_weekly_pos_variance_detail(tmp_path, weekly_variances, source_label=_source_label(micros_path))
        saved_path = _save_with_locked_workbook_fallback(tmp_path, output_path)
    if cleanup_archive_root is not None:
        cleanup_monthly_close_sources(
            input_dir=input_dir,
            archive_root=cleanup_archive_root,
            store=store,
            fiscal_period=fiscal_period or fiscal_period_for_label(period),
        )
    return saved_path, result, weekly_variances


def resolve_micros_export_dir(micros_path: Path, work_dir: Path) -> Path:
    micros_path = Path(micros_path)
    if micros_path.is_dir():
        return micros_path
    if not micros_path.exists():
        raise ParseError(f"Micros export path not found: {micros_path}")

    suffix = micros_path.suffix.lower()
    if suffix == ".7z":
        seven_zip = _find_7z()
        if seven_zip is None:
            raise ParseError("Could not find 7-Zip. Install 7-Zip or pass an already extracted Micros export folder to --micros-path.")
        output_dir = Path(work_dir) / micros_path.stem
        output_dir.mkdir(parents=True, exist_ok=True)
        subprocess.run([seven_zip, "x", "-y", f"-o{output_dir}", str(micros_path)], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        return output_dir

    if suffix == ".zip":
        import zipfile

        output_dir = Path(work_dir) / micros_path.stem
        output_dir.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(micros_path) as zf:
            zf.extractall(output_dir)
        return output_dir

    raise ParseError(f"Unsupported Micros export path: {micros_path}. Use an extracted folder, .7z, or .zip.")


def parse_micros_daily_pos_controls(micros_dir: Path) -> list[MicrosDailyPosControl]:
    system_path = Path(micros_dir) / SYSTEM_TOTALS_FILE
    if not system_path.exists():
        raise ParseError(f"Micros system totals file not found: {system_path}")

    buckets: dict[date, dict[str, Decimal]] = defaultdict(lambda: {"issue": Decimal("0.00"), "payment": Decimal("0.00")})
    with system_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        for line_no, row in enumerate(reader, start=1):
            if not row:
                continue
            business_date = parse_micros_date(row[0])
            if business_date is None:
                raise ParseError(f"Could not parse Micros date on {SYSTEM_TOTALS_FILE} line {line_no}: {row[0]!r}")
            buckets[business_date]["issue"] += _money_at(row, ISSUE_AMOUNT_INDEX, system_path, line_no)
            buckets[business_date]["payment"] += _money_at(row, PAYMENT_AMOUNT_INDEX, system_path, line_no)

    if not buckets:
        raise ParseError(f"No POS rows parsed from {system_path}")
    return [
        MicrosDailyPosControl(
            business_date=business_date,
            pos_gift_card_issue=money(values["issue"]),
            pos_gift_card_payment=money(values["payment"]),
        )
        for business_date, values in sorted(buckets.items())
    ]


def build_weekly_pos_variances(
    activities: list[ActivityFileData],
    conversion_promo_codes: set[str],
    daily_controls: list[MicrosDailyPosControl],
    *,
    period_start: date,
    period_end: date,
    adjust_boundary_weeks: bool = True,
) -> list[WeeklyPosVariance]:
    daily_by_date = {row.business_date: row for row in daily_controls}
    rows: list[WeeklyPosVariance] = []
    for activity in sorted(activities, key=lambda item: (_activity_bounds(item)[1] or date.max, item.source_file.name)):
        report_begin, report_end = _activity_bounds(activity)
        if report_begin is None or report_end is None:
            raise ParseError(f"Could not determine report date range for {activity.source_file.name}")

        rollup = rollup_activity_file(activity, conversion_promo_codes)
        expected_dates = set(_date_range(report_begin, report_end))
        available_dates = expected_dates & set(daily_by_date)
        missing_dates = expected_dates - available_dates
        missing_inside_period = {d for d in missing_dates if period_start <= d <= period_end}
        unexpected_missing_inside_period = {d for d in missing_inside_period if d.weekday() not in CLOSED_BUSINESS_WEEKDAYS}
        closed_missing_inside_period = missing_inside_period - unexpected_missing_inside_period

        activity_issue = money(rollup.net_activations)
        activity_payment = money(abs(rollup.net_redemptions))
        if missing_dates and adjust_boundary_weeks and not missing_inside_period:
            pos_issue = activity_issue
            pos_payment = activity_payment
            coverage_status = "Boundary week adjusted to activity totals"
        else:
            pos_issue = money(sum((daily_by_date[d].pos_gift_card_issue for d in available_dates), Decimal("0.00")))
            pos_payment = money(sum((daily_by_date[d].pos_gift_card_payment for d in available_dates), Decimal("0.00")))
            if not missing_dates:
                coverage_status = "Full Micros POS coverage"
            elif closed_missing_inside_period and not unexpected_missing_inside_period:
                coverage_status = "Closed days omitted from Micros POS coverage"
            else:
                coverage_status = "Partial Micros POS coverage"

        issue_variance = money(pos_issue - activity_issue)
        payment_variance = money(pos_payment - activity_payment)
        rows.append(
            WeeklyPosVariance(
                week_ending=report_end,
                report_begin=report_begin,
                report_end=report_end,
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
    if not rows:
        raise ParseError("No weekly activity files were available for monthly close.")
    return rows


def validate_tender_payment_totals(
    micros_dir: Path,
    daily_controls: list[MicrosDailyPosControl],
    allowed_dates: set[date] | None = None,
) -> list[tuple[str, str]]:
    tender_path = Path(micros_dir) / TENDER_DETAIL_FILE
    if not tender_path.exists():
        return []

    tender_totals: dict[date, Decimal] = defaultdict(lambda: Decimal("0.00"))
    with tender_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) < 4:
                continue
            tender_name = _clean_micros_text(row[3])
            if tender_name not in GIFT_CARD_PAYMENT_TENDERS:
                continue
            business_date = parse_micros_date(row[0])
            if business_date is None:
                continue
            if allowed_dates is not None and business_date not in allowed_dates:
                continue
            tender_totals[business_date] += money(row[1])

    exceptions: list[tuple[str, str]] = []
    control_by_date = {row.business_date: row.pos_gift_card_payment for row in daily_controls}
    for business_date, tender_total in sorted(tender_totals.items()):
        control_total = control_by_date.get(business_date)
        if control_total is not None and money(control_total - tender_total) != Decimal("0.00"):
            exceptions.append(
                (
                    "Review",
                    f"{business_date:%Y-%m-%d} G C Payment tender total {tender_total:,.2f} does not match {SYSTEM_TOTALS_FILE} column 103 {control_total:,.2f}.",
                )
            )
    return exceptions


def parse_micros_date(value: object) -> date | None:
    text = _clean_micros_text(value)
    if " " in text:
        text = text.split()[0]
    return parse_date(text)


def _clean_micros_text(value: object) -> str:
    return str(value or "").strip().strip("'").strip()


def month_bounds(period: str) -> tuple[date, date]:
    fiscal_period = fiscal_period_for_label(period)
    return fiscal_period.start_date, fiscal_period.end_date


def cleanup_monthly_close_sources(
    *,
    input_dir: Path,
    archive_root: Path,
    store: str,
    fiscal_period: FiscalPeriod,
) -> list[Path]:
    input_dir = Path(input_dir)
    archive_base = Path(archive_root) / "monthly-close" / str(store) / fiscal_period.folder_name
    moved: list[Path] = []
    for source in _monthly_summary_paths(input_dir):
        destination = _move_if_needed(source, archive_base / "summary" / source.name)
        if destination is not None:
            moved.append(destination)
    for source in _dedupe_preserving_order(_activity_file_candidates(input_dir / "activity") + _activity_file_candidates(input_dir)):
        destination = _move_if_needed(source, archive_base / "activity" / source.name)
        if destination is not None:
            moved.append(destination)
    return moved


def _monthly_summary_paths(input_dir: Path) -> list[Path]:
    input_dir = Path(input_dir)
    candidates = sorted((input_dir / "summary").glob("*Gift Card Summary*.xlsx"))
    candidates.extend(sorted(input_dir.glob("*Gift Card Summary*.xlsx")))
    return _dedupe_preserving_order(candidates)


def _stage_monthly_summary_files_for_period(
    *,
    store: str,
    period_end: date,
    store_monthly_dir: Path,
    summary_dir: Path,
) -> list[Path]:
    store_monthly_dir = Path(store_monthly_dir)
    summary_dir = Path(summary_dir)
    if not store_monthly_dir.exists():
        return []
    staged: list[Path] = []
    for source in sorted(store_monthly_dir.glob(f"*{store}*Gift Card Summary*.xlsx")):
        if source.parent == summary_dir:
            continue
        report_end = _summary_report_end(source)
        if report_end != period_end:
            continue
        destination = _move_if_needed(source, summary_dir / source.name)
        if destination is not None:
            staged.append(destination)
    return staged


def _monthly_activity_by_week_end(input_dir: Path) -> dict[date, Path]:
    activity_by_week_end: dict[date, Path] = {}
    for path in _activity_file_candidates(Path(input_dir) / "activity") + _activity_file_candidates(Path(input_dir)):
        report_end = _activity_report_end(path)
        if report_end is None:
            continue
        activity_by_week_end.setdefault(report_end, path)
    return activity_by_week_end


def _activity_file_candidates(folder: Path) -> list[Path]:
    folder = Path(folder)
    if not folder.exists():
        return []
    return sorted(folder.glob("**/*Gift Card Activity*.xls")) + sorted(folder.glob("**/*Gift Card Activity*.xlsx"))


def _activity_report_end(path: Path) -> date | None:
    try:
        activity = parse_activity_file(path)
    except ParseError:
        activity = None
    if activity and activity.report_end:
        return activity.report_end

    match = re.match(r"(\d{2})\.(\d{2})\.(\d{4})\s+", Path(path).name)
    if match:
        return parse_date(f"{match.group(1)}/{match.group(2)}/{match.group(3)}")
    return None


def _summary_report_end(path: Path) -> date | None:
    match = re.match(r"(\d{2})\.(\d{2})\.(\d{4})\s+", Path(path).name)
    if match:
        return parse_date(f"{match.group(1)}/{match.group(2)}/{match.group(3)}")
    return None


def _copy_if_needed(source: Path, destination: Path) -> Path | None:
    source = Path(source)
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        try:
            source_stat = source.stat()
            destination_stat = destination.stat()
        except OSError:
            return None
        if source_stat.st_size == destination_stat.st_size:
            return None
        destination = _available_destination(destination, source)
        if destination.exists() and _same_size(source, destination):
            return None
    shutil.copy2(source, destination)
    return destination


def _move_if_needed(source: Path, destination: Path) -> Path | None:
    source = Path(source)
    destination = Path(destination)
    if not source.exists():
        return None
    destination.parent.mkdir(parents=True, exist_ok=True)
    if source.resolve() == destination.resolve():
        return destination
    if destination.exists() and _same_size(source, destination):
        source.unlink()
        return destination
    if destination.exists():
        destination = _available_destination(destination, source)
        if destination.exists() and _same_size(source, destination):
            source.unlink()
            return destination
    shutil.move(str(source), str(destination))
    return destination


def _available_destination(destination: Path, source: Path) -> Path:
    for idx in range(2, 1000):
        candidate = destination.with_name(f"{destination.stem}_{idx}{destination.suffix}")
        if not candidate.exists() or _same_size(source, candidate):
            return candidate
    raise ParseError(f"Could not find an available archive name for {source.name}.")


def _same_size(left: Path, right: Path) -> bool:
    try:
        return Path(left).stat().st_size == Path(right).stat().st_size
    except OSError:
        return False


def _dedupe_preserving_order(paths: Sequence[Path]) -> list[Path]:
    seen: set[Path] = set()
    result: list[Path] = []
    for path in paths:
        resolved = Path(path).resolve()
        if resolved in seen:
            continue
        seen.add(resolved)
        result.append(Path(path))
    return result


def _micros_preflight(*, micros_path: Path, micros_work_dir: Path, period_end: date) -> tuple[bool, str, list[Path]]:
    micros_path = Path(micros_path)
    if not micros_path.exists():
        return False, f"missing at {micros_path}", [micros_path]
    try:
        micros_dir = resolve_micros_export_dir(micros_path, micros_work_dir)
        daily_controls = parse_micros_daily_pos_controls(micros_dir)
    except ParseError as exc:
        return False, str(exc), [micros_path]

    latest_date = max((row.business_date for row in daily_controls), default=None)
    if latest_date is None:
        return False, f"no business dates found in {micros_path}", [micros_path]
    if latest_date < period_end:
        return False, f"present, but latest POS date is {latest_date:%Y-%m-%d}; expected through {period_end:%Y-%m-%d}", [micros_path]
    return True, f"present through {latest_date:%Y-%m-%d} at {micros_path}", []


def _activity_report_dates(activities: list[ActivityFileData]) -> set[date]:
    dates: set[date] = set()
    for activity in activities:
        report_begin, report_end = _activity_bounds(activity)
        if report_begin is None or report_end is None:
            continue
        dates.update(_date_range(report_begin, report_end))
    return dates


def _activity_bounds(activity: ActivityFileData) -> tuple[date | None, date | None]:
    if activity.report_begin and activity.report_end:
        return activity.report_begin, activity.report_end
    dates = sorted(row.business_date for row in activity.rows if row.business_date is not None)
    if not dates:
        return None, None
    return dates[0], dates[-1]


def _date_range(start: date, end: date) -> Iterable[date]:
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def _money_at(row: list[str], index: int, path: Path, line_no: int) -> Decimal:
    if len(row) <= index:
        raise ParseError(f"{path.name} line {line_no} has {len(row)} columns; expected at least {index + 1}.")
    return money(row[index])


def _coverage_exceptions(rows: list[WeeklyPosVariance]) -> list[tuple[str, str]]:
    exceptions: list[tuple[str, str]] = []
    for row in rows:
        if row.coverage_status == "Partial Micros POS coverage":
            week = row.week_ending.strftime("%m/%d/%Y") if row.week_ending else "Unknown"
            exceptions.append(("Review", f"Week ending {week} has partial Micros POS date coverage. Variance uses available POS dates only."))
    return exceptions


def _find_7z() -> str | None:
    for candidate in [
        shutil.which("7z"),
        shutil.which("7za"),
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
    ]:
        if candidate and Path(candidate).exists():
            return str(candidate)
    return None


def _save_with_locked_workbook_fallback(tmp_path: Path, output_path: Path) -> Path:
    try:
        shutil.copyfile(tmp_path, output_path)
        return output_path
    except PermissionError:
        fallback = _fallback_output_path(output_path)
        shutil.copyfile(tmp_path, fallback)
        return fallback


def _fallback_output_path(output_path: Path) -> Path:
    base = output_path.with_name(f"{output_path.stem}_with_weekly_variance{output_path.suffix}")
    if not base.exists():
        return base
    for idx in range(2, 100):
        candidate = output_path.with_name(f"{output_path.stem}_with_weekly_variance_{idx}{output_path.suffix}")
        if not candidate.exists():
            return candidate
    raise ParseError(f"Could not find an available fallback output name for {output_path}")


def _source_label(micros_path: Path) -> str:
    name = Path(micros_path).name
    return f"{name} / {SYSTEM_TOTALS_FILE}" if name else SYSTEM_TOTALS_FILE


if __name__ == "__main__":
    raise SystemExit(main())
