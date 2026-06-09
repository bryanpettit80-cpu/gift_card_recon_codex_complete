from __future__ import annotations

import argparse
import csv
import re
import shutil
import subprocess
import tempfile
from collections import defaultdict
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path
from typing import Iterable

from gift_card_recon.excel_writer import append_weekly_pos_variance_detail, write_reconciliation_workbook
from gift_card_recon.models import ActivityFileData, MicrosDailyPosControl, PosControls, WeeklyPosVariance
from gift_card_recon.parsers import ParseError, discover_input_files, parse_activity_file, parse_summary
from gift_card_recon.reconcile import build_reconciliation, rollup_activity_file
from gift_card_recon.utils import money, parse_date

SYSTEM_TOTALS_FILE = "DLYSYSTT.TXT"
TENDER_DETAIL_FILE = "TENDER_DETAIL.TXT"

# Zero-based indexes in DLYSYSTT.TXT. These are the validated 1-based columns
# 121 and 103 in the Micros 3700 system totals export used for store 9355.
ISSUE_AMOUNT_INDEX = 120
PAYMENT_AMOUNT_INDEX = 102


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    store = str(args.store)
    period = str(args.period)
    try:
        period_start, default_period_end = month_bounds(period)
    except ParseError as exc:
        raise SystemExit(str(exc)) from exc
    period_end = parse_date(args.period_end) if args.period_end else default_period_end
    if period_end is None:
        raise SystemExit(f"Could not parse --period-end value: {args.period_end!r}")

    input_dir = Path(args.input_dir) if args.input_dir else Path("input") / store / period
    output_dir = Path(args.output_dir)
    output_path = Path(args.output_file) if args.output_file else output_dir / f"Gift_Card_Reconciliation_{store}_{period}.xlsx"

    try:
        saved_path, result, weekly_variances = run_monthly_close(
            store=store,
            period=period,
            period_start=period_start,
            period_end=period_end,
            input_dir=input_dir,
            output_path=output_path,
            micros_path=Path(args.micros_path),
            micros_work_dir=Path(args.micros_work_dir),
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
    parser.add_argument("--period", default="2026-06", help="Monthly accounting period in YYYY-MM format. Defaults to 2026-06.")
    parser.add_argument("--period-end", default=None, help="Optional period end date. Defaults to the last day of --period.")
    parser.add_argument("--input-dir", default=None, help="Input folder containing summary/ and activity/. Defaults to input/<store>/<period>.")
    parser.add_argument("--output-dir", default="output", help="Folder for the generated workbook.")
    parser.add_argument("--output-file", default=None, help="Optional explicit output .xlsx path.")
    parser.add_argument("--micros-path", default="_inspect_micros3700", help="Micros export folder or .7z archive. Defaults to _inspect_micros3700.")
    parser.add_argument("--micros-work-dir", default="tmp/monthly_close_micros", help="Extraction folder used when --micros-path is an archive.")
    parser.add_argument(
        "--no-boundary-adjustment",
        action="store_true",
        help="Do not hold boundary weeks to activity totals when Micros dates do not cover dates outside the monthly period.",
    )
    return parser


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

        activity_issue = money(rollup.net_activations)
        activity_payment = money(abs(rollup.net_redemptions))
        if missing_dates and adjust_boundary_weeks and not missing_inside_period:
            pos_issue = activity_issue
            pos_payment = activity_payment
            coverage_status = "Boundary week adjusted to activity totals"
        else:
            pos_issue = money(sum((daily_by_date[d].pos_gift_card_issue for d in available_dates), Decimal("0.00")))
            pos_payment = money(sum((daily_by_date[d].pos_gift_card_payment for d in available_dates), Decimal("0.00")))
            coverage_status = "Full Micros POS coverage" if not missing_dates else "Partial Micros POS coverage"

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
            tender_name = str(row[3]).strip()
            if tender_name not in {"G C Payment", "Gift Card Payment"}:
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
    text = str(value or "").strip().strip("'")
    if " " in text:
        text = text.split()[0]
    return parse_date(text)


def month_bounds(period: str) -> tuple[date, date]:
    match = re.fullmatch(r"(\d{4})-(\d{2})", str(period).strip())
    if not match:
        raise ParseError(f"Monthly close period must use YYYY-MM format: {period!r}")
    year = int(match.group(1))
    month = int(match.group(2))
    start = date(year, month, 1)
    if month == 12:
        next_month = date(year + 1, 1, 1)
    else:
        next_month = date(year, month + 1, 1)
    return start, next_month - timedelta(days=1)


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
