from __future__ import annotations

import argparse
import csv
import shutil
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Sequence

from gift_card_recon.excel_writer import write_reconciliation_workbook
from gift_card_recon.models import ActivityFileData
from gift_card_recon.monthly_close import stage_activity_files_for_month
from gift_card_recon.parsers import ParseError, discover_input_files, parse_activity_file, parse_pos_controls, parse_summary
from gift_card_recon.reconcile import build_reconciliation

POS_TOTAL_FIELDS = ["pos_gift_card_issue", "pos_gift_card_payment"]


@dataclass(frozen=True)
class AutoRunReport:
    store: str
    input_dir: Path
    status: str
    message: str
    period: str | None = None
    period_end: date | None = None
    output_path: Path | None = None


@dataclass(frozen=True)
class WeeklyActivitySelection:
    activity_paths: list[Path]
    activities: list[ActivityFileData]
    period_end: date
    archive_candidates: list[tuple[Path, date]]


def run_weekly_reconciliations(
    *,
    input_root: Path,
    output_dir: Path,
    stores: Sequence[str] | None = None,
) -> list[AutoRunReport]:
    input_root = Path(input_root)
    output_dir = Path(output_dir)
    wanted_stores = {str(store).strip() for store in stores or [] if str(store).strip()}

    reports: list[AutoRunReport] = []
    for input_dir in _weekly_input_dirs(input_root, wanted_stores):
        store = input_dir.parent.name
        reports.append(_run_one_weekly(store=store, input_dir=input_dir, output_dir=output_dir))
    return reports


def _weekly_input_dirs(input_root: Path, wanted_stores: set[str]) -> list[Path]:
    if not input_root.exists():
        return []
    candidates = []
    for store_dir in sorted(path for path in input_root.iterdir() if path.is_dir()):
        if wanted_stores and store_dir.name not in wanted_stores:
            continue
        weekly_dir = store_dir / "weekly"
        if weekly_dir.exists():
            candidates.append(weekly_dir)
    return candidates


def _run_one_weekly(*, store: str, input_dir: Path, output_dir: Path) -> AutoRunReport:
    try:
        summary_path, activity_paths, pos_path = discover_input_files(input_dir, mode="weekly")
        if pos_path is None:
            return AutoRunReport(store, input_dir, "skipped", f"Fill in {input_dir / 'pos_controls.csv'} before running.")

        summary = parse_summary(summary_path, store=store) if summary_path else None
        conversion_promo_codes = summary.conversion_promo_codes if summary else set()
        all_activities = [parse_activity_file(path, conversion_promo_codes) for path in activity_paths]
        selection = _select_latest_weekly_activities(activity_paths=activity_paths, activities=all_activities)
        activity_paths = selection.activity_paths
        activities = selection.activities
        period_end = selection.period_end
        period = iso_week_period(period_end)
        pos_controls = parse_pos_controls(pos_path, store=store, period=period)

        result = build_reconciliation(
            store=store,
            period=period,
            period_end=period_end,
            summary=summary,
            activities=activities,
            pos_controls=pos_controls,
            mode="weekly",
        )
        output_path = output_dir / f"Gift_Card_Reconciliation_{store}_{period}.xlsx"
        try:
            write_reconciliation_workbook(result, output_path)
            message = f"Created {output_path.name}"
        except PermissionError:
            output_path = output_dir / f"Gift_Card_Reconciliation_{store}_{period}_{datetime.now():%Y%m%d-%H%M%S}.xlsx"
            write_reconciliation_workbook(result, output_path)
            message = f"Created {output_path.name} because the standard output file is open."
        archive_message = _archive_older_weekly_activity_files(input_dir=input_dir, candidates=selection.archive_candidates)
        if archive_message:
            message = f"{message} {archive_message}"
        stage_message = _stage_monthly_close_activity(input_dir=input_dir, store=store, period_end=period_end, activity_paths=activity_paths)
        if stage_message:
            message = f"{message} {stage_message}"
        clear_message = _clear_pos_controls_after_success(pos_path, store=store, period=period)
        if clear_message:
            message = f"{message} {clear_message}"
        return AutoRunReport(store, input_dir, "created", message, period=period, period_end=period_end, output_path=output_path)
    except ParseError as exc:
        return AutoRunReport(store, input_dir, "skipped", str(exc))
    except RuntimeError as exc:
        return AutoRunReport(store, input_dir, "skipped", str(exc))


def _select_latest_weekly_activities(*, activity_paths: Sequence[Path], activities: Sequence[ActivityFileData]) -> WeeklyActivitySelection:
    dated_activities = [
        (path, activity, _activity_report_end(activity))
        for path, activity in zip(activity_paths, activities, strict=True)
    ]
    if not dated_activities:
        raise ParseError("No Gift Card Activity files were available for weekly reconciliation.")

    latest_report_end = max(report_end for _path, _activity, report_end in dated_activities)
    kept_paths = []
    kept_activities = []
    archive_candidates = []
    for path, activity, report_end in dated_activities:
        if report_end == latest_report_end:
            kept_paths.append(path)
            kept_activities.append(activity)
        else:
            archive_candidates.append((path, report_end))

    return WeeklyActivitySelection(
        activity_paths=kept_paths,
        activities=kept_activities,
        period_end=latest_report_end,
        archive_candidates=archive_candidates,
    )


def _activity_report_end(activity: ActivityFileData) -> date:
    if activity.report_end is not None:
        return activity.report_end
    business_dates = {row.business_date for row in activity.rows if row.business_date is not None}
    if business_dates:
        return max(business_dates)
    raise ParseError(f"Could not determine the week-ending date from {activity.source_file.name}.")


def _archive_older_weekly_activity_files(*, input_dir: Path, candidates: Sequence[tuple[Path, date]]) -> str | None:
    archived = 0
    for source, report_end in candidates:
        if not source.exists():
            continue
        archive_dir = input_dir / "archive" / iso_week_period(report_end)
        archive_dir.mkdir(parents=True, exist_ok=True)
        destination = _archive_destination(source, archive_dir)
        if destination.exists() and _same_size(source, destination):
            source.unlink()
        else:
            shutil.move(str(source), str(destination))
        archived += 1

    if not archived:
        return None
    plural = "s" if archived != 1 else ""
    return f"Archived {archived} older weekly activity file{plural}."


def _archive_destination(source: Path, archive_dir: Path) -> Path:
    destination = archive_dir / source.name
    if not destination.exists() or _same_size(source, destination):
        return destination

    for idx in range(2, 1000):
        candidate = archive_dir / f"{source.stem}_{idx}{source.suffix}"
        if not candidate.exists() or _same_size(source, candidate):
            return candidate
    raise ParseError(f"Could not find an available archive name for {source.name}.")


def _same_size(left: Path, right: Path) -> bool:
    try:
        return Path(left).stat().st_size == Path(right).stat().st_size
    except OSError:
        return False


def _stage_monthly_close_activity(
    *,
    input_dir: Path,
    store: str,
    period_end: date,
    activity_paths: Sequence[Path],
) -> str | None:
    period = f"{period_end:%Y-%m}"
    monthly_activity_dir = input_dir.parent / period / "activity"
    try:
        staged = stage_activity_files_for_month(
            store=store,
            period=period,
            monthly_activity_dir=monthly_activity_dir,
            activity_paths=activity_paths,
        )
    except OSError as exc:
        return f"Could not stage monthly close activity: {exc}"
    if not staged:
        return None
    plural = "s" if len(staged) != 1 else ""
    return f"Staged {len(staged)} activity file{plural} for monthly close in {monthly_activity_dir}."


def _clear_pos_controls_after_success(path: Path, *, store: str, period: str) -> str | None:
    if path.suffix.lower() != ".csv":
        return None
    try:
        cleared = _clear_csv_pos_control_totals(path, store=store, period=period)
    except OSError as exc:
        return f"Could not clear POS totals in {path.name}: {exc}"
    if cleared:
        return f"Cleared POS totals in {path.name}."
    return None


def _clear_csv_pos_control_totals(path: Path, *, store: str, period: str) -> bool:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames
        if not fieldnames or not set(POS_TOTAL_FIELDS).issubset(fieldnames):
            return False
        rows = list(reader)

    if not rows:
        return False

    cleared = False
    for row in rows:
        if not _is_successful_pos_row(row, store=store, period=period, only_row=len(rows) == 1):
            continue
        for field in POS_TOTAL_FIELDS:
            if row.get(field) not in (None, ""):
                row[field] = ""
                cleared = True

    if not cleared:
        return False

    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    return True


def _is_successful_pos_row(row: dict[str, str], *, store: str, period: str, only_row: bool) -> bool:
    if only_row:
        return True
    row_store = str(row.get("store", "")).strip()
    row_period = str(row.get("period", "")).strip()
    return row_store == str(store) and row_period in {str(period), "auto"}


def iso_week_period(period_end: date) -> str:
    iso = period_end.isocalendar()
    return f"{iso.year}-W{iso.week:02d}"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run weekly gift card reconciliations from the simple input folders.")
    parser.add_argument("--input-root", default="input", help="Folder containing store folders.")
    parser.add_argument("--output-dir", default="output", help="Folder for generated reconciliation workbooks.")
    parser.add_argument("--store", action="append", default=None, help="Optional store number to run. Can be repeated.")
    args = parser.parse_args(argv)

    reports = run_weekly_reconciliations(input_root=Path(args.input_root), output_dir=Path(args.output_dir), stores=args.store)
    if not reports:
        print("No weekly store folders found. Use input\\<store>\\weekly\\activity and input\\<store>\\weekly\\pos_controls.csv.")
        return 1

    created = 0
    print("Gift card weekly reconciliation")
    for report in reports:
        ending = f" ending {report.period_end:%m/%d/%Y}" if report.period_end else ""
        period = f" ({report.period})" if report.period else ""
        print(f"- Store {report.store}{period}{ending}: {report.status} - {report.message}")
        if report.status == "created":
            created += 1

    if created:
        print(f"\nDone. Created {created} workbook(s) in {Path(args.output_dir)}.")
        return 0
    print("\nNo workbooks were created. Fill in the skipped item(s), then run again.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
