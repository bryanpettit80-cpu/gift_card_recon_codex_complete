from __future__ import annotations

import argparse
import re
import shutil
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Sequence

from gift_card_recon.excel_io import ExcelReadError
from gift_card_recon.parsers import ParseError, parse_activity_file


ACTIVITY_PATTERN = "*Gift Card Activity*.xls*"


@dataclass(frozen=True)
class ActivityImportReport:
    source_path: Path
    status: str
    message: str
    store: str | None = None
    report_end: date | None = None
    monthly_path: Path | None = None
    weekly_path: Path | None = None


@dataclass(frozen=True)
class ValidActivityDownload:
    source_path: Path
    store: str
    report_end: date
    canonical_name: str
    monthly_path: Path
    monthly_changed: bool
    weekly_path: Path | None = None
    weekly_changed: bool = False
    weekly_message: str | None = None


def import_gmail_activity_downloads(
    *,
    source_dir: Path,
    input_root: Path,
    stores: Sequence[str] | None = None,
    update_weekly: bool = True,
) -> list[ActivityImportReport]:
    """Import selected Gmail-downloaded Gift Card Activity attachments.

    The workflow intentionally does not connect to Gmail. Operators download the
    specific RPA Bot activity attachments into source_dir, then this function
    validates and places them in the existing local folder structure.
    """

    source_dir = Path(source_dir)
    input_root = Path(input_root)
    allowed_stores = _allowed_stores(input_root, stores)

    if not source_dir.exists():
        return [
            ActivityImportReport(
                source_path=source_dir,
                status="skipped",
                message=f"Source folder not found: {source_dir}",
            )
        ]

    reports: list[ActivityImportReport] = []
    valid_downloads: list[ValidActivityDownload] = []
    for source in _activity_file_candidates(source_dir):
        validated, report = _validate_and_stage_monthly(source, input_root=input_root, allowed_stores=allowed_stores)
        if validated is None:
            reports.append(report)
        else:
            valid_downloads.append(validated)

    weekly_updates = _stage_latest_weekly_files(valid_downloads, input_root=input_root) if update_weekly else {}
    for download in valid_downloads:
        weekly_path, weekly_changed, weekly_message = weekly_updates.get(
            download.source_path.resolve(),
            (None, False, "Weekly activity was not updated."),
        )
        monthly_message = "copied" if download.monthly_changed else "already present"
        message = f"Monthly close {monthly_message}: {download.monthly_path}"
        if weekly_message:
            message = f"{message}. {weekly_message}"
        reports.append(
            ActivityImportReport(
                source_path=download.source_path,
                status="imported",
                message=message,
                store=download.store,
                report_end=download.report_end,
                monthly_path=download.monthly_path,
                weekly_path=weekly_path if weekly_changed or weekly_path else None,
            )
        )

    if not reports:
        reports.append(
            ActivityImportReport(
                source_path=source_dir,
                status="skipped",
                message=f"No Gift Card Activity .xls/.xlsx files found in {source_dir}",
            )
        )
    return sorted(reports, key=_report_sort_key)


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    reports = import_gmail_activity_downloads(
        source_dir=Path(args.source_dir),
        input_root=Path(args.input_root),
        stores=args.store,
        update_weekly=not args.no_weekly,
    )

    print("Gmail gift card activity import")
    print(f"Source folder: {Path(args.source_dir)}")
    imported = 0
    for report in reports:
        store = f" Store {report.store}" if report.store else ""
        ending = f" ending {report.report_end:%m/%d/%Y}" if report.report_end else ""
        print(f"-{store}{ending}: {report.status} - {report.message}")
        if report.status == "imported":
            imported += 1

    if imported:
        print(f"\nDone. Imported {imported} activity file(s).")
        return 0
    print("\nNo activity files were imported.")
    return 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="gift-card-import-gmail-activity",
        description="Import selected Gmail-downloaded RPA Bot Gift Card Activity attachments into the local input folders.",
    )
    parser.add_argument("--source-dir", default="input/gmail_activity", help="Folder containing downloaded Gmail activity attachments.")
    parser.add_argument("--input-root", default="input", help="Folder containing store input folders.")
    parser.add_argument("--store", action="append", default=None, help="Optional store number to allow. Can be repeated.")
    parser.add_argument("--no-weekly", action="store_true", help="Only copy files to monthly-close folders; do not update weekly/activity.")
    return parser


def _validate_and_stage_monthly(
    source: Path,
    *,
    input_root: Path,
    allowed_stores: set[str],
) -> tuple[ValidActivityDownload | None, ActivityImportReport]:
    store = _store_from_filename(source.name)
    if store is None:
        return None, ActivityImportReport(source, "skipped", "Could not find a store number in the filename.")
    if allowed_stores and store not in allowed_stores:
        return None, ActivityImportReport(source, "skipped", f"Store {store} is not one of the configured stores.", store=store)

    try:
        activity = parse_activity_file(source)
    except (ExcelReadError, OSError, ParseError) as exc:
        return None, ActivityImportReport(source, "skipped", f"Could not read activity file: {exc}", store=store)

    report_end = activity.report_end or max((row.business_date for row in activity.rows if row.business_date), default=None)
    if report_end is None:
        return None, ActivityImportReport(source, "skipped", "Could not determine the report ending date.", store=store)

    suffix = ".xlsx" if source.suffix.lower() == ".xlsx" else ".xls"
    canonical_name = f"{report_end:%m.%d.%Y} {store} Gift Card Activity{suffix}"
    monthly_path = input_root / store / f"{report_end:%Y-%m}" / "activity" / canonical_name
    monthly_changed = _copy_if_needed(source, monthly_path)
    return (
        ValidActivityDownload(
            source_path=source,
            store=store,
            report_end=report_end,
            canonical_name=canonical_name,
            monthly_path=monthly_path,
            monthly_changed=monthly_changed,
        ),
        ActivityImportReport(source, "imported", "", store=store, report_end=report_end, monthly_path=monthly_path),
    )


def _stage_latest_weekly_files(
    downloads: Sequence[ValidActivityDownload],
    *,
    input_root: Path,
) -> dict[Path, tuple[Path | None, bool, str]]:
    updates: dict[Path, tuple[Path | None, bool, str]] = {}
    by_store: dict[str, list[ValidActivityDownload]] = {}
    for download in downloads:
        by_store.setdefault(download.store, []).append(download)

    for store, store_downloads in by_store.items():
        weekly_activity_dir = input_root / store / "weekly" / "activity"
        current_latest = _latest_report_end(_activity_file_candidates(weekly_activity_dir))
        latest_download = max(store_downloads, key=lambda item: (item.report_end, item.source_path.name))

        if current_latest and latest_download.report_end < current_latest:
            message = f"Weekly activity left unchanged; it already has a newer week ending {current_latest:%m/%d/%Y}."
            for download in store_downloads:
                updates[download.source_path.resolve()] = (None, False, message)
            continue

        archived_count = _archive_weekly_files_except(weekly_activity_dir, keep_report_end=latest_download.report_end)
        destination = weekly_activity_dir / latest_download.canonical_name
        changed = _copy_if_needed(latest_download.source_path, destination)
        action = "copied" if changed else "already present"
        archive_message = f" Archived {archived_count} older weekly file(s)." if archived_count else ""
        latest_message = f"Weekly activity {action}: {destination}.{archive_message}"
        for download in store_downloads:
            if download.source_path.resolve() == latest_download.source_path.resolve():
                updates[download.source_path.resolve()] = (destination, changed, latest_message)
            else:
                updates[download.source_path.resolve()] = (
                    None,
                    False,
                    f"Weekly activity left on latest imported week ending {latest_download.report_end:%m/%d/%Y}.",
                )
    return updates


def _activity_file_candidates(folder: Path) -> list[Path]:
    folder = Path(folder)
    if not folder.exists():
        return []
    return sorted(path for path in folder.glob(f"**/{ACTIVITY_PATTERN}") if path.is_file())


def _allowed_stores(input_root: Path, stores: Sequence[str] | None) -> set[str]:
    explicit = {str(store).strip() for store in stores or [] if str(store).strip()}
    if explicit:
        return explicit
    root = Path(input_root)
    if not root.exists():
        return set()
    return {path.name for path in root.iterdir() if path.is_dir() and path.name.isdigit()}


def _store_from_filename(filename: str) -> str | None:
    match = re.search(r"(?:^|\s)(\d{4})(?=\s+Gift Card Activity\b)", filename, flags=re.IGNORECASE)
    return match.group(1) if match else None


def _latest_report_end(paths: Sequence[Path]) -> date | None:
    dates = [_report_end_from_file(path) for path in paths]
    dates = [value for value in dates if value is not None]
    return max(dates, default=None)


def _archive_weekly_files_except(weekly_activity_dir: Path, *, keep_report_end: date) -> int:
    archived = 0
    for source in _activity_file_candidates(weekly_activity_dir):
        report_end = _report_end_from_file(source)
        if report_end is None or report_end == keep_report_end:
            continue
        archive_dir = weekly_activity_dir.parent / "archive" / _iso_week_period(report_end)
        archive_dir.mkdir(parents=True, exist_ok=True)
        destination = archive_dir / source.name
        if destination.exists() and _same_size(source, destination):
            source.unlink()
        else:
            shutil.move(str(source), str(destination))
        archived += 1
    return archived


def _report_end_from_file(path: Path) -> date | None:
    try:
        activity = parse_activity_file(path)
    except (ExcelReadError, OSError, ParseError):
        activity = None
    if activity and activity.report_end:
        return activity.report_end

    match = re.match(r"(\d{2})\.(\d{2})\.(\d{4})\s+", path.name)
    if not match:
        return None
    return date(int(match.group(3)), int(match.group(1)), int(match.group(2)))


def _copy_if_needed(source: Path, destination: Path) -> bool:
    source = Path(source)
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    if source.resolve() == destination.resolve():
        return False
    if destination.exists() and _same_size(source, destination):
        return False
    shutil.copy2(source, destination)
    return True


def _same_size(left: Path, right: Path) -> bool:
    try:
        return Path(left).stat().st_size == Path(right).stat().st_size
    except OSError:
        return False


def _iso_week_period(value: date) -> str:
    iso = value.isocalendar()
    return f"{iso.year}-W{iso.week:02d}"


def _report_sort_key(report: ActivityImportReport) -> tuple[str, date, str]:
    return (report.store or "", report.report_end or date.min, str(report.source_path))


if __name__ == "__main__":
    raise SystemExit(main())
