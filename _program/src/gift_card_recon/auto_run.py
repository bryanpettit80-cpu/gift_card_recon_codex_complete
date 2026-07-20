from __future__ import annotations

import argparse
from dataclasses import dataclass, replace
from datetime import date
from pathlib import Path
from typing import Sequence

from gift_card_recon.parsers import ParseError
from gift_card_recon.store_config import STORE_CONFIGS, StoreConfig, get_store_config
from gift_card_recon.weekly_service import (
    WeeklyDuplicate,
    WeeklyPublication,
    iso_week_period,
    publish_weekly_reconciliation,
)


@dataclass(frozen=True)
class AutoRunReport:
    store: str
    input_dir: Path
    status: str
    message: str
    period: str | None = None
    period_end: date | None = None
    output_path: Path | None = None
    close_status: str | None = None
    variance_explanation_path: Path | None = None


def run_weekly_reconciliations(
    *,
    input_root: Path,
    output_dir: Path,
    stores: Sequence[str] | None = None,
    monthly_close_root: Path | None = None,
    archive_root: Path | None = None,
    review_root: Path | None = None,
    operations_root: Path | None = None,
) -> list[AutoRunReport]:
    """Run every discovered store independently; one blocked store never hides another."""

    input_root = Path(input_root)
    operations_root = Path(operations_root) if operations_root is not None else input_root
    output_dir = Path(output_dir)
    monthly_close_root = (
        Path(monthly_close_root)
        if monthly_close_root is not None
        else operations_root / "02 Monthly Close Inputs"
    )
    archive_root = (
        Path(archive_root)
        if archive_root is not None
        else operations_root / "04 Archive" / "Weekly Reconciliation"
    )
    review_root = (
        Path(review_root)
        if review_root is not None
        else operations_root / "_automation_runs" / "review"
    )
    wanted_stores = {str(store).strip() for store in stores or [] if str(store).strip()}

    reports: list[AutoRunReport] = []
    for store, input_dir in _weekly_input_dirs(input_root, wanted_stores):
        config = _operations_store_config(store, operations_root)
        reports.append(
            _run_one_weekly(
                store=store,
                input_dir=input_dir,
                output_dir=output_dir,
                monthly_close_root=monthly_close_root,
                archive_root=archive_root,
                review_root=review_root,
                config=config,
            )
        )
    return reports


def _weekly_input_dirs(input_root: Path, wanted_stores: set[str]) -> list[tuple[str, Path]]:
    if not input_root.exists():
        return []
    candidates: dict[str, Path] = {}
    for folder in sorted(path for path in input_root.iterdir() if path.is_dir()):
        store = _store_from_folder(folder)
        if store is None or wanted_stores and store not in wanted_stores:
            continue
        candidates.setdefault(store, folder)

    # Preserve compatibility with the low-level store/weekly layout.
    for store in sorted(STORE_CONFIGS):
        if wanted_stores and store not in wanted_stores:
            continue
        weekly_dir = input_root / store / "weekly"
        if weekly_dir.exists():
            candidates.setdefault(store, weekly_dir)
    return sorted(candidates.items())


def _store_from_folder(path: Path) -> str | None:
    leading = path.name.split(maxsplit=1)[0]
    if leading in STORE_CONFIGS:
        return leading
    suffix = " - Weekly"
    if path.name.endswith(suffix):
        store = path.name[: -len(suffix)].strip()
        return store if store in STORE_CONFIGS else None
    return None


def _run_one_weekly(
    *,
    store: str,
    input_dir: Path,
    output_dir: Path,
    monthly_close_root: Path,
    archive_root: Path,
    review_root: Path,
    config: StoreConfig,
) -> AutoRunReport:
    activity_dir = input_dir if input_dir.name.casefold() == "activity" else input_dir / "activity"
    activity_paths = _activity_paths(activity_dir)
    if not activity_paths:
        return AutoRunReport(store, input_dir, "no-op", f"No Activity report is waiting in {activity_dir}.")
    if len(activity_paths) != 1:
        file_list = "\n".join(f"  - {path.name}" for path in activity_paths)
        return AutoRunReport(
            store,
            input_dir,
            "skipped",
            f"Expected exactly one weekly Gift Card Activity file in {activity_dir}. "
            f"Found {len(activity_paths)}.\n{file_list}",
        )

    activity_path = activity_paths[0]
    try:
        # Parse once to establish the canonical destination; the service repeats strict
        # validation before any output is written.
        from gift_card_recon.parsers import parse_activity_file

        activity = parse_activity_file(activity_path)
        if activity.report_end is None:
            raise ParseError(f"Could not determine the week-ending date from {activity_path.name}.")
        period = iso_week_period(activity.report_end)
        store_folder = _store_folder_label(config)
        output_path = (
            output_dir
            / store_folder
            / str(activity.report_end.isocalendar().year)
            / f"Gift_Card_Reconciliation_{store}_{period}.xlsx"
        )
        published = publish_weekly_reconciliation(
            store=store,
            activity_path=activity_path,
            config=config,
            output_path=output_path,
            archive_root=archive_root,
            monthly_close_root=monthly_close_root,
            review_root=review_root,
        )
        if isinstance(published, WeeklyDuplicate):
            return AutoRunReport(
                store,
                input_dir,
                "duplicate",
                published.message,
                period=published.period,
                period_end=published.period_end,
                output_path=published.output_path,
            )
        assert isinstance(published, WeeklyPublication)
        return AutoRunReport(
            store,
            input_dir,
            "created",
            published.message,
            period=published.period,
            period_end=published.period_end,
            output_path=published.output_path,
            close_status=published.status,
            variance_explanation_path=published.variance_explanation_path,
        )
    except (ParseError, RuntimeError, OSError, ValueError) as exc:
        return AutoRunReport(store, input_dir, "skipped", str(exc))


def _operations_store_config(store: str, operations_root: Path) -> StoreConfig:
    config = get_store_config(store)
    micros_path = config.micros_default_path
    if not micros_path.is_absolute():
        micros_path = (Path(operations_root) / micros_path).resolve()
    return replace(config, micros_default_path=micros_path)


def _store_folder_label(config: StoreConfig) -> str:
    return f"{config.store} {config.location_name}"


def _activity_paths(activity_dir: Path) -> list[Path]:
    if not activity_dir.is_dir():
        return []
    candidates = [
        path
        for path in activity_dir.iterdir()
        if path.is_file()
        and path.suffix.casefold() in {".xls", ".xlsx"}
    ]
    return sorted(candidates)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Run automatic weekly gift card reconciliation from Activity-report inboxes."
    )
    parser.add_argument("--operations-root", default=".", help="Parent operator workspace.")
    parser.add_argument(
        "--input-root",
        default="01 Weekly Gift Card Activity Reports",
        help="Folder containing store Activity-report inboxes.",
    )
    parser.add_argument(
        "--output-dir",
        default=str(Path("03 Finished Reports") / "Weekly"),
        help="Root folder for canonical weekly workbooks.",
    )
    parser.add_argument(
        "--monthly-close-root",
        default="02 Monthly Close Inputs",
        help="Root folder where verified Activity reports are copied for monthly close.",
    )
    parser.add_argument(
        "--archive-root",
        default=str(Path("04 Archive") / "Weekly Reconciliation"),
        help="Root folder for immutable weekly evidence packages.",
    )
    parser.add_argument(
        "--review-root",
        default=str(Path("_automation_runs") / "review"),
        help="Runtime review folder, including duplicate-input quarantine.",
    )
    parser.add_argument("--store", action="append", default=None, help="Optional store number to run. Can be repeated.")
    args = parser.parse_args(argv)

    operations_root = Path(args.operations_root)
    reports = run_weekly_reconciliations(
        operations_root=operations_root,
        input_root=_rooted(operations_root, args.input_root),
        output_dir=_rooted(operations_root, args.output_dir),
        stores=args.store,
        monthly_close_root=_rooted(operations_root, args.monthly_close_root),
        archive_root=_rooted(operations_root, args.archive_root),
        review_root=_rooted(operations_root, args.review_root),
    )
    if not reports:
        print("No weekly store inboxes found.")
        return 1

    created = 0
    blocked = 0
    print("Gift card weekly reconciliation")
    for report in reports:
        ending = f" ending {report.period_end:%m/%d/%Y}" if report.period_end else ""
        period = f" ({report.period})" if report.period else ""
        close_status = f" [{report.close_status}]" if report.close_status else ""
        print(f"- Store {report.store}{period}{ending}: {report.status}{close_status} - {report.message}")
        if report.status == "created":
            created += 1
        elif report.status == "skipped":
            blocked += 1

    if blocked:
        print(f"\nCompleted independent store processing, but {blocked} store(s) require review.")
        return 1
    if created:
        print(f"\nDone. Created {created} workbook(s) in {Path(args.output_dir)}.")
    else:
        print("\nNothing new to reconcile; inboxes were empty or contained exact duplicates.")
    return 0


def _rooted(operations_root: Path, value: str) -> Path:
    path = Path(value)
    return path if path.is_absolute() else operations_root / path


if __name__ == "__main__":
    raise SystemExit(main())
