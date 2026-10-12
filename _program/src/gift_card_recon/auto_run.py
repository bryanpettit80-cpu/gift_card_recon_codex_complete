from __future__ import annotations

import argparse
import re
import uuid
from dataclasses import dataclass, replace
from datetime import date
from pathlib import Path
from typing import Mapping, Sequence

from gift_card_recon.parsers import ParseError, spreadsheet_paths
from gift_card_recon.store_config import (
    STORE_CONFIGS,
    StoreConfig,
    get_operations_store_config,
)
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
    input_set_id: str | None = None
    result_id: str | None = None
    correction_token: str | None = None
    recovery_action: str | None = None
    changed_facts: Mapping | None = None
    run_id: str | None = None
    archive_path: Path | None = None
    pdf_status: str = "not_requested"


def run_weekly_reconciliations(
    *,
    input_root: Path,
    output_dir: Path,
    stores: Sequence[str] | None = None,
    monthly_close_root: Path | None = None,
    archive_root: Path | None = None,
    review_root: Path | None = None,
    operations_root: Path | None = None,
    dropbox_root: Path | None = None,
    approved_corrections: Mapping[str, str] | None = None,
) -> list[AutoRunReport]:
    """Run every discovered store independently; one blocked store never hides another."""

    input_root = Path(input_root)
    operations_root = Path(operations_root) if operations_root is not None else input_root
    dropbox_root = Path(dropbox_root) if dropbox_root is not None else None
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
    input_dirs, issues = weekly_input_inventory(input_root, wanted_stores)
    reports.extend(AutoRunReport(store, input_root, "skipped", message) for store, message in issues)
    for store, input_dir in input_dirs:
        config = get_operations_store_config(
            store,
            operations_root=operations_root,
            dropbox_root=dropbox_root,
        )
        reports.append(
            _run_one_weekly(
                store=store,
                input_dir=input_dir,
                output_dir=output_dir,
                monthly_close_root=monthly_close_root,
                archive_root=archive_root,
                review_root=review_root,
                config=config,
                correction_token=(approved_corrections or {}).get(store),
            )
        )
    run_id = uuid.uuid4().hex
    return sorted((replace(report, run_id=run_id) for report in reports), key=lambda report: report.store)


def _weekly_input_dirs(input_root: Path, wanted_stores: set[str]) -> list[tuple[str, Path]]:
    candidates, issues = weekly_input_inventory(input_root, wanted_stores)
    if issues:
        raise ParseError("; ".join(message for _store, message in issues))
    return candidates


def weekly_input_inventory(input_root: Path, wanted_stores: set[str]) -> tuple[list[tuple[str, Path]], list[tuple[str, str]]]:
    """Discover independently, preserving missing and ambiguous-store failures."""
    input_root = Path(input_root)
    if not input_root.is_dir():
        return [], [(store, f"Weekly input root is missing: {input_root}.")
                    for store in sorted(wanted_stores or STORE_CONFIGS)]
    candidates: dict[str, list[Path]] = {}
    try:
        folders_to_inspect = sorted(path for path in input_root.iterdir() if path.is_dir())
    except OSError as exc:
        return [], [(store, f"Could not read weekly input root {input_root}: {exc}")
                    for store in sorted(wanted_stores or STORE_CONFIGS)]
    for folder in folders_to_inspect:
        store = _store_from_folder(folder)
        if store is None or wanted_stores and store not in wanted_stores:
            continue
        # A legacy store container is not itself a weekly inbox.
        selected = folder / "weekly" if folder.name == store and (folder / "weekly").is_dir() else folder
        candidates.setdefault(store, []).append(selected)
    found: list[tuple[str, Path]] = []
    issues: list[tuple[str, str]] = []
    for store in sorted(set(candidates) | wanted_stores):
        folders = candidates.get(store, [])
        if len(folders) != 1:
            issues.append((store, f"Expected exactly one weekly folder for store {store}; found {len(folders)}: "
                           + ", ".join(str(path) for path in folders)))
        else:
            found.append((store, folders[0]))
    return found, issues


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
    correction_token: str | None = None,
) -> AutoRunReport:
    activity_dir = input_dir if input_dir.name.casefold() == "activity" else input_dir / "activity"
    if not activity_dir.is_dir():
        return AutoRunReport(store, input_dir, "skipped", f"Required Activity folder is missing: {activity_dir}.")
    try:
        activity_paths = _activity_paths(activity_dir)
    except OSError as exc:
        return AutoRunReport(store, input_dir, "skipped", f"Could not read Activity folder {activity_dir}: {exc}")
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
    period = None
    period_end = None
    try:
        # Parse once to establish the canonical destination; the service repeats strict
        # validation before any output is written.
        from gift_card_recon.parsers import parse_activity_file

        activity = parse_activity_file(activity_path)
        if activity.report_end is None:
            raise ParseError(f"Could not determine the week-ending date from {activity_path.name}.")
        period = iso_week_period(activity.report_end)
        period_end = activity.report_end
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
            **({"correction_token": correction_token} if correction_token is not None else {}),
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
                input_set_id=getattr(published, "input_set_id", None),
                result_id=getattr(published, "result_id", None),
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
            input_set_id=getattr(published, "input_set_id", None),
            result_id=getattr(published, "result_id", None),
            archive_path=published.archive_path,
            recovery_action=getattr(published, "recovery_action", None),
        )
    except (ParseError, RuntimeError, OSError, ValueError) as exc:
        return AutoRunReport(store, input_dir, "correction-required" if getattr(exc, "token", None) else "skipped",
                             str(exc), period=period, period_end=period_end,
                             correction_token=getattr(exc, "token", None),
                             changed_facts=getattr(exc, "changes", None),
                             input_set_id=getattr(exc, "input_set_id", None),
                             recovery_action="approve_correction" if getattr(exc, "token", None) else "resolve_input")


def _store_folder_label(config: StoreConfig) -> str:
    return f"{config.store} {config.location_name}"


def _activity_paths(activity_dir: Path) -> list[Path]:
    return spreadsheet_paths(activity_dir)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Run automatic weekly gift card reconciliation from Activity-report inboxes."
    )
    parser.add_argument("--operations-root", default=".", help="Parent operator workspace.")
    parser.add_argument(
        "--dropbox-root",
        default=None,
        help=(
            "Dropbox root containing the external micros_data and GETLinkedData-VB "
            "folders. Defaults to the legacy operations-relative layout."
        ),
    )
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
    parser.add_argument("--approve-correction", action="append", default=[], metavar="STORE=TOKEN",
                        help="Approve the exact correction token shown by a previous run.")
    parser.add_argument("--restore-revision", action="append", default=[], metavar="STORE=PERIOD/REVISION",
                        help="Inspect or restore a retained weekly revision; replacement needs its exact correction token.")
    parser.add_argument("--json", action="store_true", help="Print structured execution results.")
    args = parser.parse_args(argv)
    corrections: dict[str, str] = {}
    for item in args.approve_correction:
        store, separator, token = item.partition("=")
        if not separator or store not in STORE_CONFIGS or not token:
            parser.error("--approve-correction requires STORE=TOKEN for a configured restaurant")
        corrections[store] = token

    restores: list[tuple[str, str, str]] = []
    for item in args.restore_revision:
        store, separator, selection = item.partition("=")
        period, slash, revision = selection.partition("/")
        valid = store in STORE_CONFIGS and separator and slash and re.fullmatch(r"\d{4}-W\d{2}", period)
        valid = valid and re.fullmatch(r"(?:r\d{4,}-[0-9a-f]{16}|[0-9a-f]{64})", revision)
        if valid:
            try:
                date.fromisocalendar(int(period[:4]), int(period[-2:]), 7)
            except ValueError:
                valid = False
        if not valid or any(existing_store == store for existing_store, _, _ in restores):
            parser.error("--restore-revision requires one STORE=YYYY-Www/REVISION selector per restaurant")
        restores.append((store, period, revision))

    operations_root = Path(args.operations_root)
    if restores:
        from gift_card_recon.weekly_service import restore_weekly_revision
        reports = []
        run_id = uuid.uuid4().hex
        for store, period, revision in restores:
            archive_root = _rooted(operations_root, args.archive_root)
            try:
                published = restore_weekly_revision(store=store, period=period, revision_id=revision,
                    output_root=_rooted(operations_root, args.output_dir), archive_root=archive_root,
                    monthly_close_root=_rooted(operations_root, args.monthly_close_root),
                    correction_token=corrections.get(store))
                reports.append(AutoRunReport(store, archive_root,
                    "duplicate" if isinstance(published, WeeklyDuplicate) else "restored", published.message,
                    period=published.period, period_end=published.period_end, output_path=published.output_path,
                    close_status=getattr(published, "status", None),
                    variance_explanation_path=getattr(published, "variance_explanation_path", None),
                    input_set_id=published.input_set_id, result_id=published.result_id, run_id=run_id,
                    archive_path=getattr(published, "archive_path", None),
                    recovery_action=getattr(published, "recovery_action", None)))
            except (ParseError, RuntimeError, OSError, ValueError) as exc:
                token = getattr(exc, "token", None)
                reports.append(AutoRunReport(store, archive_root, "correction-required" if token else "skipped",
                    str(exc), period=period, correction_token=token, changed_facts=getattr(exc, "changes", None),
                    input_set_id=getattr(exc, "input_set_id", None), run_id=run_id,
                    recovery_action="approve_correction" if token else "resolve_input"))
    else:
        reports = run_weekly_reconciliations(
            operations_root=operations_root,
            dropbox_root=Path(args.dropbox_root) if args.dropbox_root else None,
            input_root=_rooted(operations_root, args.input_root),
            output_dir=_rooted(operations_root, args.output_dir),
            stores=args.store,
            monthly_close_root=_rooted(operations_root, args.monthly_close_root),
            archive_root=_rooted(operations_root, args.archive_root),
            review_root=_rooted(operations_root, args.review_root),
            approved_corrections=corrections,
        )
    if args.json:
        import json
        from dataclasses import asdict
        print(json.dumps([asdict(report) for report in reports], default=str, sort_keys=True))
        return int(not reports or any(report.status in {"skipped", "correction-required"} for report in reports))
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
        if report.status in {"created", "restored"}:
            created += 1
        elif report.status in {"skipped", "correction-required"}:
            blocked += 1

    if blocked:
        print(f"\nCompleted independent store processing, but {blocked} store(s) require review.")
        return 1
    if created:
        print(f"\nDone. Published {created} workbook(s) in {Path(args.output_dir)}.")
    else:
        print("\nNothing new to reconcile; inboxes were empty or contained exact duplicates.")
    return 0


def _rooted(operations_root: Path, value: str) -> Path:
    path = Path(value)
    return path if path.is_absolute() else operations_root / path


if __name__ == "__main__":
    raise SystemExit(main())
