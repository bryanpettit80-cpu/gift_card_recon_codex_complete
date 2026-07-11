from __future__ import annotations

import argparse
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence

from gift_card_recon.darden import parse_darden_credit_memo
from gift_card_recon.fiscal_calendar import FiscalPeriod, fiscal_period_for_date, fiscal_period_for_label
from gift_card_recon.models import DardenCreditMemo
from gift_card_recon.monthly_close_service import (
    CloseBlockedError,
    MonthlyCloseRunResult,
    assess_monthly_close_inputs,
    run_monthly_close_service,
    write_monthly_close_diagnostic,
)
from gift_card_recon.parsers import ParseError
from gift_card_recon.store_config import get_store_config


SHARED_DARDEN_INBOX = "Darden Reports - Drop Here"


@dataclass(frozen=True)
class CloseJob:
    store: str
    fiscal_period: FiscalPeriod
    darden_path: Path
    darden_report: DardenCreditMemo


@dataclass(frozen=True)
class DiscoveryIssue:
    message: str
    store: str | None = None
    period: str | None = None

    def __str__(self) -> str:
        return self.message


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="gift-card-monthly-close",
        description=(
            "Close all Darden reports in the shared inbox, or rerun one explicit store-period."
        ),
    )
    parser.add_argument("--store", default=None, help="Optional store number for a single rerun.")
    parser.add_argument("--period", default=None, help="Optional fiscal period, such as FY27-M01.")
    parser.add_argument("--input-root", default="Monthly Close", help="Monthly-close input root.")
    parser.add_argument("--input-dir", default=None, help="Explicit single-job Summary/activity folder.")
    parser.add_argument("--output-dir", default="Output", help="Canonical output root.")
    parser.add_argument("--output-file", default=None, help="Explicit single-job workbook path.")
    parser.add_argument("--darden-path", default=None, help="Explicit Darden PDF for a single rerun.")
    parser.add_argument("--micros-path", default=None, help="Explicit Micros folder or archive for a single rerun.")
    parser.add_argument(
        "--micros-work-dir",
        default=str(Path.home() / "AppData" / "Local" / "GiftCardRecon" / "micros-extract"),
        help="Local extraction workspace for Micros archives.",
    )
    parser.add_argument("--archive-root", default="Archive - Old Files", help="Evidence archive root.")
    parser.add_argument("--prepare-only", action="store_true", help="Stage weekly inputs and report readiness without closing.")
    parser.add_argument("--no-stage-weekly", action="store_true", help="Do not stage weekly activity reports.")
    parser.add_argument("--no-cleanup", action="store_true", help="Retain live source files after verified publication.")
    parser.add_argument("--no-boundary-adjustment", action="store_true", help=argparse.SUPPRESS)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    input_root = Path(args.input_root)
    archive_root = Path(args.archive_root)
    output_root = Path(args.output_dir)
    inbox = input_root / SHARED_DARDEN_INBOX
    inbox.mkdir(parents=True, exist_ok=True)

    try:
        jobs, discovery_errors = discover_close_jobs(
            inbox=inbox,
            explicit_darden=Path(args.darden_path) if args.darden_path else None,
            store=args.store,
            period=args.period,
        )
    except (ParseError, OSError, ValueError, RuntimeError) as exc:
        print(f"REVIEW REQUIRED: {exc}")
        return 1

    if discovery_errors:
        for issue in discovery_errors:
            print(f"REVIEW REQUIRED: {issue.message}")
            if issue.store and issue.period:
                try:
                    workbook, pdf = write_monthly_close_diagnostic(
                        store=issue.store,
                        period=issue.period,
                        output_root=output_root,
                        message=issue.message,
                    )
                    print(f"Diagnostic workbook: {workbook.resolve()}")
                    if pdf is not None:
                        print(f"Diagnostic PDF: {pdf.resolve()}")
                except (OSError, RuntimeError, ValueError):
                    pass
    if not jobs:
        if not discovery_errors:
            print(f"No Darden PDF files were found. Add them to: {inbox.resolve()}")
        return 1
    if len(jobs) > 1 and any((args.input_dir, args.output_file, args.micros_path)):
        print("REVIEW REQUIRED: explicit input, output, and Micros paths can be used only for a single job.")
        return 1

    failures = len(discovery_errors)
    for job in jobs:
        try:
            input_dir = _resolve_input_dir(
                job,
                input_root=input_root,
                archive_root=archive_root,
                explicit_input=Path(args.input_dir) if args.input_dir else None,
                stage_weekly=not args.no_stage_weekly,
            )
            if args.prepare_only:
                config = get_store_config(job.store)
                micros_path = Path(args.micros_path) if args.micros_path else config.micros_default_path
                assessment = assess_monthly_close_inputs(
                    store=job.store,
                    period=job.fiscal_period.period_key,
                    input_dir=input_dir,
                    micros_path=micros_path,
                    micros_work_dir=Path(args.micros_work_dir),
                    archive_root=archive_root,
                    darden_path=job.darden_path,
                    fiscal_period=job.fiscal_period,
                    allow_unconfigured_micros=False,
                )
                print(f"{job.store} {job.fiscal_period.period_key}: {assessment.status.value}")
                if not assessment.can_publish_close:
                    failures += 1
                continue
            config = get_store_config(job.store)
            micros_path = Path(args.micros_path) if args.micros_path else config.micros_default_path
            result = run_monthly_close_service(
                store=job.store,
                period=job.fiscal_period.period_key,
                input_dir=input_dir,
                micros_path=micros_path,
                micros_work_dir=Path(args.micros_work_dir),
                archive_root=archive_root,
                output_root=output_root,
                output_path=Path(args.output_file) if args.output_file else None,
                darden_path=job.darden_path,
                fiscal_period=job.fiscal_period,
                cleanup_sources=not args.no_cleanup,
                allow_unconfigured_micros=False,
            )
            _print_success(result)
        except CloseBlockedError as exc:
            failures += 1
            print(f"{job.store} {job.fiscal_period.period_key}: REVIEW REQUIRED - {exc}")
            if exc.review_workbook is not None:
                print(f"Diagnostic workbook: {exc.review_workbook.resolve()}")
            if exc.review_pdf is not None:
                print(f"Diagnostic PDF: {exc.review_pdf.resolve()}")
        except (ParseError, OSError, ValueError, RuntimeError) as exc:
            failures += 1
            print(f"{job.store} {job.fiscal_period.period_key}: REVIEW REQUIRED - {exc}")
    return 1 if failures else 0


def discover_close_jobs(
    *,
    inbox: Path,
    explicit_darden: Path | None,
    store: str | None,
    period: str | None,
) -> tuple[list[CloseJob], list[DiscoveryIssue]]:
    candidates = [Path(explicit_darden)] if explicit_darden is not None else sorted(Path(inbox).glob("*.pdf"))
    parsed: list[CloseJob] = []
    errors: list[DiscoveryIssue] = []
    for path in candidates:
        try:
            report = parse_darden_credit_memo(path)
            fiscal_period = fiscal_period_for_date(report.period_end)
            if (
                report.period_start != fiscal_period.start_date
                or report.period_end != fiscal_period.end_date
            ):
                raise ParseError(
                    f"{path.name} service dates do not exactly match {fiscal_period.period_key}."
                )
            config = get_store_config(report.store)
            if store is not None and config.store != str(store):
                if explicit_darden is None:
                    continue
                raise ParseError(
                    f"{path.name} is for store {config.store}; explicit store {store} was requested."
                )
            if period is not None:
                requested = fiscal_period_for_label(period)
                if requested.period_key != fiscal_period.period_key:
                    if explicit_darden is None:
                        continue
                    raise ParseError(
                        f"{path.name} is for {fiscal_period.period_key}; {requested.period_key} was requested."
                    )
            parsed.append(
                CloseJob(
                    store=config.store,
                    fiscal_period=fiscal_period,
                    darden_path=path,
                    darden_report=report,
                )
            )
        except (ParseError, OSError, ValueError) as exc:
            known_period = None
            if period is not None:
                try:
                    known_period = fiscal_period_for_label(period).period_key
                except ParseError:
                    known_period = None
            errors.append(
                DiscoveryIssue(
                    f"{path.name}: {exc}",
                    store=str(store) if store is not None else None,
                    period=known_period,
                )
            )

    grouped: dict[tuple[str, str], list[CloseJob]] = defaultdict(list)
    for job in parsed:
        grouped[(job.store, job.fiscal_period.period_key)].append(job)
    jobs: list[CloseJob] = []
    for (job_store, job_period), matches in sorted(grouped.items()):
        if len(matches) != 1:
            names = ", ".join(item.darden_path.name for item in matches)
            errors.append(
                DiscoveryIssue(
                    f"Store {job_store} {job_period} has {len(matches)} Darden PDFs ({names}); exactly one is required.",
                    store=job_store,
                    period=job_period,
                )
            )
            continue
        jobs.append(matches[0])
    return jobs, errors


def _resolve_input_dir(
    job: CloseJob,
    *,
    input_root: Path,
    archive_root: Path,
    explicit_input: Path | None,
    stage_weekly: bool,
) -> Path:
    if explicit_input is not None:
        return explicit_input
    live = input_root / job.store / job.fiscal_period.folder_name
    live_has_material = _contains_any_close_input(live) or bool(
        list(live.parent.glob(f"*{job.store}*Gift Card Summary*.xlsx"))
    )
    if live_has_material:
        _stage_live_inputs(job, live=live, input_root=input_root, stage_weekly=stage_weekly)
        if _contains_close_inputs(live):
            return live

    # Archived evidence is never selected implicitly. A new inbox PDF must use
    # the current live month; completed evidence is available only through an
    # explicit --input-dir rerun so stale sources cannot be greenlit.
    del archive_root
    _stage_live_inputs(job, live=live, input_root=input_root, stage_weekly=stage_weekly)
    return live


def _stage_live_inputs(
    job: CloseJob,
    *,
    live: Path,
    input_root: Path,
    stage_weekly: bool,
) -> None:
    live.mkdir(parents=True, exist_ok=True)
    (live / "summary").mkdir(exist_ok=True)
    (live / "activity").mkdir(exist_ok=True)

    from gift_card_recon.monthly_close import (
        _stage_monthly_summary_files_for_period,
        stage_weekly_activity_files_for_month,
    )

    _stage_monthly_summary_files_for_period(
        store=job.store,
        period_end=job.fiscal_period.end_date,
        store_monthly_dir=live.parent,
        summary_dir=live / "summary",
    )
    if stage_weekly:
        stage_weekly_activity_files_for_month(
            store=job.store,
            period=job.fiscal_period.period_key,
            fiscal_period=job.fiscal_period,
            period_start=job.fiscal_period.start_date,
            period_end=job.fiscal_period.end_date,
            input_root=input_root,
            monthly_activity_dir=live / "activity",
        )


def _contains_close_inputs(path: Path) -> bool:
    return (
        len(list((path / "summary").glob("*Gift Card Summary*.xlsx"))) == 1
        and bool(list((path / "activity").glob("*Gift Card Activity*.xls*")))
    )


def _contains_any_close_input(path: Path) -> bool:
    return bool(
        list((path / "summary").glob("*Gift Card Summary*.xlsx"))
        or list((path / "activity").glob("*Gift Card Activity*.xls*"))
    )


def _format_readiness(job: CloseJob, input_dir: Path) -> str:
    summary_count = len(list((input_dir / "summary").glob("*Gift Card Summary*.xlsx")))
    activity_count = len(list((input_dir / "activity").glob("*Gift Card Activity*.xls*")))
    expected = len(job.fiscal_period.expected_week_endings)
    status = "READY" if summary_count == 1 and activity_count == expected else "NOT READY"
    return (
        f"{job.store} {job.fiscal_period.period_key}: {status}\n"
        f"  Input: {input_dir.resolve()}\n"
        f"  Summary files: {summary_count} of 1\n"
        f"  Activity files: {activity_count} of {expected}\n"
        f"  Darden: {job.darden_path.resolve()}"
    )


def _print_success(result: MonthlyCloseRunResult) -> None:
    print(
        f"{result.assessment.store} {result.reconciliation.period}: "
        f"{result.assessment.status.value}"
    )
    print(f"Workbook: {result.workbook_path.resolve()}")
    print(f"PDF: {result.pdf_path.resolve()}")
    print(f"Manifest: {result.manifest_path.resolve()}")
