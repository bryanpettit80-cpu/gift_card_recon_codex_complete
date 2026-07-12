from __future__ import annotations

import argparse
import json
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
from gift_card_recon.utils import sha256_file


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


@dataclass(frozen=True)
class ArchiveReissue:
    job: CloseJob
    input_dir: Path
    micros_path: Path


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
    parser.add_argument(
        "--reissue-from-archive",
        action="store_true",
        help="Reissue one store-period only after verifying every source against its close manifest.",
    )
    parser.add_argument("--no-boundary-adjustment", action="store_true", help=argparse.SUPPRESS)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    input_root = Path(args.input_root)
    archive_root = Path(args.archive_root)
    output_root = Path(args.output_dir)
    inbox = input_root / SHARED_DARDEN_INBOX

    archive_reissue: ArchiveReissue | None = None
    if args.reissue_from_archive:
        if not args.store or not args.period:
            print("REVIEW REQUIRED: --reissue-from-archive requires both --store and --period.")
            return 1
        conflicts = [
            option
            for option, value in (
                ("--input-dir", args.input_dir),
                ("--darden-path", args.darden_path),
                ("--micros-path", args.micros_path),
            )
            if value is not None
        ]
        if conflicts:
            print(
                "REVIEW REQUIRED: --reissue-from-archive derives verified evidence and cannot be combined with "
                + ", ".join(conflicts)
                + "."
            )
            return 1
        try:
            archive_reissue = _resolve_archive_reissue(
                archive_root=archive_root,
                store=args.store,
                period=args.period,
            )
        except (ParseError, OSError, ValueError, RuntimeError) as exc:
            print(f"REVIEW REQUIRED: {exc}")
            return 1
        jobs, discovery_errors = [archive_reissue.job], []
    else:
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
                    diagnostic = write_monthly_close_diagnostic(
                        store=issue.store,
                        period=issue.period,
                        output_root=output_root,
                        message=issue.message,
                    )
                    workbook, pdf = diagnostic
                    print(f"Diagnostic workbook: {workbook.resolve()}")
                    if pdf is not None:
                        print(f"Diagnostic PDF: {pdf.resolve()}")
                    if diagnostic.pdf_error is not None:
                        print(
                            f"Diagnostic PDF unavailable: {diagnostic.pdf_error}; "
                            "use the diagnostic workbook."
                        )
                except (OSError, RuntimeError, ValueError) as exc:
                    print(f"Diagnostic publication failed: {exc}")
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
            input_dir = (
                archive_reissue.input_dir
                if archive_reissue is not None
                else _resolve_input_dir(
                    job,
                    input_root=input_root,
                    archive_root=archive_root,
                    explicit_input=Path(args.input_dir) if args.input_dir else None,
                    stage_weekly=not args.no_stage_weekly,
                )
            )
            if args.prepare_only:
                config = get_store_config(job.store)
                micros_path = (
                    archive_reissue.micros_path
                    if archive_reissue is not None
                    else Path(args.micros_path) if args.micros_path else config.micros_default_path
                )
                assessment = assess_monthly_close_inputs(
                    store=job.store,
                    period=job.fiscal_period.period_key,
                    input_dir=input_dir,
                    micros_path=micros_path,
                    micros_work_dir=Path(args.micros_work_dir),
                    archive_root=archive_root,
                    darden_path=job.darden_path,
                    fiscal_period=job.fiscal_period,
                    allow_unconfigured_micros=archive_reissue is not None,
                )
                print(f"{job.store} {job.fiscal_period.period_key}: {assessment.status.value}")
                if not assessment.can_publish_close:
                    failures += 1
                continue
            config = get_store_config(job.store)
            micros_path = (
                archive_reissue.micros_path
                if archive_reissue is not None
                else Path(args.micros_path) if args.micros_path else config.micros_default_path
            )
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
                cleanup_sources=False if archive_reissue is not None else not args.no_cleanup,
                allow_unconfigured_micros=archive_reissue is not None,
            )
            _print_success(result)
        except CloseBlockedError as exc:
            failures += 1
            print(f"{job.store} {job.fiscal_period.period_key}: REVIEW REQUIRED - {exc}")
            if exc.review_workbook is not None:
                print(f"Diagnostic workbook: {exc.review_workbook.resolve()}")
            if exc.review_pdf is not None:
                print(f"Diagnostic PDF: {exc.review_pdf.resolve()}")
            if exc.review_pdf_error is not None:
                print(
                    f"Diagnostic PDF unavailable: {exc.review_pdf_error}; "
                    "use the diagnostic workbook."
                )
            if exc.review_publication_error is not None:
                print(f"Diagnostic publication failed: {exc.review_publication_error}")
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


def _resolve_archive_reissue(
    *,
    archive_root: Path,
    store: str | int,
    period: str,
) -> ArchiveReissue:
    """Resolve and hash-verify the canonical evidence package for a safe reissue."""

    config = get_store_config(store)
    fiscal_period = fiscal_period_for_label(period)
    root = Path(archive_root).resolve(strict=False)
    input_dir = root / "Monthly Close" / config.store / fiscal_period.folder_name
    manifest_path = input_dir / "close_manifest.json"
    if not manifest_path.is_file():
        raise ParseError(f"Archived close manifest was not found: {manifest_path}")
    try:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ParseError(f"Archived close manifest is unreadable: {manifest_path}: {exc}") from exc
    if payload.get("store") != config.store or payload.get("period") != fiscal_period.period_key:
        raise ParseError(
            f"Archived close manifest identity does not match store {config.store} {fiscal_period.period_key}."
        )
    sources = payload.get("sources")
    if not isinstance(sources, list) or not sources:
        raise ParseError(f"Archived close manifest contains no source records: {manifest_path}")

    by_role: dict[str, list[Path]] = defaultdict(list)
    for index, row in enumerate(sources, start=1):
        if not isinstance(row, dict):
            raise ParseError(f"Archived close manifest source row {index} is invalid.")
        role = row.get("role")
        relative = row.get("archive_path")
        digest = row.get("sha256")
        size_bytes = row.get("size_bytes")
        if not isinstance(role, str) or not isinstance(relative, str):
            raise ParseError(f"Archived close manifest source row {index} is missing role or archive_path.")
        source_path = (root / Path(relative)).resolve(strict=False)
        if not _is_within(source_path, input_dir):
            raise ParseError(
                f"Archived source escapes the canonical store-period package: {relative}"
            )
        if not source_path.is_file():
            raise ParseError(f"Archived source is missing: {source_path}")
        if not isinstance(size_bytes, int) or source_path.stat().st_size != size_bytes:
            raise ParseError(f"Archived source size does not match its close manifest: {source_path}")
        if not isinstance(digest, str) or sha256_file(source_path) != digest.lower():
            raise ParseError(f"Archived source hash does not match its close manifest: {source_path}")
        by_role[role].append(source_path)

    required_counts = {
        "Gift Card Summary": 1,
        "Weekly Gift Card Activity": len(fiscal_period.expected_week_endings),
        "Darden Credit Memo": 1,
        "Micros Daily System Totals": 1,
        "Micros Tender Detail": 1,
    }
    unknown_roles = sorted(set(by_role) - set(required_counts))
    if unknown_roles:
        raise ParseError(
            "Archived close manifest contains unsupported source role(s): "
            + ", ".join(unknown_roles)
            + "."
        )
    for role, expected_count in required_counts.items():
        found = len(by_role.get(role, ()))
        if found != expected_count:
            raise ParseError(
                f"Archived close manifest requires {expected_count} {role} source(s); found {found}."
            )

    expected_folders = {
        "Gift Card Summary": input_dir / "summary",
        "Weekly Gift Card Activity": input_dir / "activity",
        "Darden Credit Memo": input_dir / "darden",
        "Micros Daily System Totals": input_dir / "micros",
        "Micros Tender Detail": input_dir / "micros",
    }
    for role, folder in expected_folders.items():
        expected_parent = folder.resolve(strict=False)
        if any(path.parent != expected_parent for path in by_role[role]):
            raise ParseError(f"Archived {role} evidence is outside its canonical folder: {folder}")

    expected_micros_names = {
        "Micros Daily System Totals": config.micros_system_totals_file,
        "Micros Tender Detail": config.micros_tender_detail_file,
    }
    for role, expected_name in expected_micros_names.items():
        if by_role[role][0].name.casefold() != expected_name.casefold():
            raise ParseError(
                f"Archived {role} must be named {expected_name}: {by_role[role][0]}"
            )

    darden_path = by_role["Darden Credit Memo"][0]
    darden_report = parse_darden_credit_memo(darden_path)
    if darden_report.store != config.store:
        raise ParseError(
            f"Archived Darden report is for store {darden_report.store}; expected {config.store}."
        )
    if (
        darden_report.period_start != fiscal_period.start_date
        or darden_report.period_end != fiscal_period.end_date
    ):
        raise ParseError(
            f"Archived Darden service dates do not exactly match {fiscal_period.period_key}."
        )
    return ArchiveReissue(
        job=CloseJob(config.store, fiscal_period, darden_path, darden_report),
        input_dir=input_dir,
        micros_path=input_dir / "micros",
    )


def _is_within(path: Path, root: Path) -> bool:
    try:
        Path(path).resolve(strict=False).relative_to(Path(root).resolve(strict=False))
        return True
    except ValueError:
        return False


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
