from __future__ import annotations

import os
import re
import tempfile
import uuid
from dataclasses import dataclass, replace
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Callable, Iterable, Mapping, Sequence

from gift_card_recon.close_assessment import (
    CloseAssessment,
    CloseStatus,
    ControlDisposition,
    ControlOutcome,
    assess_monthly_close,
    build_close_assessment,
    integrity_control,
)
from gift_card_recon.darden import build_monthly_close_certification, parse_darden_credit_memo
from gift_card_recon.evidence_archive import (
    ArchiveError,
    ArchiveRecord,
    EvidenceItem,
    cleanup_after_publish,
    copy_and_verify_evidence,
    execute_archive_plan,
    plan_evidence_archive,
    write_close_manifest_atomic,
)
from gift_card_recon.excel_writer import write_reconciliation_workbook
from gift_card_recon.fiscal_calendar import FiscalPeriod, fiscal_period_for_label
from gift_card_recon.micros import (
    MicrosEvidence,
    build_weekly_pos_variances,
    load_micros_evidence,
    period_tender_variance,
    resolve_micros_export_dir,
    validate_micros_source,
    weekly_tender_variances,
)
from gift_card_recon.models import (
    DardenCreditMemo,
    PosControls,
    ReconciliationResult,
    SourceFileAudit,
    WeeklyPosVariance,
)
from gift_card_recon.monthly_report import (
    DEFAULT_EVIDENCE_LABELS,
    MonthlyCloseReportData,
    WeeklyCloseReportRow,
    write_monthly_close_report_workbook,
)
from gift_card_recon.parsers import ParseError, discover_input_files, parse_activity_file, parse_summary
from gift_card_recon.pdf_export import PdfExportError, export_monthly_close_report_pdf
from gift_card_recon.reconcile import build_reconciliation
from gift_card_recon.source_validation import validate_activity_evidence
from gift_card_recon.store_config import StoreConfig, get_store_config
from gift_card_recon.utils import file_modified_at, money, sha256_file


PdfExporter = Callable[..., Path]


class CloseBlockedError(RuntimeError):
    """A close was not published because one or more required controls failed."""

    def __init__(
        self,
        message: str,
        *,
        assessment: CloseAssessment,
        review_workbook: Path | None = None,
        review_pdf: Path | None = None,
    ) -> None:
        super().__init__(message)
        self.assessment = assessment
        self.review_workbook = review_workbook
        self.review_pdf = review_pdf


@dataclass(frozen=True)
class MonthlyCloseRunResult:
    workbook_path: Path
    pdf_path: Path
    reconciliation: ReconciliationResult
    weekly_variances: tuple[WeeklyPosVariance, ...]
    assessment: CloseAssessment
    manifest_path: Path
    archive_records: tuple[ArchiveRecord, ...]

    def __iter__(self):
        """Compatibility with the former three-value result."""

        yield self.workbook_path
        yield self.reconciliation
        yield list(self.weekly_variances)


def canonical_output_paths(
    output_root: Path,
    *,
    config: StoreConfig,
    fiscal_period: FiscalPeriod,
) -> tuple[Path, Path]:
    folder = Path(output_root) / "Monthly Close" / fiscal_period.folder_name
    base = f"{config.output_slug}_{fiscal_period.period_key}_Monthly_Close"
    return folder / f"{base}.xlsx", folder / f"{base}.pdf"


def review_output_paths(
    output_root: Path,
    *,
    config: StoreConfig,
    fiscal_period: FiscalPeriod,
) -> tuple[Path, Path]:
    folder = Path(output_root) / "Review Required"
    base = f"{config.output_slug}_{fiscal_period.period_key}_Review_Required"
    return folder / f"{base}.xlsx", folder / f"{base}.pdf"


def run_monthly_close_service(
    *,
    store: str | int,
    period: str,
    input_dir: Path,
    micros_path: Path | None,
    micros_work_dir: Path,
    archive_root: Path,
    output_root: Path,
    output_path: Path | None = None,
    darden_path: Path | None = None,
    darden_report: DardenCreditMemo | None = None,
    fiscal_period: FiscalPeriod | None = None,
    cleanup_sources: bool = True,
    allow_unconfigured_micros: bool = False,
    generated_at: datetime | None = None,
    pdf_exporter: PdfExporter | None = None,
) -> MonthlyCloseRunResult:
    """Assess, render, archive, publish, and clean up one store-period close."""

    config = get_store_config(store)
    canonical_period = fiscal_period_for_label(period)
    if fiscal_period is not None and (
        fiscal_period.period_key != canonical_period.period_key
        or fiscal_period.folder_name != canonical_period.folder_name
        or fiscal_period.start_date != canonical_period.start_date
        or fiscal_period.end_date != canonical_period.end_date
    ):
        raise ParseError(f"Fiscal period {period!r} does not match the supplied period object.")
    fiscal_period = canonical_period
    generated_at = generated_at.astimezone() if generated_at is not None else datetime.now().astimezone()
    pdf_exporter = pdf_exporter or export_monthly_close_report_pdf
    input_dir = Path(input_dir)
    archive_root = Path(archive_root)
    output_root = Path(output_root)
    canonical_xlsx, canonical_pdf = canonical_output_paths(
        output_root,
        config=config,
        fiscal_period=fiscal_period,
    )
    if output_path is not None:
        canonical_xlsx = Path(output_path)
        if canonical_xlsx.suffix.lower() != ".xlsx":
            raise ParseError(f"Monthly-close workbook output must end in .xlsx: {canonical_xlsx}")
        canonical_pdf = canonical_xlsx.with_suffix(".pdf")

    try:
        close_data = _build_close_data(
            config=config,
            fiscal_period=fiscal_period,
            input_dir=input_dir,
            micros_path=Path(micros_path) if micros_path is not None else config.micros_default_path,
            micros_work_dir=Path(micros_work_dir),
            archive_root=archive_root,
            darden_path=Path(darden_path) if darden_path is not None else None,
            darden_report=darden_report,
            cleanup_sources=cleanup_sources,
            allow_unconfigured_micros=allow_unconfigured_micros,
        )
    except (ParseError, ArchiveError, OSError, ValueError, RuntimeError) as exc:
        assessment = _failure_assessment(config.store, "evidence_validation", "Evidence validation", str(exc))
        review_xlsx, review_pdf = _write_review_diagnostic(
            output_root=output_root,
            config=config,
            fiscal_period=fiscal_period,
            assessment=assessment,
            generated_at=generated_at,
            message=str(exc),
            pdf_exporter=pdf_exporter,
        )
        raise CloseBlockedError(
            str(exc),
            assessment=assessment,
            review_workbook=review_xlsx,
            review_pdf=review_pdf,
        ) from exc

    assessment = close_data.assessment
    if not assessment.can_publish_close:
        review_xlsx, review_pdf = _write_detailed_review(
            output_root=output_root,
            config=config,
            fiscal_period=fiscal_period,
            close_data=close_data,
            generated_at=generated_at,
            pdf_exporter=pdf_exporter,
        )
        blockers = "; ".join(control.message for control in assessment.blockers)
        raise CloseBlockedError(
            blockers or "One or more monthly-close controls require review.",
            assessment=assessment,
            review_workbook=review_xlsx,
            review_pdf=review_pdf,
        )

    canonical_xlsx.parent.mkdir(parents=True, exist_ok=True)
    manifest_path = (
        archive_root
        / "Monthly Close"
        / config.store
        / fiscal_period.folder_name
        / "close_manifest.json"
    )
    archive_preexisting = {
        record.archive_path.resolve(strict=False): record.archive_path.exists()
        for record in close_data.archive_records
    }
    try:
        for publish_target in (canonical_xlsx, canonical_pdf, manifest_path):
            _assert_publishable(publish_target)
        with tempfile.TemporaryDirectory(prefix="monthly-close-", dir=str(canonical_xlsx.parent)) as temp_dir:
            temp_root = Path(temp_dir)
            temp_xlsx = temp_root / canonical_xlsx.name
            temp_pdf = temp_root / canonical_pdf.name
            archived_result = _with_archive_source_audits(
                close_data.reconciliation,
                close_data.archive_records,
            )
            _write_full_workbook(
                output_path=temp_xlsx,
                result=archived_result,
                close_data=close_data,
                generated_at=generated_at,
                archive_published=True,
            )
            pdf_exporter(
                workbook_path=temp_xlsx,
                pdf_path=temp_pdf,
                expected_location_label=config.report_heading,
            )
            execute_archive_plan(close_data.archive_records)

            def finalize() -> None:
                write_close_manifest_atomic(
                    manifest_path,
                    store=config.store,
                    location=config.location_name,
                    period=fiscal_period.period_key,
                    status=assessment.status.value,
                    source_records=close_data.archive_records,
                    artifacts={"workbook": canonical_xlsx, "pdf": canonical_pdf},
                    archive_root=archive_root,
                    generated_at=generated_at,
                )
                if cleanup_sources:
                    cleanup_after_publish(
                        close_data.archive_records,
                        prune_period_dirs=(input_dir,),
                    )

            _publish_pair_transactional(
                temp_xlsx=temp_xlsx,
                temp_pdf=temp_pdf,
                canonical_xlsx=canonical_xlsx,
                canonical_pdf=canonical_pdf,
                manifest_path=manifest_path,
                finalize=finalize,
            )
        _archive_stale_review_artifacts(
            output_root=output_root,
            archive_root=archive_root,
            config=config,
            fiscal_period=fiscal_period,
        )
    except (ArchiveError, PdfExportError, OSError, ValueError, RuntimeError) as exc:
        _rollback_new_archive_copies(
            close_data.archive_records,
            archive_preexisting=archive_preexisting,
            archive_root=archive_root,
        )
        failed = _with_blocking_control(
            assessment,
            code="publication_integrity",
            label="Publication and archive transaction",
            message=str(exc),
        )
        failed_data = replace(close_data, assessment=failed)
        review_xlsx, review_pdf = _write_detailed_review(
            output_root=output_root,
            config=config,
            fiscal_period=fiscal_period,
            close_data=failed_data,
            generated_at=generated_at,
            pdf_exporter=pdf_exporter,
        )
        raise CloseBlockedError(
            str(exc),
            assessment=failed,
            review_workbook=review_xlsx,
            review_pdf=review_pdf,
        ) from exc

    return MonthlyCloseRunResult(
        workbook_path=canonical_xlsx,
        pdf_path=canonical_pdf,
        reconciliation=archived_result,
        weekly_variances=close_data.weekly_variances,
        assessment=assessment,
        manifest_path=manifest_path,
        archive_records=close_data.archive_records,
    )


def assess_monthly_close_inputs(
    *,
    store: str | int,
    period: str,
    input_dir: Path,
    micros_path: Path | None,
    micros_work_dir: Path,
    archive_root: Path,
    darden_path: Path | None = None,
    fiscal_period: FiscalPeriod | None = None,
    allow_unconfigured_micros: bool = False,
) -> CloseAssessment:
    """Run the same strict controls without copying, publishing, or cleanup."""

    config = get_store_config(store)
    canonical = fiscal_period_for_label(period)
    if fiscal_period is not None and (
        fiscal_period.period_key != canonical.period_key
        or fiscal_period.folder_name != canonical.folder_name
        or fiscal_period.start_date != canonical.start_date
        or fiscal_period.end_date != canonical.end_date
    ):
        raise ParseError(f"Fiscal period {period!r} does not match the supplied period object.")
    data = _build_close_data(
        config=config,
        fiscal_period=canonical,
        input_dir=Path(input_dir),
        micros_path=Path(micros_path) if micros_path is not None else config.micros_default_path,
        micros_work_dir=Path(micros_work_dir),
        archive_root=Path(archive_root),
        darden_path=Path(darden_path) if darden_path is not None else None,
        darden_report=None,
        cleanup_sources=False,
        allow_unconfigured_micros=allow_unconfigured_micros,
    )
    return data.assessment


def write_monthly_close_diagnostic(
    *,
    store: str | int,
    period: str,
    output_root: Path,
    message: str,
    generated_at: datetime | None = None,
    pdf_exporter: PdfExporter | None = None,
) -> tuple[Path, Path | None]:
    """Create a red diagnostic for a discovery failure with known identity."""

    config = get_store_config(store)
    fiscal_period = fiscal_period_for_label(period)
    assessment = _failure_assessment(
        config.store,
        "darden_discovery",
        "Darden report discovery",
        message,
    )
    timestamp = generated_at.astimezone() if generated_at else datetime.now().astimezone()
    return _write_review_diagnostic(
        output_root=Path(output_root),
        config=config,
        fiscal_period=fiscal_period,
        assessment=assessment,
        generated_at=timestamp,
        message=message,
        pdf_exporter=pdf_exporter or export_monthly_close_report_pdf,
    )


@dataclass(frozen=True)
class _CloseData:
    certification: object
    reconciliation: ReconciliationResult
    weekly_variances: tuple[WeeklyPosVariance, ...]
    weekly_tender: Mapping[str, Decimal]
    period_tender: Decimal
    assessment: CloseAssessment
    archive_records: tuple[ArchiveRecord, ...]
    micros_evidence: MicrosEvidence


def _build_close_data(
    *,
    config: StoreConfig,
    fiscal_period: FiscalPeriod,
    input_dir: Path,
    micros_path: Path,
    micros_work_dir: Path,
    archive_root: Path,
    darden_path: Path | None,
    darden_report: DardenCreditMemo | None,
    cleanup_sources: bool,
    allow_unconfigured_micros: bool,
) -> _CloseData:
    summary_path, activity_paths, _ = discover_input_files(input_dir, mode="monthly")
    if summary_path is None:
        raise ParseError("Monthly close requires exactly one Gift Card Summary workbook.")
    _validate_summary_period(summary_path, fiscal_period)
    validate_micros_source(
        micros_path,
        config,
        allow_unconfigured_source=allow_unconfigured_micros,
    )
    resolved_micros_path = resolve_micros_export_dir(micros_path, micros_work_dir)
    if darden_report is None:
        chosen_darden = darden_path or _one_darden_path(input_dir)
    else:
        chosen_darden = darden_report.source_file
        if darden_path is not None and Path(darden_path).resolve() != Path(chosen_darden).resolve():
            raise ParseError("The supplied Darden report and Darden path identify different files.")
    initial_paths = [
        summary_path,
        *activity_paths,
        Path(chosen_darden),
        _required_source_file(resolved_micros_path, config.micros_system_totals_file),
        _required_source_file(resolved_micros_path, config.micros_tender_detail_file),
    ]
    initial_hashes = {
        path.resolve(): sha256_file(path)
        for path in initial_paths
    }

    summary = parse_summary(summary_path, store=config.store)
    activities = [
        parse_activity_file(path, summary.conversion_promo_codes) for path in activity_paths
    ]
    activity_evidence = validate_activity_evidence(
        activities,
        store=config.store,
        period_start=fiscal_period.start_date,
        period_end=fiscal_period.end_date,
        expected_week_endings=fiscal_period.expected_week_endings,
    )

    if darden_report is None:
        darden_report = parse_darden_credit_memo(chosen_darden)
    certification = build_monthly_close_certification(
        store=config.store,
        period=fiscal_period.period_key,
        period_start=fiscal_period.start_date,
        period_end=fiscal_period.end_date,
        summary=summary,
        darden_credit_memo=darden_report,
    )

    micros_evidence = load_micros_evidence(
        resolved_micros_path,
        config=config,
        activity_evidence=activity_evidence,
        period_start=fiscal_period.start_date,
        period_end=fiscal_period.end_date,
        validate_source=False,
        allow_unconfigured_source=allow_unconfigured_micros,
    )
    weekly_variances = tuple(
        build_weekly_pos_variances(
            activity_evidence,
            micros_evidence,
            conversion_promo_codes=summary.conversion_promo_codes,
        )
    )
    pos_controls = PosControls(
        store=config.store,
        period=fiscal_period.period_key,
        pos_gift_card_issue=money(
            sum((row.pos_issue for row in weekly_variances), Decimal("0.00"))
        ),
        pos_gift_card_payment=money(
            sum((row.pos_payment for row in weekly_variances), Decimal("0.00"))
        ),
    )
    reconciliation = build_reconciliation(
        store=config.store,
        period=fiscal_period.period_key,
        period_end=fiscal_period.end_date,
        summary=summary,
        activities=list(activity_evidence.files),
        pos_controls=pos_controls,
        mode="monthly",
        additional_source_files=(
            darden_report.source_file,
            micros_evidence.system_totals_path,
            micros_evidence.tender_detail_path,
        ),
        strict_nonzero_review=True,
    )

    weekly_tender = weekly_tender_variances(
        micros_evidence,
        week_endings=fiscal_period.expected_week_endings,
    )
    period_tender = period_tender_variance(micros_evidence)
    evidence_items = _evidence_items(
        config=config,
        fiscal_period=fiscal_period,
        summary_path=summary_path,
        activity_paths=activity_paths,
        darden_path=darden_report.source_file,
        micros_evidence=micros_evidence,
        archive_root=archive_root,
        cleanup_sources=cleanup_sources,
    )
    archive_records = tuple(plan_evidence_archive(evidence_items, archive_root=archive_root))
    for record in archive_records:
        initial_hash = initial_hashes.get(record.source_path.resolve())
        if initial_hash is None or initial_hash != record.sha256:
            raise ArchiveError(
                f"Evidence changed while the close was being calculated: {record.source_path}. "
                "No report was published; rerun with a stable source file."
            )

    integrity_controls = (
        integrity_control(
            code="summary_identity",
            label="Summary identity and required values",
            passed=True,
            pass_message="Exactly one Summary row matched this store and required money parsed.",
            failure_message="Summary identity validation failed.",
        ),
        integrity_control(
            code="activity_identity",
            label="Activity report identity",
            passed=True,
            pass_message="Every activity report identifies the configured store.",
            failure_message="Activity identity validation failed.",
        ),
        integrity_control(
            code="activity_coverage",
            label="Activity weekly coverage",
            passed=True,
            pass_message="Exactly one non-overlapping Monday-Sunday report covers every expected week.",
            failure_message="Activity coverage validation failed.",
        ),
        integrity_control(
            code="darden_identity",
            label="Darden identity and period",
            passed=True,
            pass_message="The Darden memo store and service dates match this close.",
            failure_message="Darden identity validation failed.",
        ),
        integrity_control(
            code="micros_source",
            label="Micros source and layout",
            passed=True,
            pass_message=f"Validated {config.micros_source_label} and configured control columns.",
            failure_message="Micros source validation failed.",
        ),
        integrity_control(
            code="micros_coverage",
            label="Micros date coverage",
            passed=True,
            pass_message="Every fiscal date is present or is an evidence-confirmed scheduled closure.",
            failure_message="Micros coverage validation failed.",
        ),
        integrity_control(
            code="tender_evidence",
            label="Tender evidence",
            passed=True,
            pass_message="Tender Detail is present, well formed, normalized, and assessed.",
            failure_message="Tender evidence validation failed.",
        ),
        integrity_control(
            code="archive_integrity",
            label="Archive plan and source hashes",
            passed=True,
            pass_message="Every required source was hashed and assigned a canonical archive path.",
            failure_message="Archive preflight failed.",
        ),
    )
    summary_activity = {
        line.metric: line.activity_variance for line in reconciliation.lines
    }
    weekly_pos: dict[str, Decimal] = {}
    for row in weekly_variances:
        week = row.week_ending.strftime("%m/%d/%Y") if row.week_ending else "Unknown"
        weekly_pos[f"Week ending {week} POS issue"] = row.issue_variance
        weekly_pos[f"Week ending {week} POS payment"] = row.payment_variance
        weekly_pos[f"Week ending {week} POS net"] = row.net_variance
    period_pos = {
        f"Period POS {line.metric}": line.pos_variance for line in reconciliation.lines
    }
    assessment = assess_monthly_close(
        store=config.store,
        darden_variance=certification.variance,
        summary_activity_variances=summary_activity,
        weekly_pos_variances=weekly_pos,
        period_pos_variances=period_pos,
        weekly_tender_variances=weekly_tender,
        period_tender_variances={"Period tender": period_tender},
        integrity_controls=integrity_controls,
        additional_required_integrity_codes=("archive_integrity",),
        expected_week_count=len(fiscal_period.expected_week_endings),
    )
    assessment_exceptions = [
        (
            control.disposition.value,
            f"{control.label}: {control.message}",
        )
        for control in assessment.controls
        if not control.passed
    ]
    reconciliation = replace(reconciliation, exceptions=assessment_exceptions)
    return _CloseData(
        certification=certification,
        reconciliation=reconciliation,
        weekly_variances=weekly_variances,
        weekly_tender=weekly_tender,
        period_tender=period_tender,
        assessment=assessment,
        archive_records=archive_records,
        micros_evidence=micros_evidence,
    )


def _evidence_items(
    *,
    config: StoreConfig,
    fiscal_period: FiscalPeriod,
    summary_path: Path,
    activity_paths: Sequence[Path],
    darden_path: Path,
    micros_evidence: MicrosEvidence,
    archive_root: Path,
    cleanup_sources: bool,
) -> list[EvidenceItem]:
    base = Path("Monthly Close") / config.store / fiscal_period.folder_name

    def removable(path: Path) -> bool:
        return cleanup_sources and not _is_within(path, archive_root)

    items = [
        EvidenceItem("Gift Card Summary", summary_path, str(base / "summary"), removable(summary_path)),
        *[
            EvidenceItem("Weekly Gift Card Activity", path, str(base / "activity"), removable(path))
            for path in activity_paths
        ],
        EvidenceItem("Darden Credit Memo", darden_path, str(base / "darden"), removable(darden_path)),
        EvidenceItem("Micros Daily System Totals", micros_evidence.system_totals_path, str(base / "micros"), False),
        EvidenceItem("Micros Tender Detail", micros_evidence.tender_detail_path, str(base / "micros"), False),
    ]
    return items


def _write_full_workbook(
    *,
    output_path: Path,
    result: ReconciliationResult,
    close_data: _CloseData,
    generated_at: datetime,
    archive_published: bool = False,
) -> None:
    assessment = close_data.assessment
    weekly_rows, weekly_codes = _weekly_report_rows(close_data)
    period_pos = next(
        (line.pos_variance for line in result.lines if line.metric == "Net Gift Card Impact"),
        None,
    )
    write_reconciliation_workbook(
        result,
        output_path,
        monthly_close_certification=close_data.certification,
        close_assessment=assessment,
        weekly_pos_variances=list(close_data.weekly_variances),
        weekly_close_rows=weekly_rows,
        period_pos_net_variance=period_pos,
        period_pos_disposition=_worst_disposition(
            control for control in assessment.controls if control.code.startswith("period_pos_")
        ),
        period_tender_variance=close_data.period_tender,
        period_tender_disposition=_worst_disposition(
            control for control in assessment.controls if control.code.startswith("period_tender_")
        ),
        evidence_notes=(
            (
                "Canonical archive-relative paths and SHA-256 hashes are recorded in Source Files and the close manifest."
                if archive_published
                else "Source hashes and archive destinations were planned, but no canonical archive or close manifest was published for this diagnostic."
            ),
            "Scheduled Mondays are accepted only when both activity and tender evidence are zero; existing Monday POS is included normally.",
        ),
        source_labels=DEFAULT_EVIDENCE_LABELS,
        weekly_control_codes=frozenset(weekly_codes),
        generated_at=generated_at,
        micros_source_label=assessment.store_config.micros_source_label,
    )


def _weekly_report_rows(
    close_data: _CloseData,
) -> tuple[list[WeeklyCloseReportRow], set[str]]:
    assessment = close_data.assessment
    control_by_label = {control.label: control for control in assessment.controls}
    result: list[WeeklyCloseReportRow] = []
    codes: set[str] = set()
    for row in close_data.weekly_variances:
        week = row.week_ending.strftime("%m/%d/%Y") if row.week_ending else "Unknown"
        labels = (
            f"Week ending {week} POS issue",
            f"Week ending {week} POS payment",
            f"Week ending {week} POS net",
            f"Week ending {week} tender",
        )
        controls = [control_by_label[label] for label in labels]
        codes.update(control.code for control in controls)
        non_pass = [
            f"{control.label.rsplit(' ', 2)[-2]} {control.label.rsplit(' ', 1)[-1]} "
            f"{control.variance:+,.2f}"
            if control.variance is not None
            else control.label
            for control in controls
            if not control.passed
        ]
        result.append(
            WeeklyCloseReportRow(
                week_ending=row.week_ending,
                coverage=row.coverage_status,
                pos_issue_variance=row.issue_variance,
                pos_payment_variance=row.payment_variance,
                pos_net_variance=row.net_variance,
                tender_variance=close_data.weekly_tender[labels[-1]],
                disposition=_worst_disposition(controls),
                evidence_note=(
                    "Review required: " + "; ".join(non_pass)
                    if non_pass
                    else "Evidence complete; all weekly controls passed."
                ),
            )
        )
    return result, codes


def _with_archive_source_audits(
    result: ReconciliationResult,
    records: Sequence[ArchiveRecord],
) -> ReconciliationResult:
    audits: list[SourceFileAudit] = []
    for record in records:
        relative_path = Path(record.archive_category) / record.archive_path.name
        audits.append(
            SourceFileAudit(
                path=relative_path,
                file_type=record.source_path.suffix.lower().lstrip("."),
                size_bytes=record.size_bytes,
                modified_at=file_modified_at(record.source_path),
                sha256=record.sha256,
            )
        )
    return replace(result, source_files=audits)


def _write_detailed_review(
    *,
    output_root: Path,
    config: StoreConfig,
    fiscal_period: FiscalPeriod,
    close_data: _CloseData,
    generated_at: datetime,
    pdf_exporter: PdfExporter,
) -> tuple[Path, Path | None]:
    review_xlsx, review_pdf = review_output_paths(
        output_root,
        config=config,
        fiscal_period=fiscal_period,
    )
    review_xlsx.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="review-required-", dir=str(review_xlsx.parent)) as temp_dir:
        temp_xlsx = Path(temp_dir) / review_xlsx.name
        temp_pdf = Path(temp_dir) / review_pdf.name
        _write_full_workbook(
            output_path=temp_xlsx,
            result=close_data.reconciliation,
            close_data=close_data,
            generated_at=generated_at,
        )
        exported = _try_export_review(
            temp_xlsx,
            temp_pdf,
            config=config,
            pdf_exporter=pdf_exporter,
        )
        if exported:
            _publish_pair_transactional(
                temp_xlsx=temp_xlsx,
                temp_pdf=temp_pdf,
                canonical_xlsx=review_xlsx,
                canonical_pdf=review_pdf,
            )
            return review_xlsx, review_pdf
        _publish_single(temp_xlsx, review_xlsx)
        return review_xlsx, None


def _write_review_diagnostic(
    *,
    output_root: Path,
    config: StoreConfig,
    fiscal_period: FiscalPeriod,
    assessment: CloseAssessment,
    generated_at: datetime,
    message: str,
    pdf_exporter: PdfExporter,
) -> tuple[Path, Path | None]:
    review_xlsx, review_pdf = review_output_paths(
        output_root,
        config=config,
        fiscal_period=fiscal_period,
    )
    review_xlsx.parent.mkdir(parents=True, exist_ok=True)
    report_data = MonthlyCloseReportData(
        assessment=assessment,
        period=fiscal_period.period_key,
        period_start=fiscal_period.start_date,
        period_end=fiscal_period.end_date,
        generated_at=generated_at,
        explicit_exceptions=(("BLOCK", message),),
        evidence_notes=("No inputs were archived and no canonical close report was published.",),
    )
    with tempfile.TemporaryDirectory(prefix="review-required-", dir=str(review_xlsx.parent)) as temp_dir:
        temp_xlsx = Path(temp_dir) / review_xlsx.name
        temp_pdf = Path(temp_dir) / review_pdf.name
        write_monthly_close_report_workbook(report_data, temp_xlsx)
        exported = _try_export_review(
            temp_xlsx,
            temp_pdf,
            config=config,
            pdf_exporter=pdf_exporter,
        )
        if exported:
            _publish_pair_transactional(
                temp_xlsx=temp_xlsx,
                temp_pdf=temp_pdf,
                canonical_xlsx=review_xlsx,
                canonical_pdf=review_pdf,
            )
            return review_xlsx, review_pdf
        _publish_single(temp_xlsx, review_xlsx)
        return review_xlsx, None


def _try_export_review(
    workbook: Path,
    pdf: Path,
    *,
    config: StoreConfig,
    pdf_exporter: PdfExporter,
) -> bool:
    try:
        pdf_exporter(
            workbook_path=workbook,
            pdf_path=pdf,
            expected_location_label=config.report_heading,
        )
        return True
    except (PdfExportError, OSError, RuntimeError):
        return False


def _publish_pair_transactional(
    *,
    temp_xlsx: Path,
    temp_pdf: Path,
    canonical_xlsx: Path,
    canonical_pdf: Path,
    manifest_path: Path | None = None,
    finalize: Callable[[], None] | None = None,
) -> None:
    for source in (temp_xlsx, temp_pdf):
        if not source.is_file() or source.stat().st_size <= 0:
            raise OSError(f"Verified publication artifact is missing or empty: {source}")
    targets = [canonical_xlsx, canonical_pdf]
    if manifest_path is not None:
        targets.append(manifest_path)
    for target in targets:
        target.parent.mkdir(parents=True, exist_ok=True)
        _assert_publishable(target)

    backups: dict[Path, Path] = {}
    published: list[Path] = []
    try:
        for target in targets:
            if target.exists():
                backup = target.with_name(f".{target.name}.{uuid.uuid4().hex}.backup")
                os.replace(target, backup)
                backups[target] = backup
        os.replace(temp_xlsx, canonical_xlsx)
        published.append(canonical_xlsx)
        os.replace(temp_pdf, canonical_pdf)
        published.append(canonical_pdf)
        if finalize is not None:
            finalize()
        for backup in backups.values():
            try:
                backup.unlink(missing_ok=True)
            except OSError:
                # Backup disposal is post-commit housekeeping. The canonical
                # artifacts, manifest, and verified archive are already valid.
                pass
    except Exception:
        for target in reversed(published):
            try:
                target.unlink(missing_ok=True)
            except OSError:
                pass
        if manifest_path is not None and manifest_path not in backups:
            try:
                manifest_path.unlink(missing_ok=True)
            except OSError:
                pass
        for target, backup in backups.items():
            if backup.exists():
                try:
                    os.replace(backup, target)
                except OSError:
                    pass
        raise


def _publish_single(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    _assert_publishable(destination)
    try:
        os.replace(source, destination)
    except PermissionError as exc:
        raise PermissionError(
            f"Cannot replace locked report {destination}. Close it and rerun; no alternate file was created."
        ) from exc


def _assert_publishable(path: Path) -> None:
    if not path.exists():
        return
    try:
        with path.open("r+b"):
            pass
    except OSError as exc:
        raise PermissionError(
            f"Cannot replace locked canonical output {path}. Close it and rerun; "
            "no alternate filename will be created."
        ) from exc


def _failure_assessment(store: str, code: str, label: str, message: str) -> CloseAssessment:
    return build_close_assessment(
        store=store,
        darden_variance=None,
        controls=(
            ControlOutcome(
                code=code,
                label=label,
                disposition=ControlDisposition.BLOCK,
                message=message or "Required evidence validation failed.",
            ),
        ),
    )


def _with_blocking_control(
    assessment: CloseAssessment,
    *,
    code: str,
    label: str,
    message: str,
) -> CloseAssessment:
    controls = tuple(control for control in assessment.controls if control.code != code)
    return CloseAssessment(
        store_config=assessment.store_config,
        darden_matched=assessment.darden_matched,
        controls=(
            *controls,
            ControlOutcome(
                code=code,
                label=label,
                disposition=ControlDisposition.BLOCK,
                message=message or "Publication integrity failed.",
            ),
        ),
    )


def _worst_disposition(controls: Iterable[ControlOutcome]) -> ControlDisposition:
    dispositions = {control.disposition for control in controls}
    if ControlDisposition.BLOCK in dispositions:
        return ControlDisposition.BLOCK
    if ControlDisposition.REVIEW in dispositions:
        return ControlDisposition.REVIEW
    return ControlDisposition.PASS


def _one_darden_path(input_dir: Path) -> Path:
    candidates = sorted((Path(input_dir) / "darden").glob("*.pdf"))
    candidates.extend(sorted(Path(input_dir).glob("*Darden*.pdf")))
    unique = list(dict.fromkeys(path.resolve() for path in candidates))
    if len(unique) != 1:
        raise ParseError(
            f"Monthly close requires exactly one Darden PDF in {input_dir}; found {len(unique)}."
        )
    return Path(unique[0])


def _validate_summary_period(path: Path, fiscal_period: FiscalPeriod) -> None:
    match = re.match(r"(\d{2})\.(\d{2})\.(\d{4})\s+", Path(path).name)
    if match is None:
        raise ParseError(
            f"Summary filename must begin with its report end date (MM.DD.YYYY): {Path(path).name}"
        )
    report_end = date(int(match.group(3)), int(match.group(1)), int(match.group(2)))
    if report_end != fiscal_period.end_date:
        raise ParseError(
            f"Summary {Path(path).name} is for report end {report_end:%Y-%m-%d}; "
            f"{fiscal_period.period_key} ends {fiscal_period.end_date:%Y-%m-%d}."
        )


def _required_source_file(folder: Path, expected_name: str) -> Path:
    matches = [
        path
        for path in Path(folder).iterdir()
        if path.is_file() and path.name.casefold() == expected_name.casefold()
    ] if Path(folder).is_dir() else []
    if len(matches) != 1:
        raise ParseError(
            f"Expected exactly one {expected_name} in {folder}; found {len(matches)}."
        )
    return matches[0]


def _is_within(path: Path, root: Path) -> bool:
    try:
        Path(path).resolve().relative_to(Path(root).resolve())
        return True
    except ValueError:
        return False


def _rollback_new_archive_copies(
    records: Sequence[ArchiveRecord],
    *,
    archive_preexisting: Mapping[Path, bool],
    archive_root: Path,
) -> None:
    root = Path(archive_root).resolve(strict=False)
    parents: set[Path] = set()
    for record in records:
        path = record.archive_path
        resolved = path.resolve(strict=False)
        if archive_preexisting.get(resolved, False) or not path.is_file():
            continue
        try:
            resolved.relative_to(root)
            if path.stat().st_size == record.size_bytes and sha256_file(path) == record.sha256:
                path.unlink()
                parents.add(path.parent)
        except (OSError, ValueError):
            continue
    for parent in sorted(parents, key=lambda value: len(value.parts), reverse=True):
        current = parent
        while current != root:
            try:
                current.rmdir()
            except OSError:
                break
            current = current.parent


def _archive_stale_review_artifacts(
    *,
    output_root: Path,
    archive_root: Path,
    config: StoreConfig,
    fiscal_period: FiscalPeriod,
) -> None:
    review_xlsx, review_pdf = review_output_paths(
        output_root,
        config=config,
        fiscal_period=fiscal_period,
    )
    candidates = [path for path in (review_xlsx, review_pdf) if path.is_file()]
    if not candidates:
        return
    category = str(
        Path("Generated Reports")
        / "Diagnostics"
        / config.store
        / fiscal_period.period_key
    )
    try:
        records = copy_and_verify_evidence(
            [
                EvidenceItem(
                    role="Superseded review diagnostic",
                    source_path=path,
                    archive_category=category,
                    remove_after_publish=True,
                )
                for path in candidates
            ],
            archive_root=archive_root,
        )
        cleanup_after_publish(records, prune_period_dirs=(review_xlsx.parent,))
    except ArchiveError:
        # The canonical close and its evidence manifest are already committed.
        # Leave a locked diagnostic intact rather than invalidating that close.
        return
