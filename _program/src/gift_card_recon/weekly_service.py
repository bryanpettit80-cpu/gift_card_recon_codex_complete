from __future__ import annotations

import csv
import errno
import json
import os
import shutil
import tempfile
import time
import uuid
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date, datetime, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any, Mapping, TypeVar

from gift_card_recon.excel_writer import write_reconciliation_workbook
from gift_card_recon.fiscal_calendar import fiscal_period_for_date
from gift_card_recon.micros import MicrosEvidence, load_micros_evidence, period_tender_variance
from gift_card_recon.models import ActivityFileData, PosControls, ReconciliationResult
from gift_card_recon.parsers import ParseError, parse_activity_file
from gift_card_recon.reconcile import build_reconciliation
from gift_card_recon.source_validation import ActivityEvidence, validate_activity_evidence
from gift_card_recon.store_config import StoreConfig
from gift_card_recon.utils import money, sha256_file


_T = TypeVar("_T")
_TRANSIENT_FILE_RETRY_DELAYS = (0.10, 0.25, 0.50, 1.00, 2.00)
_TRANSIENT_WINERRORS = frozenset({5, 32, 33, 148, 170})
_TRANSIENT_ERRNOS = frozenset({errno.EACCES, errno.EBUSY})


@dataclass(frozen=True)
class WeeklyPublication:
    period: str
    period_end: date
    status: str
    output_path: Path
    archive_path: Path
    monthly_activity_path: Path
    message: str


@dataclass(frozen=True)
class WeeklyDuplicate:
    period: str
    period_end: date
    quarantined_path: Path
    output_path: Path
    message: str


@dataclass(frozen=True)
class SourceFingerprint:
    path: Path
    size_bytes: int
    sha256: str


@dataclass(frozen=True)
class WeeklyPreparation:
    activity: ActivityFileData
    activity_evidence: ActivityEvidence
    micros: MicrosEvidence
    result: ReconciliationResult
    tender_variance: Decimal
    review_items: tuple[str, ...]
    source_fingerprints: tuple[SourceFingerprint, ...]


def prepare_weekly_reconciliation(
    *,
    store: str,
    activity_path: Path,
    config: StoreConfig,
) -> WeeklyPreparation:
    """Parse and validate one week, then derive all controls from live Micros evidence."""

    activity, activity_evidence, activity_source = _prepare_activity_evidence(
        store=store,
        activity_path=activity_path,
    )
    return _prepare_weekly_from_activity(
        store=store,
        config=config,
        activity=activity,
        activity_evidence=activity_evidence,
        activity_source=activity_source,
    )


def _prepare_activity_evidence(
    *,
    store: str,
    activity_path: Path,
) -> tuple[ActivityFileData, ActivityEvidence, SourceFingerprint]:
    """Validate the Activity report without touching the live Micros export."""

    activity_source = _capture_source_fingerprint(activity_path)
    activity = parse_activity_file(activity_path)
    _verify_source_stability((activity_source,), "parsing the Activity report")
    if activity.report_begin is None or activity.report_end is None:
        raise ParseError(f"{activity_path.name} is missing its report date range.")
    if activity.report_begin.weekday() != 0 or activity.report_end.weekday() != 6:
        raise ParseError(
            f"{activity_path.name} covers {activity.report_begin.isoformat()} through "
            f"{activity.report_end.isoformat()}; expected an exact Monday-Sunday week."
        )
    activity_evidence = validate_activity_evidence(
        [activity],
        store=store,
        period_start=activity.report_begin,
        period_end=activity.report_end,
        expected_week_endings=[activity.report_end],
    )
    return activity, activity_evidence, activity_source


def _prepare_weekly_from_activity(
    *,
    store: str,
    config: StoreConfig,
    activity: ActivityFileData,
    activity_evidence: ActivityEvidence,
    activity_source: SourceFingerprint,
) -> WeeklyPreparation:
    """Derive weekly controls after the archive has ruled out a duplicate."""

    assert activity.report_begin is not None and activity.report_end is not None
    system_path = _required_source_path(
        config.micros_default_path,
        config.micros_system_totals_file,
    )
    tender_path = _required_source_path(
        config.micros_default_path,
        config.micros_tender_detail_file,
    )
    micros_sources = (
        _capture_source_fingerprint(system_path),
        _capture_source_fingerprint(tender_path),
    )
    micros_evidence = load_micros_evidence(
        config.micros_default_path,
        config=config,
        activity_evidence=activity_evidence,
        period_start=activity.report_begin,
        period_end=activity.report_end,
    )
    source_fingerprints = (activity_source, *micros_sources)
    _verify_source_stability(source_fingerprints, "loading Micros evidence")
    pos_controls = PosControls(
        store=store,
        period=iso_week_period(activity.report_end),
        pos_gift_card_issue=money(
            sum((row.pos_gift_card_issue for row in micros_evidence.daily_pos), Decimal("0.00"))
        ),
        pos_gift_card_payment=money(
            sum((row.pos_gift_card_payment for row in micros_evidence.daily_pos), Decimal("0.00"))
        ),
    )
    tender_variance = period_tender_variance(micros_evidence)
    preliminary = build_reconciliation(
        store=store,
        period=pos_controls.period,
        period_end=activity.report_end,
        summary=None,
        activities=[activity],
        pos_controls=pos_controls,
        mode="weekly",
        strict_nonzero_review=True,
        additional_source_files=(
            micros_evidence.system_totals_path,
            micros_evidence.tender_detail_path,
        ),
    )
    review_items = _review_items(preliminary, tender_variance)
    exceptions = [("Review", item) for item in review_items]
    result = build_reconciliation(
        store=store,
        period=pos_controls.period,
        period_end=activity.report_end,
        summary=None,
        activities=[activity],
        pos_controls=pos_controls,
        mode="weekly",
        strict_nonzero_review=True,
        exceptions=exceptions,
        additional_source_files=(
            micros_evidence.system_totals_path,
            micros_evidence.tender_detail_path,
        ),
    )
    return WeeklyPreparation(
        activity=activity,
        activity_evidence=activity_evidence,
        micros=micros_evidence,
        result=result,
        tender_variance=tender_variance,
        review_items=review_items,
        source_fingerprints=source_fingerprints,
    )


def publish_weekly_reconciliation(
    *,
    store: str,
    activity_path: Path,
    config: StoreConfig,
    output_path: Path,
    archive_root: Path,
    monthly_close_root: Path,
    review_root: Path,
) -> WeeklyPublication | WeeklyDuplicate:
    """Create the weekly report and evidence package as one rollback-safe publication."""

    activity_path = Path(activity_path)
    activity, activity_evidence, activity_source = _prepare_activity_evidence(
        store=store,
        activity_path=activity_path,
    )
    assert activity.report_begin is not None and activity.report_end is not None
    output_path = Path(output_path)
    period = iso_week_period(activity.report_end)
    iso_year = activity.report_end.isocalendar().year
    store_folder = f"{config.store} {config.location_name}"
    package_path = Path(archive_root) / store_folder / str(iso_year) / period
    duplicate = _handle_existing_package(
        package_path=package_path,
        store=store,
        activity_path=activity_path,
        activity_source=activity_source,
        output_path=output_path,
        period=period,
        period_begin=activity.report_begin,
        period_end=activity.report_end,
        review_root=review_root,
    )
    if duplicate is not None:
        return duplicate

    prepared = _prepare_weekly_from_activity(
        store=store,
        config=config,
        activity=activity,
        activity_evidence=activity_evidence,
        activity_source=activity_source,
    )
    micros = prepared.micros
    result = prepared.result
    tender_variance = prepared.tender_variance
    review_items = prepared.review_items
    source_fingerprints = prepared.source_fingerprints

    monthly_period = fiscal_period_for_date(activity.report_end)
    monthly_path = (
        Path(monthly_close_root)
        / store_folder
        / monthly_period.folder_name
        / "activity"
        / activity_path.name
    )
    _preflight_destination(output_path, "weekly report")
    monthly_already_staged = _preflight_monthly_destination(monthly_path, activity_path)

    archive_root = Path(archive_root)
    created_directories: list[Path] = []
    _ensure_directory(archive_root, created_directories)
    stage_root: Path | None = None
    output_temp: Path | None = None
    monthly_temp: Path | None = None
    archive_committed = False
    output_committed = False
    monthly_committed = False
    phase = "creating the staging workspace"
    try:
        stage_root = _create_staging_root(archive_root)
        package_stage = stage_root / "package"
        archived_activity = package_stage / "activity" / activity_path.name
        archived_report = package_stage / "report" / output_path.name
        evidence_path = package_stage / "pos" / "weekly_pos_tender_evidence.csv"
        for directory in (archived_activity.parent, archived_report.parent, evidence_path.parent):
            directory.mkdir(parents=True, exist_ok=True)

        phase = "building staged evidence"
        shutil.copy2(activity_path, archived_activity)
        if sha256_file(archived_activity) != source_fingerprints[0].sha256:
            raise ValueError("the archived Activity copy does not match the bytes used for calculation")
        write_reconciliation_workbook(result, archived_report)
        _write_daily_evidence(
            evidence_path,
            activity_evidence=activity_evidence,
            micros=micros,
        )
        _verify_source_stability(source_fingerprints, "building the report and evidence package")
        generated_at = datetime.now(timezone.utc)
        status = "REVIEW" if review_items else "PASS"
        manifest = _build_manifest(
            store=store,
            period=period,
            activity=activity,
            micros=micros,
            source_fingerprints=source_fingerprints,
            result=result,
            tender_variance=tender_variance,
            review_items=review_items,
            status=status,
            generated_at=generated_at,
            package_path=package_path,
            archived_activity=archived_activity,
            archived_report=archived_report,
            evidence_path=evidence_path,
            output_path=output_path,
            monthly_path=monthly_path,
        )
        manifest_path = package_stage / "weekly_manifest.json"
        _write_json(manifest_path, manifest)
        _verify_staged_package(package_stage, manifest)
        _verify_source_stability(source_fingerprints, "finalizing the manifest")

        phase = "preparing canonical and monthly copies"
        _ensure_directory(output_path.parent, created_directories)
        output_temp = _temporary_copy_path(stage_root, output_path, "canonical")
        shutil.copy2(archived_report, output_temp)
        if not monthly_already_staged:
            _ensure_directory(monthly_path.parent, created_directories)
            monthly_temp = _temporary_copy_path(stage_root, monthly_path, "monthly")
            shutil.copy2(archived_activity, monthly_temp)

        phase = "committing evidence package"
        _ensure_directory(package_path.parent, created_directories)
        _retry_transient_file_operation(lambda: os.replace(package_stage, package_path))
        archive_committed = True
        phase = "committing canonical workbook"
        _retry_transient_file_operation(lambda: os.replace(output_temp, output_path))
        output_temp = None
        output_committed = True
        if monthly_temp is not None:
            phase = "committing monthly-close activity"
            _retry_transient_file_operation(lambda: os.replace(monthly_temp, monthly_path))
            monthly_temp = None
            monthly_committed = True

        phase = "verifying published artifacts"
        _verify_published_artifacts(
            package_path=package_path,
            output_path=output_path,
            monthly_path=monthly_path,
            manifest=manifest,
        )
        _verify_source_stability(source_fingerprints, "publishing the weekly reconciliation")
        phase = "removing the temporary staging folder"
        _remove_tree_strict(stage_root)
        phase = "removing processed inbox activity"
        _retry_transient_file_operation(lambda: activity_path.unlink(missing_ok=True))
        message = (
            f"Created {output_path.name}; archived immutable weekly evidence; "
            f"copied activity to monthly close and removed it from the inbox."
        )
        return WeeklyPublication(
            period=period,
            period_end=activity.report_end,
            status=status,
            output_path=output_path,
            archive_path=package_path,
            monthly_activity_path=monthly_path,
            message=message,
        )
    except (OSError, RuntimeError, ValueError, ParseError) as exc:
        cleanup_issues: list[str] = []
        if monthly_committed:
            cleanup_issues.extend(_cleanup_file(monthly_path))
        if output_committed:
            cleanup_issues.extend(_cleanup_file(output_path))
        if archive_committed:
            cleanup_issues.extend(_cleanup_tree(package_path))
        if output_temp is not None:
            cleanup_issues.extend(_cleanup_file(output_temp))
        if monthly_temp is not None:
            cleanup_issues.extend(_cleanup_file(monthly_temp))
        if stage_root is not None:
            cleanup_issues.extend(_cleanup_tree(stage_root))
        cleanup_issues.extend(_prune_created_directories(created_directories))
        cleanup_text = ""
        if cleanup_issues:
            cleanup_text = " Rollback incomplete: " + "; ".join(dict.fromkeys(cleanup_issues))
        raise ParseError(
            f"Weekly publication failed while {phase}: {exc}."
            + (
                cleanup_text
                if cleanup_issues
                else " Rollback verified; no partial publication was retained."
            )
        ) from exc


def iso_week_period(period_end: date) -> str:
    iso = period_end.isocalendar()
    return f"{iso.year}-W{iso.week:02d}"


def _review_items(result: ReconciliationResult, tender_variance: Decimal) -> tuple[str, ...]:
    items: list[str] = []
    for line in result.lines[:2]:
        if line.pos_variance is not None and money(line.pos_variance) != Decimal("0.00"):
            items.append(f"{line.metric} POS variance is {line.pos_variance:+,.2f}.")
    if money(tender_variance) != Decimal("0.00"):
        items.append(
            f"Micros gift-card payment versus tender-detail variance is {tender_variance:+,.2f}."
        )
    return tuple(items)


def _write_daily_evidence(
    path: Path,
    *,
    activity_evidence: ActivityEvidence,
    micros: MicrosEvidence,
) -> None:
    pos_by_date = micros.daily_pos_by_date
    activity_magnitude = activity_evidence.daily_activity_magnitude
    accepted_closed = set(micros.accepted_closed_dates)
    fieldnames = [
        "business_date",
        "pos_gift_card_issue",
        "pos_gift_card_payment",
        "tender_gift_card_payment",
        "activity_magnitude",
        "tender_date_observed",
        "coverage_status",
    ]
    with path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=fieldnames)
        writer.writeheader()
        current = activity_evidence.period_start
        while current <= activity_evidence.period_end:
            pos = pos_by_date.get(current)
            writer.writerow(
                {
                    "business_date": current.isoformat(),
                    "pos_gift_card_issue": _decimal_text(pos.pos_gift_card_issue if pos else Decimal("0.00")),
                    "pos_gift_card_payment": _decimal_text(pos.pos_gift_card_payment if pos else Decimal("0.00")),
                    "tender_gift_card_payment": _decimal_text(micros.daily_tender.get(current, Decimal("0.00"))),
                    "activity_magnitude": _decimal_text(activity_magnitude.get(current, Decimal("0.00"))),
                    "tender_date_observed": str(current in micros.tender_observed_dates).lower(),
                    "coverage_status": "accepted scheduled closure" if current in accepted_closed else "complete",
                }
            )
            current = date.fromordinal(current.toordinal() + 1)


def _build_manifest(
    *,
    store: str,
    period: str,
    activity: ActivityFileData,
    micros: MicrosEvidence,
    source_fingerprints: tuple[SourceFingerprint, ...],
    result: ReconciliationResult,
    tender_variance: Decimal,
    review_items: tuple[str, ...],
    status: str,
    generated_at: datetime,
    package_path: Path,
    archived_activity: Path,
    archived_report: Path,
    evidence_path: Path,
    output_path: Path,
    monthly_path: Path,
) -> dict[str, Any]:
    assert activity.report_begin is not None and activity.report_end is not None
    pos_by_date = micros.daily_pos_by_date
    daily: list[dict[str, Any]] = []
    current = activity.report_begin
    while current <= activity.report_end:
        pos = pos_by_date.get(current)
        daily.append(
            {
                "business_date": current.isoformat(),
                "pos_gift_card_issue": _decimal_text(pos.pos_gift_card_issue if pos else Decimal("0.00")),
                "pos_gift_card_payment": _decimal_text(pos.pos_gift_card_payment if pos else Decimal("0.00")),
                "tender_gift_card_payment": _decimal_text(micros.daily_tender.get(current, Decimal("0.00"))),
                "tender_date_observed": current in micros.tender_observed_dates,
                "accepted_scheduled_closure": current in set(micros.accepted_closed_dates),
            }
        )
        current = date.fromordinal(current.toordinal() + 1)

    archived_activity_record = _artifact_record(
        archived_activity,
        "activity",
        Path("activity") / archived_activity.name,
    )
    archived_report_record = _artifact_record(archived_report, "report", Path("report") / archived_report.name)
    evidence_record = _artifact_record(evidence_path, "weekly_pos_tender_evidence", Path("pos") / evidence_path.name)
    return {
        "schema_version": 1,
        "store": store,
        "period": period,
        "week": {"start": activity.report_begin.isoformat(), "end": activity.report_end.isoformat()},
        "status": status,
        "generated_at": generated_at.isoformat().replace("+00:00", "Z"),
        "source_files": {
            "activity": _source_record(source_fingerprints[0]),
            "micros_system_totals": _source_record(source_fingerprints[1]),
            "micros_tender_detail": _source_record(source_fingerprints[2]),
        },
        "daily_pos_tender": daily,
        "totals": {
            "activity_issue": _decimal_text(result.activity_total_activations),
            "activity_payment": _decimal_text(abs(result.activity_total_redemptions)),
            "pos_gift_card_issue": _decimal_text(result.pos_controls.pos_gift_card_issue),
            "pos_gift_card_payment": _decimal_text(result.pos_controls.pos_gift_card_payment),
            "tender_gift_card_payment": _decimal_text(sum(micros.daily_tender.values(), Decimal("0.00"))),
            "tender_variance": _decimal_text(tender_variance),
        },
        "review_items": list(review_items),
        "artifacts": {
            "archived_activity": archived_activity_record,
            "weekly_pos_tender_evidence": evidence_record,
            "archived_workbook": archived_report_record,
            "canonical_workbook": {
                **archived_report_record,
                "path": str(output_path.resolve()),
            },
            "monthly_staged_activity": {
                **archived_activity_record,
                "path": str(monthly_path.resolve()),
            },
        },
        "archive_path": str(package_path.resolve()),
    }


def _artifact_record(path: Path, role: str, relative_path: Path) -> dict[str, Any]:
    return {
        "role": role,
        "relative_path": relative_path.as_posix(),
        "size_bytes": path.stat().st_size,
        "sha256": sha256_file(path),
    }


def _source_record(source: SourceFingerprint) -> dict[str, Any]:
    return {
        "path": str(source.path.resolve()),
        "name": source.path.name,
        "size_bytes": source.size_bytes,
        "sha256": source.sha256,
    }


def _write_json(path: Path, payload: Mapping[str, Any]) -> None:
    path.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def _verify_staged_package(package_path: Path, manifest: Mapping[str, Any]) -> None:
    manifest_path = package_path / "weekly_manifest.json"
    if not manifest_path.is_file():
        raise ValueError("weekly_manifest.json was not created")
    for role in ("archived_activity", "weekly_pos_tender_evidence", "archived_workbook"):
        record = manifest["artifacts"][role]
        path = _package_artifact_path(package_path, record, role)
        _verify_record(path, record, role)


def _package_artifact_path(
    package_path: Path,
    record: Mapping[str, Any],
    role: str,
) -> Path:
    """Resolve only the role-specific durable path inside the evidence package."""

    relative_path = Path(str(record["relative_path"]))
    expected_parent = {
        "archived_activity": Path("activity"),
        "weekly_pos_tender_evidence": Path("pos"),
        "archived_workbook": Path("report"),
    }[role]
    if relative_path.is_absolute() or relative_path.parent != expected_parent:
        raise ValueError(f"{role} relative_path is outside its expected package folder")
    package_resolved = package_path.resolve()
    candidate = (package_resolved / relative_path).resolve()
    try:
        candidate.relative_to(package_resolved)
    except ValueError as exc:
        raise ValueError(f"{role} relative_path escapes the weekly evidence package") from exc
    return candidate


def _verify_published_artifacts(
    *,
    package_path: Path,
    output_path: Path,
    monthly_path: Path,
    manifest: Mapping[str, Any],
) -> None:
    _verify_staged_package(package_path, manifest)
    _verify_record(output_path, manifest["artifacts"]["canonical_workbook"], "canonical_workbook")
    _verify_record(monthly_path, manifest["artifacts"]["monthly_staged_activity"], "monthly_staged_activity")


def _verify_record(path: Path, record: Mapping[str, Any], role: str) -> None:
    if not path.is_file():
        raise ValueError(f"{role} is missing: {path}")
    if path.stat().st_size != int(record["size_bytes"]) or sha256_file(path) != record["sha256"]:
        raise ValueError(f"{role} failed size/SHA-256 verification: {path}")


def _handle_existing_package(
    *,
    package_path: Path,
    store: str,
    activity_path: Path,
    activity_source: SourceFingerprint,
    output_path: Path,
    period: str,
    period_begin: date,
    period_end: date,
    review_root: Path,
) -> WeeklyDuplicate | None:
    if not package_path.exists():
        return None
    manifest_path = package_path / "weekly_manifest.json"
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        _verify_staged_package(package_path, manifest)
        activity_record = manifest["source_files"]["activity"]
        archived_activity_record = manifest["artifacts"]["archived_activity"]
        _verify_activity_record_consistency(activity_record, archived_activity_record)
        archived_workbook_record = manifest["artifacts"]["archived_workbook"]
        canonical_record = manifest["artifacts"]["canonical_workbook"]
        _verify_canonical_record_consistency(archived_workbook_record, canonical_record)
        canonical_path = Path(canonical_record["path"])
        if canonical_path.resolve() != output_path.resolve():
            raise ValueError("canonical_workbook path does not match the expected weekly output path")
        _verify_record(canonical_path, canonical_record, "canonical_workbook")
        if manifest.get("schema_version") != 1:
            raise ValueError("manifest schema version is unsupported")
        if manifest.get("store") != store or manifest.get("period") != period:
            raise ValueError("manifest store/period identity does not match its archive folder")
        if manifest.get("week", {}).get("start") != period_begin.isoformat():
            raise ValueError("manifest week-start date does not match the Activity report")
        if manifest.get("week", {}).get("end") != period_end.isoformat():
            raise ValueError("manifest week-ending date does not match the Activity report")
        if Path(str(manifest.get("archive_path", ""))).resolve() != package_path.resolve():
            raise ValueError("manifest archive_path does not match its evidence package")
        _verify_source_stability((activity_source,), "checking the archived weekly package")
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        raise ParseError(f"Existing weekly evidence package is incomplete or invalid: {package_path}: {exc}") from exc
    if (
        activity_source.size_bytes != int(archived_activity_record["size_bytes"])
        or activity_source.sha256 != archived_activity_record["sha256"]
    ):
        raise ParseError(
            f"A different Activity report already exists for store {manifest.get('store')} {period}; "
            f"the new file was left in place for review."
        )

    quarantine_dir = Path(review_root) / "duplicate-inputs" / str(manifest.get("store")) / period
    quarantine_dir.mkdir(parents=True, exist_ok=True)
    quarantine_path = quarantine_dir / (
        f"{datetime.now():%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:8]}-{activity_path.name}"
    )
    _verify_source_stability((activity_source,), "quarantining the exact duplicate")
    try:
        _retry_transient_file_operation(lambda: os.replace(activity_path, quarantine_path))
    except OSError as exc:
        raise ParseError(f"Could not quarantine exact duplicate {activity_path}: {exc}") from exc
    return WeeklyDuplicate(
        period=period,
        period_end=period_end,
        quarantined_path=quarantine_path,
        output_path=canonical_path,
        message=f"Quarantined exact duplicate input at {quarantine_path}.",
    )


def _verify_activity_record_consistency(
    source_record: Mapping[str, Any],
    archived_record: Mapping[str, Any],
) -> None:
    """Bind duplicate identity to the verified Activity artifact in the package."""

    if (
        int(source_record["size_bytes"]) != int(archived_record["size_bytes"])
        or source_record["sha256"] != archived_record["sha256"]
        or source_record["name"] != Path(str(archived_record["relative_path"])).name
    ):
        raise ValueError(
            "manifest Activity source does not match the durable archived_activity artifact"
        )


def _verify_canonical_record_consistency(
    archived_record: Mapping[str, Any],
    canonical_record: Mapping[str, Any],
) -> None:
    """Bind the canonical workbook to the verified durable workbook copy."""

    for field in ("size_bytes", "sha256", "relative_path", "role"):
        if archived_record[field] != canonical_record[field]:
            raise ValueError(
                "canonical_workbook record does not match the durable archived_workbook artifact"
            )


def _preflight_destination(path: Path, label: str) -> None:
    if path.exists():
        raise ParseError(
            f"The canonical {label} already exists without a completed matching evidence package: {path}."
        )


def _preflight_monthly_destination(destination: Path, source: Path) -> bool:
    if not destination.exists():
        return False
    if not destination.is_file() or sha256_file(destination) != sha256_file(source):
        raise ParseError(f"Monthly-close staging conflict: {destination} is not identical to {source.name}.")
    return True


def _decimal_text(value: Decimal) -> str:
    return f"{money(value):.2f}"


def _capture_source_fingerprint(path: Path) -> SourceFingerprint:
    path = Path(path)
    try:
        stat = path.stat()
        digest = sha256_file(path)
    except OSError as exc:
        raise ParseError(f"Could not fingerprint source file {path}: {exc}") from exc
    return SourceFingerprint(path=path, size_bytes=stat.st_size, sha256=digest)


def _verify_source_stability(
    sources: tuple[SourceFingerprint, ...],
    phase: str,
) -> None:
    changed: list[str] = []
    for source in sources:
        try:
            stat = source.path.stat()
            digest = sha256_file(source.path)
        except OSError as exc:
            changed.append(f"{source.path} ({exc})")
            continue
        if stat.st_size != source.size_bytes or digest != source.sha256:
            changed.append(str(source.path))
    if changed:
        raise ParseError(
            f"Source evidence changed while weekly reconciliation was running during {phase}: "
            + "; ".join(changed)
        )


def _required_source_path(folder: Path, expected_name: str) -> Path:
    folder = Path(folder)
    if not folder.is_dir():
        raise ParseError(f"Micros export folder not found: {folder}")
    matches = [
        item
        for item in folder.iterdir()
        if item.is_file() and item.name.casefold() == expected_name.casefold()
    ]
    if len(matches) != 1:
        raise ParseError(
            f"Expected exactly one {expected_name} in {folder}; found {len(matches)}."
        )
    return matches[0]


def _remove_tree_strict(path: Path) -> None:
    path = Path(path)

    def remove_if_present() -> None:
        if path.exists():
            shutil.rmtree(path)

    _retry_transient_file_operation(remove_if_present)
    if path.exists():
        raise OSError(f"Temporary staging path was retained: {path}")


def _cleanup_file(path: Path) -> list[str]:
    path = Path(path)
    errors: list[str] = []
    try:
        _retry_transient_file_operation(lambda: path.unlink(missing_ok=True))
    except OSError as exc:
        errors.append(f"could not remove {path}: {exc}")
    if path.exists():
        errors.append(f"retained path {path}")
    return errors


def _cleanup_tree(path: Path) -> list[str]:
    path = Path(path)
    errors: list[str] = []

    def remove_if_present() -> None:
        if path.exists():
            shutil.rmtree(path)

    try:
        _retry_transient_file_operation(remove_if_present)
    except OSError as exc:
        errors.append(f"could not remove {path}: {exc}")
    if path.exists():
        errors.append(f"retained path {path}")
    return errors


def _ensure_directory(path: Path, created_directories: list[Path]) -> None:
    missing: list[Path] = []
    current = Path(path)
    while not current.exists():
        missing.append(current)
        if current.parent == current:
            break
        current = current.parent
    path.mkdir(parents=True, exist_ok=True)
    for item in missing:
        if item not in created_directories:
            created_directories.append(item)


def _prune_created_directories(created_directories: list[Path]) -> list[str]:
    errors: list[str] = []
    for path in created_directories:
        try:
            _retry_transient_file_operation(lambda path=path: _remove_empty_directory(path))
        except FileNotFoundError:
            continue
        except OSError as exc:
            errors.append(f"could not remove created directory {path}: {exc}")
        if path.exists():
            errors.append(f"retained path {path}")
    return errors


def _create_staging_root(archive_root: Path) -> Path:
    """Build outside Dropbox when the runtime temp folder is on the same volume."""

    stage_root = Path(tempfile.mkdtemp(prefix="gift-card-weekly-"))
    try:
        if stage_root.stat().st_dev == Path(archive_root).stat().st_dev:
            return stage_root
    except OSError:
        _remove_tree_strict(stage_root)
        raise

    _remove_tree_strict(stage_root)
    return Path(tempfile.mkdtemp(prefix=".weekly-staging-", dir=archive_root))


def _temporary_copy_path(stage_root: Path, destination: Path, label: str) -> Path:
    """Prefer local staging, but stay on the destination volume for atomic replacement."""

    if Path(stage_root).stat().st_dev == Path(destination).parent.stat().st_dev:
        return Path(stage_root) / f"{label}-{uuid.uuid4().hex[:8]}.tmp"
    return Path(destination).parent / f".gc-{uuid.uuid4().hex[:8]}.tmp"


def _remove_empty_directory(path: Path) -> None:
    if path.exists():
        path.rmdir()


def _retry_transient_file_operation(operation: Callable[[], _T]) -> _T:
    """Retry short-lived sharing and sync-provider locks, then preserve the original error."""

    for delay in _TRANSIENT_FILE_RETRY_DELAYS:
        try:
            return operation()
        except OSError as exc:
            if not _is_transient_file_error(exc):
                raise
            time.sleep(delay)
    return operation()


def _is_transient_file_error(exc: OSError) -> bool:
    winerror = getattr(exc, "winerror", None)
    if winerror is not None:
        return winerror in _TRANSIENT_WINERRORS
    return exc.errno in _TRANSIENT_ERRNOS
