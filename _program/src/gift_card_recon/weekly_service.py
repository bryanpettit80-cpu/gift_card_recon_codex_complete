from __future__ import annotations

import csv
import json
import os
import shutil
import tempfile
import time
import uuid
from collections.abc import Callable
from dataclasses import dataclass, replace
from datetime import date, datetime, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any, Mapping, TypeVar
from zipfile import BadZipFile

from gift_card_recon.business_identity import (
    CorrectionRequired, activity_identity, activity_changes, correction_token as make_correction_token,
    identity, verify_workbook_values,
)
from gift_card_recon.publication import (
    PublicationTransaction, recover_pending, support_root_for, writer_lock,
    retry_file_operation, file_state,
)
from gift_card_recon.excel_writer import write_reconciliation_workbook
from gift_card_recon.fiscal_calendar import fiscal_period_for_date
from gift_card_recon.micros import MicrosEvidence, load_micros_evidence, period_tender_variance
from gift_card_recon.models import ActivityFileData, PosControls, ReconciliationResult
from gift_card_recon.parsers import ParseError, parse_activity_file
from gift_card_recon.reconcile import build_reconciliation
from gift_card_recon.source_validation import ActivityEvidence, validate_activity_evidence
from gift_card_recon.store_config import REVIEW_VARIANCE_LIMIT, StoreConfig, get_store_config
from gift_card_recon.utils import money, sha256_file
from gift_card_recon.weekly_report_revision import resolve_weekly_report_revision
from gift_card_recon.variance_explanation import (
    EXPLANATION_INPUT_CELL,
    EXPLANATION_SHEET,
    WeeklyVarianceExplanation,
    read_variance_explanation_workbook,
    variance_control_mismatch_details,
    variance_explanation_path,
    find_legacy_variance_explanation,
    verify_weekly_report_explanation_edit,
)


_T = TypeVar("_T")
_TRANSIENT_FILE_RETRY_DELAYS = (0.10, 0.25, 0.50, 1.00, 2.00)


@dataclass(frozen=True)
class WeeklyPublication:
    period: str
    period_end: date
    status: str
    output_path: Path
    archive_path: Path
    monthly_activity_path: Path
    variance_explanation_path: Path | None
    message: str
    input_set_id: str = ""
    result_id: str = ""
    revision: int = 1
    recovery_action: str | None = None


@dataclass(frozen=True)
class WeeklyDuplicate:
    period: str
    period_end: date
    quarantined_path: Path
    output_path: Path
    message: str
    input_set_id: str = ""
    result_id: str = ""


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
    *, store: str, activity_path: Path, config: StoreConfig, output_path: Path,
    archive_root: Path, monthly_close_root: Path, review_root: Path,
    correction_token: str | None = None,
) -> WeeklyPublication | WeeklyDuplicate:
    """Validate under the workspace writer lock and recover before changing files."""
    support = support_root_for(archive_root)
    with writer_lock(support):
        recovery = recover_pending(support)
        result = _publish_weekly_reconciliation(
            store=store, activity_path=activity_path, config=config, output_path=output_path,
            archive_root=archive_root, monthly_close_root=monthly_close_root,
            review_root=review_root, correction_token=correction_token, support=support,
        )
        if recovery and isinstance(result, WeeklyPublication):
            result = replace(result, recovery_action=" ".join(recovery))
        return result


def _daily_controls(activity: ActivityFileData, micros: MicrosEvidence) -> list[dict[str, Any]]:
    assert activity.report_begin is not None and activity.report_end is not None
    daily = []
    current = activity.report_begin
    while current <= activity.report_end:
        pos = micros.daily_pos_by_date.get(current)
        daily.append({
            "business_date": current.isoformat(),
            "pos_gift_card_issue": _decimal_text(pos.pos_gift_card_issue if pos else Decimal("0.00")),
            "pos_gift_card_payment": _decimal_text(pos.pos_gift_card_payment if pos else Decimal("0.00")),
            "tender_gift_card_payment": _decimal_text(micros.daily_tender.get(current, Decimal("0.00"))),
            "tender_date_observed": current in micros.tender_observed_dates,
            "accepted_scheduled_closure": current in set(micros.accepted_closed_dates),
        })
        current = date.fromordinal(current.toordinal() + 1)
    return daily


def _weekly_input_id(activity: ActivityFileData, daily: list[dict[str, Any]]) -> str:
    return identity({"activity": activity_identity(activity), "daily_pos_tender": daily})


def weekly_preparation_identity(prepared: WeeklyPreparation) -> str:
    return _weekly_input_id(prepared.activity, _daily_controls(prepared.activity, prepared.micros))


def historical_controls_available(config: StoreConfig, activity: ActivityFileData) -> bool:
    """Distinguish absent/rotated historical exports from incomplete current ones."""
    from gift_card_recon.micros import _parse_system_totals, _parse_tender_detail
    folder = Path(config.micros_default_path)
    if not folder.is_dir():
        return False
    names = {item.name.casefold() for item in folder.iterdir() if item.is_file()}
    required = {config.micros_system_totals_file.casefold(), config.micros_tender_detail_file.casefold()}
    if not names.intersection(required):
        return False
    system = _required_source_path(folder, config.micros_system_totals_file)
    tender = _required_source_path(folder, config.micros_tender_detail_file)
    sources = (_capture_source_fingerprint(system), _capture_source_fingerprint(tender))
    pos = _parse_system_totals(system, config)
    _, _, tender_dates = _parse_tender_detail(tender, config)
    observed = {row.business_date for row in pos} | set(tender_dates)
    _verify_source_stability(sources, "checking historical Micros coverage")
    return any(activity.report_begin <= day <= activity.report_end for day in observed)


def _weekly_totals(prepared: WeeklyPreparation) -> dict[str, str]:
    return {
        "activity_issue": _decimal_text(prepared.result.activity_total_activations),
        "activity_payment": _decimal_text(abs(prepared.result.activity_total_redemptions)),
        "pos_gift_card_issue": _decimal_text(prepared.result.pos_controls.pos_gift_card_issue),
        "pos_gift_card_payment": _decimal_text(prepared.result.pos_controls.pos_gift_card_payment),
        "tender_gift_card_payment": _decimal_text(sum(prepared.micros.daily_tender.values(), Decimal("0.00"))),
        "tender_variance": _decimal_text(prepared.tender_variance),
    }


def _publish_weekly_reconciliation(
    *, store: str, activity_path: Path, config: StoreConfig, output_path: Path,
    archive_root: Path, monthly_close_root: Path, review_root: Path,
    correction_token: str | None, support: Path,
) -> WeeklyPublication | WeeklyDuplicate:
    activity_path, output_path = Path(activity_path), Path(output_path)
    activity, activity_evidence, activity_source = _prepare_activity_evidence(store=store, activity_path=activity_path)
    assert activity.report_begin is not None and activity.report_end is not None
    period = iso_week_period(activity.report_end)
    store_folder = f"{config.store} {config.location_name}"
    package_path = Path(archive_root) / store_folder / str(activity.report_end.isocalendar().year) / period
    preimages: dict[Path, str | None] = {}
    existing = _load_existing_package(
        package_path=package_path, store=store, output_path=output_path,
        period=period, period_begin=activity.report_begin, period_end=activity.report_end,
        preimages=preimages,
    )
    previous_id = None
    previous_manifest = None
    previous_revision = None
    if existing is not None:
        previous_manifest, archived_activity, output_path = existing
        previous_input_id = _weekly_input_id(archived_activity, previous_manifest["daily_pos_tender"])
        previous_id = previous_manifest.get("result_id") or identity({"input_set_id": previous_input_id, "rules": "weekly-v2"})
        unchanged_activity = activity_identity(activity) == activity_identity(archived_activity)
        # A duplicate can be recognized from durable evidence when historical live
        # Micros exports are no longer available. If current evidence exists it must
        # be validated, so POS-only corrections are not silently ignored.
        live_available = historical_controls_available(config, activity) if unchanged_activity else True
        if unchanged_activity and not live_available:
            return _quarantine_duplicate(activity_source, store, period, activity.report_end,
                                         output_path, review_root, support, previous_input_id, previous_id)
    prepared = _prepare_weekly_from_activity(
        store=store, config=config, activity=activity, activity_evidence=activity_evidence,
        activity_source=activity_source,
    )
    input_id = _weekly_input_id(activity, _daily_controls(activity, prepared.micros))
    result_id = identity({"input_set_id": input_id, "rules": "weekly-v2"})
    revision = 1
    if previous_manifest is not None:
        if input_id == previous_input_id:
            return _quarantine_duplicate(activity_source, store, period, activity.report_end,
                                         output_path, review_root, support, input_id, previous_id)
        changes = {name: f"{previous_manifest['totals'].get(name)} -> {value}"
                   for name, value in _weekly_totals(prepared).items()
                   if previous_manifest['totals'].get(name) != value}
        changes["activity_records"] = f"{len(archived_activity.rows)} -> {len(activity.rows)} (record details and multiplicity compared)"
        if activity_identity(archived_activity) != activity_identity(activity):
            changes["changed_transactions"] = activity_changes(archived_activity, activity)
        changes["input_set"] = f"{previous_input_id} -> {input_id}"
        proposal = CorrectionRequired(previous_id, input_id, changes)
        if correction_token != proposal.token:
            raise proposal
        revision = int(previous_manifest.get("revision", 1)) + 1
        previous_revision = package_path.with_name(period + ".revisions") / f"r{revision-1:04d}-{previous_id[:16]}"
        if previous_revision.exists():
            raise ParseError(f"Previous revision destination already exists: {previous_revision}")
    else:
        _preflight_destination(output_path, "weekly report")
        preimages[output_path] = None

    result_id = identity({"input_set_id": input_id, "rules": "weekly-v2", "revision": revision,
                          "previous_result_id": previous_id})

    monthly_period = fiscal_period_for_date(activity.report_end)
    monthly_period_dir = Path(monthly_close_root) / store_folder / monthly_period.folder_name
    monthly_path = monthly_period_dir / "activity" / activity_path.name
    preimages[monthly_path] = file_state(monthly_path)
    old_monthly_path = None
    for candidate in sorted(monthly_path.parent.glob("*")):
        if candidate.suffix.lower() not in {".xls", ".xlsx"} or candidate.name.startswith("~$"):
            continue
        candidate_before = file_state(candidate)
        candidate_activity = parse_activity_file(candidate)
        _require_file_state(candidate, candidate_before, "validating monthly Activity")
        if candidate_activity.store == store and candidate_activity.report_end == activity.report_end:
            if old_monthly_path is not None:
                raise ParseError(f"Multiple monthly Activity inputs match {store} {period}.")
            permitted = activity_identity(archived_activity if existing else activity)
            if activity_identity(candidate_activity) != permitted:
                raise ParseError(f"Monthly-close staging conflict: {candidate} changed independently.")
            old_monthly_path = candidate
            preimages[candidate] = candidate_before
    if monthly_path.exists() and old_monthly_path != monthly_path:
        raise ParseError(f"Monthly-close staging conflict: {monthly_path}")
    variance_explanation = _required_variance_explanation(prepared)
    explanation_path = output_path if variance_explanation is not None else None
    legacy_explanation_path = find_legacy_variance_explanation(monthly_period_dir, store, activity.report_end)
    explanation_source = output_path if existing and previous_manifest.get("variance_explanation", {}).get("storage") == "weekly_report" else legacy_explanation_path
    if variance_explanation is not None and explanation_source is not None and explanation_source.exists():
        saved = read_variance_explanation_workbook(explanation_source, expected_store=store,
            expected_week_start=activity.report_begin, expected_week_end=activity.report_end, require_text=False)
        mismatch = variance_control_mismatch_details(saved, **variance_explanation.controls)
        if mismatch and not existing:
            raise ParseError("Existing weekly variance explanation does not match the current weekly controls: " + "; ".join(mismatch))
        if not mismatch:
            variance_explanation = replace(variance_explanation, explanation=saved.explanation)

    archive_root = Path(archive_root)
    created_directories: list[Path] = []
    _ensure_directory(archive_root, created_directories)
    stage_root = _create_staging_root(archive_root)
    phase = "building staged evidence"
    transaction = PublicationTransaction(support)
    try:
        package_stage = stage_root / "package"
        archived_activity_path = package_stage / "activity" / activity_path.name
        archived_report = package_stage / "report" / output_path.name
        evidence_path = package_stage / "pos" / "weekly_pos_tender_evidence.csv"
        for directory in (archived_activity_path.parent, archived_report.parent, evidence_path.parent):
            directory.mkdir(parents=True, exist_ok=True)
        shutil.copy2(activity_path, archived_activity_path)
        if sha256_file(archived_activity_path) != activity_source.sha256:
            raise ValueError("Archived Activity copy does not match the calculation source.")
        write_reconciliation_workbook(prepared.result, archived_report,
            weekly_variance_explanation_path=explanation_path, weekly_variance_explanation=variance_explanation)
        _write_daily_evidence(evidence_path, activity_evidence=activity_evidence, micros=prepared.micros)
        _verify_source_stability(prepared.source_fingerprints, "building the report and evidence package")
        status = "REVIEW" if prepared.review_items else "PASS"
        manifest = _build_manifest(
            store=store, period=period, activity=activity, micros=prepared.micros,
            source_fingerprints=prepared.source_fingerprints, result=prepared.result,
            tender_variance=prepared.tender_variance, review_items=prepared.review_items, status=status,
            generated_at=datetime.now(timezone.utc), package_path=package_path,
            archived_activity=archived_activity_path, archived_report=archived_report,
            evidence_path=evidence_path, output_path=output_path, monthly_path=monthly_path,
            variance_explanation=variance_explanation, explanation_path=explanation_path,
        )
        manifest.update(input_set_id=input_id, result_id=result_id, revision=revision,
                        previous_result_id=previous_id,
                        previous_revision=str(previous_revision) if previous_revision else None)
        _write_json(package_stage / "weekly_manifest.json", manifest)
        _verify_staged_package(package_stage, manifest)
        output_temp = stage_root / "canonical.xlsx"
        monthly_temp = stage_root / ("monthly" + activity_path.suffix)
        shutil.copy2(archived_report, output_temp)
        shutil.copy2(archived_activity_path, monthly_temp)
        if previous_revision is not None:
            previous_stage = stage_root / "previous"
            shutil.copytree(package_path, previous_stage)
            _require_file_state(previous_stage, preimages[package_path], "preserving the previous evidence package")
            shutil.copy2(output_path, previous_stage / "operator_workbook.xlsx")
            _require_file_state(previous_stage / "operator_workbook.xlsx", preimages[output_path],
                                "preserving the previous operator workbook")
            transaction.replace(previous_stage, previous_revision, expected_before=None)
        phase = "committing evidence package"
        transaction.replace(package_stage, package_path, expected_before=preimages[package_path])
        phase = "committing canonical workbook"
        transaction.replace(output_temp, output_path, expected_before=preimages[output_path])
        phase = "committing monthly-close activity"
        if old_monthly_path is not None and old_monthly_path != monthly_path:
            transaction.remove(old_monthly_path, expected_before=preimages[old_monthly_path])
        transaction.replace(monthly_temp, monthly_path, expected_before=preimages[monthly_path])
        phase = "verifying published artifacts"
        _verify_published_artifacts(package_path=package_path, output_path=output_path,
                                   monthly_path=monthly_path, manifest=manifest)
        if explanation_path is not None:
            read_variance_explanation_workbook(explanation_path, expected_store=store,
                expected_week_start=activity.report_begin, expected_week_end=activity.report_end, require_text=False)
        _verify_source_stability(prepared.source_fingerprints, "publishing the weekly reconciliation")
        warnings = transaction.commit(cleanup=(activity_path, stage_root),
                                      expected_cleanup={activity_path: activity_source.sha256})
        message = f"Created {output_path.name}; revision {revision}; archived weekly evidence and staged monthly Activity."
        if explanation_path is not None:
            message += f" ACTION REQUIRED: enter an explanation in {EXPLANATION_SHEET}!{EXPLANATION_INPUT_CELL} for controls exceeding $5.00."
        if warnings:
            message += " " + " ".join(warnings)
        return WeeklyPublication(period=period, period_end=activity.report_end, status=status,
            output_path=output_path, archive_path=package_path, monthly_activity_path=monthly_path,
            variance_explanation_path=explanation_path, message=message, input_set_id=input_id,
            result_id=result_id, revision=revision, recovery_action=" ".join(warnings) or None)
    except (OSError, RuntimeError, ValueError, ParseError) as exc:
        if transaction.committed:
            recovery = f"Cleanup or recovery pending: {exc}. Recovery journal: {transaction.path}"
            return WeeklyPublication(period=period, period_end=activity.report_end, status=status,
                output_path=output_path, archive_path=package_path, monthly_activity_path=monthly_path,
                variance_explanation_path=explanation_path,
                message=f"Committed {output_path.name}; revision {revision}. {recovery}",
                input_set_id=input_id, result_id=result_id, revision=revision, recovery_action=recovery)
        cleanup_issues = []
        try:
            transaction.rollback()
        except (OSError, RuntimeError, ValueError) as rollback_error:
            cleanup_issues.append(str(rollback_error))
        cleanup_issues.extend(_cleanup_tree(stage_root))
        if not cleanup_issues:
            cleanup_issues.extend(_prune_created_directories(created_directories))
        detail = " Rollback incomplete; retained recovery journal: " + "; ".join(cleanup_issues) if cleanup_issues else " Rollback verified; no partial publication was retained."
        raise ParseError(f"Weekly publication failed while {phase}: {exc}." + detail) from exc


def _quarantine_duplicate(source, store, period, period_end, output_path, review_root, support, input_id, result_id):
    quarantine_dir = Path(review_root) / "duplicate-inputs" / store / period
    quarantine_dir.mkdir(parents=True, exist_ok=True)
    quarantine = quarantine_dir / f"{datetime.now():%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:8]}-{source.path.name}"
    staged = quarantine.with_suffix(quarantine.suffix + ".stage")
    try:
        shutil.copy2(source.path, staged)
        _verify_source_stability((source,), "preserving a semantic duplicate")
        if sha256_file(staged) != source.sha256:
            raise ParseError("Duplicate source changed during capture.")
        with PublicationTransaction(support) as transaction:
            transaction.replace(staged, quarantine)
            warnings = transaction.commit(cleanup=(source.path,), expected_cleanup={source.path: source.sha256})
    finally:
        staged.unlink(missing_ok=True)
    return WeeklyDuplicate(period=period, period_end=period_end, quarantined_path=quarantine,
        output_path=output_path, message=f"No change; preserved duplicate input at {quarantine}. " + " ".join(warnings),
        input_set_id=input_id, result_id=result_id)


def restore_weekly_revision(
    *, store: str, period: str, revision_id: str, output_root: Path,
    archive_root: Path, monthly_close_root: Path, correction_token: str | None = None,
) -> WeeklyPublication:
    """Select preserved evidence as a new revision, with one bound approval."""
    import re
    if not re.fullmatch(r"\d{4}-W\d{2}", period):
        raise ValueError("Weekly restore period must be YYYY-Www.")
    if not re.fullmatch(r"(?:r\d{4}-[0-9a-f]{16}|[0-9a-f]{64})", revision_id):
        raise ValueError("Select an existing revision folder name or full result ID.")
    year, week = int(period[:4]), int(period[-2:])
    ending = date.fromisocalendar(year, week, 7)
    beginning = date.fromordinal(ending.toordinal() - 6)
    config = get_store_config(store)
    store_folder = f"{store} {config.location_name}"
    package = Path(archive_root) / store_folder / str(year) / period
    output = Path(output_root) / store_folder / str(year) / f"Gift_Card_Reconciliation_{store}_{period}.xlsx"
    support = support_root_for(archive_root)
    with writer_lock(support):
        recovery = recover_pending(support)
        preimages: dict[Path, str | None] = {}
        current = _load_existing_package(package_path=package, store=store, output_path=output,
            period=period, period_begin=beginning, period_end=ending, preimages=preimages)
        if current is None:
            raise ParseError("No current weekly publication exists for this restore.")
        current_manifest, current_activity, output = current
        matches = []
        for candidate in package.with_name(period + ".revisions").glob("r*-*"):
            if candidate.name == revision_id:
                matches.append(candidate)
            elif len(revision_id) == 64 and (candidate / "weekly_manifest.json").is_file():
                data = json.loads((candidate / "weekly_manifest.json").read_text(encoding="utf-8"))
                if data.get("result_id") == revision_id:
                    matches.append(candidate)
        if len(matches) != 1:
            raise ParseError(f"Expected one preserved revision matching {revision_id}; found {len(matches)}.")
        selected = matches[0]
        selected_manifest, selected_activity, saved_output = _load_existing_package(
            package_path=selected, store=store, output_path=selected / "operator_workbook.xlsx",
            period=period, period_begin=beginning, period_end=ending, preimages=preimages)
        current_input = _weekly_input_id(current_activity, current_manifest["daily_pos_tender"])
        current_id = current_manifest.get("result_id") or identity({"input_set_id": current_input, "rules": "weekly-v2"})
        selected_input = _weekly_input_id(selected_activity, selected_manifest["daily_pos_tender"])
        selected_result = identity({"input_set_id": selected_input, "rules": "weekly-v2"})
        proposal = CorrectionRequired(current_id, selected_input,
            {"restore_revision": selected.name, "previous_totals": current_manifest["totals"],
             "selected_totals": selected_manifest["totals"],
             "changed_transactions": activity_changes(current_activity, selected_activity)})
        if correction_token != proposal.token:
            raise proposal
        revision = int(current_manifest.get("revision", 1)) + 1
        selected_result = identity({"input_set_id": selected_input, "rules": "weekly-v2",
                                    "revision": revision, "previous_result_id": current_id})
        preserved = package.with_name(period + ".revisions") / f"r{revision-1:04d}-{current_id[:16]}"
        if preserved.exists():
            raise ParseError(f"Previous revision destination already exists: {preserved}")
        monthly_dir = Path(monthly_close_root) / store_folder / fiscal_period_for_date(ending).folder_name / "activity"
        old_staged = []
        for candidate in monthly_dir.glob("*"):
            if candidate.suffix.lower() not in {".xls", ".xlsx"} or candidate.name.startswith("~$"):
                continue
            candidate_before = file_state(candidate)
            parsed = parse_activity_file(candidate)
            _require_file_state(candidate, candidate_before, "validating monthly Activity for restore")
            if parsed.store == store and parsed.report_end == ending:
                if activity_identity(parsed) != activity_identity(current_activity):
                    raise ParseError(f"Monthly Activity changed independently: {candidate}")
                old_staged.append(candidate)
                preimages[candidate] = candidate_before
        if len(old_staged) > 1:
            raise ParseError("Multiple monthly Activity inputs match the selected restore week.")
        activity_record = selected_manifest["artifacts"]["archived_activity"]
        monthly_path = monthly_dir / Path(activity_record["relative_path"]).name
        if monthly_path.exists() and monthly_path not in old_staged:
            raise ParseError(f"Monthly staging destination conflicts with restore: {monthly_path}")
        if monthly_path not in preimages:
            preimages[monthly_path] = None
        watched = tuple(_capture_source_fingerprint(p) for p in (
            package / "weekly_manifest.json", output, selected / "weekly_manifest.json", saved_output))
        with tempfile.TemporaryDirectory(prefix="weekly-restore-", dir=support) as temp:
            stage = Path(temp)
            previous_stage = stage / "previous"
            shutil.copytree(package, previous_stage)
            _require_file_state(previous_stage, preimages[package], "preserving the current evidence package")
            shutil.copy2(output, previous_stage / "operator_workbook.xlsx")
            _require_file_state(previous_stage / "operator_workbook.xlsx", preimages[output],
                                "preserving the current operator workbook")
            selected_stage = stage / "selected"
            shutil.copytree(selected, selected_stage)
            _require_file_state(selected_stage, preimages[selected], "copying the selected revision")
            # A maintenance layout revision becomes this revision's explicit
            # report baseline. Original evidence remains in its preserved version.
            (selected_stage / "weekly_report_revision.json").unlink(missing_ok=True)
            archived_report = selected_stage / "report" / output.name
            shutil.copy2(saved_output, archived_report)
            _require_file_state(archived_report, preimages[saved_output], "copying the selected operator workbook")
            record = _artifact_record(archived_report, "report", Path("report") / output.name)
            restored_manifest = dict(selected_manifest)
            artifacts = dict(selected_manifest["artifacts"])
            artifacts["archived_workbook"] = record
            artifacts["canonical_workbook"] = {**record, "path": str(output.resolve())}
            artifacts["monthly_staged_activity"] = {**activity_record, "path": str(monthly_path.resolve())}
            restored_manifest.update(artifacts=artifacts, input_set_id=selected_input,
                result_id=selected_result, revision=revision, previous_result_id=current_id,
                previous_revision=str(preserved), restored_from_revision=selected.name,
                archive_path=str(package.resolve()), generated_at=datetime.now(timezone.utc).isoformat())
            _write_json(selected_stage / "weekly_manifest.json", restored_manifest)
            _verify_staged_package(selected_stage, restored_manifest)
            _verify_source_stability(watched, "preparing a previous revision restore")
            with PublicationTransaction(support) as tx:
                tx.replace(previous_stage, preserved, expected_before=None)
                tx.replace(selected_stage / "report" / output.name, output, expected_before=preimages[output])
                for old in old_staged:
                    if old != monthly_path:
                        tx.remove(old, expected_before=preimages[old])
                tx.replace(selected_stage / activity_record["relative_path"], monthly_path,
                           expected_before=preimages[monthly_path])
                tx.replace(selected_stage, package, expected_before=preimages[package])
                _verify_published_artifacts(package_path=package, output_path=output,
                    monthly_path=monthly_path, manifest=restored_manifest)
                warnings = tx.commit()
        return WeeklyPublication(period=period, period_end=ending, status=restored_manifest["status"],
            output_path=output, archive_path=package, monthly_activity_path=monthly_path,
            variance_explanation_path=output if restored_manifest.get("variance_explanation", {}).get("storage") == "weekly_report" else None,
            message=f"Restored {selected.name} as revision {revision}; previous result retained.",
            input_set_id=selected_input, result_id=selected_result, revision=revision,
            recovery_action=" ".join(recovery + warnings) or None)


def iso_week_period(period_end: date) -> str:
    iso = period_end.isocalendar()
    return f"{iso.year}-W{iso.week:02d}"


def _required_variance_explanation(
    prepared: WeeklyPreparation,
) -> WeeklyVarianceExplanation | None:
    line_by_metric = {line.metric: line for line in prepared.result.lines}

    def variance(metric: str) -> Decimal:
        value = line_by_metric[metric].pos_variance
        return money(value if value is not None else Decimal("0.00"))

    pos_issue = variance("Gift Card Issue / Activations")
    pos_payment = variance("Gift Card Payment / Redemptions")
    pos_net = variance("Net Gift Card Impact")
    tender = money(prepared.tender_variance)
    if max(abs(value) for value in (pos_issue, pos_payment, pos_net, tender)) <= REVIEW_VARIANCE_LIMIT:
        return None

    assert prepared.activity.report_begin is not None
    assert prepared.activity.report_end is not None
    return WeeklyVarianceExplanation(
        store=prepared.result.store,
        week_start=prepared.activity.report_begin,
        week_end=prepared.activity.report_end,
        pos_issue_variance=pos_issue,
        pos_payment_variance=pos_payment,
        pos_net_variance=pos_net,
        tender_variance=tender,
    )


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
    variance_explanation: WeeklyVarianceExplanation | None,
    explanation_path: Path | None,
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
        "variance_explanation": {
            "required": variance_explanation is not None,
            "threshold": _decimal_text(REVIEW_VARIANCE_LIMIT),
            "storage": "weekly_report" if explanation_path else None,
            "input_path": str(explanation_path.resolve()) if explanation_path else None,
            "input_sheet": EXPLANATION_SHEET if explanation_path else None,
            "input_cell": EXPLANATION_INPUT_CELL if explanation_path else None,
        },
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


def _load_existing_package(
    *, package_path: Path, store: str, output_path: Path, period: str,
    period_begin: date, period_end: date,
    preimages: dict[Path, str | None] | None = None,
):
    package_before = file_state(package_path)
    if preimages is not None:
        preimages[package_path] = package_before
    if package_before is None:
        return None
    try:
        manifest = json.loads((package_path / "weekly_manifest.json").read_text(encoding="utf-8"))
        _verify_staged_package(package_path, manifest)
        activity_record = manifest["artifacts"]["archived_activity"]
        _verify_activity_record_consistency(manifest["source_files"]["activity"], activity_record)
        workbook_record = manifest["artifacts"]["archived_workbook"]
        _verify_canonical_record_consistency(workbook_record, manifest["artifacts"]["canonical_workbook"])
        if manifest.get("schema_version") != 1 or manifest.get("store") != store or manifest.get("period") != period:
            raise ValueError("Manifest store/period identity does not match its archive folder.")
        if manifest.get("week") != {"start": period_begin.isoformat(), "end": period_end.isoformat()}:
            raise ValueError("Manifest week dates do not match the Activity report.")
        revised_baseline = resolve_weekly_report_revision(package_path, manifest)
        baseline = revised_baseline or _package_artifact_path(package_path, workbook_record, "archived_workbook")
        embedded = manifest.get("variance_explanation", {}).get("storage") == "weekly_report" or revised_baseline is not None
        output_states: dict[Path, str | None] = {}
        def verify(candidate):
            candidate_before = file_state(candidate)
            if embedded:
                verify_weekly_report_explanation_edit(candidate, baseline)
            else:
                verify_workbook_values(candidate, baseline)
            _require_file_state(candidate, candidate_before, "validating the current weekly workbook")
            output_states[candidate] = candidate_before
        if output_path.is_file():
            verify(output_path)
        else:
            matches = []
            for candidate in sorted(output_path.parent.glob("*.xlsx")):
                if candidate.name.startswith("~$"):
                    continue
                try:
                    verify(candidate)
                    matches.append(candidate)
                except (ValueError, OSError):
                    continue
            if len(matches) != 1:
                raise ValueError(f"Expected one matching current weekly workbook; found {len(matches)}.")
            output_path = matches[0]
        archived_activity = parse_activity_file(_package_artifact_path(package_path, activity_record, "archived_activity"))
        validate_activity_evidence([archived_activity], store=store, period_start=period_begin,
            period_end=period_end, expected_week_endings=[period_end])
        _verify_manifest_controls(package_path, manifest, archived_activity)
        _require_file_state(package_path, package_before, "validating the previous evidence package")
        _require_file_state(output_path, output_states[output_path], "validating the current weekly workbook")
        if preimages is not None:
            preimages[output_path] = output_states[output_path]
        return manifest, archived_activity, output_path
    except (OSError, ValueError, KeyError, TypeError, BadZipFile, json.JSONDecodeError) as exc:
        raise ParseError(f"Existing weekly evidence package is incomplete or invalid: {package_path}: {exc}") from exc


def _verify_manifest_controls(package_path, manifest, activity):
    """Bind readable manifest controls to its verified, durable daily evidence."""
    from gift_card_recon.reconcile import rollup_activity_file
    evidence = _package_artifact_path(package_path,
        manifest["artifacts"]["weekly_pos_tender_evidence"], "weekly_pos_tender_evidence")
    with evidence.open(encoding="utf-8", newline="") as stream:
        rows = list(csv.DictReader(stream))
    daily = [{"business_date": row["business_date"],
              "pos_gift_card_issue": row["pos_gift_card_issue"],
              "pos_gift_card_payment": row["pos_gift_card_payment"],
              "tender_gift_card_payment": row["tender_gift_card_payment"],
              "tender_date_observed": row["tender_date_observed"] == "true",
              "accepted_scheduled_closure": row["coverage_status"] == "accepted scheduled closure"}
             for row in rows]
    if daily != manifest["daily_pos_tender"]:
        raise ValueError("Manifest daily controls do not match the archived POS/tender evidence.")
    rollup = rollup_activity_file(activity, set())
    pos_issue = sum((Decimal(row["pos_gift_card_issue"]) for row in rows), Decimal("0"))
    pos_payment = sum((Decimal(row["pos_gift_card_payment"]) for row in rows), Decimal("0"))
    tender = sum((Decimal(row["tender_gift_card_payment"]) for row in rows), Decimal("0"))
    totals = {"activity_issue": _decimal_text(rollup.net_activations),
              "activity_payment": _decimal_text(abs(rollup.net_redemptions)),
              "pos_gift_card_issue": _decimal_text(pos_issue),
              "pos_gift_card_payment": _decimal_text(pos_payment),
              "tender_gift_card_payment": _decimal_text(tender),
              "tender_variance": _decimal_text(pos_payment - tender)}
    if manifest["totals"] != totals:
        raise ValueError("Manifest totals do not match archived Activity and POS/tender evidence.")
    expected_status = "REVIEW" if any((pos_issue - rollup.net_activations,
        pos_payment - abs(rollup.net_redemptions), pos_payment - tender)) else "PASS"
    if manifest["status"] != expected_status:
        raise ValueError("Manifest business disposition does not match archived evidence.")
    input_id = _weekly_input_id(activity, daily)
    if manifest.get("input_set_id", input_id) != input_id:
        raise ValueError("Manifest input-set identity does not match archived evidence.")


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


def _require_file_state(path: Path, expected: str | None, phase: str) -> None:
    if file_state(path) != expected:
        raise ParseError(f"Existing evidence changed while {phase}: {path}")


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


def _unlink_fingerprinted_source(source: SourceFingerprint) -> None:
    """Remove a processed live source only while it still matches its archived fingerprint."""

    def unlink_if_unchanged() -> None:
        try:
            stat = source.path.stat()
            digest = sha256_file(source.path)
        except FileNotFoundError:
            return
        if stat.st_size != source.size_bytes or digest != source.sha256:
            raise ParseError(
                "Source evidence changed while weekly reconciliation was running during "
                f"removing processed inbox activity: {source.path}"
            )
        source.path.unlink(missing_ok=True)

    _retry_transient_file_operation(unlink_if_unchanged)


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
    """Compatibility wrapper around the shared publication retry policy."""
    return retry_file_operation(operation)
