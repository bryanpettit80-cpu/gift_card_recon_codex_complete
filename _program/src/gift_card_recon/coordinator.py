"""Read-only coordinator boundary. Intake, scheduling and messages live elsewhere.

Readiness reads already available source files. It never creates a folder,
extracts an archive, moves an input, recovers a transaction or publishes a report.
The execution entrypoints remain auto_run and monthly_close_service.
"""
from __future__ import annotations

import argparse
import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Mapping, Sequence

from gift_card_recon.auto_run import _activity_paths, _store_from_folder, weekly_input_inventory
from gift_card_recon.business_identity import (
    activity_changes, activity_identity, correction_token, identity, workbook_values_identity,
)
from gift_card_recon.fiscal_calendar import fiscal_period_for_label
from gift_card_recon.parsers import (ParseError, parse_activity_file, parse_summary, spreadsheet_paths,
    spreadsheet_role, summary_contains_store, summary_report_end)
from gift_card_recon.publication import support_root_for
from gift_card_recon.store_config import STORE_CONFIGS, get_operations_store_config
from gift_card_recon.utils import sha256_file


@dataclass(frozen=True)
class ReadinessResult:
    mode: str
    store: str
    period: str | None
    ready: bool
    disposition: str
    reasons: tuple[str, ...]
    input_set_id: str | None
    result_id: str
    artifact_paths: tuple[str, ...] = ()
    pdf_status: str = "not_requested"
    recovery_action: str | None = None
    correction_token: str | None = None
    changed_facts: Mapping | None = None
    schema_version: int = 1

    def to_dict(self) -> dict:
        return asdict(self)


def _result(mode: str, store: str, period: str | None, *, ready: bool = False,
            disposition: str = "BLOCKED", reasons: Sequence[str] = (), input_set_id: str | None = None,
            recovery_action: str | None = None, correction_token: str | None = None,
            changed_facts: Mapping | None = None, artifact_paths: Sequence[str] = (),
            pdf_status: str = "not_requested") -> ReadinessResult:
    result_id = identity({"kind": "readiness-v1", "mode": mode, "store": store, "period": period,
                          "input_set_id": input_set_id, "ready": ready, "disposition": disposition,
                          "reasons": list(reasons), "recovery_action": recovery_action,
                          "correction_token": correction_token, "changed_facts": changed_facts})
    return ReadinessResult(mode, store, period, ready, disposition, tuple(reasons), input_set_id,
                           result_id, artifact_paths=tuple(artifact_paths), pdf_status=pdf_status,
                           recovery_action=recovery_action, correction_token=correction_token,
                           changed_facts=changed_facts)


def _pending_recovery(archive_root: Path) -> tuple[str, ...]:
    transactions = support_root_for(archive_root) / "transactions"
    reasons: list[str] = []
    for path in sorted(transactions.glob("*/transaction.json")):
        try:
            state = json.loads(path.read_text(encoding="utf-8"))["state"]
        except (OSError, ValueError, KeyError, TypeError) as exc:
            reasons.append(f"Publication recovery record cannot be read: {path.parent.name}: {exc}")
            continue
        if state not in {"completed", "rolled_back"}:
            reasons.append(f"Publication transaction {path.parent.name} requires recovery ({state}).")
    return tuple(reasons)


def _monthly_live_sources(*, store: str, fiscal, input_root: Path, intended_input: Path) -> dict[str, list[Path]]:
    """Select the same waiting intake used by staging, without creating its destination."""
    from gift_card_recon.monthly_close import weekly_activity_source_dirs

    summaries: list[Path] = []
    activities: list[Path] = []
    for folder in (intended_input, intended_input / "summary", intended_input / "activity"):
        for path in spreadsheet_paths(folder):
            role = spreadsheet_role(path)
            if role == "summary":
                summaries.append(path)
            elif role == "activity":
                activities.append(path)
            elif folder != intended_input:
                raise ParseError(f"Unrecognized report in {folder}: {path.name}.")
    for path in spreadsheet_paths(intended_input.parent):
        if spreadsheet_role(path) != "summary":
            continue
        end = summary_report_end(path)
        if end is None:
            raise ParseError(f"Summary {path.name} has no report period or recorded intake date.")
        if end == fiscal.end_date and summary_contains_store(path, store):
            summaries.append(path)
    for folder in weekly_activity_source_dirs(input_root=input_root, store=store):
        for path in spreadsheet_paths(folder):
            if spreadsheet_role(path) != "activity":
                continue
            activity = parse_activity_file(path)
            if activity.report_end is None:
                raise ParseError(f"Activity {path.name} has no report period.")
            if fiscal.start_date <= activity.report_end <= fiscal.end_date:
                activities.append(path)
    # Equivalent deliveries can wait in both a weekly inbox and the month folder.
    # Preserve every transaction within each report; deduplicate only whole reports.
    summaries_by_id: dict[str, Path] = {}
    for path in dict.fromkeys(summaries):
        summaries_by_id.setdefault(identity(parse_summary(path, store)), path)
    activities_by_id: dict[str, Path] = {}
    for path in dict.fromkeys(activities):
        activities_by_id.setdefault(activity_identity(parse_activity_file(path)), path)
    if len(summaries_by_id) != 1:
        raise ParseError(f"Expected exactly one Summary for store {store} {fiscal.period_key}; found {len(summaries_by_id)}.")
    return {"Gift Card Summary": list(summaries_by_id.values()),
            "Weekly Gift Card Activity": list(activities_by_id.values())}


def assess_weekly_readiness(*, operations_root: Path, input_root: Path | None = None,
                            stores: Sequence[str] | None = None, dropbox_root: Path | None = None,
                            archive_root: Path | None = None, output_root: Path | None = None) -> list[ReadinessResult]:
    from gift_card_recon.weekly_service import (_load_existing_package, _prepare_activity_evidence,
        _prepare_weekly_from_activity, _weekly_input_id, _weekly_totals, historical_controls_available,
        iso_week_period, weekly_preparation_identity)

    operations_root = Path(operations_root)
    input_root = Path(input_root) if input_root is not None else operations_root / "01 Weekly Gift Card Activity Reports"
    archive_root = Path(archive_root) if archive_root is not None else operations_root / "04 Archive" / "Weekly Reconciliation"
    output_root = Path(output_root) if output_root is not None else operations_root / "03 Finished Reports" / "Weekly"
    wanted = {str(store) for store in stores} if stores else set(STORE_CONFIGS)
    found, issues = weekly_input_inventory(input_root, wanted)
    results = [_result("weekly", store, None, reasons=[message], recovery_action="resolve_input")
               for store, message in issues]
    recovery = _pending_recovery(archive_root)
    for store, folder in found:
        period = None
        try:
            if recovery:
                results.append(_result("weekly", store, None, reasons=recovery, recovery_action="run_recovery"))
                continue
            activity_dir = folder if folder.name.casefold() == "activity" else folder / "activity"
            if not activity_dir.is_dir():
                raise ParseError(f"Required Activity folder is missing: {activity_dir}.")
            paths = _activity_paths(activity_dir)
            if not paths:
                results.append(_result("weekly", store, None, disposition="NO_INPUT",
                                       reasons=["No Activity report is waiting."], recovery_action="supply_activity"))
                continue
            if len(paths) != 1:
                raise ParseError(f"Expected exactly one weekly Activity report; found {len(paths)}.")
            config = get_operations_store_config(store, operations_root=operations_root, dropbox_root=dropbox_root)
            activity, activity_evidence, source = _prepare_activity_evidence(store=store, activity_path=paths[0])
            period = iso_week_period(activity.report_end)
            store_folder = f"{store} {config.location_name}"
            year = str(activity.report_end.isocalendar().year)
            output_path = output_root / store_folder / year / f"Gift_Card_Reconciliation_{store}_{period}.xlsx"
            previous = _load_existing_package(package_path=archive_root / store_folder / year / period,
                                             store=store, output_path=output_path, period=period,
                                             period_begin=activity.report_begin, period_end=activity.report_end)
            previous_input_id = previous_result_id = None
            if previous:
                manifest, archived_activity, output_path = previous
                previous_input_id = _weekly_input_id(archived_activity, manifest["daily_pos_tender"])
                previous_result_id = manifest.get("result_id") or identity({"input_set_id": previous_input_id, "rules": "weekly-v2"})
                unchanged_activity = activity_identity(activity) == activity_identity(archived_activity)
                live_present = historical_controls_available(config, activity) if unchanged_activity else True
                if unchanged_activity and not live_present:
                    results.append(_result("weekly", store, period, ready=True, disposition="DUPLICATE",
                                           input_set_id=previous_input_id, artifact_paths=[str(output_path)]))
                    continue
            data = _prepare_weekly_from_activity(store=store, config=config, activity=activity,
                                                 activity_evidence=activity_evidence, activity_source=source)
            input_id = weekly_preparation_identity(data)
            if previous:
                if input_id == previous_input_id:
                    results.append(_result("weekly", store, period, ready=True, disposition="DUPLICATE",
                                           input_set_id=input_id, artifact_paths=[str(output_path)]))
                else:
                    changes = {name: {"previous": manifest["totals"].get(name), "current": value}
                               for name, value in _weekly_totals(data).items() if manifest["totals"].get(name) != value}
                    changes["activity_records"] = {"previous": len(archived_activity.rows), "current": len(activity.rows)}
                    changes["transactions"] = activity_changes(archived_activity, activity)
                    results.append(_result("weekly", store, period, disposition="CORRECTION_REQUIRED",
                                           reasons=["The changed input set requires one deliberate replacement decision."],
                                           input_set_id=input_id, recovery_action="approve_correction",
                                           correction_token=correction_token(previous_result_id, input_id), changed_facts=changes))
                continue
            results.append(_result("weekly", store, period, ready=True,
                                   disposition="REVIEW" if data.review_items else "PASS",
                                   reasons=data.review_items, input_set_id=input_id))
        except (RuntimeError, OSError, ValueError) as exc:
            results.append(_result("weekly", store, period, reasons=[str(exc)], recovery_action="resolve_input"))
    return sorted(results, key=lambda item: item.store)


def assess_monthly_readiness(*, operations_root: Path, period: str, stores: Sequence[str] | None = None,
                             input_root: Path | None = None, dropbox_root: Path | None = None,
                             archive_root: Path | None = None, output_root: Path | None = None,
                             input_dirs: Mapping[str, Path] | None = None,
                             micros_paths: Mapping[str, Path] | None = None,
                             darden_paths: Mapping[str, Path] | None = None,
                             allow_unconfigured_micros: bool = False) -> list[ReadinessResult]:
    from gift_card_recon.monthly_close_service import (prepare_monthly_close_readiness,
        _read_previous_close, _legacy_monthly_identity, _monthly_changed_facts,
        canonical_output_paths, resolve_monthly_workbook, _resolve_monthly_pdf,
        _previous_archive_records)
    from gift_card_recon.monthly_close_cli import SHARED_DARDEN_INBOX, discover_close_jobs

    operations_root = Path(operations_root)
    input_root = Path(input_root) if input_root is not None else operations_root / "02 Monthly Close Inputs"
    archive_root = Path(archive_root) if archive_root is not None else operations_root / "04 Archive"
    output_root = Path(output_root) if output_root is not None else operations_root / "03 Finished Reports"
    wanted = sorted({str(store) for store in stores} if stores else STORE_CONFIGS)
    results: list[ReadinessResult] = []
    recovery = _pending_recovery(archive_root)
    for store in wanted:
        canonical_period = period
        try:
            fiscal = fiscal_period_for_label(period)
            canonical_period = fiscal.period_key
            if recovery:
                results.append(_result("monthly", store, canonical_period, reasons=recovery, recovery_action="run_recovery"))
                continue
            config = get_operations_store_config(store, operations_root=operations_root, dropbox_root=dropbox_root)
            chosen_input = (input_dirs or {}).get(store)
            live_sources = None
            if chosen_input is None:
                folders = [folder for folder in input_root.iterdir()
                           if folder.is_dir() and _store_from_folder(folder) == store] if input_root.is_dir() else []
                if len(folders) != 1:
                    raise ParseError(f"Expected exactly one monthly input folder for store {store}; found {len(folders)}.")
                periods = [folder for folder in folders[0].iterdir() if folder.is_dir()
                           and folder.name in {fiscal.folder_name, fiscal.period_key, period}]
                if len(periods) > 1:
                    raise ParseError(f"Expected exactly one input folder for {fiscal.period_key}; found {len(periods)}.")
                chosen_input = periods[0] if periods else folders[0] / fiscal.folder_name
                live_sources = _monthly_live_sources(store=store, fiscal=fiscal, input_root=input_root,
                                                     intended_input=chosen_input)
            chosen_darden = (darden_paths or {}).get(store)
            if chosen_darden is None and not list((Path(chosen_input) / "darden").glob("*.pdf")):
                jobs, issues = discover_close_jobs(inbox=input_root / SHARED_DARDEN_INBOX,
                                                  explicit_darden=None, store=store, period=canonical_period)
                if issues:
                    raise ParseError("; ".join(issue.message for issue in issues))
                if len(jobs) == 1:
                    chosen_darden = jobs[0].darden_path
            data = prepare_monthly_close_readiness(
                store=store, store_config=config, period=canonical_period, input_dir=Path(chosen_input),
                micros_path=(micros_paths or {}).get(store, config.micros_default_path),
                micros_work_dir=operations_root / "_automation_runs" / "support" / "unused-readiness-extraction",
                archive_root=archive_root, darden_path=chosen_darden, output_root=output_root,
                allow_unconfigured_micros=allow_unconfigured_micros,
                live_source_paths=live_sources,
            )
            manifest_path = archive_root / "Monthly Close" / store / fiscal.folder_name / "close_manifest.json"
            previous = _read_previous_close(manifest_path, store=store, period=canonical_period)
            if previous:
                expected_workbook, expected_pdf = canonical_output_paths(output_root, config=config, fiscal_period=fiscal)
                workbook = resolve_monthly_workbook(previous, expected_workbook)
                previous_input_id = previous.get("input_set_id") or _legacy_monthly_identity(
                    config=config, fiscal_period=fiscal, archive_root=archive_root,
                    micros_work_dir=operations_root / "_automation_runs" / "support" / "unused-readiness-extraction")
                previous_result_id = previous.get("result_id") or sha256_file(manifest_path)
                unchanged_report = (not previous.get("workbook_values_identity")
                                    or workbook_values_identity(workbook) == previous["workbook_values_identity"])
                if previous_input_id != data.input_set_id or not unchanged_report:
                    changes = _monthly_changed_facts(previous, data)
                    if not unchanged_report:
                        changes["report_explanations"] = (
                            "The current report narrative differs from the recorded result; "
                            "publishing the selected evidence will replace that narrative."
                        )
                    results.append(_result("monthly", store, canonical_period, disposition="CORRECTION_REQUIRED",
                                           reasons=["The changed input set requires one deliberate replacement decision.",
                                                    *data.assessment.follow_up_items],
                                           input_set_id=data.input_set_id, recovery_action="approve_correction",
                                           correction_token=correction_token(previous_result_id, data.input_set_id),
                                           changed_facts=changes))
                    continue
                _previous_archive_records(previous, archive_root=archive_root)
                artifact_paths = [str(workbook), str(manifest_path)]
                pdf_status = "pending"
                pdf_record = next((item for item in previous.get("artifacts", []) if item.get("role") == "pdf"), None)
                if previous.get("pdf_status", "complete") == "complete" and pdf_record:
                    try:
                        pdf = _resolve_monthly_pdf(previous, expected_pdf)
                        if pdf.is_file() and sha256_file(pdf) == pdf_record.get("sha256"):
                            pdf_status = "complete"
                            artifact_paths.append(str(pdf))
                    except OSError:
                        pass  # Workbook completion does not depend on PDF availability.
                results.append(_result("monthly", store, canonical_period, ready=True, disposition="DUPLICATE",
                                       input_set_id=data.input_set_id, artifact_paths=artifact_paths,
                                       pdf_status=pdf_status, recovery_action="retry_pdf" if pdf_status == "pending" else None))
                continue
            results.append(_result("monthly", store, canonical_period, ready=data.assessment.can_publish_close,
                                   disposition=str(data.assessment.status), reasons=data.assessment.follow_up_items,
                                   input_set_id=data.input_set_id,
                                   recovery_action=None if data.assessment.can_publish_close else "resolve_controls"))
        except (RuntimeError, OSError, ValueError) as exc:
            results.append(_result("monthly", store, canonical_period, reasons=[str(exc)], recovery_action="resolve_input"))
    return results


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Read-only gift-card readiness. No inputs are staged or published.")
    parser.add_argument("mode", choices=("weekly", "monthly"))
    parser.add_argument("--operations-root", type=Path, default=Path("."))
    parser.add_argument("--dropbox-root", type=Path)
    parser.add_argument("--input-root", type=Path)
    parser.add_argument("--archive-root", type=Path)
    parser.add_argument("--output-root", type=Path)
    parser.add_argument("--store", action="append")
    parser.add_argument("--period")
    args = parser.parse_args(argv)
    for name in ("input_root", "archive_root", "output_root"):
        value = getattr(args, name)
        if value is not None and not value.is_absolute():
            setattr(args, name, args.operations_root / value)
    kwargs = dict(operations_root=args.operations_root, dropbox_root=args.dropbox_root,
                  input_root=args.input_root, archive_root=args.archive_root, output_root=args.output_root, stores=args.store)
    if args.mode == "monthly":
        if not args.period:
            parser.error("--period is required for monthly readiness")
        results = assess_monthly_readiness(period=args.period, **kwargs)
    else:
        results = assess_weekly_readiness(**kwargs)
    print(json.dumps([result.to_dict() for result in results], sort_keys=True))
    return int(any(not result.ready and result.disposition != "NO_INPUT" for result in results))


if __name__ == "__main__":
    raise SystemExit(main())
