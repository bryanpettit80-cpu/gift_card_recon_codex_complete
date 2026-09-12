from datetime import date
from decimal import Decimal
import json
from pathlib import Path
import shutil

import pytest
from openpyxl import Workbook, load_workbook

from gift_card_recon import weekly_report_revision
from gift_card_recon.utils import sha256_file
from gift_card_recon.variance_explanation import WeeklyVarianceExplanation, add_variance_explanation_sheets
from gift_card_recon.weekly_report_revision import publish_weekly_report_revision, resolve_weekly_report_revision


@pytest.fixture
def revision_inputs(tmp_path: Path):
    package = tmp_path / "archive" / "2026-W35"
    original = package / "report" / "Gift_Card_Reconciliation_9355_2026-W35.xlsx"
    original.parent.mkdir(parents=True)
    workbook = Workbook()
    workbook.active.title = "Reconciliation"
    workbook.active["C6"] = 870
    workbook.active["A1"] = "Weekly gift card reconciliation"
    workbook.save(original)
    workbook.close()
    canonical = tmp_path / "weekly" / original.name
    canonical.parent.mkdir()
    shutil.copy2(original, canonical)
    manifest = {
        "schema_version": 1, "store": "9355", "period": "2026-W35",
        "week": {"start": "2026-08-24", "end": "2026-08-30"},
        "artifacts": {"archived_workbook": {
            "relative_path": "report/" + original.name,
            "sha256": sha256_file(original), "size_bytes": original.stat().st_size,
        }},
    }
    (package / "weekly_manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
    candidate = tmp_path / "candidate.xlsx"
    workbook = load_workbook(canonical)
    workbook.active["C6"] = 920
    workbook.security = None  # Excel-saved workbooks can omit workbook protection.
    add_variance_explanation_sheets(workbook, WeeklyVarianceExplanation(
        "9355", date(2026, 8, 24), date(2026, 8, 30), Decimal("0"), Decimal("35"),
        Decimal("-35"), Decimal("0"), "The reload is included in issuance; payment follow-up remains.",
    ))
    workbook.save(candidate)
    workbook.close()
    return package, original, canonical, candidate


def test_report_revision_retains_original_evidence_and_has_verified_baseline(revision_inputs):
    package, original, canonical, candidate = revision_inputs
    original_hash = sha256_file(original)
    manifest_bytes = (package / "weekly_manifest.json").read_bytes()
    sidecar = publish_weekly_report_revision(package, canonical, candidate)
    assert sidecar.is_file()
    assert sha256_file(original) == original_hash
    assert (package / "weekly_manifest.json").read_bytes() == manifest_bytes
    assert sha256_file(canonical) == sha256_file(candidate)
    baseline = resolve_weekly_report_revision(package)
    assert baseline is not None
    assert baseline.parent.parent.name == "report-revisions"
    assert sha256_file(baseline) == sha256_file(candidate)


def test_report_revision_accepts_reviewed_presentation_only_changes(revision_inputs):
    package, original, canonical, candidate = revision_inputs
    workbook = load_workbook(canonical)
    workbook.active.column_dimensions["C"].width = 35
    workbook.save(canonical)
    workbook.close()
    current_hash = sha256_file(canonical)
    assert current_hash != sha256_file(original)
    with pytest.raises(ValueError, match="differs from its original archive"):
        publish_weekly_report_revision(package, canonical, candidate)
    publish_weekly_report_revision(package, canonical, candidate, expected_canonical_sha256=current_hash)
    assert resolve_weekly_report_revision(package).is_file()


def test_report_revision_refuses_changed_values_even_with_reviewed_hash(revision_inputs):
    package, original, canonical, candidate = revision_inputs
    workbook = load_workbook(canonical)
    workbook.active["C6"] = 999
    workbook.save(canonical)
    workbook.close()
    current_hash = sha256_file(canonical)
    with pytest.raises(ValueError, match="values or formulas changed"):
        publish_weekly_report_revision(package, canonical, candidate, expected_canonical_sha256=current_hash)
    assert sha256_file(canonical) == current_hash
    assert not (package / "weekly_report_revision.json").exists()


def test_report_revision_rolls_back_current_presentation_when_sidecar_publish_fails(
    revision_inputs, monkeypatch,
):
    package, original, canonical, candidate = revision_inputs
    workbook = load_workbook(canonical)
    workbook.active.column_dimensions["C"].width = 35
    workbook.save(canonical)
    workbook.close()
    prior_hash = sha256_file(canonical)
    real_replace = weekly_report_revision.os.replace
    def fail_sidecar(source, destination):
        if Path(destination).name == "weekly_report_revision.json":
            raise PermissionError("Simulated Dropbox lock")
        return real_replace(source, destination)
    monkeypatch.setattr(weekly_report_revision.os, "replace", fail_sidecar)
    with pytest.raises(PermissionError, match="Dropbox lock"):
        publish_weekly_report_revision(package, canonical, candidate, expected_canonical_sha256=prior_hash)
    assert sha256_file(canonical) == prior_hash
    assert not (package / "weekly_report_revision.json").exists()
    assert not list((package / "report-revisions").rglob("*.xlsx"))


def test_report_revision_rejects_tampered_baseline_and_path_traversal(revision_inputs):
    package, original, canonical, candidate = revision_inputs
    sidecar = publish_weekly_report_revision(package, canonical, candidate)
    baseline = resolve_weekly_report_revision(package)
    baseline.write_bytes(b"changed evidence")
    with pytest.raises(ValueError, match="integrity check failed"):
        resolve_weekly_report_revision(package)
    payload = json.loads(sidecar.read_text(encoding="utf-8"))
    payload["revised_workbook"]["relative_path"] = "report-revisions/../" + original.name
    sidecar.write_text(json.dumps(payload), encoding="utf-8")
    with pytest.raises(ValueError, match="outside its revision folder"):
        resolve_weekly_report_revision(package)


@pytest.fixture
def embedded_revision_inputs(revision_inputs):
    package, original, canonical, candidate = revision_inputs
    shutil.copy2(candidate, original)
    shutil.copy2(original, canonical)
    manifest_path = package / "weekly_manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["artifacts"]["archived_workbook"].update(
        sha256=sha256_file(original), size_bytes=original.stat().st_size,
    )
    manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
    return package, original, canonical, candidate


def test_revision_preserves_operator_explanation_edit(embedded_revision_inputs):
    package, original, canonical, candidate = embedded_revision_inputs
    narrative = "Payment difference was traced to a separate settlement entry."
    for path in (canonical, candidate):
        workbook = load_workbook(path)
        workbook["Variance Explanation"]["B15"] = narrative
        workbook.save(path)
        workbook.close()
    original_hash = sha256_file(original)
    publish_weekly_report_revision(
        package, canonical, candidate, expected_canonical_sha256=sha256_file(canonical),
    )
    assert sha256_file(original) == original_hash
    for path in (canonical, resolve_weekly_report_revision(package)):
        workbook = load_workbook(path)
        try:
            assert workbook["Variance Explanation"]["B15"].value == narrative
        finally:
            workbook.close()


@pytest.mark.parametrize("replacement", [None, "An unrelated replacement explanation."])
def test_revision_rejects_candidate_that_loses_existing_explanation(
    embedded_revision_inputs, replacement,
):
    package, original, canonical, candidate = embedded_revision_inputs
    workbook = load_workbook(candidate)
    workbook["Variance Explanation"]["B15"] = replacement
    workbook.save(candidate)
    workbook.close()
    prior_hash = sha256_file(canonical)
    with pytest.raises(ValueError, match="retain the current operator explanation"):
        publish_weekly_report_revision(
            package, canonical, candidate, expected_canonical_sha256=prior_hash,
        )
    assert sha256_file(canonical) == prior_hash
    assert not (package / "weekly_report_revision.json").exists()


def test_embedded_revision_rejects_edits_outside_explanation(embedded_revision_inputs):
    package, original, canonical, candidate = embedded_revision_inputs
    workbook = load_workbook(canonical)
    workbook["Reconciliation"]["C6"] = 999
    workbook["Variance Explanation"]["B15"] = "Operator explanation."
    workbook.save(canonical)
    workbook.close()
    prior_hash = sha256_file(canonical)
    with pytest.raises(ValueError, match="values or formulas changed outside"):
        publish_weekly_report_revision(
            package, canonical, candidate, expected_canonical_sha256=prior_hash,
        )
    assert sha256_file(canonical) == prior_hash
    assert not (package / "weekly_report_revision.json").exists()


def test_failed_rollback_retains_verified_operator_report(revision_inputs, monkeypatch):
    package, original, canonical, candidate = revision_inputs
    workbook = load_workbook(canonical)
    workbook.active.column_dimensions["C"].width = 35
    workbook.save(canonical)
    workbook.close()
    prior_bytes = canonical.read_bytes()
    real_replace = weekly_report_revision.os.replace
    real_copy = weekly_report_revision.shutil.copy2

    def fail_sidecar(source, destination):
        if Path(destination).name == "weekly_report_revision.json":
            raise PermissionError("Simulated sidecar lock")
        return real_replace(source, destination)

    def fail_restore(source, destination):
        if Path(source).name.startswith(".weekly-before-revision-") and Path(destination) == canonical:
            raise OSError("Simulated full disk during rollback")
        return real_copy(source, destination)

    monkeypatch.setattr(weekly_report_revision.os, "replace", fail_sidecar)
    monkeypatch.setattr(weekly_report_revision.shutil, "copy2", fail_restore)
    with pytest.raises(RuntimeError, match="verified prior copy retained"):
        publish_weekly_report_revision(
            package, canonical, candidate, expected_canonical_sha256=sha256_file(canonical),
        )
    backups = list(canonical.parent.glob(".weekly-before-revision-*.xlsx"))
    assert len(backups) == 1
    assert backups[0].read_bytes() == prior_bytes
    assert sha256_file(canonical) == sha256_file(candidate)
    assert list((package / "report-revisions").rglob("*.xlsx"))
