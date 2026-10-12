from __future__ import annotations

import errno
import hashlib
import json
import os
from dataclasses import replace
from datetime import date, datetime, timedelta
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook
import pytest

from gift_card_recon.business_identity import CorrectionRequired
from gift_card_recon.close_assessment import CloseStatus
from gift_card_recon.evidence_archive import ArchiveError, write_close_manifest_atomic
from gift_card_recon.fiscal_calendar import fiscal_period_for_label
from gift_card_recon.models import DardenCreditMemo
from gift_card_recon.monthly_close_service import (
    ArchivedVarianceExplanationSource,
    CloseBlockedError,
    _publication_staging_directory,
    assess_monthly_close_inputs,
    canonical_output_paths,
    review_output_paths,
    run_monthly_close_service,
    retry_monthly_pdf,
)
from gift_card_recon.pdf_export import PdfExportError
from gift_card_recon.parsers import ParseError
from gift_card_recon.store_config import get_store_config
from gift_card_recon.utils import sha256_file
from gift_card_recon.variance_explanation import (
    WeeklyVarianceExplanation,
    variance_explanation_path,
    write_variance_explanation_workbook,
)


@pytest.mark.parametrize(
    ("store", "payment_variances", "expected_status", "expected_period_net"),
    [
        (
            "9354",
            [Decimal("-0.01"), Decimal("0.00"), Decimal("2.44"), Decimal("0.00"), Decimal("0.00")],
            CloseStatus.CLOSED_WITH_REVIEW,
            Decimal("-2.43"),
        ),
        (
            "9355",
            [Decimal("0.00")] * 5,
            CloseStatus.CLOSED,
            Decimal("0.00"),
        ),
    ],
)
def test_realistic_five_week_close_status_and_report_flow(
    tmp_path: Path,
    store: str,
    payment_variances: list[Decimal],
    expected_status: CloseStatus,
    expected_period_net: Decimal,
) -> None:
    setup = _build_period(tmp_path, store=store, payment_variances=payment_variances)

    run = run_monthly_close_service(
        store=store,
        period=setup["period"].period_key,
        input_dir=setup["input_dir"],
        micros_path=setup["micros_dir"],
        micros_work_dir=tmp_path / "extract",
        archive_root=setup["archive_root"],
        output_root=setup["output_root"],
        darden_report=setup["darden"],
        fiscal_period=setup["period"],
        cleanup_sources=False,
        allow_unconfigured_micros=True,
        generated_at=datetime(2026, 7, 6, 9, 30),
        pdf_exporter=_fake_pdf_exporter,
    )

    assert run.assessment.status is expected_status
    assert sum((row.net_variance for row in run.weekly_variances), Decimal("0.00")) == expected_period_net
    assert run.workbook_path.exists()
    assert run.pdf_path.exists()
    expected_xlsx, expected_pdf = canonical_output_paths(
        setup["output_root"],
        config=get_store_config(store),
        fiscal_period=setup["period"],
    )
    assert (run.workbook_path, run.pdf_path) == (expected_xlsx, expected_pdf)

    workbook = load_workbook(run.workbook_path, data_only=False)
    report = workbook["Monthly Close Report"]
    assert report["A1"].value == get_store_config(store).report_heading
    assert report["A4"].value == expected_status.value
    assert report["G9"].value == "MATCHED"
    assert len(report.row_breaks.brk) == 1
    assert report.page_setup.fitToHeight == 0
    assert report.page_setup.scale == 85
    reconciliation_sheet = workbook["Reconciliation"]
    assert reconciliation_sheet["A3"].value.startswith(
        f"Authoritative monthly close status: {expected_status.value}."
    )
    assert "does not determine overall close status" in reconciliation_sheet["A3"].value
    assert any(
        reconciliation_sheet.cell(row, 1).value == "Darden Settlement Control"
        for row in range(1, reconciliation_sheet.max_row + 1)
    )
    source_sheet = workbook["Source Files"]
    source_paths = [
        source_sheet.cell(row, 6).value for row in range(4, source_sheet.max_row + 1)
    ]
    normalized_source_paths = [str(value).replace("\\", "/") for value in source_paths]
    assert all(value.startswith(f"Monthly Close/{store}") for value in normalized_source_paths)
    assert any(str(value).endswith("DLYSYSTT.TXT") for value in source_paths)
    assert any(str(value).endswith("TENDER_DETAIL.TXT") for value in source_paths)
    exception_sheet = workbook["Exception Log"]
    if expected_status is CloseStatus.CLOSED_WITH_REVIEW:
        assert exception_sheet["A4"].value == "REVIEW"
        assert "No parsing or validation exceptions" not in str(exception_sheet["B4"].value)

    manifest = json.loads(run.manifest_path.read_text(encoding="utf-8"))
    assert manifest["status"] == expected_status.value
    assert len(manifest["sources"]) == 9
    assert {item["role"] for item in manifest["artifacts"]} == {"workbook", "pdf"}
    assert all(len(item["sha256"]) == 64 for item in manifest["sources"])


def test_relocated_store_config_is_used_for_strict_source_validation(tmp_path: Path) -> None:
    setup = _build_period(tmp_path, store="9355")
    relocated = replace(
        get_store_config("9355"),
        micros_default_path=Path(setup["micros_dir"]).resolve(strict=False),
    )

    run = _run(
        setup,
        store_config=relocated,
        allow_unconfigured_micros=False,
    )

    assert run.assessment.status is CloseStatus.CLOSED


def test_relocated_store_config_is_used_for_prepare_only_validation(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    setup = _build_period(tmp_path, store="9355")
    monkeypatch.setattr(
        "gift_card_recon.monthly_close_service.parse_darden_credit_memo",
        lambda _path: setup["darden"],
    )
    relocated = replace(
        get_store_config("9355"),
        micros_default_path=Path(setup["micros_dir"]).resolve(strict=False),
    )

    assessment = assess_monthly_close_inputs(
        store="9355",
        store_config=relocated,
        period=setup["period"].period_key,
        input_dir=setup["input_dir"],
        micros_path=setup["micros_dir"],
        micros_work_dir=tmp_path / "extract",
        archive_root=setup["archive_root"],
        darden_path=setup["darden"].source_file,
        fiscal_period=setup["period"],
        allow_unconfigured_micros=False,
    )

    assert assessment.status is CloseStatus.CLOSED


def test_relocated_store_config_rejects_a_different_store(tmp_path: Path) -> None:
    setup = _build_period(tmp_path, store="9355")
    wrong_store = replace(
        get_store_config("9354"),
        micros_default_path=Path(setup["micros_dir"]).resolve(strict=False),
    )

    with pytest.raises(ValueError, match="does not match requested store"):
        _run(setup, store_config=wrong_store)




def test_publication_staging_ignores_system_temp_and_stays_on_destination_volume(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    system_temp = tmp_path / "system-temp"
    system_temp.mkdir()
    monkeypatch.setattr("gift_card_recon.monthly_close_service.tempfile.tempdir", str(system_temp))
    destination_dir = tmp_path / "shared-reports" / "Monthly Close" / "FY27 M01"
    destination_dir.mkdir(parents=True)

    with _publication_staging_directory(
        destination_dir,
        prefix="monthly-close-",
    ) as temp_dir:
        staging_dir = Path(temp_dir)
        assert staging_dir.parent == destination_dir.parent
        assert destination_dir not in staging_dir.parents
        assert system_temp not in staging_dir.parents
        assert staging_dir.stat().st_dev == destination_dir.stat().st_dev
        assert staging_dir.name.startswith(".gift-card-monthly-close-")

    assert not staging_dir.exists()






def test_monthly_manifest_publication_retries_transient_dropbox_lock(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from gift_card_recon import monthly_close_service as service

    setup = _build_period(tmp_path, store="9355")
    manifest_path = (
        setup["archive_root"]
        / "Monthly Close"
        / "9355"
        / setup["period"].folder_name
        / "close_manifest.json"
    )
    real_replace = service.os.replace
    manifest_attempts = 0
    retry_delays: list[float] = []

    def transient_first_manifest_replace(source: Path | str, destination: Path | str) -> None:
        nonlocal manifest_attempts
        if Path(destination) == manifest_path:
            manifest_attempts += 1
            if manifest_attempts == 1:
                raise _sharing_violation("simulated transient Dropbox manifest lock")
        real_replace(source, destination)

    monkeypatch.setattr(service.os, "replace", transient_first_manifest_replace)
    monkeypatch.setattr("gift_card_recon.publication.time.sleep", retry_delays.append)

    run = _run(setup)

    assert manifest_attempts == 2
    assert retry_delays == [0.10]
    assert run.manifest_path == manifest_path
    assert manifest_path.is_file()
    assert run.workbook_path.is_file()
    assert run.pdf_path.is_file()












def test_review_output_paths_use_monthly_specific_operator_folder(tmp_path: Path) -> None:
    output_root = tmp_path / "03 Finished Reports"
    workbook, pdf = review_output_paths(
        output_root,
        config=get_store_config("9355"),
        fiscal_period=fiscal_period_for_label("FY27-M01"),
    )

    expected_folder = output_root / "Monthly Close - Review Required"
    assert workbook == expected_folder / "Virginia_Beach_9355_FY27-M01_Review_Required.xlsx"
    assert pdf == expected_folder / "Virginia_Beach_9355_FY27-M01_Review_Required.pdf"
    assert workbook.parent != output_root / "Review Required"


def test_missing_required_weekly_variance_explanation_blocks_with_actionable_path(
    tmp_path: Path,
) -> None:
    setup = _build_period(
        tmp_path,
        store="9355",
        issue_variances=[Decimal("7.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00")],
    )
    week_end = setup["period"].expected_week_endings[0]
    expected_path = variance_explanation_path(setup["input_dir"], "9355", week_end)
    expected_message = (
        "Weekly variance explanation is required for store 9355 week ending "
        "06/07/2026 because POS issue +7.00, POS net +7.00 exceeds the $5.00 "
        f"limit. Expected {expected_path}. The weekly runner creates this form automatically."
    )

    with pytest.raises(CloseBlockedError) as exc_info:
        _run(setup)

    assert expected_message in str(exc_info.value)
    assert exc_info.value.assessment.status is CloseStatus.REVIEW_REQUIRED
    assert exc_info.value.review_workbook is not None
    assert not expected_path.exists()
    report = load_workbook(exc_info.value.review_workbook)["Monthly Close Report"]
    values = [cell.value for row in report for cell in row]
    assert "Weekly Variance Detail" in values
    assert 7 in values


def test_blank_required_weekly_variance_explanation_blocks_on_input_cell(
    tmp_path: Path,
) -> None:
    setup = _build_period(
        tmp_path,
        store="9355",
        issue_variances=[Decimal("7.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00")],
    )
    explanation_path = _write_weekly_variance_explanation(
        setup,
        issue_variance=Decimal("7.00"),
        explanation="",
    )

    with pytest.raises(CloseBlockedError) as exc_info:
        _run(setup)

    assert "Variance Explanation!B15" in str(exc_info.value)
    assert str(explanation_path) in str(exc_info.value)
    assert exc_info.value.assessment.status is CloseStatus.REVIEW_REQUIRED
    assert explanation_path.exists()


def test_stale_weekly_variance_explanation_controls_block_monthly_close(
    tmp_path: Path,
) -> None:
    setup = _build_period(
        tmp_path,
        store="9355",
        issue_variances=[Decimal("7.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00")],
    )
    explanation_path = _write_weekly_variance_explanation(
        setup,
        issue_variance=Decimal("6.00"),
        explanation="Earlier discrepancy under review.",
    )

    with pytest.raises(CloseBlockedError) as exc_info:
        _run(setup)

    message = str(exc_info.value)
    assert "controls do not match current source evidence" in message
    assert "POS gift card issue: form +6.00, current +7.00" in message
    assert "POS net: form +6.00, current +7.00" in message
    assert str(explanation_path) in message
    assert explanation_path.exists()


def test_explained_larger_variance_closes_and_preserves_amount_and_evidence(tmp_path: Path) -> None:
    setup = _build_period(
        tmp_path,
        store="9355",
        issue_variances=[Decimal("7.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00")],
    )
    explanation_text = "A duplicate activation was identified and is pending correction."
    explanation_path = _write_weekly_variance_explanation(
        setup,
        issue_variance=Decimal("7.00"),
        explanation=explanation_text,
    )

    run = _run(setup, cleanup_sources=True)
    assert run.assessment.status is CloseStatus.CLOSED_WITH_REVIEW
    assert run.workbook_path.exists() and run.pdf_path.exists()
    assert all(control.disposition.value == "EXPLAINED" for control in run.assessment.controls if control.variance == Decimal("7.00"))
    review_workbook = load_workbook(run.workbook_path)
    report = review_workbook["Monthly Close Report"]
    assert report["A4"].value == "CLOSED WITH REVIEW"
    report_text = "\n".join(
        str(cell.value)
        for row in report.iter_rows()
        for cell in row
        if cell.value is not None
    )
    assert "Variance Explanations" in review_workbook.sheetnames
    assert review_workbook["Variance Explanations"]["D7"].value == explanation_text
    assert explanation_text not in report_text
    assert "Explanation recorded; see Variance Explanations worksheet." in report_text
    weekly_title_row = next(
        cell.row
        for row in report.iter_rows()
        for cell in row
        if cell.value == "Weekly Variance Detail"
    )
    assert report.cell(weekly_title_row + 2, 3).value == 7
    canonical, canonical_pdf = canonical_output_paths(
        setup["output_root"],
        config=get_store_config("9355"),
        fiscal_period=setup["period"],
    )
    assert canonical.exists()
    assert canonical_pdf.exists()
    assert (setup["archive_root"] / "Monthly Close").exists()
    assert not setup["summary_path"].exists()
    assert not setup["darden"].source_file.exists()
    assert not explanation_path.exists()
    retained = [record for record in run.archive_records if record.role == "Weekly Variance Explanation"]
    assert len(retained) == 1 and retained[0].archive_path.is_file()


def test_monthly_monetary_explanations_close_with_real_mismatch_and_archived_proof(tmp_path: Path) -> None:
    setup = _build_period(tmp_path, store="9355")
    workbook = load_workbook(setup["summary_path"])
    workbook["Summary"]["D3"] = 250
    workbook["Summary"]["H3"] = -50
    workbook.save(setup["summary_path"])
    with pytest.raises(CloseBlockedError) as blocked:
        _run(setup)
    source = blocked.value.review_workbook
    workbook = load_workbook(source)
    reasons = workbook["Monthly Variance Explanations"]
    for row in range(7, reasons.max_row + 1):
        reasons.cell(row, 6, "The month-end corporate adjustment was reviewed against the settlement support.")
    workbook.save(source)
    original_hash = sha256_file(source)
    run = _run(setup)
    assert run.assessment.status is CloseStatus.CLOSED_WITH_REVIEW
    assert not run.assessment.darden_matched
    darden = next(control for control in run.assessment.controls if control.code == "darden_summary_match")
    assert darden.disposition.value == "EXPLAINED" and darden.variance == Decimal("-150.00")
    retained = [record for record in run.archive_records if record.role == "monthly_variance_explanation"]
    assert len(retained) == 1
    assert retained[0].sha256 == original_hash == sha256_file(retained[0].archive_path)
    assert not retained[0].remove_after_publish
    report = load_workbook(run.workbook_path)["Monthly Close Report"]
    values = [cell.value for row in report for cell in row]
    assert "MISMATCHED" in values and "EXPLAINED" in values and -150 in values
    archive_input = setup["archive_root"] / "Monthly Close" / "9355" / setup["period"].folder_name
    archived_source = ArchivedVarianceExplanationSource(
        path=retained[0].archive_path, sha256=retained[0].sha256,
        size_bytes=retained[0].size_bytes,
    )
    # A reissue consumes the manifest-selected proof, not a newer live narrative.
    current = load_workbook(run.workbook_path)
    current["Monthly Variance Explanations"]["F7"] = "A later unarchived note must not replace retained proof."
    current.save(run.workbook_path)
    with pytest.raises(CorrectionRequired) as proposed:
        _run(
        setup, input_dir=archive_input, micros_path=archive_input / "micros",
        darden_report=replace(setup["darden"], source_file=archive_input / "darden" / setup["darden"].source_file.name),
        archived_variance_explanations={}, archived_monthly_variance_explanation=archived_source,
    )
    reissue = _run(
        setup, approve_correction=proposed.value.token, input_dir=archive_input, micros_path=archive_input / "micros",
        darden_report=replace(setup["darden"], source_file=archive_input / "darden" / setup["darden"].source_file.name),
        archived_variance_explanations={}, archived_monthly_variance_explanation=archived_source,
    )
    assert reissue.assessment.status is CloseStatus.CLOSED_WITH_REVIEW
    assert not reissue.assessment.darden_matched
    assert all("later unarchived" not in control.explanation for control in reissue.assessment.controls)
    retained_again = [record for record in reissue.archive_records if record.role == "monthly_variance_explanation"]
    assert len(retained_again) == 1 and retained_again[0].archive_path == retained[0].archive_path


def test_explicit_output_workbook_supplies_monthly_explanations(tmp_path: Path) -> None:
    setup = _build_period(tmp_path, store="9355")
    workbook = load_workbook(setup["summary_path"])
    workbook["Summary"]["D3"] = 250
    workbook["Summary"]["H3"] = -50
    workbook.save(setup["summary_path"])
    workbook.close()
    with pytest.raises(CloseBlockedError) as blocked:
        _run(setup)

    explicit_output = tmp_path / "custom-output" / "requested-close.xlsx"
    explicit_output.parent.mkdir(parents=True)
    workbook = load_workbook(blocked.value.review_workbook)
    reasons = workbook["Monthly Variance Explanations"]
    for row in range(7, reasons.max_row + 1):
        reasons.cell(row, 6, "The requested-output explanation was reviewed against settlement support.")
    workbook.save(explicit_output)
    workbook.close()
    blocked.value.review_workbook.unlink()

    run = _run(setup, output_path=explicit_output)

    assert run.workbook_path == explicit_output
    assert run.assessment.status is CloseStatus.CLOSED_WITH_REVIEW
    assert any(
        control.explanation.startswith("The requested-output explanation")
        for control in run.assessment.controls
    )


def test_period_reason_cannot_waive_missing_required_weekly_reason(tmp_path: Path) -> None:
    setup = _build_period(tmp_path, store="9355", issue_variances=[Decimal("7.00")] + [Decimal("0.00")] * 4)
    with pytest.raises(CloseBlockedError) as blocked:
        _run(setup)
    workbook = load_workbook(blocked.value.review_workbook)
    reasons = workbook["Monthly Variance Explanations"]
    for row in range(7, reasons.max_row + 1):
        reasons.cell(row, 6, "The period total was reviewed.")
    workbook.save(blocked.value.review_workbook)
    with pytest.raises(CloseBlockedError) as still_blocked:
        _run(setup)
    assert any(control.code.startswith("weekly_explanation_") for control in still_blocked.value.assessment.blockers)
    assert any(control.disposition.value == "EXPLAINED" for control in still_blocked.value.assessment.controls)
    assert not (setup["archive_root"] / "Monthly Close").exists()


def test_weekly_text_does_not_waive_unexplained_contribution_to_period(tmp_path: Path) -> None:
    setup = _build_period(tmp_path, store="9355", issue_variances=[Decimal("7.00"), Decimal("1.00")] + [Decimal("0.00")] * 3)
    _write_weekly_variance_explanation(setup, issue_variance=Decimal("7.00"), explanation="This first-week issue was reviewed.")
    with pytest.raises(CloseBlockedError) as blocked:
        _run(setup)
    period_controls = [control for control in blocked.value.assessment.controls if control.code.startswith("period_pos_") and control.variance == Decimal("8.00")]
    assert period_controls and all(control.is_blocking for control in period_controls)


def test_missing_tender_evidence_is_blocking_and_never_archived(tmp_path: Path) -> None:
    setup = _build_period(tmp_path, store="9355")
    (setup["micros_dir"] / "TENDER_DETAIL.TXT").unlink()

    with pytest.raises(CloseBlockedError, match="TENDER_DETAIL.TXT") as exc_info:
        _run(setup)

    assert exc_info.value.review_workbook is not None
    assert not (setup["archive_root"] / "Monthly Close").exists()
    assert setup["summary_path"].exists()


def test_pdf_export_failure_completes_workbook_and_retries_without_inputs(tmp_path: Path) -> None:
    setup = _build_period(tmp_path, store="9355")

    def fail_export(**_kwargs):
        raise PdfExportError("simulated Excel PDF export failure")

    run = _run(setup, pdf_exporter=fail_export, cleanup_sources=True)
    assert run.workbook_path.is_file()
    assert run.assessment.status is CloseStatus.CLOSED
    assert run.pdf_path is None and run.pdf_status == "pending"
    assert run.pdf_error == "simulated Excel PDF export failure"
    manifest = json.loads(run.manifest_path.read_text(encoding="utf-8"))
    assert manifest["pdf_status"] == "pending"
    assert {item["role"] for item in manifest["artifacts"]} == {"workbook"}
    assert not setup["summary_path"].exists()
    original_workbook = sha256_file(run.workbook_path)
    pdf = retry_monthly_pdf(archive_root=setup["archive_root"], store="9355",
                            period=setup["period"].period_key, pdf_exporter=_fake_pdf_exporter)
    updated = json.loads(run.manifest_path.read_text(encoding="utf-8"))
    assert pdf.is_file() and updated["pdf_status"] == "complete"
    assert updated["result_id"] == run.result_id
    assert sha256_file(run.workbook_path) == original_workbook


def test_locked_stale_diagnostic_preserves_old_pair_without_masking_blocker(
    tmp_path: Path,
    monkeypatch,
) -> None:
    setup = _build_period(
        tmp_path,
        store="9355",
        issue_variances=[Decimal("7.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00"), Decimal("0.00")],
    )
    review_xlsx, review_pdf = review_output_paths(
        setup["output_root"],
        config=get_store_config("9355"),
        fiscal_period=setup["period"],
    )
    review_xlsx.parent.mkdir(parents=True)
    review_xlsx.write_bytes(b"older diagnostic workbook")
    review_pdf.write_bytes(b"older diagnostic PDF")

    from gift_card_recon import monthly_close_service as service

    original = service._assert_publishable

    def locked_pdf(path: Path) -> None:
        if Path(path) == review_pdf:
            raise PermissionError("simulated locked diagnostic PDF")
        original(path)

    def fail_export(**_kwargs):
        raise PdfExportError("simulated diagnostic PDF failure")

    monkeypatch.setattr(service, "_assert_publishable", locked_pdf)
    with pytest.raises(CloseBlockedError) as exc_info:
        _run(setup, pdf_exporter=fail_export)

    error = exc_info.value
    assert "7.00" in str(error)
    assert error.review_workbook is None
    assert error.review_pdf is None
    assert error.review_publication_error == "simulated locked diagnostic PDF"
    assert review_xlsx.read_bytes() == b"older diagnostic workbook"
    assert review_pdf.read_bytes() == b"older diagnostic PDF"


def test_archive_failure_blocks_publication_and_preserves_sources(
    tmp_path: Path,
    monkeypatch,
) -> None:
    setup = _build_period(tmp_path, store="9355")

    from gift_card_recon.publication import PublicationTransaction
    real_replace = PublicationTransaction.replace

    def fail_archive(transaction, source, destination, **kwargs):
        if "Monthly Close" in Path(destination).parts and Path(destination).is_relative_to(setup["archive_root"]):
            raise ArchiveError("simulated archive failure")
        return real_replace(transaction, source, destination, **kwargs)

    monkeypatch.setattr(PublicationTransaction, "replace", fail_archive)
    with pytest.raises(CloseBlockedError, match="simulated archive failure"):
        _run(setup, cleanup_sources=True)

    canonical, canonical_pdf = canonical_output_paths(
        setup["output_root"],
        config=get_store_config("9355"),
        fiscal_period=setup["period"],
    )
    assert not canonical.exists() and not canonical_pdf.exists()
    assert setup["summary_path"].exists()
    assert setup["darden"].source_file.exists()


def test_locked_canonical_output_fails_without_alternate_or_archiving(
    tmp_path: Path,
    monkeypatch,
) -> None:
    setup = _build_period(tmp_path, store="9355")
    canonical, _ = canonical_output_paths(
        setup["output_root"],
        config=get_store_config("9355"),
        fiscal_period=setup["period"],
    )
    canonical.parent.mkdir(parents=True)
    canonical.write_bytes(b"older canonical workbook")

    from gift_card_recon import monthly_close_service as service

    original = service._assert_publishable

    def locked_only(path: Path) -> None:
        if Path(path) == canonical:
            raise PermissionError("simulated locked canonical output")
        original(path)

    monkeypatch.setattr(service, "_assert_publishable", locked_only)
    with pytest.raises(CloseBlockedError, match="simulated locked canonical output"):
        _run(setup, cleanup_sources=True)

    assert canonical.read_bytes() == b"older canonical workbook"
    assert not list(canonical.parent.glob("*with_weekly_variance*"))
    assert not (setup["archive_root"] / "Monthly Close").exists()
    assert setup["summary_path"].exists()


def test_successful_close_removes_only_live_inputs_after_verified_archive(tmp_path: Path) -> None:
    setup = _build_period(tmp_path, store="9355")

    run = _run(setup, cleanup_sources=True)

    assert run.workbook_path.exists() and run.pdf_path.exists()
    assert not setup["summary_path"].exists()
    assert not setup["darden"].source_file.exists()
    assert not list((setup["input_dir"] / "activity").glob("*.xlsx"))
    assert (setup["micros_dir"] / "DLYSYSTT.TXT").exists()
    assert (setup["micros_dir"] / "TENDER_DETAIL.TXT").exists()
    archive_base = (
        setup["archive_root"]
        / "Monthly Close"
        / "9355"
        / setup["period"].folder_name
    )
    assert len(list((archive_base / "activity").glob("*.xlsx"))) == 5
    assert (archive_base / "summary" / setup["summary_path"].name).exists()
    assert (archive_base / "darden" / setup["darden"].source_file.name).exists()


def test_present_weekly_variance_explanation_is_archived_and_hashed_on_success(
    tmp_path: Path,
) -> None:
    setup = _build_period(tmp_path, store="9355")
    explanation_text = "A timing question was reviewed; no unresolved variance remains."
    explanation_path = _write_weekly_variance_explanation(
        setup,
        explanation=explanation_text,
    )
    explanation_hash = sha256_file(explanation_path)

    run = _run(setup, cleanup_sources=True)

    explanation_records = [
        record
        for record in run.archive_records
        if record.role == "Weekly Variance Explanation"
    ]
    assert len(explanation_records) == 1
    record = explanation_records[0]
    assert record.sha256 == explanation_hash
    assert record.archive_path.exists()
    assert record.archive_path.parent.name == "Variance Explanations"
    assert sha256_file(record.archive_path) == explanation_hash
    assert not explanation_path.exists()

    manifest = json.loads(run.manifest_path.read_text(encoding="utf-8"))
    manifest_sources = [
        source
        for source in manifest["sources"]
        if source["role"] == "Weekly Variance Explanation"
    ]
    assert len(manifest_sources) == 1
    assert manifest_sources[0]["sha256"] == explanation_hash
    assert manifest_sources[0]["archive_category"].endswith(
        "Variance Explanations"
    )

    monthly_workbook = load_workbook(run.workbook_path)
    report = monthly_workbook["Monthly Close Report"]
    report_text = "\n".join(
        str(cell.value)
        for row in report.iter_rows()
        for cell in row
        if cell.value is not None
    )
    assert "Variance Explanations" in monthly_workbook.sheetnames
    assert monthly_workbook["Variance Explanations"]["D7"].value == explanation_text
    assert explanation_text not in report_text


def test_archive_reissue_consumes_exact_legacy_collision_explanation_without_duplicate(
    tmp_path: Path,
) -> None:
    archive_root = tmp_path / "Archive - Old Files"
    setup = _build_period(archive_root, store="9355")
    setup["archive_root"] = archive_root
    setup["output_root"] = tmp_path / "Output"
    week_end = setup["period"].expected_week_endings[0]
    explanation_text = "The manifest-selected legacy explanation was reviewed."
    canonical_name = variance_explanation_path(
        setup["input_dir"],
        setup["store"],
        week_end,
    ).name
    legacy_path = setup["input_dir"] / "variance explanations" / canonical_name
    selected_explanation = WeeklyVarianceExplanation(
        store=setup["store"],
        week_start=week_end - timedelta(days=6),
        week_end=week_end,
        pos_issue_variance=Decimal("0.00"),
        pos_payment_variance=Decimal("0.00"),
        pos_net_variance=Decimal("0.00"),
        tender_variance=Decimal("0.00"),
        explanation=explanation_text,
    )
    write_variance_explanation_workbook(
        selected_explanation,
        legacy_path,
    )
    explanation_hash = sha256_file(legacy_path)
    collision_path = legacy_path.with_name(
        f"{legacy_path.stem}__{explanation_hash[:12]}{legacy_path.suffix}"
    )
    legacy_path.replace(collision_path)
    write_variance_explanation_workbook(
        replace(
            selected_explanation,
            explanation="This unmanifested canonical form must not be consumed.",
        ),
        legacy_path,
    )
    archived_source = ArchivedVarianceExplanationSource(
        path=collision_path,
        sha256=explanation_hash,
        size_bytes=collision_path.stat().st_size,
    )

    run = _run(
        setup,
        archived_variance_explanations={week_end: archived_source},
    )

    explanation_records = [
        record
        for record in run.archive_records
        if record.role == "Weekly Variance Explanation"
    ]
    assert len(explanation_records) == 1
    assert explanation_records[0].source_path.resolve() == collision_path.resolve()
    assert explanation_records[0].archive_path.resolve() == collision_path.resolve()
    assert legacy_path.exists()
    matching_folders = [
        path.name
        for path in setup["input_dir"].iterdir()
        if path.is_dir() and path.name.casefold() == "variance explanations"
    ]
    assert matching_folders == ["variance explanations"]

    manifest = json.loads(run.manifest_path.read_text(encoding="utf-8"))
    explanation_sources = [
        source
        for source in manifest["sources"]
        if source["role"] == "Weekly Variance Explanation"
    ]
    assert len(explanation_sources) == 1
    assert explanation_sources[0]["archive_path"].endswith(collision_path.name)
    assert "/variance explanations/" in explanation_sources[0]["archive_path"]

    workbook = load_workbook(run.workbook_path)
    assert workbook["Variance Explanations"]["D7"].value == explanation_text


def test_archive_reissue_ignores_unmanifested_variance_explanation(
    tmp_path: Path,
) -> None:
    setup = _build_period(tmp_path, store="9355")
    unmanifested_path = _write_weekly_variance_explanation(
        setup,
        explanation="This form is not listed in the verified close manifest.",
    )

    run = _run(setup, archived_variance_explanations={})

    assert unmanifested_path.exists()
    assert not any(
        record.role == "Weekly Variance Explanation"
        for record in run.archive_records
    )


def test_archive_reissue_blocks_when_only_unmanifested_required_explanation_exists(
    tmp_path: Path,
) -> None:
    setup = _build_period(
        tmp_path,
        store="9355",
        payment_variances=[Decimal("6.00")] + [Decimal("0.00")] * 4,
    )
    unmanifested_path = _write_weekly_variance_explanation(
        setup,
        payment_variance=Decimal("6.00"),
        explanation="This valid form is deliberately absent from the archive manifest.",
    )

    with pytest.raises(
        CloseBlockedError,
        match="verified archive manifest contains no explanation form",
    ):
        _run(setup, archived_variance_explanations={})

    assert unmanifested_path.exists()


@pytest.mark.parametrize(
    ("same_size", "message"),
    [
        (False, "size no longer matches"),
        (True, "hash no longer matches"),
    ],
)
def test_archive_reissue_rechecks_manifest_variance_explanation_integrity(
    tmp_path: Path,
    same_size: bool,
    message: str,
) -> None:
    setup = _build_period(tmp_path, store="9355")
    week_end = setup["period"].expected_week_endings[0]
    explanation_path = _write_weekly_variance_explanation(
        setup,
        explanation="Verified before the simulated file change.",
    )
    archived_source = ArchivedVarianceExplanationSource(
        path=explanation_path,
        sha256=sha256_file(explanation_path),
        size_bytes=explanation_path.stat().st_size,
    )
    changed = bytearray(explanation_path.read_bytes())
    if same_size:
        changed[0] ^= 1
    else:
        changed.extend(b"changed")
    explanation_path.write_bytes(changed)

    with pytest.raises(CloseBlockedError, match=message):
        _run(
            setup,
            archived_variance_explanations={week_end: archived_source},
        )


def test_summary_period_mismatch_is_blocking(tmp_path: Path) -> None:
    setup = _build_period(tmp_path, store="9355")
    wrong = setup["summary_path"].with_name("06.28.2026 9355 Gift Card Summary.xlsx")
    setup["summary_path"].replace(wrong)

    with pytest.raises(CloseBlockedError, match="report end 2026-06-28"):
        _run(setup)


@pytest.mark.parametrize(
    ("wrong_store", "period_start", "period_end", "message"),
    [
        ("9354", date(2026, 6, 1), date(2026, 7, 5), "expected store 9355"),
        ("9355", date(2026, 6, 2), date(2026, 7, 5), "does not match FY27-M01"),
    ],
)
def test_wrong_darden_identity_is_blocking(
    tmp_path: Path,
    wrong_store: str,
    period_start: date,
    period_end: date,
    message: str,
) -> None:
    setup = _build_period(tmp_path, store="9355")
    setup["darden"] = DardenCreditMemo(
        source_file=setup["darden"].source_file,
        store=wrong_store,
        period_start=period_start,
        period_end=period_end,
        total=Decimal("-200.00"),
    )

    with pytest.raises(CloseBlockedError, match=message):
        _run(setup)


def test_invalid_workbook_extension_is_rejected(tmp_path: Path) -> None:
    setup = _build_period(tmp_path, store="9355")

    with pytest.raises(ParseError, match="must end in .xlsx"):
        _run(setup, output_path=tmp_path / "Output" / "wrong.pdf")


def test_supplied_fiscal_period_object_must_match_canonical_calendar(tmp_path: Path) -> None:
    setup = _build_period(tmp_path, store="9355")
    altered = replace(setup["period"], start_date=date(2026, 6, 2))

    with pytest.raises(ParseError, match="does not match the supplied period object"):
        _run(setup, fiscal_period=altered)


def test_manifest_failure_rolls_back_new_archive_copies_and_canonical_outputs(
    tmp_path: Path,
    monkeypatch,
) -> None:
    setup = _build_period(tmp_path, store="9355")

    def fail_manifest(*_args, **_kwargs):
        raise ArchiveError("simulated manifest failure")

    monkeypatch.setattr(
        "gift_card_recon.monthly_close_service.write_close_manifest_atomic",
        fail_manifest,
    )
    with pytest.raises(CloseBlockedError, match="simulated manifest failure"):
        _run(setup, cleanup_sources=True)

    canonical, canonical_pdf = canonical_output_paths(
        setup["output_root"],
        config=get_store_config("9355"),
        fiscal_period=setup["period"],
    )
    assert not canonical.exists() and not canonical_pdf.exists()
    assert not any(path.is_file() for path in (setup["archive_root"] / "Monthly Close").rglob("*"))
    assert setup["summary_path"].exists()
    assert setup["darden"].source_file.exists()


def test_evidence_change_during_calculation_blocks_archive_and_publication(
    tmp_path: Path,
    monkeypatch,
) -> None:
    setup = _build_period(tmp_path, store="9355")
    from gift_card_recon import monthly_close_service as service

    original_plan = service.plan_evidence_archive

    def mutate_then_plan(items, *, archive_root):
        setup["summary_path"].write_bytes(b"changed after parsing")
        return original_plan(items, archive_root=archive_root)

    monkeypatch.setattr(service, "plan_evidence_archive", mutate_then_plan)
    with pytest.raises(CloseBlockedError, match="Evidence changed while the close was being calculated"):
        _run(setup)

    assert not (setup["archive_root"] / "Monthly Close").exists()


def test_success_archives_and_removes_superseded_review_artifact(tmp_path: Path) -> None:
    setup = _build_period(tmp_path, store="9355")
    review = (
        setup["output_root"]
        / "Monthly Close - Review Required"
        / "Virginia_Beach_9355_FY27-M01_Review_Required.xlsx"
    )
    review.parent.mkdir(parents=True)
    review.write_bytes(b"older diagnostic")

    _run(setup)

    assert not review.exists()
    archived = (
        setup["archive_root"]
        / "Generated Reports"
        / "Diagnostics"
        / "9355"
        / "FY27-M01"
        / review.name
    )
    assert archived.exists()


def test_success_archives_and_removes_legacy_numbered_review_pair(tmp_path: Path) -> None:
    setup = _build_period(tmp_path, store="9355")
    current_xlsx, current_pdf = review_output_paths(
        setup["output_root"],
        config=get_store_config("9355"),
        fiscal_period=setup["period"],
    )
    legacy_folder = setup["output_root"] / "Review Required"
    legacy_xlsx = legacy_folder / current_xlsx.name
    legacy_pdf = legacy_folder / current_pdf.name
    legacy_folder.mkdir(parents=True)
    legacy_xlsx.write_bytes(b"legacy diagnostic workbook")
    legacy_pdf.write_bytes(b"legacy diagnostic PDF")

    _run(setup)

    assert not legacy_xlsx.exists()
    assert not legacy_pdf.exists()
    assert not legacy_folder.exists()
    archive_folder = (
        setup["archive_root"]
        / "Generated Reports"
        / "Diagnostics"
        / "9355"
        / "FY27-M01"
    )
    assert (archive_folder / legacy_xlsx.name).read_bytes() == b"legacy diagnostic workbook"
    assert (archive_folder / legacy_pdf.name).read_bytes() == b"legacy diagnostic PDF"


def test_success_archives_both_review_folders_with_content_safe_deduplication(
    tmp_path: Path,
) -> None:
    setup = _build_period(tmp_path, store="9355")
    current_xlsx, current_pdf = review_output_paths(
        setup["output_root"],
        config=get_store_config("9355"),
        fiscal_period=setup["period"],
    )
    legacy_folder = setup["output_root"] / "Review Required"
    legacy_xlsx = legacy_folder / current_xlsx.name
    legacy_pdf = legacy_folder / current_pdf.name
    current_xlsx.parent.mkdir(parents=True)
    legacy_folder.mkdir(parents=True)
    current_xlsx.write_bytes(b"current diagnostic workbook")
    legacy_xlsx.write_bytes(b"legacy diagnostic workbook")
    current_pdf.write_bytes(b"identical diagnostic PDF")
    legacy_pdf.write_bytes(b"identical diagnostic PDF")

    _run(setup)

    assert not current_xlsx.parent.exists()
    assert not legacy_folder.exists()
    archive_folder = (
        setup["archive_root"]
        / "Generated Reports"
        / "Diagnostics"
        / "9355"
        / "FY27-M01"
    )
    archived = [path for path in archive_folder.iterdir() if path.is_file()]
    assert len(archived) == 3
    assert {path.read_bytes() for path in archived} == {
        b"current diagnostic workbook",
        b"legacy diagnostic workbook",
        b"identical diagnostic PDF",
    }
    assert (archive_folder / current_xlsx.name).read_bytes() == b"current diagnostic workbook"
    collision_workbooks = [
        path
        for path in archived
        if path.suffix == ".xlsx" and path.name != current_xlsx.name
    ]
    assert len(collision_workbooks) == 1
    assert collision_workbooks[0].stem.startswith(f"{current_xlsx.stem}__")
    assert collision_workbooks[0].read_bytes() == b"legacy diagnostic workbook"






def _sharing_violation(message: str) -> PermissionError:
    exc = PermissionError(errno.EACCES, message)
    exc.winerror = 32
    return exc


def _run(setup: dict[str, object], **overrides):
    arguments = {
        "store": setup["store"],
        "period": setup["period"].period_key,
        "input_dir": setup["input_dir"],
        "micros_path": setup["micros_dir"],
        "micros_work_dir": setup["root"] / "extract",
        "archive_root": setup["archive_root"],
        "output_root": setup["output_root"],
        "darden_report": setup["darden"],
        "fiscal_period": setup["period"],
        "cleanup_sources": False,
        "allow_unconfigured_micros": True,
        "generated_at": datetime(2026, 7, 6, 9, 30),
        "pdf_exporter": _fake_pdf_exporter,
    }
    arguments.update(overrides)
    return run_monthly_close_service(**arguments)


def _write_weekly_variance_explanation(
    setup: dict[str, object],
    *,
    week_index: int = 0,
    issue_variance: Decimal = Decimal("0.00"),
    payment_variance: Decimal = Decimal("0.00"),
    tender_variance: Decimal = Decimal("0.00"),
    explanation: str,
) -> Path:
    week_end = setup["period"].expected_week_endings[week_index]
    week_start = week_end - timedelta(days=6)
    output_path = variance_explanation_path(
        setup["input_dir"],
        setup["store"],
        week_end,
    )
    return write_variance_explanation_workbook(
        WeeklyVarianceExplanation(
            store=setup["store"],
            week_start=week_start,
            week_end=week_end,
            pos_issue_variance=issue_variance,
            pos_payment_variance=payment_variance,
            pos_net_variance=issue_variance - payment_variance,
            tender_variance=tender_variance,
            explanation=explanation,
        ),
        output_path,
    )


def _build_period(
    root: Path,
    *,
    store: str,
    issue_variances: list[Decimal] | None = None,
    payment_variances: list[Decimal] | None = None,
) -> dict[str, object]:
    fiscal_period = fiscal_period_for_label("FY27-M01")
    input_dir = root / "Monthly Close" / store / fiscal_period.folder_name
    summary_path = input_dir / "summary" / f"07.05.2026 {store} Gift Card Summary.xlsx"
    activity_dir = input_dir / "activity"
    micros_dir = root / f"micros-{store}"
    summary_path.parent.mkdir(parents=True)
    activity_dir.mkdir()
    micros_dir.mkdir()
    _write_summary(summary_path, store=store)

    issue_variances = issue_variances or [Decimal("0.00")] * 5
    payment_variances = payment_variances or [Decimal("0.00")] * 5
    system_rows: list[tuple[date, Decimal, Decimal]] = []
    for index, week_end in enumerate(fiscal_period.expected_week_endings):
        week_start = week_end - timedelta(days=6)
        activation = Decimal("100.00") if index == 0 else Decimal("0.00")
        redemption = Decimal("-300.00") if index == 0 else Decimal("0.00")
        _write_activity(
            activity_dir / f"{week_end:%m.%d.%Y} {store} Gift Card Activity.xlsx",
            store=store,
            begin=week_start,
            end=week_end,
            transaction_date=week_start,
            activation=activation,
            redemption=redemption,
        )
        for business_date in _dates(week_start, week_end):
            system_rows.append((business_date, Decimal("0.00"), Decimal("0.00")))
        issue_total = activation + issue_variances[index]
        payment_total = abs(redemption) + payment_variances[index]
        first_row_index = len(system_rows) - 7
        system_rows[first_row_index] = (week_start, issue_total, payment_total)
    _write_micros(micros_dir, system_rows)

    darden_path = input_dir / "darden" / f"FY27-M01 {store} Darden Credit Memo.pdf"
    darden_path.parent.mkdir()
    darden_path.write_bytes(b"synthetic Darden evidence")
    darden = DardenCreditMemo(
        source_file=darden_path,
        store=store,
        period_start=fiscal_period.start_date,
        period_end=fiscal_period.end_date,
        total=Decimal("-200.00"),
        invoice_number="FY27 M01 synthetic",
        invoice_date=date(2026, 7, 6),
    )
    return {
        "root": root,
        "store": store,
        "period": fiscal_period,
        "input_dir": input_dir,
        "summary_path": summary_path,
        "micros_dir": micros_dir,
        "darden": darden,
        "archive_root": root / "Archive - Old Files",
        "output_root": root / "Output",
    }


def _write_summary(path: Path, *, store: str) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Summary"
    sheet.append(["SUMMARY"])
    sheet.append(
        [
            "Franchise Partner",
            "Store Number",
            "Gift Card Franchise Fee Rate",
            "Total Activations",
            "Total Redemptions",
            "Payable Redemptions",
            "GCDR",
            "Net Settlement",
        ]
    )
    sheet.append(["Sorensen", int(store), 0.1, 100, -300, -300, 0, -200])
    workbook.save(path)


def _write_activity(
    path: Path,
    *,
    store: str,
    begin: date,
    end: date,
    transaction_date: date,
    activation: Decimal,
    redemption: Decimal,
) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet 1"
    sheet.append(
        [
            f"All GC Activity BEGIN DATE: '{begin:%d-%b-%Y}', END DATE: "
            f"'{end:%d-%b-%Y}', Rest Number Parameter 1: '{store}'"
        ]
    )
    sheet.append(
        [
            "Card No",
            "Request",
            "Request Code Listing",
            "Business Date",
            "Transaction No",
            "Amount SUM",
            "Promocode",
            "Authorization Code",
        ]
    )
    sheet.append(["0001", 100, "Activation", transaction_date, 1, float(activation), None, 111])
    sheet.append(["0002", 202, "Redemption No Nsf", transaction_date, 2, float(redemption), None, 222])
    workbook.save(path)


def _write_micros(folder: Path, rows: list[tuple[date, Decimal, Decimal]]) -> None:
    config = get_store_config("9355")
    system_lines: list[str] = []
    tender_lines: list[str] = []
    for business_date, issue, payment in rows:
        row = ["0"] * 132
        row[0] = f"{business_date:%Y-%m-%d} 00:00:00.000"
        row[config.micros_issue_column_index] = f"{issue:.2f}"
        row[config.micros_payment_column_index] = f"{payment:.2f}"
        system_lines.append(",".join(row))
        tender_lines.append(
            f"'{business_date:%Y-%m-%d}',{payment:.2f},350,'G C Payment','T'"
        )
    (folder / "DLYSYSTT.TXT").write_text("\n".join(system_lines), encoding="utf-8")
    (folder / "TENDER_DETAIL.TXT").write_text("\n".join(tender_lines), encoding="utf-8")


def _dates(start: date, end: date):
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def _fake_pdf_exporter(*, workbook_path: Path, pdf_path: Path, expected_location_label: str):
    assert Path(workbook_path).is_file()
    assert expected_location_label in {
        "RICHMOND - STORE 9354",
        "VIRGINIA BEACH - STORE 9355",
    }
    Path(pdf_path).write_bytes(b"%PDF-1.4 mocked two-page Excel export")
    return Path(pdf_path)
