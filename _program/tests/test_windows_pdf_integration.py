from __future__ import annotations

import sys
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest
from pypdf import PdfReader

from gift_card_recon.close_assessment import (
    ControlDisposition, ControlOutcome, build_close_assessment,
    exact_match_control, variance_control,
)
from gift_card_recon.models import DardenCreditMemo, MonthlyCloseCertification
from gift_card_recon.monthly_report import (
    MonthlyCloseReportData,
    WeeklyCloseReportRow,
    write_monthly_close_report_workbook,
)
from gift_card_recon.pdf_export import export_monthly_close_report_pdf


@pytest.mark.skipif(sys.platform != "win32", reason="Windows Excel COM integration test")
def test_excel_exports_and_validates_exact_two_page_monthly_close_pdf(tmp_path: Path) -> None:
    assessment = build_close_assessment(
        store="9355",
        darden_variance=Decimal("0.00"),
        controls=(
            ControlOutcome(
                code="evidence_integrity",
                label="Evidence integrity",
                disposition=ControlDisposition.PASS,
                message="All evidence controls passed.",
            ),
        ),
    )
    memo_path = tmp_path / "Darden.pdf"
    memo_path.write_bytes(b"integration fixture")
    memo = DardenCreditMemo(
        source_file=memo_path,
        store="9355",
        period_start=date(2026, 6, 1),
        period_end=date(2026, 7, 5),
        total=Decimal("-200.00"),
    )
    certification = MonthlyCloseCertification(
        store="9355",
        period="FY27-M01",
        period_start=memo.period_start,
        period_end=memo.period_end,
        summary_net_settlement=Decimal("-200.00"),
        darden_credit_memo=memo,
        variance=Decimal("0.00"),
    )
    workbook = tmp_path / "close.xlsx"
    pdf = tmp_path / "close.pdf"
    write_monthly_close_report_workbook(
        MonthlyCloseReportData(
            assessment=assessment,
            period="FY27-M01",
            period_start=memo.period_start,
            period_end=memo.period_end,
            generated_at=datetime(2026, 7, 6, 9, 30),
            certification=certification,
            evidence_notes=("Windows Excel integration validation.",),
        ),
        workbook,
    )

    result = export_monthly_close_report_pdf(
        workbook_path=workbook,
        pdf_path=pdf,
        expected_location_label="VIRGINIA BEACH - STORE 9355",
    )

    assert result == pdf.resolve()
    assert pdf.stat().st_size > 0
    metadata = PdfReader(pdf, strict=False).metadata
    assert metadata.title == (
        "VIRGINIA BEACH - STORE 9355 FY27-M01 Monthly Close Report"
    )
    assert metadata.subject == "Gift Card Monthly Close Reconciliation"
    assert metadata.author == "Gift Card Reconciliation Close Control"


@pytest.mark.skipif(sys.platform != "win32", reason="Windows Excel COM integration test")
def test_realistic_review_report_with_service_exception_duplicates_stays_two_pages(
    tmp_path: Path,
) -> None:
    controls: list[ControlOutcome] = [
        ControlOutcome(
            code=code,
            label=code.replace("_", " ").title(),
            disposition=ControlDisposition.PASS,
            message="Passed.",
        )
        for code in (
            "summary_identity",
            "activity_coverage",
            "micros_coverage",
            "tender_evidence",
            "archive_integrity",
        )
    ]
    weekly_codes: set[str] = set()
    weekly_rows: list[WeeklyCloseReportRow] = []
    for week_index, week_ending in enumerate(
        (
            date(2026, 6, 7),
            date(2026, 6, 14),
            date(2026, 6, 21),
            date(2026, 6, 28),
            date(2026, 7, 5),
        ),
        start=1,
    ):
        reviewed = week_index in {2, 4}
        disposition = (
            ControlDisposition.REVIEW if reviewed else ControlDisposition.PASS
        )
        weekly_rows.append(
            WeeklyCloseReportRow(
                week_ending=week_ending,
                coverage="Complete",
                pos_issue_variance=Decimal("0.00"),
                pos_payment_variance=Decimal("2.43") if reviewed else Decimal("0.00"),
                pos_net_variance=Decimal("-2.43") if reviewed else Decimal("0.00"),
                tender_variance=Decimal("0.00"),
                disposition=disposition,
            )
        )
        for metric in ("POS issue", "POS payment", "POS net", "Tender"):
            code = (
                f"{'tender' if metric == 'Tender' else 'pos'}_week_{week_index}_"
                f"{metric.lower().replace(' ', '_')}"
            )
            weekly_codes.add(code)
            control_disposition = (
                ControlDisposition.REVIEW
                if reviewed and metric in {"POS payment", "POS net"}
                else ControlDisposition.PASS
            )
            controls.append(
                ControlOutcome(
                    code=code,
                    label=f"Week ending {week_ending:%m/%d/%Y} {metric}",
                    disposition=control_disposition,
                    message=(
                        "Review the small weekly variance."
                        if control_disposition is ControlDisposition.REVIEW
                        else "No variance."
                    ),
                    variance=(
                        Decimal("2.43")
                        if control_disposition is ControlDisposition.REVIEW
                        else Decimal("0.00")
                    ),
                )
            )
    controls.extend(
        (
            ControlOutcome(
                "pos_period_payment",
                "Period POS payment",
                ControlDisposition.REVIEW,
                "Review +2.43.",
                Decimal("2.43"),
            ),
            ControlOutcome(
                "pos_period_net",
                "Period POS net",
                ControlDisposition.REVIEW,
                "Review +2.43.",
                Decimal("2.43"),
            ),
            ControlOutcome(
                "tender_period",
                "Period tender",
                ControlDisposition.PASS,
                "No variance.",
                Decimal("0.00"),
            ),
        )
    )
    assessment = build_close_assessment(
        store="9354",
        darden_variance=Decimal("0.00"),
        controls=controls,
    )
    memo_path = tmp_path / "Darden.pdf"
    memo_path.write_bytes(b"integration fixture")
    memo = DardenCreditMemo(
        source_file=memo_path,
        store="9354",
        period_start=date(2026, 6, 1),
        period_end=date(2026, 7, 5),
        total=Decimal("-26722.95"),
    )
    certification = MonthlyCloseCertification(
        store="9354",
        period="FY27-M01",
        period_start=memo.period_start,
        period_end=memo.period_end,
        summary_net_settlement=Decimal("-26722.95"),
        darden_credit_memo=memo,
        variance=Decimal("0.00"),
    )
    duplicated_exceptions = tuple(
        (
            control.disposition.value,
            f"{control.label}: {control.message}",
        )
        for control in assessment.controls
        if not control.passed
    )
    workbook = tmp_path / "richmond-review.xlsx"
    pdf = tmp_path / "richmond-review.pdf"
    write_monthly_close_report_workbook(
        MonthlyCloseReportData(
            assessment=assessment,
            period="FY27-M01",
            period_start=memo.period_start,
            period_end=memo.period_end,
            generated_at=datetime(2026, 7, 11, 10, 30),
            certification=certification,
            weekly_rows=tuple(weekly_rows),
            period_pos_net_variance=Decimal("2.43"),
            period_pos_disposition=ControlDisposition.REVIEW,
            period_tender_variance=Decimal("0.00"),
            period_tender_disposition=ControlDisposition.PASS,
            explicit_exceptions=duplicated_exceptions,
            weekly_control_codes=frozenset(weekly_codes),
        ),
        workbook,
    )

    export_monthly_close_report_pdf(
        workbook_path=workbook,
        pdf_path=pdf,
        expected_location_label="RICHMOND - STORE 9354",
    )

    reader = PdfReader(pdf, strict=False)
    assert len(reader.pages) == 2
    page_two_text = reader.pages[1].extract_text() or ""
    assert "Week ending 06/14/2026" in page_two_text
    assert "Week ending 06/28/2026" in page_two_text
    assert "POS payment +$2.43; POS net -$2.43" in page_two_text
    assert "Exception" not in page_two_text


@pytest.mark.skipif(sys.platform != "win32", reason="Windows Excel COM integration test")
@pytest.mark.parametrize("explanation_scope", ["none", "summary", "summary_and_period"])
def test_dense_monthly_variances_keep_two_readable_pages(
    tmp_path: Path, explanation_scope: str,
) -> None:
    """Exercise the report shape that previously spilled an empty third page."""
    reason = "Synthetic adjustment reviewed against the corresponding source records."
    controls = [
        ControlOutcome(code, label, ControlDisposition.PASS, "Source evidence verified.")
        for code, label in (
            ("summary_identity", "Summary identity and required values"),
            ("activity_coverage", "Activity weekly coverage"),
            ("micros_coverage", "Micros date coverage"),
            ("tender_evidence", "Tender evidence"),
            ("archive_integrity", "Archive plan and source hashes"),
        )
    ]
    for code, label, amount in (
        ("summary_activity_gift_card_payment_redemptions", "Gift Card Payment / Redemptions", "125"),
        ("summary_activity_net_gift_card_impact", "Net Gift Card Impact", "-125"),
    ):
        controls.append(exact_match_control(
            code=code, label=label, variance=Decimal(amount),
            explanation=reason if explanation_scope != "none" else "",
        ))
    for code, label, amount in (
        ("period_pos_period_pos_gift_card_payment_redemptions", "Period POS Gift Card Payment / Redemptions", "155"),
        ("period_pos_period_pos_net_gift_card_impact", "Period POS Net Gift Card Impact", "-155"),
        ("period_tender_period_tender", "Period tender", "0"),
    ):
        controls.append(variance_control(
            code=code, label=label, variance=Decimal(amount),
            explanation=reason if explanation_scope == "summary_and_period" else "",
        ))
    weekly_codes = set()
    weekly_rows = []
    for day, amount in ((9, "0"), (16, "125"), (23, "0"), (30, "30")):
        end = date(2026, 8, day)
        payment = Decimal(amount)
        disposition = (
            ControlDisposition.EXPLAINED if day == 16
            else ControlDisposition.BLOCK if day == 30 else ControlDisposition.PASS
        )
        weekly_rows.append(WeeklyCloseReportRow(
            week_ending=end,
            coverage="Complete - all business dates present",
            pos_issue_variance=Decimal("0"), pos_payment_variance=payment,
            pos_net_variance=-payment, tender_variance=Decimal("0"),
            disposition=disposition, variance_explanation=reason if day == 16 else "",
        ))
        for metric, value in (("issue", Decimal("0")), ("payment", payment), ("net", -payment)):
            code = f"weekly_pos_week_ending_08_{day}_2026_pos_{metric}"
            weekly_codes.add(code)
            controls.append(variance_control(
                code=code, label=f"Week ending {end:%m/%d/%Y} POS {metric}",
                variance=value, explanation=reason if day == 16 else "",
            ))
    missing_code = "weekly_explanation_20260830"
    weekly_codes.add(missing_code)
    controls.append(ControlOutcome(
        missing_code, "Week ending 08/30/2026 variance explanation",
        ControlDisposition.BLOCK, "Enter the required explanation in the weekly report.",
    ))
    assessment = build_close_assessment(store="9355", darden_variance=Decimal("0"), controls=controls)
    period_disposition = (
        ControlDisposition.EXPLAINED if explanation_scope == "summary_and_period"
        else ControlDisposition.BLOCK
    )
    workbook = tmp_path / "dense-variances.xlsx"
    pdf = tmp_path / "dense-variances.pdf"
    write_monthly_close_report_workbook(MonthlyCloseReportData(
        assessment=assessment, period="FY27-M03",
        period_start=date(2026, 8, 3), period_end=date(2026, 8, 30),
        generated_at=datetime(2026, 9, 1, 9, 0), weekly_rows=tuple(weekly_rows),
        period_pos_net_variance=Decimal("-155"), period_pos_disposition=period_disposition,
        period_tender_variance=Decimal("0"), period_tender_disposition=ControlDisposition.PASS,
        source_labels=(
            "Gift Card Summary", "Weekly Gift Card Activity Reports", "Micros Daily System Totals",
            "Micros Tender Detail", "Darden Credit Memo", "Weekly Variance Explanations",
        ),
        evidence_notes=(
            "Source hashes and archive destinations were planned, but no canonical archive or close manifest was published for this diagnostic.",
            "Scheduled Mondays are accepted only when both activity and tender evidence are zero; existing Monday POS is included normally.",
        ),
        explicit_exceptions=tuple((c.disposition.value, f"{c.label}: {c.message}") for c in assessment.controls if not c.passed),
        weekly_control_codes=frozenset(weekly_codes),
    ), workbook)

    export_monthly_close_report_pdf(
        workbook_path=workbook, pdf_path=pdf,
        expected_location_label="VIRGINIA BEACH - STORE 9355",
    )
    pages = [(page.extract_text() or "") for page in PdfReader(pdf).pages]
    assert len(pages) == 2
    assert all("VIRGINIA BEACH - STORE 9355" in text for text in pages)
    assert "REVIEW REQUIRED" in pages[0]
    for label in ("Gift Card Payment / Redemptions", "Net Gift Card Impact", "Period POS"):
        assert label in pages[0] and label in pages[1]
    for amount in ("125.00", "155.00"):
        assert amount in pages[0] and amount in pages[1]
    assert "30.00" in pages[1]
    assert "Tender: PASS" in pages[1]
    assert f"POS: {period_disposition.value}" in " ".join(pages[1].split())
    assert "Week ending 08/16/2026" in pages[1]
    assert "Week ending 08/30/2026" in pages[1]
    assert "Exception" not in pages[1]
