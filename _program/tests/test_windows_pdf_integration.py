from __future__ import annotations

import sys
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest
from pypdf import PdfReader

from gift_card_recon.close_assessment import ControlDisposition, ControlOutcome, build_close_assessment
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
