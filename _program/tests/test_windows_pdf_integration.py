from __future__ import annotations

import sys
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest

from gift_card_recon.close_assessment import ControlDisposition, ControlOutcome, build_close_assessment
from gift_card_recon.models import DardenCreditMemo, MonthlyCloseCertification
from gift_card_recon.monthly_report import MonthlyCloseReportData, write_monthly_close_report_workbook
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
        expected_location_label="VIRGINIA BEACH — STORE 9355",
    )

    assert result == pdf.resolve()
    assert pdf.stat().st_size > 0
