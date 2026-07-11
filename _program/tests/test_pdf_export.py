from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from gift_card_recon.pdf_export import (
    PdfExportError,
    export_monthly_close_report_pdf,
    validate_monthly_close_pdf,
)


class FakePage:
    def __init__(self, text: str):
        self.text = text

    def extract_text(self) -> str:
        return self.text


def test_validate_monthly_close_pdf_accepts_two_pages_and_location_label(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
):
    pdf_path = tmp_path / "report.pdf"
    pdf_path.write_bytes(b"synthetic nonempty pdf")
    fake_reader = SimpleNamespace(
        pages=[
            FakePage(
                "RICHMOND \u2014 STORE 9354\nDarden Final Checkbox\n"
                "Close Control Matrix\nOpen Actions\nPage 1 of 2"
            ),
            FakePage(
                "RICHMOND - STORE 9354\nWeekly Variances and Coverage\n"
                "Evidence Notes\nPage 2 of 2"
            ),
        ]
    )
    monkeypatch.setattr("gift_card_recon.pdf_export.PdfReader", lambda *_args, **_kwargs: fake_reader)

    result = validate_monthly_close_pdf(
        pdf_path,
        expected_location_label="RICHMOND \u2014 STORE 9354",
    )

    assert result == pdf_path.resolve()


@pytest.mark.parametrize(
    ("pages", "expected_message"),
    [
        ([FakePage("VIRGINIA BEACH \u2014 STORE 9355")], "exactly 2 pages"),
        ([FakePage("Wrong location"), FakePage("Evidence")], "expected location heading"),
    ],
)
def test_validate_monthly_close_pdf_rejects_bad_structure_or_heading(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    pages: list[FakePage],
    expected_message: str,
):
    pdf_path = tmp_path / "report.pdf"
    pdf_path.write_bytes(b"synthetic nonempty pdf")
    monkeypatch.setattr(
        "gift_card_recon.pdf_export.PdfReader",
        lambda *_args, **_kwargs: SimpleNamespace(pages=pages),
    )

    with pytest.raises(PdfExportError, match=expected_message):
        validate_monthly_close_pdf(
            pdf_path,
            expected_location_label="VIRGINIA BEACH \u2014 STORE 9355",
        )


def test_export_surfaces_powershell_failure_and_removes_partial_pdf(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
):
    workbook_path = tmp_path / "close.xlsx"
    workbook_path.write_bytes(b"synthetic workbook")
    script_path = tmp_path / "export.ps1"
    script_path.write_text("exit 1", encoding="utf-8")
    pdf_path = tmp_path / "report.pdf"

    def failed_run(command, **kwargs):
        pdf_path.write_bytes(b"partial")
        return SimpleNamespace(returncode=7, stdout="", stderr="Excel export failed")

    monkeypatch.setattr("gift_card_recon.pdf_export.subprocess.run", failed_run)

    with pytest.raises(PdfExportError, match="exit code 7: Excel export failed"):
        export_monthly_close_report_pdf(
            workbook_path=workbook_path,
            pdf_path=pdf_path,
            expected_location_label="RICHMOND \u2014 STORE 9354",
            powershell_executable="powershell.exe",
            script_path=script_path,
        )

    assert not pdf_path.exists()


def test_export_invokes_versioned_helper_and_validates_result(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
):
    workbook_path = tmp_path / "close.xlsx"
    workbook_path.write_bytes(b"synthetic workbook")
    script_path = tmp_path / "export_monthly_close_pdf_v1.ps1"
    script_path.write_text("exit 0", encoding="utf-8")
    pdf_path = tmp_path / "report.pdf"
    observed: dict[str, object] = {}

    def successful_run(command, **kwargs):
        observed["command"] = command
        observed["kwargs"] = kwargs
        pdf_path.write_bytes(b"synthetic nonempty pdf")
        return SimpleNamespace(returncode=0, stdout="", stderr="")

    def successful_validation(path, *, expected_location_label):
        observed["validated"] = (path, expected_location_label)
        return Path(path)

    monkeypatch.setattr("gift_card_recon.pdf_export.subprocess.run", successful_run)
    monkeypatch.setattr(
        "gift_card_recon.pdf_export.validate_monthly_close_pdf",
        successful_validation,
    )

    result = export_monthly_close_report_pdf(
        workbook_path=workbook_path,
        pdf_path=pdf_path,
        expected_location_label="VIRGINIA BEACH \u2014 STORE 9355",
        powershell_executable="powershell.exe",
        script_path=script_path,
    )

    command = observed["command"]
    assert result == pdf_path.resolve()
    assert str(script_path.resolve()) in command
    assert str(workbook_path.resolve()) in command
    assert str(pdf_path.resolve()) in command
    assert command[-2:] == ["-WorksheetName", "Monthly Close Report"]
    assert observed["validated"] == (
        pdf_path.resolve(),
        "VIRGINIA BEACH \u2014 STORE 9355",
    )
    assert observed["kwargs"]["check"] is False
    assert observed["kwargs"]["timeout"] == 120
