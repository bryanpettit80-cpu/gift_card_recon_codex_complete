from __future__ import annotations

import re
import shutil
import subprocess
import unicodedata
from importlib import resources
from pathlib import Path

from pypdf import PdfReader


REPORT_SHEET_NAME = "Monthly Close Report"
EXPORT_SCRIPT_NAME = "export_monthly_close_pdf_v1.ps1"
EXPECTED_PAGE_COUNT = 2


class PdfExportError(RuntimeError):
    """Raised when a monthly-close PDF cannot be exported or validated."""


def export_monthly_close_report_pdf(
    *,
    workbook_path: Path,
    pdf_path: Path,
    expected_location_label: str,
    powershell_executable: str | None = None,
    script_path: Path | None = None,
    timeout_seconds: int = 120,
) -> Path:
    """Export and validate the two-page monthly-close report worksheet.

    Excel performs the export so the PDF uses the workbook's print settings and
    remains visually identical to the canonical report worksheet. The caller is
    expected to provide a temporary ``pdf_path`` and publish it only after this
    function returns successfully.
    """

    workbook = Path(workbook_path).expanduser().resolve()
    destination = Path(pdf_path).expanduser().resolve()
    location_label = str(expected_location_label or "").strip()

    if not workbook.is_file():
        raise PdfExportError(f"Workbook not found: {workbook}")
    if destination.suffix.lower() != ".pdf":
        raise PdfExportError(f"PDF destination must end in .pdf: {destination}")
    if workbook == destination:
        raise PdfExportError("Workbook and PDF destination must be different files.")
    if not location_label:
        raise PdfExportError("An expected location label is required for PDF validation.")
    if timeout_seconds <= 0:
        raise PdfExportError("PDF export timeout must be greater than zero seconds.")

    destination.parent.mkdir(parents=True, exist_ok=True)
    _remove_stale_output(destination)

    executable = powershell_executable or _find_powershell()
    if script_path is not None:
        helper = Path(script_path).expanduser().resolve()
        if not helper.is_file():
            raise PdfExportError(f"PDF export helper not found: {helper}")
        _run_excel_export(
            executable=executable,
            script_path=helper,
            workbook_path=workbook,
            pdf_path=destination,
            timeout_seconds=timeout_seconds,
        )
    else:
        script_resource = resources.files("gift_card_recon.data").joinpath(EXPORT_SCRIPT_NAME)
        try:
            with resources.as_file(script_resource) as bundled_helper:
                _run_excel_export(
                    executable=executable,
                    script_path=bundled_helper,
                    workbook_path=workbook,
                    pdf_path=destination,
                    timeout_seconds=timeout_seconds,
                )
        except (FileNotFoundError, ModuleNotFoundError) as exc:
            raise PdfExportError(
                f"Bundled PDF export helper is unavailable: {EXPORT_SCRIPT_NAME}"
            ) from exc

    try:
        return validate_monthly_close_pdf(
            destination,
            expected_location_label=location_label,
        )
    except PdfExportError:
        _remove_partial_output(destination)
        raise


def validate_monthly_close_pdf(
    pdf_path: Path,
    *,
    expected_location_label: str,
) -> Path:
    """Require a readable, nonempty, two-page PDF with the location heading."""

    path = Path(pdf_path).expanduser().resolve()
    location_label = str(expected_location_label or "").strip()
    if not location_label:
        raise PdfExportError("An expected location label is required for PDF validation.")
    if not path.is_file():
        raise PdfExportError(f"Excel did not create the requested PDF: {path}")
    try:
        size = path.stat().st_size
    except OSError as exc:
        raise PdfExportError(f"Could not inspect exported PDF {path}: {exc}") from exc
    if size <= 0:
        raise PdfExportError(f"Excel created an empty PDF: {path}")

    try:
        reader = PdfReader(path, strict=False)
        page_count = len(reader.pages)
        if page_count != EXPECTED_PAGE_COUNT:
            raise PdfExportError(
                f"Monthly-close PDF must contain exactly {EXPECTED_PAGE_COUNT} pages; "
                f"found {page_count}: {path}"
            )
        page_text = [page.extract_text() or "" for page in reader.pages]
    except PdfExportError:
        raise
    except Exception as exc:
        raise PdfExportError(f"Could not read exported PDF {path}: {exc}") from exc

    missing_heading_pages = [
        index
        for index, text in enumerate(page_text, start=1)
        if _normalize_text(location_label) not in _normalize_text(text)
    ]
    if missing_heading_pages:
        raise PdfExportError(
            f"Exported PDF does not contain the expected location heading "
            f"{location_label!r} on page(s) {missing_heading_pages}: {path}"
        )
    required_sections = {
        1: ("Darden Final Checkbox", "Close Control Matrix", "Open Actions", "Page 1 of 2"),
        2: ("Weekly Variances and Coverage", "Evidence Notes", "Page 2 of 2"),
    }
    for page_number, labels in required_sections.items():
        normalized_page = _normalize_text(page_text[page_number - 1])
        missing = [label for label in labels if _normalize_text(label) not in normalized_page]
        if missing:
            raise PdfExportError(
                f"Monthly-close PDF page {page_number} is missing required section(s): "
                f"{', '.join(missing)}: {path}"
            )
    return path


def _run_excel_export(
    *,
    executable: str,
    script_path: Path,
    workbook_path: Path,
    pdf_path: Path,
    timeout_seconds: int,
) -> None:
    command = [
        executable,
        "-NoLogo",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script_path),
        "-WorkbookPath",
        str(workbook_path),
        "-PdfPath",
        str(pdf_path),
        "-WorksheetName",
        REPORT_SHEET_NAME,
    ]
    try:
        completed = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except subprocess.TimeoutExpired as exc:
        _remove_partial_output(pdf_path)
        raise PdfExportError(
            f"Excel PDF export timed out after {timeout_seconds} seconds."
        ) from exc
    except OSError as exc:
        _remove_partial_output(pdf_path)
        raise PdfExportError(f"Could not start PowerShell for Excel PDF export: {exc}") from exc

    if completed.returncode != 0:
        _remove_partial_output(pdf_path)
        diagnostic = _subprocess_diagnostic(completed.stdout, completed.stderr)
        raise PdfExportError(
            f"Excel PDF export failed with exit code {completed.returncode}: {diagnostic}"
        )

    try:
        pdf_is_nonempty = pdf_path.is_file() and pdf_path.stat().st_size > 0
    except OSError as exc:
        _remove_partial_output(pdf_path)
        raise PdfExportError(f"Could not inspect Excel PDF output {pdf_path}: {exc}") from exc
    if not pdf_is_nonempty:
        _remove_partial_output(pdf_path)
        raise PdfExportError(f"Excel did not create a nonempty PDF: {pdf_path}")


def _find_powershell() -> str:
    for candidate in ("powershell.exe", "pwsh.exe", "pwsh"):
        executable = shutil.which(candidate)
        if executable:
            return executable
    raise PdfExportError(
        "PowerShell is required for Excel PDF export but was not found on PATH."
    )


def _remove_stale_output(path: Path) -> None:
    if not path.exists():
        return
    try:
        path.unlink()
    except OSError as exc:
        raise PdfExportError(
            f"Cannot replace the requested temporary PDF; close or remove it first: {path} ({exc})"
        ) from exc


def _remove_partial_output(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except OSError:
        # Preserve the primary export failure. The caller's temporary workspace
        # cleanup can make a second best-effort attempt.
        pass


def _normalize_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", str(value))
    normalized = normalized.translate(
        str.maketrans(
            {
                "\u2010": "-",
                "\u2011": "-",
                "\u2012": "-",
                "\u2013": "-",
                "\u2014": "-",
                "\u2212": "-",
            }
        )
    )
    return re.sub(r"\s+", " ", normalized).strip().casefold()


def _subprocess_diagnostic(stdout: str | None, stderr: str | None) -> str:
    message = str(stderr or "").strip() or str(stdout or "").strip()
    if not message:
        return "PowerShell returned no diagnostic message."
    return re.sub(r"\s+", " ", message)[:1000]
