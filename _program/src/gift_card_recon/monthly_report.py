from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import TYPE_CHECKING, Iterable, Sequence

from gift_card_recon.close_assessment import (
    CloseAssessment,
    CloseStatus,
    ControlDisposition,
    ControlOutcome,
    REQUIRED_CLOSE_INTEGRITY_CODES,
)
from gift_card_recon.models import MonthlyCloseCertification, ReconciliationResult

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet


ACCOUNTING_FMT = '$#,##0.00;[Red]($#,##0.00);$-'
DATE_FMT = "mm/dd/yyyy"

_NAVY = "17365D"
_BLUE = "4472C4"
_LIGHT_BLUE = "D9EAF7"
_GREEN = "E2F0D9"
_DARK_GREEN = "375623"
_AMBER = "FFF2CC"
_DARK_AMBER = "7F6000"
_RED = "F4CCCC"
_DARK_RED = "9C0006"
_GRAY = "F2F2F2"
_TEXT = "333333"

DEFAULT_EVIDENCE_LABELS = (
    "Gift Card Summary",
    "Weekly Gift Card Activity Reports",
    "Micros Daily System Totals",
    "Micros Tender Detail",
    "Darden Credit Memo",
)


@dataclass(frozen=True)
class WeeklyCloseReportRow:
    """One assessed week for page 2 of the monthly close report.

    The caller supplies ``disposition``. The renderer intentionally does not
    infer a status from the monetary values.
    """

    week_ending: date | None
    coverage: str
    pos_issue_variance: Decimal | None
    pos_payment_variance: Decimal | None
    pos_net_variance: Decimal | None
    tender_variance: Decimal | None
    disposition: ControlDisposition
    evidence_note: str = ""


@dataclass(frozen=True)
class MonthlyCloseReportData:
    """Presentation-only data for a two-page close certificate or diagnostic."""

    assessment: CloseAssessment
    period: str
    period_start: date
    period_end: date
    generated_at: datetime
    certification: MonthlyCloseCertification | None = None
    result: ReconciliationResult | None = None
    weekly_rows: tuple[WeeklyCloseReportRow, ...] = ()
    period_pos_net_variance: Decimal | None = None
    period_pos_disposition: ControlDisposition | None = None
    period_tender_variance: Decimal | None = None
    period_tender_disposition: ControlDisposition | None = None
    evidence_notes: tuple[str, ...] = ()
    source_labels: tuple[str, ...] = field(default_factory=lambda: DEFAULT_EVIDENCE_LABELS)
    explicit_exceptions: tuple[tuple[str, str], ...] = ()
    weekly_control_codes: frozenset[str] = frozenset()

    def __post_init__(self) -> None:
        store = self.assessment.store
        if self.period_end < self.period_start:
            raise ValueError("The monthly close report end date precedes its start date.")
        if self.certification is not None:
            if self.certification.store != store:
                raise ValueError("Certification store does not match the close assessment.")
            if self.certification.period != self.period:
                raise ValueError("Certification period does not match the report period.")
        if self.result is not None:
            if self.result.store != store:
                raise ValueError("Reconciliation store does not match the close assessment.")
            if self.result.period != self.period:
                raise ValueError("Reconciliation period does not match the report period.")

    @property
    def exceptions(self) -> tuple[tuple[str, str], ...]:
        result_exceptions: Iterable[tuple[str, str]] = ()
        if self.result is not None:
            result_exceptions = self.result.exceptions
        return tuple(result_exceptions) + self.explicit_exceptions


def write_monthly_close_report_workbook(
    data: MonthlyCloseReportData,
    output_path: Path,
) -> Path:
    """Write a report-only workbook, including blocked-run diagnostics."""

    try:
        from openpyxl import Workbook
    except ImportError as exc:
        raise RuntimeError(
            "openpyxl is required to write the monthly close report workbook."
        ) from exc

    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    write_monthly_close_report(workbook.active, data)
    workbook.save(destination)
    return destination


def write_monthly_close_report(ws: Worksheet, data: MonthlyCloseReportData) -> None:
    """Render the authoritative two-page monthly report into ``ws``.

    Overall and row-level dispositions are consumed from assessed inputs. This
    module performs presentation calculations only (for example, identifying
    the largest absolute weekly variance).
    """

    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.page import PageMargins
    from openpyxl.worksheet.pagebreak import Break

    ws.title = "Monthly Close Report"
    ws.sheet_view.showGridLines = False
    thin = Side(style="thin", color="C9D5E3")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)

    heading = data.assessment.store_config.report_heading
    _write_page_heading(ws, 1, heading, data)

    ws.merge_cells("A4:H5")
    status_cell = ws["A4"]
    status_cell.value = data.assessment.status.value
    status_cell.font = Font(
        bold=True,
        color=_status_text(data.assessment.status),
        size=18,
    )
    status_cell.fill = PatternFill("solid", fgColor=_status_fill(data.assessment.status))
    status_cell.alignment = Alignment(horizontal="center", vertical="center")
    status_cell.border = border
    ws.row_dimensions[4].height = 25
    ws.row_dimensions[5].height = 25

    _section_title(ws, 7, "Darden Final Checkbox")
    certification = data.certification
    summary_total = certification.summary_net_settlement if certification else None
    darden_total = certification.darden_credit_memo.total if certification else None
    darden_variance = certification.variance if certification else _darden_variance(data.assessment)
    darden_evaluated = certification is not None or darden_variance is not None
    darden_result = (
        "MATCHED"
        if data.assessment.darden_matched
        else "MISMATCHED" if darden_evaluated else "NOT EVALUATED"
    )
    cards = (
        ("Summary Net Settlement", summary_total, ACCOUNTING_FMT),
        ("Darden Credit Memo", darden_total, ACCOUNTING_FMT),
        ("Darden Variance", darden_variance, ACCOUNTING_FMT),
        ("Darden Result", darden_result, None),
    )
    for index, (label, value, number_format) in enumerate(cards):
        start_col = 1 + (index * 2)
        ws.merge_cells(start_row=8, start_column=start_col, end_row=8, end_column=start_col + 1)
        ws.merge_cells(start_row=9, start_column=start_col, end_row=10, end_column=start_col + 1)
        label_cell = ws.cell(8, start_col, label)
        value_cell = ws.cell(9, start_col, _excel_money(value) if number_format else value)
        label_cell.fill = PatternFill("solid", fgColor=_BLUE)
        label_cell.font = Font(bold=True, color="FFFFFF", size=9)
        value_fill = (
            _GREEN if data.assessment.darden_matched else _RED if darden_evaluated else _GRAY
        )
        value_text = (
            _DARK_GREEN
            if data.assessment.darden_matched
            else _DARK_RED if darden_evaluated else _TEXT
        )
        value_cell.fill = PatternFill("solid", fgColor=value_fill)
        value_cell.font = Font(
            bold=True,
            color=value_text,
            size=12,
        )
        for cell in (label_cell, value_cell):
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        if number_format:
            value_cell.number_format = number_format

    _section_title(ws, 12, "Close Control Matrix")
    _write_cells(ws, 13, ("Control", "", "", "Disposition", "Variance", "Result", "", ""))
    ws.merge_cells("A13:C13")
    ws.merge_cells("F13:H13")
    _style_header(ws, 13, 8)
    executive_controls = _executive_controls(data)
    control_start = 14
    for row_index, control in enumerate(executive_controls, start=control_start):
        ws.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=3)
        ws.merge_cells(start_row=row_index, start_column=6, end_row=row_index, end_column=8)
        ws.cell(row_index, 1, control.label)
        ws.cell(row_index, 4, control.disposition.value)
        ws.cell(row_index, 5, _excel_money(control.variance))
        ws.cell(row_index, 6, control.message)
        ws.cell(row_index, 5).number_format = ACCOUNTING_FMT
        _style_assessed_row(ws, row_index, control.disposition, border)

    next_row = control_start + len(executive_controls) + 1
    _section_title(ws, next_row, "Open Actions")
    next_row += 1
    open_controls = tuple(control for control in data.assessment.controls if not control.passed)
    if open_controls:
        sequence = 1
        weekly_open = tuple(
            control
            for control in open_controls
            if _is_weekly_child_control(control, data.weekly_control_codes)
        )
        nonweekly_open = tuple(control for control in open_controls if control not in weekly_open)
        if weekly_open:
            disposition = _worst_disposition(weekly_open)
            affected_weeks = {
                control.label.split(" POS ", 1)[0].split(" tender", 1)[0]
                for control in weekly_open
            }
            action_text = (
                f"{len(weekly_open)} weekly control(s) across {len(affected_weeks)} week(s) "
                "require review. See page 2 for every assessed amount and follow-up."
            )
            ws.merge_cells(start_row=next_row, start_column=2, end_row=next_row, end_column=8)
            ws.cell(next_row, 1, f"{disposition.value} {sequence}")
            ws.cell(next_row, 2, action_text)
            _style_action_row(ws, next_row, disposition, border)
            ws.row_dimensions[next_row].height = 27
            next_row += 1
            sequence += 1
        for disposition, controls in _group_open_controls(nonweekly_open):
            for chunk in _chunks(controls, 3):
                action_text = " | ".join(
                    f"{control.label}: {control.message}" for control in chunk
                )
                ws.merge_cells(start_row=next_row, start_column=2, end_row=next_row, end_column=8)
                ws.cell(next_row, 1, f"{disposition.value} {sequence}")
                ws.cell(next_row, 2, action_text)
                _style_action_row(ws, next_row, disposition, border)
                ws.row_dimensions[next_row].height = min(42, 24 + (len(action_text) // 130) * 9)
                next_row += 1
                sequence += 1
    else:
        ws.merge_cells(start_row=next_row, start_column=1, end_row=next_row, end_column=8)
        ws.cell(next_row, 1, "No open actions. Every close control passed.")
        ws.cell(next_row, 1).fill = PatternFill("solid", fgColor=_GREEN)
        ws.cell(next_row, 1).font = Font(color=_DARK_GREEN)
        ws.cell(next_row, 1).alignment = Alignment(wrap_text=True, vertical="center")
        ws.cell(next_row, 1).border = border
        next_row += 1

    page_one_end = next_row
    page_two_start = page_one_end + 2
    ws.row_breaks.append(Break(id=page_two_start - 1))

    _write_page_heading(ws, page_two_start, heading, data)
    weekly_title = page_two_start + 3
    _section_title(ws, weekly_title, "Weekly Variances and Coverage")
    weekly_header = weekly_title + 1
    headers = (
        "Week Ending",
        "Coverage",
        "POS Issue Variance",
        "POS Payment Variance",
        "POS Net Variance",
        "Tender Variance",
        "Status",
        "Evidence / Action",
    )
    _write_cells(ws, weekly_header, headers)
    _style_header(ws, weekly_header, 8)
    row_cursor = weekly_header + 1
    if data.weekly_rows:
        for weekly in data.weekly_rows:
            _write_cells(
                ws,
                row_cursor,
                (
                    weekly.week_ending,
                    weekly.coverage,
                    _excel_money(weekly.pos_issue_variance),
                    _excel_money(weekly.pos_payment_variance),
                    _excel_money(weekly.pos_net_variance),
                    _excel_money(weekly.tender_variance),
                    weekly.disposition.value,
                    weekly.evidence_note,
                ),
            )
            ws.cell(row_cursor, 1).number_format = DATE_FMT
            for column in range(3, 7):
                ws.cell(row_cursor, column).number_format = ACCOUNTING_FMT
            _style_weekly_row(ws, row_cursor, weekly.disposition, border)
            row_cursor += 1
    else:
        ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=8)
        ws.cell(row_cursor, 1, "No weekly variance rows were available; see the control matrix for the blocking assessment.")
        ws.cell(row_cursor, 1).fill = PatternFill("solid", fgColor=_RED)
        ws.cell(row_cursor, 1).font = Font(color=_DARK_RED)
        ws.cell(row_cursor, 1).alignment = Alignment(wrap_text=True, vertical="center")
        ws.cell(row_cursor, 1).border = border
        row_cursor += 1

    period_status = _period_status_text(data)
    _write_cells(
        ws,
        row_cursor,
        (
            "PERIOD NET",
            "Independent period controls",
            None,
            None,
            _excel_money(data.period_pos_net_variance),
            _excel_money(data.period_tender_variance),
            period_status,
            "Period controls are assessed independently from weekly controls.",
        ),
    )
    for column in range(3, 7):
        ws.cell(row_cursor, column).number_format = ACCOUNTING_FMT
    _style_total_row(ws, row_cursor, border)
    row_cursor += 2

    _section_title(ws, row_cursor, "Variance Highlights")
    row_cursor += 1
    largest = _largest_weekly_variance(data.weekly_rows)
    ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=3)
    ws.merge_cells(start_row=row_cursor, start_column=5, end_row=row_cursor, end_column=8)
    ws.cell(row_cursor, 1, "Largest Weekly Absolute Variance")
    ws.cell(row_cursor, 4, _excel_money(largest[0]) if largest else None)
    ws.cell(row_cursor, 5, largest[1] if largest else "Not available")
    ws.cell(row_cursor, 4).number_format = ACCOUNTING_FMT
    _style_highlight_row(ws, row_cursor, border)
    row_cursor += 1
    ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=3)
    ws.merge_cells(start_row=row_cursor, start_column=5, end_row=row_cursor, end_column=8)
    ws.cell(row_cursor, 1, "Period-Net POS Variance")
    ws.cell(row_cursor, 4, _excel_money(data.period_pos_net_variance))
    ws.cell(row_cursor, 5, _disposition_text(data.period_pos_disposition))
    ws.cell(row_cursor, 4).number_format = ACCOUNTING_FMT
    _style_highlight_row(ws, row_cursor, border)
    row_cursor += 2

    _section_title(ws, row_cursor, "Unified Exceptions and Review Items")
    row_cursor += 1
    unified_items = _unified_items(data)
    if unified_items:
        for severity, message in unified_items:
            disposition = _disposition_from_severity(severity)
            ws.merge_cells(start_row=row_cursor, start_column=2, end_row=row_cursor, end_column=8)
            ws.cell(row_cursor, 1, severity)
            ws.cell(row_cursor, 2, message)
            _style_action_row(ws, row_cursor, disposition, border)
            row_cursor += 1
    else:
        ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=8)
        ws.cell(row_cursor, 1, "No exceptions or review items.")
        ws.cell(row_cursor, 1).fill = PatternFill("solid", fgColor=_GREEN)
        ws.cell(row_cursor, 1).font = Font(color=_DARK_GREEN)
        ws.cell(row_cursor, 1).alignment = Alignment(wrap_text=True, vertical="center")
        ws.cell(row_cursor, 1).border = border
        row_cursor += 1

    row_cursor += 1
    _section_title(ws, row_cursor, "Evidence Notes")
    row_cursor += 1
    notes = _evidence_notes(data)
    for note in notes:
        ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=8)
        ws.cell(row_cursor, 1, f"- {note}")
        ws.cell(row_cursor, 1).fill = PatternFill("solid", fgColor=_GRAY)
        ws.cell(row_cursor, 1).font = Font(color=_TEXT, size=9)
        ws.cell(row_cursor, 1).alignment = Alignment(wrap_text=True, vertical="center")
        ws.cell(row_cursor, 1).border = border
        row_cursor += 1

    widths = {"A": 15, "B": 25, "C": 16, "D": 18, "E": 16, "F": 16, "G": 17, "H": 34}
    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    ws.freeze_panes = None
    ws.print_area = f"A1:H{row_cursor}"
    ws.print_options.horizontalCentered = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 0
    ws.page_setup.fitToHeight = 0
    # A fixed, readable scale plus one explicit row break keeps Excel from
    # rebalancing content across the intended page boundary.
    ws.page_setup.scale = 70
    ws.sheet_properties.pageSetUpPr.fitToPage = False
    ws.sheet_properties.pageSetUpPr.autoPageBreaks = False
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.35, bottom=0.4, header=0.15, footer=0.2)
    ws.oddFooter.center.text = "Page &P of &N"
    ws.oddFooter.right.text = f"{data.period} | {data.assessment.store_config.location_name}"


def _write_page_heading(ws: Worksheet, row: int, heading: str, data: MonthlyCloseReportData) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    title = ws.cell(row, 1, heading)
    title.font = Font(bold=True, color="FFFFFF", size=20)
    title.fill = PatternFill("solid", fgColor=_NAVY)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 34

    ws.merge_cells(start_row=row + 1, start_column=1, end_row=row + 1, end_column=8)
    subtitle = ws.cell(
        row + 1,
        1,
        (
            f"{data.period} | {data.period_start:%B %d, %Y} to {data.period_end:%B %d, %Y} | "
            f"Generated {data.generated_at:%B %d, %Y at %I:%M %p}"
        ),
    )
    subtitle.font = Font(color="44546A", italic=True, size=9)
    subtitle.alignment = Alignment(horizontal="center", vertical="center")


def _section_title(ws: Worksheet, row: int, title: str) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    cell = ws.cell(row, 1, title)
    cell.font = Font(bold=True, color=_NAVY, size=11)
    cell.fill = PatternFill("solid", fgColor=_LIGHT_BLUE)
    cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[row].height = 20


def _style_header(ws: Worksheet, row: int, max_column: int) -> None:
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    side = Side(style="thin", color="C9D5E3")
    for column in range(1, max_column + 1):
        cell = ws.cell(row, column)
        cell.fill = PatternFill("solid", fgColor=_BLUE)
        cell.font = Font(bold=True, color="FFFFFF", size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(top=side, bottom=side, left=side, right=side)
    ws.row_dimensions[row].height = 28


def _style_assessed_row(ws: Worksheet, row: int, disposition: ControlDisposition, border) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill

    for column in (1, 4, 5, 6):
        cell = ws.cell(row, column)
        cell.border = border
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.font = Font(size=8)
    ws.cell(row, 1).font = Font(bold=True, size=8)
    status_cell = ws.cell(row, 4)
    status_cell.fill = PatternFill("solid", fgColor=_disposition_fill(disposition))
    status_cell.font = Font(bold=True, color=_disposition_text_color(disposition), size=8)
    status_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 22


def _style_action_row(ws: Worksheet, row: int, disposition: ControlDisposition, border) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill

    for column in (1, 2):
        cell = ws.cell(row, column)
        cell.border = border
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.font = Font(size=8, bold=(column == 1))
    ws.cell(row, 1).fill = PatternFill("solid", fgColor=_disposition_fill(disposition))
    ws.cell(row, 1).font = Font(bold=True, color=_disposition_text_color(disposition), size=8)
    ws.row_dimensions[row].height = 26


def _style_weekly_row(ws: Worksheet, row: int, disposition: ControlDisposition, border) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill

    for column in range(1, 9):
        cell = ws.cell(row, column)
        cell.border = border
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.font = Font(size=8)
    status_cell = ws.cell(row, 7)
    status_cell.fill = PatternFill("solid", fgColor=_disposition_fill(disposition))
    status_cell.font = Font(bold=True, color=_disposition_text_color(disposition), size=8)
    status_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    note_length = len(str(ws.cell(row, 8).value or ""))
    ws.row_dimensions[row].height = min(48, 28 + (note_length // 85) * 9)


def _style_total_row(ws: Worksheet, row: int, border) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill

    for column in range(1, 9):
        cell = ws.cell(row, column)
        cell.border = border
        cell.fill = PatternFill("solid", fgColor=_LIGHT_BLUE)
        cell.font = Font(bold=True, size=8)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[row].height = 30


def _style_highlight_row(ws: Worksheet, row: int, border) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill

    for column in (1, 4, 5):
        cell = ws.cell(row, column)
        cell.border = border
        cell.fill = PatternFill("solid", fgColor=_GRAY)
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.font = Font(bold=(column == 1), size=9)
    ws.row_dimensions[row].height = 25


def _write_cells(ws: Worksheet, row: int, values: Sequence[object]) -> None:
    for column, value in enumerate(values, start=1):
        ws.cell(row, column, value)


def _excel_money(value: Decimal | None) -> float | None:
    return None if value is None else float(value)


def _darden_variance(assessment: CloseAssessment) -> Decimal | None:
    return next(
        (control.variance for control in assessment.controls if control.code == "darden_summary_match"),
        None,
    )


def _executive_controls(data: MonthlyCloseReportData) -> tuple[ControlOutcome, ...]:
    # Page 1 is an executive certificate, while page 2 carries every weekly
    # child control. Keep a stable set of core controls and add any non-weekly
    # exception so a failure can never disappear from the matrix.
    core_codes = {
        "darden_summary_match",
        "summary_identity",
        "activity_coverage",
        "micros_coverage",
        "tender_evidence",
        "archive_integrity",
        "summary_activity_net_gift_card_impact",
        "period_pos_period_pos_net_gift_card_impact",
        "period_tender_period_tender",
    }
    return tuple(
        control
        for control in data.assessment.controls
        if control.code in core_codes
        or (
            not control.passed
            and not _is_weekly_child_control(control, data.weekly_control_codes)
        )
    )


def _worst_disposition(controls: Sequence[ControlOutcome]) -> ControlDisposition:
    dispositions = {control.disposition for control in controls}
    if ControlDisposition.BLOCK in dispositions:
        return ControlDisposition.BLOCK
    if ControlDisposition.REVIEW in dispositions:
        return ControlDisposition.REVIEW
    return ControlDisposition.PASS


def _is_weekly_child_control(
    control: ControlOutcome,
    explicit_weekly_codes: frozenset[str],
) -> bool:
    if explicit_weekly_codes:
        return control.code in explicit_weekly_codes
    if control.code in REQUIRED_CLOSE_INTEGRITY_CODES:
        return False
    if control.code == "darden_summary_match" or control.code.startswith("summary_activity_"):
        return False
    text = f"{control.code} {control.label}".lower()
    if "period" in text or "monthly" in text:
        return False
    if not control.code.startswith(("pos_", "tender_")):
        return False
    return any(token in text for token in ("week", "weekly", "week_ending", "2026", "2027"))


def _group_open_controls(
    controls: Sequence[ControlOutcome],
) -> tuple[tuple[ControlDisposition, tuple[ControlOutcome, ...]], ...]:
    groups: list[tuple[ControlDisposition, tuple[ControlOutcome, ...]]] = []
    for disposition in (ControlDisposition.BLOCK, ControlDisposition.REVIEW):
        members = tuple(control for control in controls if control.disposition is disposition)
        if members:
            groups.append((disposition, members))
    return tuple(groups)


def _chunks(
    controls: Sequence[ControlOutcome],
    size: int,
) -> tuple[tuple[ControlOutcome, ...], ...]:
    return tuple(
        tuple(controls[index : index + size])
        for index in range(0, len(controls), size)
    )


def _status_fill(status: CloseStatus) -> str:
    return {
        CloseStatus.CLOSED: _GREEN,
        CloseStatus.CLOSED_WITH_REVIEW: _AMBER,
        CloseStatus.REVIEW_REQUIRED: _RED,
    }[status]


def _status_text(status: CloseStatus) -> str:
    return {
        CloseStatus.CLOSED: _DARK_GREEN,
        CloseStatus.CLOSED_WITH_REVIEW: _DARK_AMBER,
        CloseStatus.REVIEW_REQUIRED: _DARK_RED,
    }[status]


def _disposition_fill(disposition: ControlDisposition) -> str:
    return {
        ControlDisposition.PASS: _GREEN,
        ControlDisposition.REVIEW: _AMBER,
        ControlDisposition.BLOCK: _RED,
    }[disposition]


def _disposition_text_color(disposition: ControlDisposition) -> str:
    return {
        ControlDisposition.PASS: _DARK_GREEN,
        ControlDisposition.REVIEW: _DARK_AMBER,
        ControlDisposition.BLOCK: _DARK_RED,
    }[disposition]


def _disposition_text(disposition: ControlDisposition | None) -> str:
    return "NOT AVAILABLE" if disposition is None else disposition.value


def _period_status_text(data: MonthlyCloseReportData) -> str:
    return (
        f"POS: {_disposition_text(data.period_pos_disposition)}; "
        f"Tender: {_disposition_text(data.period_tender_disposition)}"
    )


def _largest_weekly_variance(
    weekly_rows: Sequence[WeeklyCloseReportRow],
) -> tuple[Decimal, str] | None:
    candidates: list[tuple[Decimal, str]] = []
    for row in weekly_rows:
        week = row.week_ending.strftime("%m/%d/%Y") if row.week_ending else "Unknown week"
        values = (
            ("POS issue", row.pos_issue_variance),
            ("POS payment", row.pos_payment_variance),
            ("POS net", row.pos_net_variance),
            ("Tender", row.tender_variance),
        )
        for label, value in values:
            if value is not None:
                candidates.append((abs(value), f"{label} | Week ending {week} | Reported {value:+,.2f}"))
    return max(candidates, key=lambda item: item[0]) if candidates else None


def _unified_items(data: MonthlyCloseReportData) -> tuple[tuple[str, str], ...]:
    items: list[tuple[str, str]] = [
        (control.disposition.value, f"{control.label}: {control.message}")
        for control in data.assessment.controls
        if not control.passed
    ]
    items.extend(data.exceptions)
    unique: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for severity, message in items:
        item = (str(severity), str(message))
        if item not in seen:
            seen.add(item)
            unique.append(item)
    return tuple(unique)


def _disposition_from_severity(severity: str) -> ControlDisposition:
    normalized = str(severity).strip().upper()
    if normalized in {"BLOCK", "ERROR", "REVIEW REQUIRED", "FAIL", "FAILED"}:
        return ControlDisposition.BLOCK
    if normalized in {"REVIEW", "WARNING", "WARN", "MINOR VARIANCE"}:
        return ControlDisposition.REVIEW
    return ControlDisposition.PASS


def _evidence_notes(data: MonthlyCloseReportData) -> tuple[str, ...]:
    notes = [f"Evidence retained: {label}." for label in data.source_labels]
    notes.extend(str(note) for note in data.evidence_notes)
    notes.append(
        "The close disposition is the centralized assessment shown on page 1; the Darden result is one control and is not the overall close status."
    )
    return tuple(notes)
