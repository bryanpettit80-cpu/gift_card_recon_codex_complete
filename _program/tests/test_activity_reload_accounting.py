from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
import pytest

from gift_card_recon.models import ActivityRow, PosControls
from gift_card_recon.parsers import parse_activity_file
from gift_card_recon.reconcile import build_reconciliation, rollup_activity_file, rollup_daily


def test_parsed_reload_eliminates_false_weekly_issue_variance(tmp_path: Path) -> None:
    path = tmp_path / "08.30.2026 9355 Gift Card Activity.xlsx"
    _write_activity(
        path,
        [
            ["0001", 100, "Activation", "2026-08-25", 1, 870, None],
            ["0002", 202, "Redemption No Nsf", "2026-08-27", 2, -476, None],
            ["0003", 300, "Reload", "2026-08-29", 3, 50, None],
        ],
    )
    activity = parse_activity_file(path)
    result = build_reconciliation(
        store="9355",
        period="2026-W35",
        period_end=date(2026, 8, 30),
        summary=None,
        activities=[activity],
        pos_controls=PosControls("9355", "2026-W35", Decimal("920"), Decimal("511")),
        mode="weekly",
    )

    assert result.activity_total_activations == Decimal("920.00")
    assert result.lines[0].pos_variance == Decimal("0.00")
    assert result.lines[1].pos_variance == Decimal("35.00")
    assert result.lines[2].pos_variance == Decimal("-35.00")
    reload_day = next(row for row in result.daily_rollups if row.business_date == date(2026, 8, 29))
    assert reload_day.net_activations == Decimal("50.00")
    assert reload_day.net_redemptions == Decimal("0.00")
    assert reload_day.net_activity == Decimal("50.00")
    assert result.raw_rows[-1].request == 300
    assert result.raw_rows[-1].request_code_listing == "Reload"


def test_parsed_void_reload_reverses_issuance_without_affecting_redemptions(tmp_path: Path) -> None:
    path = tmp_path / "08.30.2026 9355 Gift Card Activity.xlsx"
    _write_activity(
        path,
        [
            ["0001", 300, "Reload", "2026-08-29", 1, 75, "PROMO"],
            ["0001", None, "Void Of Reload", "2026-08-29", 2, -25, "PROMO"],
            ["0002", 202, "Redemption No Nsf", "2026-08-29", 3, -40, "PROMO"],
            ["0002", 203, "Void Of Redemption", "2026-08-29", 4, 10, "PROMO"],
        ],
    )
    activity = parse_activity_file(path)
    weekly = rollup_activity_file(activity, {"PROMO"})
    daily = rollup_daily(activity.rows, {"PROMO"})

    assert weekly.gross_activations == Decimal("75.00")
    assert weekly.void_activations == Decimal("-25.00")
    assert weekly.net_activations == Decimal("50.00")
    assert weekly.net_redemptions == Decimal("-30.00")
    assert weekly.conversion_redemptions == Decimal("-30.00")
    assert weekly.net_activity == Decimal("20.00")
    assert len(daily) == 1
    assert daily[0].net_activations == weekly.net_activations
    assert daily[0].net_redemptions == weekly.net_redemptions
    assert daily[0].conversion_redemptions == weekly.conversion_redemptions
    assert daily[0].net_activity == weekly.net_activity
    assert activity.rows[1].request_code_listing == "Void Of Reload"


@pytest.mark.parametrize(
    ("label", "flags"),
    [
        ("  rElOaD  ", (True, False, False, False)),
        (" VOID  OF\tRELOAD ", (False, True, False, False)),
        ("Reload Inquiry", (False, False, False, False)),
        ("Void Of Reload Inquiry", (False, False, False, False)),
        ("Unknown", (False, False, False, False)),
        ("", (False, False, False, False)),
        ("Redemption No Nsf", (False, False, True, False)),
        ("Void Of Redemption", (False, False, False, True)),
    ],
)
def test_reload_classification_uses_supported_labels_only(label: str, flags: tuple[bool, ...]) -> None:
    row = ActivityRow("activity.xlsx", "0001", 300, label, date(2026, 8, 29), 1, Decimal("50"))
    assert (row.is_activation, row.is_void_activation, row.is_redemption, row.is_void_redemption) == flags
    assert row.request_code_listing == label


def _write_activity(path: Path, rows: list[list[object]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["BEGIN DATE: '24-AUG-2026', END DATE: '30-AUG-2026', Rest Number Parameter 1: '9355'"])
    sheet.append(["Card No", "Request", "Request Code Listing", "Business Date", "Transaction No", "Amount SUM", "Promocode"])
    for row in rows:
        sheet.append(row)
    workbook.save(path)
