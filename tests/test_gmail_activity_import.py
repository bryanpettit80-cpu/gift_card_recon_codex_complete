from __future__ import annotations

from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

from gift_card_recon.gmail_activity_import import import_gmail_activity_downloads


def test_import_downloaded_gmail_activity_places_monthly_and_latest_weekly_files(tmp_path: Path):
    source_dir = tmp_path / "gmail"
    source_dir.mkdir()
    input_root = tmp_path / "input"
    (input_root / "9354" / "weekly" / "activity").mkdir(parents=True)
    (input_root / "9355" / "weekly" / "activity").mkdir(parents=True)

    create_activity(
        source_dir / "06.14.2026 9354 Gift Card Activity.xlsx",
        store="9354",
        begin="08-JUN-2026",
        end="14-JUN-2026",
    )
    create_activity(
        source_dir / "06.28.2026 9354 Gift Card Activity.xlsx",
        store="9354",
        begin="22-JUN-2026",
        end="28-JUN-2026",
    )
    create_activity(
        source_dir / "06.28.2026 9355 Gift Card Activity.xlsx",
        store="9355",
        begin="22-JUN-2026",
        end="28-JUN-2026",
    )

    reports = import_gmail_activity_downloads(source_dir=source_dir, input_root=input_root)

    assert [report.status for report in reports] == ["imported", "imported", "imported"]
    assert (input_root / "9354" / "2026-06" / "activity" / "06.14.2026 9354 Gift Card Activity.xlsx").exists()
    assert (input_root / "9354" / "2026-06" / "activity" / "06.28.2026 9354 Gift Card Activity.xlsx").exists()
    assert (input_root / "9355" / "2026-06" / "activity" / "06.28.2026 9355 Gift Card Activity.xlsx").exists()
    assert not (input_root / "9354" / "weekly" / "activity" / "06.14.2026 9354 Gift Card Activity.xlsx").exists()
    assert (input_root / "9354" / "weekly" / "activity" / "06.28.2026 9354 Gift Card Activity.xlsx").exists()
    assert (input_root / "9355" / "weekly" / "activity" / "06.28.2026 9355 Gift Card Activity.xlsx").exists()


def test_import_leaves_weekly_activity_when_download_is_older_than_current_week(tmp_path: Path):
    source_dir = tmp_path / "gmail"
    source_dir.mkdir()
    input_root = tmp_path / "input"
    weekly_dir = input_root / "9355" / "weekly" / "activity"
    weekly_dir.mkdir(parents=True)
    create_activity(
        weekly_dir / "06.28.2026 9355 Gift Card Activity.xlsx",
        store="9355",
        begin="22-JUN-2026",
        end="28-JUN-2026",
    )
    create_activity(
        source_dir / "06.14.2026 9355 Gift Card Activity.xlsx",
        store="9355",
        begin="08-JUN-2026",
        end="14-JUN-2026",
    )

    reports = import_gmail_activity_downloads(source_dir=source_dir, input_root=input_root)

    assert len(reports) == 1
    assert reports[0].status == "imported"
    assert "newer week ending 06/28/2026" in reports[0].message
    assert (input_root / "9355" / "2026-06" / "activity" / "06.14.2026 9355 Gift Card Activity.xlsx").exists()
    assert not (weekly_dir / "06.14.2026 9355 Gift Card Activity.xlsx").exists()
    assert (weekly_dir / "06.28.2026 9355 Gift Card Activity.xlsx").exists()


def test_import_archives_older_active_week_when_download_is_newer(tmp_path: Path):
    source_dir = tmp_path / "gmail"
    source_dir.mkdir()
    input_root = tmp_path / "input"
    weekly_dir = input_root / "9354" / "weekly" / "activity"
    weekly_dir.mkdir(parents=True)
    create_activity(
        weekly_dir / "06.21.2026 9354 Gift Card Activity.xlsx",
        store="9354",
        begin="15-JUN-2026",
        end="21-JUN-2026",
    )
    create_activity(
        source_dir / "06.28.2026 9354 Gift Card Activity.xlsx",
        store="9354",
        begin="22-JUN-2026",
        end="28-JUN-2026",
    )

    reports = import_gmail_activity_downloads(source_dir=source_dir, input_root=input_root)

    assert reports[0].status == "imported"
    assert "Archived 1 older weekly file" in reports[0].message
    assert not (weekly_dir / "06.21.2026 9354 Gift Card Activity.xlsx").exists()
    assert (input_root / "9354" / "weekly" / "archive" / "2026-W25" / "06.21.2026 9354 Gift Card Activity.xlsx").exists()
    assert (weekly_dir / "06.28.2026 9354 Gift Card Activity.xlsx").exists()


def test_import_skips_unconfigured_store(tmp_path: Path):
    source_dir = tmp_path / "gmail"
    source_dir.mkdir()
    input_root = tmp_path / "input"
    (input_root / "9355" / "weekly" / "activity").mkdir(parents=True)
    create_activity(
        source_dir / "06.28.2026 9999 Gift Card Activity.xlsx",
        store="9999",
        begin="22-JUN-2026",
        end="28-JUN-2026",
    )

    reports = import_gmail_activity_downloads(source_dir=source_dir, input_root=input_root)

    assert reports[0].status == "skipped"
    assert "not one of the configured stores" in reports[0].message
    assert not (input_root / "9999").exists()


def create_activity(path: Path, *, store: str, begin: str, end: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet 1"
    ws.append([f"All GC Activity BY Rest Number and Date Range  BEGIN DATE: '{begin}', END DATE: '{end}', Rest Number Parameter 1: '{store}'"])
    ws.append(["Card No", "Request", "Request Code Listing", "Business Date", "Transaction No", "Amount SUM", "Promocode", "Authorization Code"])
    ws.append(["0001xxxx", 100, "Activation", "2026-06-28", 1, float(Decimal("10.00")), None, 111111])
    ws.append(["0002xxxx", 202, "Redemption No Nsf", "2026-06-28", 2, float(Decimal("-5.00")), None, 222222])
    wb.save(path)
