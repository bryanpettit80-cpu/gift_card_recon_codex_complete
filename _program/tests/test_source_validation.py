from __future__ import annotations

from dataclasses import replace
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path
import zipfile

import pytest

from gift_card_recon.micros import (
    load_micros_evidence,
    period_tender_variance,
    resolve_micros_export_dir,
    validate_micros_source,
    weekly_tender_variances,
)
from gift_card_recon.models import ActivityFileData, ActivityRow
from gift_card_recon.parsers import ParseError
from gift_card_recon.source_validation import validate_activity_evidence
from gift_card_recon.store_config import get_store_config


PERIOD_START = date(2026, 6, 1)
PERIOD_END = date(2026, 6, 7)


def test_activity_evidence_rejects_wrong_store() -> None:
    activity = _activity(store="9354")

    with pytest.raises(ParseError, match="identifies store 9354"):
        validate_activity_evidence(
            [activity],
            store="9355",
            period_start=PERIOD_START,
            period_end=PERIOD_END,
            expected_week_endings=[PERIOD_END],
        )


def test_activity_evidence_rejects_duplicate_and_overlapping_week() -> None:
    first = _activity(path=Path("first.xlsx"))
    duplicate = _activity(path=Path("duplicate.xlsx"))

    with pytest.raises(ParseError) as exc_info:
        validate_activity_evidence(
            [first, duplicate],
            store="9355",
            period_start=PERIOD_START,
            period_end=PERIOD_END,
            expected_week_endings=[PERIOD_END],
        )

    message = str(exc_info.value)
    assert "duplicate week-ending" in message
    assert "overlapping activity reports" in message


def test_activity_evidence_rejects_missing_and_out_of_period_weeks() -> None:
    extra_end = PERIOD_END + timedelta(days=7)
    activity = _activity(
        path=Path("extra.xlsx"),
        begin=extra_end - timedelta(days=6),
        end=extra_end,
        transaction_date=extra_end,
    )

    with pytest.raises(ParseError) as exc_info:
        validate_activity_evidence(
            [activity],
            store="9355",
            period_start=PERIOD_START,
            period_end=PERIOD_END,
            expected_week_endings=[PERIOD_END],
        )

    message = str(exc_info.value)
    assert "missing expected week-ending" in message
    assert "out-of-period week-ending" in message


def test_missing_scheduled_monday_is_accepted_only_when_activity_and_tender_are_zero(
    tmp_path: Path,
) -> None:
    activity = _activity(transaction_date=date(2026, 6, 2))
    validated = _validated_activity(activity)
    micros_dir = tmp_path / "micros"
    _write_micros_week(micros_dir, include_monday=False)
    config = replace(get_store_config("9355"), micros_default_path=micros_dir)

    evidence = load_micros_evidence(
        micros_dir,
        config=config,
        activity_evidence=validated,
        period_start=PERIOD_START,
        period_end=PERIOD_END,
    )

    assert evidence.accepted_closed_dates == (PERIOD_START,)
    assert period_tender_variance(evidence) == Decimal("0.00")
    assert weekly_tender_variances(evidence, week_endings=[PERIOD_END]) == {
        "Week ending 06/07/2026 tender": Decimal("0.00")
    }


def test_missing_scheduled_monday_blocks_when_real_activity_exists(tmp_path: Path) -> None:
    activity = _activity(transaction_date=PERIOD_START)
    validated = _validated_activity(activity)
    micros_dir = tmp_path / "micros"
    _write_micros_week(micros_dir, include_monday=False)
    config = replace(get_store_config("9355"), micros_default_path=micros_dir)

    with pytest.raises(ParseError, match="2026-06-01 is missing"):
        load_micros_evidence(
            micros_dir,
            config=config,
            activity_evidence=validated,
            period_start=PERIOD_START,
            period_end=PERIOD_END,
        )


def test_existing_monday_pos_is_included_normally(tmp_path: Path) -> None:
    activity = _activity(transaction_date=PERIOD_START)
    validated = _validated_activity(activity)
    micros_dir = tmp_path / "micros"
    _write_micros_week(micros_dir, include_monday=True)
    config = replace(get_store_config("9355"), micros_default_path=micros_dir)

    evidence = load_micros_evidence(
        micros_dir,
        config=config,
        activity_evidence=validated,
        period_start=PERIOD_START,
        period_end=PERIOD_END,
    )

    assert evidence.accepted_closed_dates == ()
    assert PERIOD_START in evidence.daily_pos_by_date


def test_missing_or_malformed_tender_evidence_blocks(tmp_path: Path) -> None:
    activity = _validated_activity(_activity(transaction_date=date(2026, 6, 2)))
    micros_dir = tmp_path / "micros"
    _write_micros_week(micros_dir, include_monday=False)
    config = replace(get_store_config("9355"), micros_default_path=micros_dir)
    (micros_dir / "TENDER_DETAIL.TXT").unlink()

    with pytest.raises(ParseError, match="TENDER_DETAIL.TXT"):
        load_micros_evidence(
            micros_dir,
            config=config,
            activity_evidence=activity,
            period_start=PERIOD_START,
            period_end=PERIOD_END,
        )

    (micros_dir / "TENDER_DETAIL.TXT").write_text("bad,row", encoding="utf-8")
    with pytest.raises(ParseError, match="expected at least 4"):
        load_micros_evidence(
            micros_dir,
            config=config,
            activity_evidence=activity,
            period_start=PERIOD_START,
            period_end=PERIOD_END,
        )


def test_wrong_live_micros_source_is_rejected() -> None:
    richmond = get_store_config("9354")
    virginia_beach = get_store_config("9355")

    with pytest.raises(ParseError, match="different location"):
        validate_micros_source(virginia_beach.micros_default_path, richmond)


def test_archive_replacement_with_same_name_cannot_reuse_stale_tender(tmp_path: Path) -> None:
    source_dir = tmp_path / "source"
    _write_micros_week(source_dir, include_monday=True)
    archive = tmp_path / "Micros3700.zip"
    with zipfile.ZipFile(archive, "w") as bundle:
        bundle.write(source_dir / "DLYSYSTT.TXT", "DLYSYSTT.TXT")
        bundle.write(source_dir / "TENDER_DETAIL.TXT", "TENDER_DETAIL.TXT")
    first = resolve_micros_export_dir(archive, tmp_path / "extract")
    assert (first / "TENDER_DETAIL.TXT").exists()

    with zipfile.ZipFile(archive, "w") as bundle:
        bundle.write(source_dir / "DLYSYSTT.TXT", "DLYSYSTT.TXT")
    second = resolve_micros_export_dir(archive, tmp_path / "extract")

    assert second != first
    assert not (second / "TENDER_DETAIL.TXT").exists()


def test_offsetting_tender_activity_does_not_qualify_as_zero_monday(tmp_path: Path) -> None:
    activity = _validated_activity(_activity(transaction_date=date(2026, 6, 2)))
    micros_dir = tmp_path / "micros"
    _write_micros_week(micros_dir, include_monday=False)
    with (micros_dir / "TENDER_DETAIL.TXT").open("a", encoding="utf-8") as stream:
        stream.write("\n'2026-06-01',100.00,350,'G C Payment','T'")
        stream.write("\n'2026-06-01',-100.00,350,'G C Payment','T'")
    config = replace(get_store_config("9355"), micros_default_path=micros_dir)

    with pytest.raises(ParseError, match="2026-06-01 is missing"):
        load_micros_evidence(
            micros_dir,
            config=config,
            activity_evidence=activity,
            period_start=PERIOD_START,
            period_end=PERIOD_END,
        )


def test_tender_detail_requires_open_day_date_coverage(tmp_path: Path) -> None:
    activity = _validated_activity(_activity(transaction_date=date(2026, 6, 2)))
    micros_dir = tmp_path / "micros"
    _write_micros_week(micros_dir, include_monday=True)
    tender_path = micros_dir / "TENDER_DETAIL.TXT"
    lines = [line for line in tender_path.read_text(encoding="utf-8").splitlines() if "2026-06-03" not in line]
    tender_path.write_text("\n".join(lines), encoding="utf-8")
    config = replace(get_store_config("9355"), micros_default_path=micros_dir)

    with pytest.raises(ParseError, match="2026-06-03 is missing from TENDER_DETAIL.TXT"):
        load_micros_evidence(
            micros_dir,
            config=config,
            activity_evidence=activity,
            period_start=PERIOD_START,
            period_end=PERIOD_END,
        )


def _validated_activity(activity: ActivityFileData):
    return validate_activity_evidence(
        [activity],
        store="9355",
        period_start=PERIOD_START,
        period_end=PERIOD_END,
        expected_week_endings=[PERIOD_END],
    )


def _activity(
    *,
    path: Path = Path("activity.xlsx"),
    store: str = "9355",
    begin: date = PERIOD_START,
    end: date = PERIOD_END,
    transaction_date: date = date(2026, 6, 2),
) -> ActivityFileData:
    row = ActivityRow(
        source_file=path.name,
        card_no="0001xxxx",
        request=100,
        request_code_listing="Activation",
        business_date=transaction_date,
        transaction_no=1,
        amount=Decimal("10.00"),
    )
    return ActivityFileData(
        source_file=path,
        report_begin=begin,
        report_end=end,
        rows=[row],
        store=store,
    )


def _write_micros_week(folder: Path, *, include_monday: bool) -> None:
    folder.mkdir(parents=True)
    config = get_store_config("9355")
    system_lines: list[str] = []
    tender_lines: list[str] = []
    start = PERIOD_START if include_monday else PERIOD_START + timedelta(days=1)
    for offset in range((PERIOD_END - start).days + 1):
        business_date = start + timedelta(days=offset)
        row = ["0"] * 132
        row[0] = f"{business_date:%Y-%m-%d} 00:00:00.000"
        row[config.micros_issue_column_index] = "0.00"
        row[config.micros_payment_column_index] = "0.00"
        system_lines.append(",".join(row))
        tender_lines.append(
            f"'{business_date:%Y-%m-%d}',0.00,350,'  g c   payment  ','T'"
        )
    (folder / "DLYSYSTT.TXT").write_text("\n".join(system_lines), encoding="utf-8")
    (folder / "TENDER_DETAIL.TXT").write_text("\n".join(tender_lines), encoding="utf-8")
