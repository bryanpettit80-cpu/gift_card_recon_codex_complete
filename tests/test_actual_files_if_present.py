from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import pytest

from gift_card_recon.parsers import discover_input_files, parse_activity_file, parse_pos_controls, parse_summary
from gift_card_recon.reconcile import build_reconciliation
from gift_card_recon.utils import parse_date


def test_may_2026_9354_if_input_files_are_present():
    input_dir = Path("input/9354/2026-05")
    if not input_dir.exists():
        pytest.skip("Actual May 2026 files are not present. Copy them into input/9354/2026-05 to enable this test.")
    try:
        summary_path, activity_paths, pos_path = discover_input_files(input_dir)
    except Exception as exc:
        pytest.skip(f"Actual files are not fully staged: {exc}")
    if pos_path is None:
        pytest.skip("pos_controls.csv not found.")

    summary = parse_summary(summary_path, "9354")
    activities = [parse_activity_file(path, summary.conversion_promo_codes) for path in activity_paths]
    pos = parse_pos_controls(pos_path, "9354", "2026-05")
    result = build_reconciliation(store="9354", period="2026-05", period_end=parse_date("2026-05-31"), summary=summary, activities=activities, pos_controls=pos)

    assert result.activity_total_activations == Decimal("11507.00")
    assert result.activity_total_redemptions == Decimal("-49867.48")
    assert result.lines[0].pos_variance == Decimal("135.00")
    assert result.lines[1].pos_variance == Decimal("2.27")
