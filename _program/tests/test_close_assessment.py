from __future__ import annotations

from decimal import Decimal

import pytest

from gift_card_recon.close_assessment import (
    CloseStatus,
    ControlDisposition,
    assess_monthly_close,
    exact_match_control,
    integrity_control,
    variance_control,
)
from gift_card_recon.store_config import REVIEW_VARIANCE_LIMIT, get_store_config


def passing_integrity_controls():
    codes = {
        "summary_identity": "Summary identity",
        "activity_identity": "Activity identity",
        "activity_coverage": "Activity coverage",
        "darden_identity": "Darden identity",
        "micros_source": "Micros source",
        "micros_coverage": "Micros coverage",
        "tender_evidence": "Tender evidence",
    }
    return [
        integrity_control(
            code=code,
            label=label,
            passed=True,
            pass_message=f"{label} passed.",
            failure_message=f"{label} failed.",
        )
        for code, label in codes.items()
    ]


def summary_group(value=Decimal("0.00")):
    return {
        "Summary issue": Decimal("0.00"),
        "Summary payment": Decimal("0.00"),
        "Summary net activity": value,
    }


def pos_group(prefix: str, value=Decimal("0.00")):
    return {
        f"{prefix} issue": Decimal("0.00"),
        f"{prefix} payment": Decimal("0.00"),
        f"{prefix} net": value,
    }


def assess(
    *,
    darden=Decimal("0.00"),
    summary=Decimal("0.00"),
    weekly_pos=Decimal("0.00"),
    period_pos=Decimal("0.00"),
    weekly_tender=Decimal("0.00"),
    period_tender=Decimal("0.00"),
):
    return assess_monthly_close(
        store="9354",
        darden_variance=darden,
        summary_activity_variances=summary_group(summary),
        weekly_pos_variances=pos_group("Week ending 07/05/2026 POS", weekly_pos),
        period_pos_variances=pos_group("Period POS", period_pos),
        weekly_tender_variances={
            "Week ending 07/05/2026 tender payment": weekly_tender
        },
        period_tender_variances={"Period tender payment": period_tender},
        integrity_controls=passing_integrity_controls(),
    )


def test_store_configs_centralize_location_and_micros_facts():
    richmond = get_store_config("9354")
    virginia_beach = get_store_config(9355)

    assert richmond.report_heading == "RICHMOND - STORE 9354"
    assert virginia_beach.report_heading == "VIRGINIA BEACH - STORE 9355"
    assert richmond.scheduled_closed_weekdays == frozenset({0})
    assert virginia_beach.scheduled_closed_weekdays == frozenset({0})
    assert richmond.micros_default_path.as_posix() == "../micros_data/RC-Richmond-current"
    assert virginia_beach.micros_default_path.as_posix() == "../GETLinkedData-VB"
    assert richmond.micros_issue_column_number == 121
    assert richmond.micros_payment_column_number == 103
    assert virginia_beach.micros_issue_column_number == 121
    assert virginia_beach.micros_payment_column_number == 103
    assert REVIEW_VARIANCE_LIMIT == Decimal("5.00")


def test_unknown_store_is_rejected_instead_of_falling_back():
    with pytest.raises(ValueError, match="Unsupported store"):
        get_store_config("9999")


def test_exact_control_blocks_missing_malformed_and_nonzero_values():
    for value in (None, "", "not-money", Decimal("0.01")):
        outcome = exact_match_control(code="exact", label="Exact control", variance=value)
        assert outcome.disposition is ControlDisposition.BLOCK

    assert exact_match_control(
        code="exact", label="Exact control", variance=Decimal("0.004")
    ).disposition is ControlDisposition.PASS


def test_variance_control_uses_inclusive_five_dollar_review_limit():
    assert variance_control(
        code="zero", label="Zero", variance=Decimal("0.00")
    ).disposition is ControlDisposition.PASS
    assert variance_control(
        code="amber", label="Amber", variance=Decimal("-5.00")
    ).disposition is ControlDisposition.REVIEW
    assert variance_control(
        code="red", label="Red", variance=Decimal("5.01")
    ).disposition is ControlDisposition.BLOCK


def test_all_controls_pass_is_closed_and_darden_is_separate():
    assessment = assess()

    assert assessment.status is CloseStatus.CLOSED
    assert str(assessment.status) == "CLOSED"
    assert assessment.darden_matched
    assert assessment.can_publish_close
    assert assessment.follow_up_items == ()


@pytest.mark.parametrize(
    "field",
    ["weekly_pos", "period_pos", "weekly_tender", "period_tender"],
)
def test_small_nonzero_pos_or_tender_variance_is_closed_with_review(field):
    assessment = assess(**{field: Decimal("2.43")})

    assert assessment.status is CloseStatus.CLOSED_WITH_REVIEW
    assert assessment.darden_matched
    assert assessment.can_publish_close
    assert len(assessment.review_items) == 1
    assert assessment.review_items[0].variance == Decimal("2.43")


@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("darden", Decimal("0.01")),
        ("summary", Decimal("0.01")),
        ("weekly_pos", Decimal("5.01")),
        ("period_pos", Decimal("-5.01")),
        ("weekly_tender", Decimal("5.01")),
        ("period_tender", Decimal("-5.01")),
    ],
)
def test_exact_mismatches_and_large_variances_require_review(field, value):
    assessment = assess(**{field: value})

    assert assessment.status is CloseStatus.REVIEW_REQUIRED
    assert not assessment.can_publish_close


def test_integrity_blocker_overrides_otherwise_green_numbers():
    failed_identity = integrity_control(
        code="summary_identity",
        label="Summary identity",
        passed=False,
        pass_message="Summary identity passed.",
        failure_message="Summary is for the wrong store.",
    )
    assessment = assess_monthly_close(
        store="9355",
        darden_variance=Decimal("0.00"),
        summary_activity_variances=summary_group(),
        weekly_pos_variances=pos_group("Weekly POS"),
        period_pos_variances=pos_group("Period POS"),
        weekly_tender_variances={"Weekly tender payment": Decimal("0.00")},
        period_tender_variances={"Period tender payment": Decimal("0.00")},
        integrity_controls=[
            failed_identity,
            *[
                item
                for item in passing_integrity_controls()
                if item.code != "summary_identity"
            ],
        ],
    )

    assert assessment.status is CloseStatus.REVIEW_REQUIRED
    assert assessment.darden_matched
    assert assessment.blockers == (failed_identity,)


@pytest.mark.parametrize(
    "missing_group",
    [
        "summary_activity_variances",
        "weekly_pos_variances",
        "period_pos_variances",
        "weekly_tender_variances",
        "period_tender_variances",
    ],
)
def test_absent_control_groups_block_instead_of_greenlighting(missing_group):
    groups = {
        "summary_activity_variances": summary_group(),
        "weekly_pos_variances": pos_group("Weekly POS"),
        "period_pos_variances": pos_group("Period POS"),
        "weekly_tender_variances": {"Weekly tender payment": Decimal("0.00")},
        "period_tender_variances": {"Period tender payment": Decimal("0.00")},
    }
    groups[missing_group] = {}
    assessment = assess_monthly_close(
        store="9354",
        darden_variance=Decimal("0.00"),
        integrity_controls=passing_integrity_controls(),
        **groups,
    )

    assert assessment.status is CloseStatus.REVIEW_REQUIRED
    assert any("No " in blocker.message for blocker in assessment.blockers)


def test_missing_required_integrity_control_blocks_close():
    assessment = assess_monthly_close(
        store="9354",
        darden_variance=Decimal("0.00"),
        summary_activity_variances=summary_group(),
        weekly_pos_variances=pos_group("Weekly POS"),
        period_pos_variances=pos_group("Period POS"),
        weekly_tender_variances={"Weekly tender payment": Decimal("0.00")},
        period_tender_variances={"Period tender payment": Decimal("0.00")},
        integrity_controls=[],
        additional_required_integrity_codes={"archive_integrity"},
    )

    assert assessment.status is CloseStatus.REVIEW_REQUIRED
    blocker_codes = {item.code for item in assessment.blockers}
    assert blocker_codes.issuperset(
        {"archive_integrity", "micros_coverage", "summary_identity"}
    )


def test_partially_populated_weekly_control_groups_cannot_close():
    assessment = assess_monthly_close(
        store="9354",
        darden_variance=Decimal("0.00"),
        summary_activity_variances=summary_group(),
        weekly_pos_variances=pos_group("Only one week"),
        period_pos_variances=pos_group("Period POS"),
        weekly_tender_variances={"Only one week tender": Decimal("0.00")},
        period_tender_variances={"Period tender": Decimal("0.00")},
        integrity_controls=passing_integrity_controls(),
        expected_week_count=5,
    )

    assert assessment.status is CloseStatus.REVIEW_REQUIRED
    blocker_codes = {control.code for control in assessment.blockers}
    assert "weekly_pos_completeness" in blocker_codes
    assert "weekly_tender_completeness" in blocker_codes


def test_required_integrity_control_cannot_be_amber():
    integrity = passing_integrity_controls()
    integrity[0] = integrity[0].__class__(
        code=integrity[0].code,
        label=integrity[0].label,
        disposition=ControlDisposition.REVIEW,
        message="Identity remains uncertain.",
    )
    assessment = assess_monthly_close(
        store="9354",
        darden_variance=Decimal("0.00"),
        summary_activity_variances=summary_group(),
        weekly_pos_variances=pos_group("Weekly POS"),
        period_pos_variances=pos_group("Period POS"),
        weekly_tender_variances={"Weekly tender": Decimal("0.00")},
        period_tender_variances={"Period tender": Decimal("0.00")},
        integrity_controls=integrity,
    )

    assert assessment.status is CloseStatus.REVIEW_REQUIRED
    assert next(control for control in assessment.controls if control.code == "summary_identity").is_blocking
