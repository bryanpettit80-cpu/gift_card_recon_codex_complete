from __future__ import annotations

import re
from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from enum import Enum

from gift_card_recon.store_config import REVIEW_VARIANCE_LIMIT, StoreConfig, get_store_config


_CENT = Decimal("0.01")
_ZERO = Decimal("0.00")
REQUIRED_CLOSE_INTEGRITY_CODES = frozenset(
    {
        "summary_identity",
        "activity_identity",
        "activity_coverage",
        "darden_identity",
        "micros_source",
        "micros_coverage",
        "tender_evidence",
    }
)


class _ValueEnum(str, Enum):
    def __str__(self) -> str:
        return self.value


class ControlDisposition(_ValueEnum):
    PASS = "PASS"
    REVIEW = "REVIEW"
    BLOCK = "BLOCK"


class CloseStatus(_ValueEnum):
    CLOSED = "CLOSED"
    CLOSED_WITH_REVIEW = "CLOSED WITH REVIEW"
    REVIEW_REQUIRED = "REVIEW REQUIRED"


@dataclass(frozen=True)
class ControlOutcome:
    code: str
    label: str
    disposition: ControlDisposition
    message: str
    variance: Decimal | None = None

    def __post_init__(self) -> None:
        if not self.code.strip() or not self.label.strip() or not self.message.strip():
            raise ValueError("Control outcomes require a code, label, and message.")
        if not isinstance(self.disposition, ControlDisposition):
            raise TypeError("Control disposition must be a ControlDisposition value.")

    @property
    def passed(self) -> bool:
        return self.disposition is ControlDisposition.PASS

    @property
    def needs_review(self) -> bool:
        return self.disposition is ControlDisposition.REVIEW

    @property
    def is_blocking(self) -> bool:
        return self.disposition is ControlDisposition.BLOCK


@dataclass(frozen=True)
class CloseAssessment:
    store_config: StoreConfig
    darden_matched: bool
    controls: tuple[ControlOutcome, ...]
    status: CloseStatus = field(init=False)

    def __post_init__(self) -> None:
        if not self.controls:
            raise ValueError("A close assessment requires at least one control outcome.")
        codes = [control.code for control in self.controls]
        if len(codes) != len(set(codes)):
            raise ValueError("Close-assessment control codes must be unique.")
        darden_controls = [
            control
            for control in self.controls
            if control.code == "darden_summary_match"
        ]
        if len(darden_controls) != 1:
            raise ValueError(
                "A close assessment requires exactly one Darden-to-Summary control."
            )
        if self.darden_matched is not darden_controls[0].passed:
            raise ValueError(
                "darden_matched must agree with the Darden-to-Summary control."
            )
        if any(control.is_blocking for control in self.controls):
            status = CloseStatus.REVIEW_REQUIRED
        elif any(control.needs_review for control in self.controls):
            status = CloseStatus.CLOSED_WITH_REVIEW
        else:
            status = CloseStatus.CLOSED
        object.__setattr__(self, "status", status)

    @property
    def store(self) -> str:
        return self.store_config.store

    @property
    def location_name(self) -> str:
        return self.store_config.location_name

    @property
    def can_publish_close(self) -> bool:
        return self.status is not CloseStatus.REVIEW_REQUIRED

    @property
    def blockers(self) -> tuple[ControlOutcome, ...]:
        return tuple(control for control in self.controls if control.is_blocking)

    @property
    def review_items(self) -> tuple[ControlOutcome, ...]:
        return tuple(control for control in self.controls if control.needs_review)

    @property
    def follow_up_items(self) -> tuple[str, ...]:
        return tuple(
            control.message for control in self.controls if not control.passed
        )


def integrity_control(
    *,
    code: str,
    label: str,
    passed: bool,
    pass_message: str,
    failure_message: str,
) -> ControlOutcome:
    return ControlOutcome(
        code=code,
        label=label,
        disposition=(ControlDisposition.PASS if passed else ControlDisposition.BLOCK),
        message=pass_message if passed else failure_message,
    )


def exact_match_control(
    *,
    code: str,
    label: str,
    variance: object,
) -> ControlOutcome:
    normalized = _strict_money(variance)
    if normalized is None:
        return ControlOutcome(
            code=code,
            label=label,
            disposition=ControlDisposition.BLOCK,
            message=f"{label} is missing or is not a valid monetary value.",
        )
    if normalized == _ZERO:
        return ControlOutcome(
            code=code,
            label=label,
            disposition=ControlDisposition.PASS,
            message=f"{label} matches to the cent.",
            variance=normalized,
        )
    return ControlOutcome(
        code=code,
        label=label,
        disposition=ControlDisposition.BLOCK,
        message=f"{label} variance is {normalized:+,.2f}; an exact match is required.",
        variance=normalized,
    )


def variance_control(
    *,
    code: str,
    label: str,
    variance: object,
    review_limit: Decimal = REVIEW_VARIANCE_LIMIT,
) -> ControlOutcome:
    normalized = _strict_money(variance)
    limit = _strict_money(review_limit)
    if limit is None or limit < _ZERO:
        raise ValueError("The review limit must be a non-negative monetary value.")
    if normalized is None:
        return ControlOutcome(
            code=code,
            label=label,
            disposition=ControlDisposition.BLOCK,
            message=f"{label} is missing or is not a valid monetary value.",
        )
    if normalized == _ZERO:
        disposition = ControlDisposition.PASS
        message = f"{label} has no variance."
    elif abs(normalized) <= limit:
        disposition = ControlDisposition.REVIEW
        message = (
            f"{label} variance is {normalized:+,.2f}; review and document "
            f"the amount before final sign-off (limit {limit:,.2f})."
        )
    else:
        disposition = ControlDisposition.BLOCK
        message = (
            f"{label} variance is {normalized:+,.2f}; it exceeds the "
            f"{limit:,.2f} close limit."
        )
    return ControlOutcome(
        code=code,
        label=label,
        disposition=disposition,
        message=message,
        variance=normalized,
    )


def build_close_assessment(
    *,
    store: str | int,
    darden_variance: object,
    controls: Iterable[ControlOutcome],
) -> CloseAssessment:
    config = get_store_config(store)
    darden = exact_match_control(
        code="darden_summary_match",
        label="Darden credit memo to Summary Net Settlement",
        variance=darden_variance,
    )
    supplied = tuple(controls)
    if not supplied:
        supplied = (
            ControlOutcome(
                code="close_evidence",
                label="Close evidence",
                disposition=ControlDisposition.BLOCK,
                message="No close controls beyond the Darden match were evaluated.",
            ),
        )
    if any(control.code == darden.code for control in supplied):
        raise ValueError(f"Control code {darden.code!r} is reserved for the Darden match.")
    return CloseAssessment(
        store_config=config,
        darden_matched=darden.passed,
        controls=(darden, *supplied),
    )


def assess_monthly_close(
    *,
    store: str | int,
    darden_variance: object,
    summary_activity_variances: Mapping[str, object],
    weekly_pos_variances: Mapping[str, object],
    period_pos_variances: Mapping[str, object],
    weekly_tender_variances: Mapping[str, object],
    period_tender_variances: Mapping[str, object],
    integrity_controls: Iterable[ControlOutcome],
    additional_required_integrity_codes: Iterable[str] = (),
    expected_week_count: int = 1,
) -> CloseAssessment:
    """Build the close decision, requiring weekly and period controls explicitly."""

    if expected_week_count <= 0:
        raise ValueError("expected_week_count must be greater than zero.")
    controls = list(integrity_controls)
    supplied_integrity_codes = {control.code for control in controls}
    required_integrity_codes = REQUIRED_CLOSE_INTEGRITY_CODES | set(
        additional_required_integrity_codes
    )
    for code in sorted(required_integrity_codes - supplied_integrity_codes):
        controls.append(
            ControlOutcome(
                code=code,
                label=_label_from_code(code),
                disposition=ControlDisposition.BLOCK,
                message=f"Required integrity control {code!r} was not evaluated.",
            )
        )
    for index, control in enumerate(controls):
        if (
            control.code in required_integrity_codes
            and control.disposition is ControlDisposition.REVIEW
        ):
            controls[index] = ControlOutcome(
                code=control.code,
                label=control.label,
                disposition=ControlDisposition.BLOCK,
                message=(
                    f"{control.label} is a required integrity control and cannot be "
                    f"left in review: {control.message}"
                ),
                variance=control.variance,
            )

    expected_counts = {
        "summary_activity": 3,
        "weekly_pos": expected_week_count * 3,
        "period_pos": 3,
        "weekly_tender": expected_week_count,
        "period_tender": 1,
    }
    supplied_groups = {
        "summary_activity": summary_activity_variances,
        "weekly_pos": weekly_pos_variances,
        "period_pos": period_pos_variances,
        "weekly_tender": weekly_tender_variances,
        "period_tender": period_tender_variances,
    }
    for prefix, expected_count in expected_counts.items():
        actual_count = len(supplied_groups[prefix])
        if actual_count != expected_count:
            controls.append(
                ControlOutcome(
                    code=f"{prefix}_completeness",
                    label=f"{_label_from_code(prefix)} completeness",
                    disposition=ControlDisposition.BLOCK,
                    message=(
                        f"Expected {expected_count} {prefix.replace('_', ' ')} control(s); "
                        f"evaluated {actual_count}."
                    ),
                )
            )

    controls.extend(
        _controls_for_group(
            prefix="summary_activity",
            group_label="Summary-to-activity",
            values=summary_activity_variances,
            exact=True,
        )
    )
    controls.extend(
        _controls_for_group(
            prefix="weekly_pos",
            group_label="weekly POS",
            values=weekly_pos_variances,
            exact=False,
        )
    )
    controls.extend(
        _controls_for_group(
            prefix="period_pos",
            group_label="period POS",
            values=period_pos_variances,
            exact=False,
        )
    )
    controls.extend(
        _controls_for_group(
            prefix="weekly_tender",
            group_label="weekly tender",
            values=weekly_tender_variances,
            exact=False,
        )
    )
    controls.extend(
        _controls_for_group(
            prefix="period_tender",
            group_label="period tender",
            values=period_tender_variances,
            exact=False,
        )
    )
    return build_close_assessment(
        store=store,
        darden_variance=darden_variance,
        controls=controls,
    )


def _controls_for_group(
    *,
    prefix: str,
    group_label: str,
    values: Mapping[str, object],
    exact: bool,
) -> list[ControlOutcome]:
    if not values:
        return [
            ControlOutcome(
                code=f"{prefix}_variance_evidence",
                label=f"{group_label} evidence",
                disposition=ControlDisposition.BLOCK,
                message=f"No {group_label.lower()} controls were evaluated.",
            )
        ]

    outcomes: list[ControlOutcome] = []
    used_codes: set[str] = set()
    for index, (label, value) in enumerate(values.items(), start=1):
        suffix = re.sub(r"[^a-z0-9]+", "_", str(label).lower()).strip("_")
        base_code = f"{prefix}_{suffix or index}"
        code = base_code
        duplicate = 2
        while code in used_codes:
            code = f"{base_code}_{duplicate}"
            duplicate += 1
        used_codes.add(code)
        outcome = (
            exact_match_control(code=code, label=str(label), variance=value)
            if exact
            else variance_control(code=code, label=str(label), variance=value)
        )
        outcomes.append(outcome)
    return outcomes


def _strict_money(value: object) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, Decimal):
        parsed = value
    elif isinstance(value, int):
        parsed = Decimal(value)
    elif isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        try:
            parsed = Decimal(text)
        except InvalidOperation:
            return None
    else:
        return None
    if not parsed.is_finite():
        return None
    return parsed.quantize(_CENT, rounding=ROUND_HALF_UP)


def _label_from_code(code: str) -> str:
    return str(code).replace("_", " ").strip().title()
