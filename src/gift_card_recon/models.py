from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path


@dataclass(frozen=True)
class SummaryData:
    store: str
    total_activations: Decimal
    total_redemptions: Decimal  # signed, usually negative
    payable_redemptions: Decimal | None = None
    gcdr: Decimal | None = None
    net_settlement: Decimal | None = None
    gift_card_franchise_fee_rate: Decimal | None = None
    non_conversion_redemptions: Decimal | None = None
    non_conversion_gcdr_redemptions: Decimal | None = None
    non_conversion_gcdr: Decimal | None = None
    pre_conversion_fee_redemptions: Decimal | None = None
    pre_conversion_gcdr: Decimal | None = None
    post_conversion_fee_redemptions: Decimal | None = None
    post_conversion_gcdr: Decimal | None = None
    cross_location_redeemed: Decimal | None = None
    conversion_promo_codes: set[str] = field(default_factory=set)
    source_file: Path | None = None

    @property
    def total_redemptions_abs(self) -> Decimal:
        return abs(self.total_redemptions)

    @property
    def net_activity(self) -> Decimal:
        return self.total_activations + self.total_redemptions

    @property
    def conversion_redemptions(self) -> Decimal | None:
        if self.non_conversion_redemptions is None:
            return None
        return self.total_redemptions - self.non_conversion_redemptions


@dataclass(frozen=True)
class ActivityRow:
    source_file: str
    card_no: str
    request: str | int | None
    request_code_listing: str
    business_date: date | None
    transaction_no: str | int | None
    amount: Decimal
    promocode: str | None = None
    authorization_code: str | int | None = None

    @property
    def request_label_lower(self) -> str:
        return (self.request_code_listing or "").lower()

    @property
    def is_activation(self) -> bool:
        label = self.request_label_lower
        return "activation" in label and "void" not in label

    @property
    def is_void_activation(self) -> bool:
        label = self.request_label_lower
        return "activation" in label and "void" in label

    @property
    def is_redemption(self) -> bool:
        label = self.request_label_lower
        return "redemption" in label and "void" not in label

    @property
    def is_void_redemption(self) -> bool:
        label = self.request_label_lower
        return "redemption" in label and "void" in label


@dataclass(frozen=True)
class ActivityFileData:
    source_file: Path
    report_begin: date | None
    report_end: date | None
    rows: list[ActivityRow]

    @property
    def report_period_label(self) -> str:
        if self.report_begin and self.report_end:
            return f"{self.report_begin:%d-%b-%Y} to {self.report_end:%d-%b-%Y}".upper()
        return "Unknown"


@dataclass(frozen=True)
class ActivityRollup:
    source_file: str
    report_period: str
    row_count: int
    gross_activations: Decimal
    void_activations: Decimal
    net_activations: Decimal
    gross_redemptions: Decimal
    void_redemptions: Decimal
    net_redemptions: Decimal
    conversion_redemptions: Decimal
    non_conversion_redemptions: Decimal
    net_activity: Decimal


@dataclass(frozen=True)
class DailyRollup:
    business_date: date | None
    source_file: str
    net_activations: Decimal
    net_redemptions: Decimal
    conversion_redemptions: Decimal
    non_conversion_redemptions: Decimal
    net_activity: Decimal


@dataclass(frozen=True)
class MicrosDailyPosControl:
    business_date: date
    pos_gift_card_issue: Decimal
    pos_gift_card_payment: Decimal


@dataclass(frozen=True)
class WeeklyPosVariance:
    week_ending: date | None
    report_begin: date | None
    report_end: date | None
    activity_issue: Decimal
    pos_issue: Decimal
    issue_variance: Decimal
    activity_payment: Decimal
    pos_payment: Decimal
    payment_variance: Decimal
    net_variance: Decimal
    coverage_status: str


@dataclass(frozen=True)
class PosControls:
    store: str
    period: str
    pos_gift_card_issue: Decimal
    pos_gift_card_payment: Decimal  # positive control from POS

    @property
    def net_impact(self) -> Decimal:
        return self.pos_gift_card_issue - self.pos_gift_card_payment


@dataclass(frozen=True)
class SourceFileAudit:
    path: Path
    file_type: str
    size_bytes: int
    modified_at: datetime | None
    sha256: str


@dataclass(frozen=True)
class ReconciliationLine:
    metric: str
    summary_value: Decimal | None
    activity_value: Decimal | None
    activity_variance: Decimal | None
    pos_value: Decimal | None
    pos_variance: Decimal | None
    status: str
    note: str


@dataclass(frozen=True)
class ReconciliationResult:
    store: str
    period: str
    period_end: date | None
    mode: str
    summary: SummaryData | None
    pos_controls: PosControls
    weekly_rollups: list[ActivityRollup]
    daily_rollups: list[DailyRollup]
    raw_rows: list[ActivityRow]
    source_files: list[SourceFileAudit]
    exceptions: list[tuple[str, str]]
    lines: list[ReconciliationLine]

    @property
    def activity_total_activations(self) -> Decimal:
        return sum((r.net_activations for r in self.weekly_rollups), Decimal("0.00"))

    @property
    def activity_total_redemptions(self) -> Decimal:
        return sum((r.net_redemptions for r in self.weekly_rollups), Decimal("0.00"))

    @property
    def activity_net_impact(self) -> Decimal:
        return self.activity_total_activations + self.activity_total_redemptions
