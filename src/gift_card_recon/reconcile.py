from __future__ import annotations

from collections import defaultdict
from datetime import date
from decimal import Decimal
from pathlib import Path

from gift_card_recon.models import (
    ActivityFileData,
    ActivityRollup,
    DailyRollup,
    PosControls,
    ReconciliationLine,
    ReconciliationResult,
    SourceFileAudit,
    SummaryData,
)
from gift_card_recon.utils import file_modified_at, money, sha256_file, variance_status


def build_reconciliation(
    *,
    store: str,
    period: str,
    period_end: date | None,
    summary: SummaryData | None,
    activities: list[ActivityFileData],
    pos_controls: PosControls,
    mode: str = "monthly",
    exceptions: list[tuple[str, str]] | None = None,
) -> ReconciliationResult:
    exceptions = exceptions or []
    mode = mode.lower()
    if mode not in {"monthly", "weekly"}:
        raise ValueError(f"Unsupported reconciliation mode: {mode}")
    if mode == "monthly" and summary is None:
        raise ValueError("Monthly reconciliation requires a Gift Card Summary file.")

    conversion_promo_codes = summary.conversion_promo_codes if summary else set()
    weekly_rollups = [rollup_activity_file(activity, conversion_promo_codes) for activity in activities]
    raw_rows = [row for activity in activities for row in activity.rows]
    daily_rollups = rollup_daily(raw_rows, conversion_promo_codes)
    source_files = audit_source_files(summary, activities)

    activity_total_activations = sum((r.net_activations for r in weekly_rollups), Decimal("0.00"))
    activity_total_redemptions = sum((r.net_redemptions for r in weekly_rollups), Decimal("0.00"))
    activity_net_impact = money(activity_total_activations + activity_total_redemptions)

    issue_activity_variance = money(summary.total_activations - activity_total_activations) if summary else None
    payment_activity_variance = money(abs(summary.total_redemptions) - abs(activity_total_redemptions)) if summary else None
    net_activity_variance = money(summary.net_activity - activity_net_impact) if summary else None

    pos_issue_variance = money(pos_controls.pos_gift_card_issue - activity_total_activations)
    pos_payment_variance = money(pos_controls.pos_gift_card_payment - abs(activity_total_redemptions))
    pos_net_variance = money(pos_controls.net_impact - activity_net_impact)

    lines = [
        ReconciliationLine(
            metric="Gift Card Issue / Activations",
            summary_value=summary.total_activations if summary else None,
            activity_value=money(activity_total_activations),
            activity_variance=issue_activity_variance,
            pos_value=pos_controls.pos_gift_card_issue,
            pos_variance=pos_issue_variance,
            status=_combined_status(issue_activity_variance, pos_issue_variance),
            note=_line_note(mode, "issue", summary is not None),
        ),
        ReconciliationLine(
            metric="Gift Card Payment / Redemptions",
            summary_value=abs(summary.total_redemptions) if summary else None,
            activity_value=money(abs(activity_total_redemptions)),
            activity_variance=payment_activity_variance,
            pos_value=pos_controls.pos_gift_card_payment,
            pos_variance=pos_payment_variance,
            status=_combined_status(payment_activity_variance, pos_payment_variance),
            note=_line_note(mode, "payment", summary is not None),
        ),
        ReconciliationLine(
            metric="Net Gift Card Impact",
            summary_value=money(summary.net_activity) if summary else None,
            activity_value=activity_net_impact,
            activity_variance=net_activity_variance,
            pos_value=pos_controls.net_impact,
            pos_variance=pos_net_variance,
            status=_combined_status(net_activity_variance, pos_net_variance),
            note=_line_note(mode, "net", summary is not None),
        ),
    ]

    return ReconciliationResult(
        store=str(store),
        period=str(period),
        period_end=period_end,
        mode=mode,
        summary=summary,
        pos_controls=pos_controls,
        weekly_rollups=weekly_rollups,
        daily_rollups=daily_rollups,
        raw_rows=raw_rows,
        source_files=source_files,
        exceptions=exceptions,
        lines=lines,
    )


def rollup_activity_file(activity: ActivityFileData, conversion_promo_codes: set[str]) -> ActivityRollup:
    gross_activations = Decimal("0.00")
    void_activations = Decimal("0.00")
    gross_redemptions = Decimal("0.00")
    void_redemptions = Decimal("0.00")
    conversion_redemptions = Decimal("0.00")

    for row in activity.rows:
        if row.is_activation:
            gross_activations += row.amount
        elif row.is_void_activation:
            void_activations += row.amount
        elif row.is_redemption:
            gross_redemptions += row.amount
        elif row.is_void_redemption:
            void_redemptions += row.amount

        if (row.is_redemption or row.is_void_redemption) and row.promocode in conversion_promo_codes:
            conversion_redemptions += row.amount

    gross_activations = money(gross_activations)
    void_activations = money(void_activations)
    net_activations = money(gross_activations + void_activations)
    gross_redemptions = money(gross_redemptions)
    void_redemptions = money(void_redemptions)
    net_redemptions = money(gross_redemptions + void_redemptions)
    conversion_redemptions = money(conversion_redemptions)
    non_conversion_redemptions = money(net_redemptions - conversion_redemptions)
    net_activity = money(net_activations + net_redemptions)

    return ActivityRollup(
        source_file=activity.source_file.name,
        report_period=activity.report_period_label,
        row_count=len(activity.rows),
        gross_activations=gross_activations,
        void_activations=void_activations,
        net_activations=net_activations,
        gross_redemptions=gross_redemptions,
        void_redemptions=void_redemptions,
        net_redemptions=net_redemptions,
        conversion_redemptions=conversion_redemptions,
        non_conversion_redemptions=non_conversion_redemptions,
        net_activity=net_activity,
    )


def rollup_daily(rows, conversion_promo_codes: set[str]) -> list[DailyRollup]:
    buckets: dict[tuple[date | None, str], dict[str, Decimal]] = defaultdict(
        lambda: {"net_activations": Decimal("0.00"), "net_redemptions": Decimal("0.00"), "conversion_redemptions": Decimal("0.00")}
    )
    for row in rows:
        key = (row.business_date, row.source_file)
        if row.is_activation or row.is_void_activation:
            buckets[key]["net_activations"] += row.amount
        elif row.is_redemption or row.is_void_redemption:
            buckets[key]["net_redemptions"] += row.amount
            if row.promocode in conversion_promo_codes:
                buckets[key]["conversion_redemptions"] += row.amount

    result: list[DailyRollup] = []
    for (business_date, source_file), totals in sorted(buckets.items(), key=lambda item: (item[0][0] or date.min, item[0][1])):
        net_activations = money(totals["net_activations"])
        net_redemptions = money(totals["net_redemptions"])
        conversion_redemptions = money(totals["conversion_redemptions"])
        non_conversion_redemptions = money(net_redemptions - conversion_redemptions)
        result.append(
            DailyRollup(
                business_date=business_date,
                source_file=source_file,
                net_activations=net_activations,
                net_redemptions=net_redemptions,
                conversion_redemptions=conversion_redemptions,
                non_conversion_redemptions=non_conversion_redemptions,
                net_activity=money(net_activations + net_redemptions),
            )
        )
    return result


def audit_source_files(summary: SummaryData | None, activities: list[ActivityFileData]) -> list[SourceFileAudit]:
    files: list[Path] = []
    if summary and summary.source_file:
        files.append(summary.source_file)
    files.extend(activity.source_file for activity in activities)

    audits: list[SourceFileAudit] = []
    for path in files:
        try:
            stat = path.stat()
            audits.append(
                SourceFileAudit(
                    path=path,
                    file_type=path.suffix.lower().lstrip("."),
                    size_bytes=stat.st_size,
                    modified_at=file_modified_at(path),
                    sha256=sha256_file(path),
                )
            )
        except OSError:
            audits.append(SourceFileAudit(path=path, file_type=path.suffix.lower().lstrip("."), size_bytes=0, modified_at=None, sha256=""))
    return audits


def _combined_status(activity_variance: Decimal | None, pos_variance: Decimal | None) -> str:
    statuses = {variance_status(activity_variance), variance_status(pos_variance)}
    if "Review" in statuses:
        return "Review"
    if "Minor variance" in statuses:
        return "Minor variance"
    return "OK"


def _line_note(mode: str, metric: str, has_summary: bool) -> str:
    if mode == "weekly" and not has_summary:
        notes = {
            "issue": "No weekly summary supplied. POS variance compares POS Gift Card Issue against activity activations.",
            "payment": "No weekly summary supplied. Payment is displayed as a positive value to match POS control reporting.",
            "net": "No weekly summary supplied. Net impact = issues less payments.",
        }
        return notes[metric]
    if metric == "issue":
        return "Summary ties to weekly activity when Activity Variance is $0. POS variance compares POS Gift Card Issue against net activity activations."
    if metric == "payment":
        return "Displayed as positive payment dollars to match POS control reporting. Source gift card redemptions are signed negative in detail."
    return "Net impact = issues less payments. Negative means redemptions exceeded activations for the period."
