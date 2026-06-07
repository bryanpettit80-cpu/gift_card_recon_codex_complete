from __future__ import annotations

import argparse
from pathlib import Path

from gift_card_recon.excel_writer import write_reconciliation_workbook
from gift_card_recon.parsers import ParseError, discover_input_files, parse_activity_file, parse_pos_controls, parse_summary, pos_controls_from_args
from gift_card_recon.reconcile import build_reconciliation
from gift_card_recon.utils import parse_date


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir)
    period_end = parse_date(args.period_end) if args.period_end else None

    try:
        summary_path, activity_paths, discovered_pos_path = discover_input_files(input_dir, mode=args.mode)
        summary = parse_summary(summary_path, store=args.store) if summary_path else None
        conversion_promo_codes = summary.conversion_promo_codes if summary else set()
        activities = [parse_activity_file(path, conversion_promo_codes) for path in activity_paths]

        if args.pos_controls:
            pos_controls = parse_pos_controls(Path(args.pos_controls), store=args.store, period=args.period)
        elif discovered_pos_path:
            pos_controls = parse_pos_controls(discovered_pos_path, store=args.store, period=args.period)
        elif args.pos_gift_card_issue is not None and args.pos_gift_card_payment is not None:
            pos_controls = pos_controls_from_args(args.store, args.period, args.pos_gift_card_issue, args.pos_gift_card_payment)
        else:
            raise SystemExit("POS controls missing. Provide --pos-controls or both --pos-gift-card-issue and --pos-gift-card-payment.")

        result = build_reconciliation(
            store=args.store,
            period=args.period,
            period_end=period_end,
            summary=summary,
            activities=activities,
            pos_controls=pos_controls,
            mode=args.mode,
        )
    except ParseError as exc:
        raise SystemExit(str(exc)) from exc

    output_path = Path(args.output_file) if args.output_file else output_dir / f"Gift_Card_Reconciliation_{args.store}_{args.period}.xlsx"
    write_reconciliation_workbook(result, output_path)

    print(f"Created: {output_path}")
    print("Primary tie-out:")
    for line in result.lines:
        activity_variance = "N/A" if line.activity_variance is None else f"{line.activity_variance:+,.2f}"
        pos_variance = "N/A" if line.pos_variance is None else f"{line.pos_variance:+,.2f}"
        print(f"  - {line.metric}: activity variance={activity_variance} | POS variance={pos_variance} | status={line.status}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="gift-card-recon", description="Reconcile gift card summary, weekly activity files, and POS controls.")
    parser.add_argument("--mode", choices=["monthly", "weekly"], default="monthly", help="Reconciliation mode. Defaults to monthly.")
    parser.add_argument("--store", required=True, help="Store number, e.g. 9354")
    parser.add_argument("--period", required=True, help="Accounting/reporting period, e.g. 2026-05")
    parser.add_argument("--period-end", default=None, help="Optional period end date, e.g. 2026-05-31")
    parser.add_argument("--input-dir", required=True, help="Input folder containing summary/, activity/, and optional pos_controls.csv")
    parser.add_argument("--output-dir", default="output", help="Folder for generated reconciliation workbook")
    parser.add_argument("--output-file", default=None, help="Optional explicit output .xlsx path")
    parser.add_argument("--pos-controls", default=None, help="Optional POS controls .csv/.xlsx path")
    parser.add_argument("--pos-gift-card-issue", default=None, help="POS Gift Card Issue control total")
    parser.add_argument("--pos-gift-card-payment", default=None, help="POS Gift Card Payment control total")
    return parser
