from __future__ import annotations

import re
from datetime import date
from decimal import Decimal
from pathlib import Path

from gift_card_recon.models import DardenCreditMemo, MonthlyCloseCertification, SummaryData
from gift_card_recon.parsers import ParseError
from gift_card_recon.utils import money, parse_date, to_decimal


def parse_darden_credit_memo(path: Path) -> DardenCreditMemo:
    path = Path(path)
    if not path.exists():
        raise ParseError(f"Darden credit memo not found: {path}")
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise ParseError("pypdf is required to read the Darden credit memo. Run: pip install -r requirements.txt") from exc

    try:
        reader = PdfReader(path)
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as exc:
        raise ParseError(f"Could not read Darden credit memo {path.name}: {exc}") from exc
    return parse_darden_credit_memo_text(text, source_file=path)


def parse_darden_credit_memo_text(text: str, *, source_file: Path) -> DardenCreditMemo:
    normalized = re.sub(r"[\u00a0\t ]+", " ", str(text or ""))
    normalized = re.sub(r"\r\n?", "\n", normalized)

    store_match = re.search(r"\bLocation\s+(\d{4,})\b", normalized, flags=re.IGNORECASE)
    if store_match is None:
        raise ParseError(f"Could not find the restaurant location in Darden credit memo {Path(source_file).name}.")

    range_match = re.search(r"\b(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})\b", normalized)
    if range_match is None:
        raise ParseError(f"Could not find the service date range in Darden credit memo {Path(source_file).name}.")
    period_start = parse_date(range_match.group(1))
    period_end = parse_date(range_match.group(2))
    if period_start is None or period_end is None:
        raise ParseError(f"Could not parse the service date range in Darden credit memo {Path(source_file).name}.")

    total_match = re.search(
        r"\bTOTAL\s+\$?\s*(\(?\s*-?\s*[\d,]+\.\d{2}\s*\)?)\s*\$?",
        normalized,
        flags=re.IGNORECASE,
    )
    if total_match is None:
        raise ParseError(f"Could not find the Total amount in Darden credit memo {Path(source_file).name}.")
    amount_text = re.sub(r"\s+", "", total_match.group(1))
    total = to_decimal(amount_text)
    if total is None:
        raise ParseError(f"Could not parse the Total amount in Darden credit memo {Path(source_file).name}.")

    invoice_number_match = re.search(
        r"INVOICE\s+NUMBER\s+(.+?)\s+INVOICE\s+DATE",
        normalized,
        flags=re.IGNORECASE | re.DOTALL,
    )
    invoice_date_match = re.search(r"INVOICE\s+DATE\s+(\d{1,2}/\d{1,2}/\d{4})", normalized, flags=re.IGNORECASE)

    return DardenCreditMemo(
        source_file=Path(source_file),
        store=store_match.group(1),
        period_start=period_start,
        period_end=period_end,
        total=money(total),
        invoice_number=re.sub(r"\s+", " ", invoice_number_match.group(1)).strip() if invoice_number_match else "",
        invoice_date=parse_date(invoice_date_match.group(1)) if invoice_date_match else None,
    )


def build_monthly_close_certification(
    *,
    store: str,
    period: str,
    period_start: date,
    period_end: date,
    summary: SummaryData,
    darden_credit_memo: DardenCreditMemo,
) -> MonthlyCloseCertification:
    if darden_credit_memo.store != str(store):
        raise ParseError(
            f"Darden credit memo is for location {darden_credit_memo.store}; expected store {store}."
        )
    if darden_credit_memo.period_start != period_start or darden_credit_memo.period_end != period_end:
        raise ParseError(
            "Darden credit memo service period "
            f"{darden_credit_memo.period_start:%m/%d/%Y}-{darden_credit_memo.period_end:%m/%d/%Y} "
            f"does not match {period} ({period_start:%m/%d/%Y}-{period_end:%m/%d/%Y})."
        )
    if summary.net_settlement is None:
        raise ParseError("Gift Card Summary does not contain a Net Settlement value for the Darden close control.")

    summary_amount = money(summary.net_settlement)
    darden_amount = money(darden_credit_memo.total)
    return MonthlyCloseCertification(
        store=str(store),
        period=str(period),
        period_start=period_start,
        period_end=period_end,
        summary_net_settlement=summary_amount,
        darden_credit_memo=darden_credit_memo,
        variance=money(darden_amount - summary_amount),
    )
