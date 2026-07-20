from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import Any, Iterable

from gift_card_recon.excel_io import first_sheet_values, workbook_values
from gift_card_recon.models import ActivityFileData, ActivityRow, PosControls, SummaryData
from gift_card_recon.utils import clean_code, money, normalize_header, normalize_text, parse_date, to_decimal


class ParseError(RuntimeError):
    pass


def parse_summary(path: Path, store: str) -> SummaryData:
    path = Path(path)
    sheets = workbook_values(path)
    if not sheets:
        raise ParseError(f"Summary workbook has no sheets: {path}")
    matrix = next(iter(sheets.values()))

    summary_section = _extract_section_row(matrix, "SUMMARY", store)
    non_conversion_section = _extract_section_row(matrix, "Non-Conversion Promo Card Calculation", store, required=False)
    pre_section = _extract_section_row(matrix, "Cards ACTIVATED PRE-Conversion", store, required=False)
    post_section = _extract_section_row(matrix, "Cards ACTIVATED POST-Conversion", store, required=False)
    in_restaurant_section = _extract_section_row(matrix, "In-Restaurant Activations and Cash Collections", store, required=False)

    promo_codes = _extract_conversion_promo_codes(matrix)

    return SummaryData(
        store=str(store),
        total_activations=_money_from_header(summary_section, "Total Activations"),
        total_redemptions=_money_from_header(summary_section, "Total Redemptions"),
        payable_redemptions=_optional_money_from_header(summary_section, "Payable Redemptions"),
        gcdr=_optional_money_from_header(summary_section, "GCDR", prefer_exact=True),
        net_settlement=_money_from_header(summary_section, "Net Settlement"),
        gift_card_franchise_fee_rate=_optional_decimal_from_header(summary_section, "Gift Card Franchise Fee Rate"),
        non_conversion_redemptions=_optional_money_from_header(non_conversion_section, "Total Redemptions") if non_conversion_section else None,
        non_conversion_gcdr_redemptions=_optional_money_from_header(non_conversion_section, "GCDR Redemptions") if non_conversion_section else None,
        non_conversion_gcdr=_optional_money_from_header(non_conversion_section, "GCDR", prefer_exact=True) if non_conversion_section else None,
        pre_conversion_fee_redemptions=_optional_money_from_header(pre_section, "Redemptions subject to 10% Fee") if pre_section else None,
        pre_conversion_gcdr=_optional_money_from_header(pre_section, "GCDR Total") if pre_section else None,
        post_conversion_fee_redemptions=_optional_money_from_header(post_section, "Redemptions subject to 10% Fee") if post_section else None,
        post_conversion_gcdr=_optional_money_from_header(post_section, "GCDR") if post_section else None,
        cross_location_redeemed=_optional_money_from_header(in_restaurant_section, "Activated in restaurant pre-conversion and Redeemed at another Ruth's Chris in current month") if in_restaurant_section else None,
        conversion_promo_codes=promo_codes,
        source_file=path,
    )


def summary_contains_store(path: Path, store: str) -> bool:
    """Check Summary identity without requiring every financial value to parse."""

    path = Path(path)
    try:
        sheets = workbook_values(path)
    except Exception as exc:
        raise ParseError(f"Could not read Gift Card Summary workbook {path.name}: {exc}") from exc
    if not sheets:
        raise ParseError(f"Summary workbook has no sheets: {path}")
    matrix = next(iter(sheets.values()))
    section = _extract_section_row(
        matrix,
        "SUMMARY",
        store,
        allow_missing_store=True,
    )
    return section is not None


def parse_activity_file(path: Path, conversion_promo_codes: set[str] | None = None) -> ActivityFileData:
    path = Path(path)
    _sheet_name, matrix = first_sheet_values(path)
    report_begin, report_end = _extract_activity_report_dates(matrix)
    store = _extract_activity_store(matrix, path)
    header_row_idx, header_map = _find_activity_header(matrix)

    rows: list[ActivityRow] = []
    for raw_row in matrix[header_row_idx + 1 :]:
        if not _row_has_value(raw_row) or _looks_like_total_or_footer(raw_row):
            continue

        card_no = _value_by_canonical(raw_row, header_map, "card_no")
        request_code_listing = _value_by_canonical(raw_row, header_map, "request_code_listing")
        amount_value = _value_by_canonical(raw_row, header_map, "amount")
        business_date_value = _value_by_canonical(raw_row, header_map, "business_date")

        required_values = {
            "Card No": card_no,
            "Request Code Listing": request_code_listing,
            "Amount": amount_value,
        }
        missing = [label for label, value in required_values.items() if value in (None, "")]
        if missing:
            transaction_signals = (
                card_no,
                request_code_listing,
                business_date_value,
                amount_value,
                _value_by_canonical(raw_row, header_map, "request"),
                _value_by_canonical(raw_row, header_map, "transaction_no"),
            )
            if any(value not in (None, "") for value in transaction_signals):
                raise ParseError(
                    f"Activity row in {path.name} is missing required field(s): "
                    f"{', '.join(missing)}."
                )
            continue

        business_date = parse_date(business_date_value)
        if business_date is None:
            raise ParseError(f"Could not parse business date {business_date_value!r} in {path.name}.")
        amount = _required_money_value(amount_value, field="Amount", path=path)

        rows.append(
            ActivityRow(
                source_file=path.name,
                card_no=str(card_no).strip(),
                request=_value_by_canonical(raw_row, header_map, "request"),
                request_code_listing=str(request_code_listing).strip(),
                business_date=business_date,
                transaction_no=_value_by_canonical(raw_row, header_map, "transaction_no"),
                amount=amount,
                promocode=clean_code(_value_by_canonical(raw_row, header_map, "promocode")),
                authorization_code=_value_by_canonical(raw_row, header_map, "authorization_code"),
            )
        )

    if not rows:
        raise ParseError(f"No activity rows parsed from {path.name}. Check report format and headers.")
    return ActivityFileData(
        source_file=path,
        report_begin=report_begin,
        report_end=report_end,
        rows=rows,
        store=store,
    )


def parse_pos_controls(path: Path, store: str, period: str) -> PosControls:
    path = Path(path)
    if path.suffix.lower() == ".csv":
        rows = _read_csv_dicts(path)
    elif path.suffix.lower() in {".xlsx", ".xls"}:
        rows = _read_excel_dicts(path)
    else:
        raise ParseError(f"Unsupported POS controls file type: {path.name}. Use .csv or .xlsx.")

    _validate_pos_headers(rows, path)

    matches = [
        row
        for row in rows
        if str(row.get("store", "")).strip() == str(store)
        and str(row.get("period", "")).strip().lower() in {str(period).lower(), "auto"}
    ]
    if len(matches) != 1:
        raise ParseError(
            f"Expected exactly one POS controls row for store={store}, period={period} (or auto) in {path.name}. Found {len(matches)}."
        )
    selected = matches[0]
    _validate_pos_identity(selected, path, store=store, period=period)
    return PosControls(
        store=str(store),
        period=str(period),
        pos_gift_card_issue=_required_pos_money(selected, "pos_gift_card_issue", path),
        pos_gift_card_payment=_required_pos_money(selected, "pos_gift_card_payment", path),
    )


def pos_controls_from_args(store: str, period: str, issue: Any, payment: Any) -> PosControls:
    return PosControls(
        store=str(store),
        period=str(period),
        pos_gift_card_issue=_required_direct_pos_money(issue, "pos_gift_card_issue"),
        pos_gift_card_payment=_required_direct_pos_money(payment, "pos_gift_card_payment"),
    )


def discover_input_files(input_dir: Path, mode: str = "monthly") -> tuple[Path | None, list[Path], Path | None]:
    input_dir = Path(input_dir)
    mode = mode.lower()
    summary_candidates = sorted(input_dir.glob("summary/*Gift Card Summary*.xlsx")) + sorted(input_dir.glob("*Gift Card Summary*.xlsx"))
    activity_candidates = sorted(input_dir.glob("activity/*Gift Card Activity*.xls")) + sorted(input_dir.glob("*Gift Card Activity*.xls"))
    activity_candidates += sorted(input_dir.glob("activity/*Gift Card Activity*.xlsx")) + sorted(input_dir.glob("*Gift Card Activity*.xlsx"))
    pos_candidates = sorted(input_dir.glob("pos_controls.csv")) + sorted(input_dir.glob("pos_controls.xlsx"))

    summary_candidates = _dedupe_paths(summary_candidates)
    activity_candidates = _dedupe_paths(activity_candidates)
    pos_candidates = _dedupe_paths(pos_candidates)

    if mode not in {"monthly", "weekly"}:
        raise ParseError(f"Unsupported mode {mode!r}. Use monthly or weekly.")
    if mode == "monthly" and len(summary_candidates) != 1:
        raise ParseError(f"Expected exactly one Gift Card Summary .xlsx file in {input_dir} or summary/. Found {len(summary_candidates)}.")
    if mode == "weekly" and len(summary_candidates) > 1:
        raise ParseError(f"Expected at most one optional weekly Gift Card Summary .xlsx file in {input_dir} or summary/. Found {len(summary_candidates)}.")
    if not activity_candidates:
        raise ParseError(f"No Gift Card Activity .xls/.xlsx files found in {input_dir} or activity/.")
    return (summary_candidates[0] if summary_candidates else None), activity_candidates, (pos_candidates[0] if pos_candidates else None)


def _dedupe_paths(paths: list[Path]) -> list[Path]:
    seen = set()
    out = []
    for p in paths:
        r = p.resolve()
        if r not in seen:
            seen.add(r)
            out.append(p)
    return out


def _read_csv_dicts(path: Path) -> list[dict[str, Any]]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        return [{_normalize_pos_header(k): v for k, v in row.items()} for row in reader]


def _validate_pos_headers(rows: list[dict[str, Any]], path: Path) -> None:
    if not rows:
        raise ParseError(f"POS controls file has no data rows: {path.name}")
    required = {"store", "period", "pos_gift_card_issue", "pos_gift_card_payment"}
    missing = sorted(required - set(rows[0].keys()))
    if missing:
        raise ParseError(f"POS controls file {path.name} is missing required field(s): {', '.join(missing)}")


def _validate_pos_identity(row: dict[str, Any], path: Path, *, store: str, period: str) -> None:
    for field in ["store", "period"]:
        if row.get(field) in (None, ""):
            raise ParseError(f"POS controls file {path.name} has a blank required field: {field}")
    if str(row.get("store", "")).strip() != str(store):
        raise ParseError(f"POS controls file {path.name} is for store {row.get('store')!r}; expected {store}.")
    row_period = str(row.get("period", "")).strip().lower()
    if row_period not in {str(period).lower(), "auto"}:
        raise ParseError(f"POS controls file {path.name} is for period {row.get('period')!r}; expected {period} or auto.")


def _required_pos_money(row: dict[str, Any], field: str, path: Path):
    value = row.get(field)
    parsed = to_decimal(value)
    if parsed is None:
        raise ParseError(f"POS controls file {path.name} has a missing or malformed value for {field}: {value!r}")
    return money(parsed)


def _required_direct_pos_money(value: Any, field: str):
    parsed = to_decimal(value)
    if parsed is None:
        raise ParseError(f"Missing or malformed direct POS value for {field}: {value!r}")
    return money(parsed)


def _read_excel_dicts(path: Path) -> list[dict[str, Any]]:
    _sheet, matrix = first_sheet_values(path)
    header_idx = None
    for idx, row in enumerate(matrix):
        headers = [_normalize_pos_header(v) for v in row]
        if {"store", "period", "pos_gift_card_issue", "pos_gift_card_payment"}.issubset(set(headers)):
            header_idx = idx
            break
    if header_idx is None:
        raise ParseError(f"Could not find POS controls header in {path.name}")
    headers = [_normalize_pos_header(v) for v in matrix[header_idx]]
    result = []
    for row in matrix[header_idx + 1 :]:
        if _row_has_value(row):
            result.append({headers[i]: row[i] if i < len(row) else None for i in range(len(headers)) if headers[i]})
    return result


def _normalize_pos_header(value: Any) -> str:
    h = normalize_header(value)
    aliases = {
        "gift card issue": "pos_gift_card_issue",
        "pos gift card issue": "pos_gift_card_issue",
        "gift card payment": "pos_gift_card_payment",
        "pos gift card payment": "pos_gift_card_payment",
        "store number": "store",
        "restaurant": "store",
        "rest number": "store",
        "period ending": "period",
    }
    return aliases.get(h, h.replace(" ", "_"))


def _extract_activity_report_dates(matrix: list[list[Any]]):
    pattern = re.compile(r"BEGIN DATE:\s*'([^']+)'\s*,\s*END DATE:\s*'([^']+)'", flags=re.IGNORECASE)
    for row in matrix[:15]:
        for value in row:
            match = pattern.search(str(value or ""))
            if match:
                return parse_date(match.group(1)), parse_date(match.group(2))
    return None, None


def _extract_activity_store(matrix: list[list[Any]], path: Path) -> str:
    patterns = (
        re.compile(r"Rest\s+Number\s+Parameter\s+1\s*:\s*'?([0-9]+)'?", re.IGNORECASE),
        re.compile(r"Rest\s+Number\s*:\s*'?([0-9]+)'?", re.IGNORECASE),
    )
    stores: set[str] = set()
    for row in matrix[:15]:
        for value in row:
            text = str(value or "")
            for pattern in patterns:
                stores.update(match.group(1) for match in pattern.finditer(text))
    if len(stores) != 1:
        found = ", ".join(sorted(stores)) if stores else "none"
        raise ParseError(
            f"Expected exactly one activity store identifier in {path.name}; found {found}."
        )
    return next(iter(stores))


def _find_activity_header(matrix: list[list[Any]]) -> tuple[int, dict[str, int]]:
    aliases = {
        "card no": "card_no",
        "card number": "card_no",
        "request": "request",
        "request code listing": "request_code_listing",
        "business date": "business_date",
        "transaction no": "transaction_no",
        "transaction number": "transaction_no",
        "amount sum": "amount",
        "amount": "amount",
        "promocode": "promocode",
        "promo code": "promocode",
        "authorization code": "authorization_code",
        "auth code": "authorization_code",
    }
    required = {"card_no", "request_code_listing", "business_date", "amount"}
    for idx, row in enumerate(matrix[:50]):
        header_map: dict[str, int] = {}
        for col_idx, value in enumerate(row):
            canonical = aliases.get(normalize_header(value))
            if canonical and canonical not in header_map:
                header_map[canonical] = col_idx
        if required.issubset(header_map.keys()):
            return idx, header_map
    raise ParseError("Could not locate activity header row. Expected Card No, Request Code Listing, Business Date, and Amount SUM.")


def _value_by_canonical(row: list[Any], header_map: dict[str, int], canonical: str) -> Any:
    idx = header_map.get(canonical)
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def _row_has_value(row: Iterable[Any]) -> bool:
    return any(value is not None and value != "" for value in row)


def _looks_like_total_or_footer(row: list[Any]) -> bool:
    text = " ".join(str(v) for v in row if v is not None).lower()
    return "grand total" in text or "page " in text or text.strip().startswith("sum:")


def _extract_section_row(
    matrix: list[list[Any]],
    title: str,
    store: str,
    required: bool = True,
    allow_missing_store: bool = False,
) -> dict[str, Any] | None:
    title_norm = normalize_text(title)
    title_idx = None
    for idx, row in enumerate(matrix):
        if any(title_norm in normalize_text(value) for value in row):
            title_idx = idx
            break
    if title_idx is None:
        if required:
            raise ParseError(f"Could not find section {title!r} in summary workbook.")
        return None
    header_idx = title_idx + 1
    if header_idx >= len(matrix):
        if required:
            raise ParseError(f"Could not find header row for section {title!r}.")
        return None
    headers = [str(v).strip() if v is not None else "" for v in matrix[header_idx]]
    store_str = str(store).strip()
    store_columns = [i for i, header in enumerate(headers) if normalize_header(header) in {"store", "store number", "restaurant", "rest number"}]
    if len(store_columns) != 1:
        if required:
            raise ParseError(f"Section {title!r} does not contain exactly one Store Number column.")
        return None
    store_col = store_columns[0]
    section_rows: list[list[Any]] = []
    for row in matrix[header_idx + 1 :]:
        if not _row_has_value(row):
            if section_rows:
                break
            continue
        section_rows.append(row)
    matches = [
        row
        for row in section_rows
        if store_col < len(row) and str(row[store_col]).strip() == store_str
    ]
    if len(matches) == 1:
        row = matches[0]
        return {headers[i]: row[i] if i < len(row) else None for i in range(len(headers)) if headers[i]}
    if len(matches) > 1:
        raise ParseError(f"Section {title!r} contains multiple rows for store {store_str}.")
    if allow_missing_store:
        return None
    if required:
        raise ParseError(f"Section {title!r} does not contain a row for store {store_str}.")
    return None


def _find_header_value(section: dict[str, Any] | None, label: str, prefer_exact: bool = False) -> Any:
    if not section:
        return None
    wanted = normalize_header(label)
    for header, value in section.items():
        if normalize_header(header) == wanted:
            return value
    if prefer_exact:
        return None
    for header, value in section.items():
        h = normalize_header(header)
        if wanted in h or h in wanted:
            return value
    return None


def _money_from_header(section: dict[str, Any], label: str):
    value = _find_header_value(section, label, prefer_exact=True)
    if value is None:
        raise ParseError(f"Could not find summary value for header {label!r}.")
    return _required_money_value(value, field=label)


def _optional_money_from_header(section: dict[str, Any] | None, label: str, prefer_exact: bool = False):
    value = _find_header_value(section, label, prefer_exact=prefer_exact)
    if value in (None, ""):
        return None
    return _required_money_value(value, field=label)


def _optional_decimal_from_header(section: dict[str, Any] | None, label: str):
    value = _find_header_value(section, label)
    if value in (None, ""):
        return None
    parsed = to_decimal(value)
    if parsed is None:
        raise ParseError(f"Could not parse required decimal value for {label!r}: {value!r}.")
    return parsed


def _required_money_value(value: Any, *, field: str, path: Path | None = None):
    parsed = to_decimal(value)
    if parsed is None:
        source = f" in {path.name}" if path is not None else ""
        raise ParseError(f"Could not parse required money value for {field!r}{source}: {value!r}.")
    return money(parsed)


def _extract_conversion_promo_codes(matrix: list[list[Any]]) -> set[str]:
    header_idx = None
    promo_col = None
    for r, row in enumerate(matrix):
        for c, value in enumerate(row):
            if "rc conversion promo codes" in normalize_text(value):
                header_idx = r
                promo_col = c
                break
        if header_idx is not None:
            break
    codes: set[str] = set()
    if header_idx is None or promo_col is None:
        return codes
    for row in matrix[header_idx + 1 :]:
        value = row[promo_col] if promo_col < len(row) else None
        code = clean_code(value)
        if code and code.isdigit():
            codes.add(code)
        elif codes:
            break
    return codes
