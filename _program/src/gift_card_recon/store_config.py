from __future__ import annotations

from dataclasses import dataclass, replace
from decimal import Decimal
from pathlib import Path
from types import MappingProxyType
from typing import Mapping


REVIEW_VARIANCE_LIMIT = Decimal("5.00")


@dataclass(frozen=True)
class StoreConfig:
    """Store-specific facts used by monthly-close validation and reporting."""

    store: str
    location_name: str
    scheduled_closed_weekdays: frozenset[int]
    micros_default_path: Path
    micros_dropbox_relative_path: Path
    micros_source_label: str
    micros_issue_column_index: int
    micros_payment_column_index: int
    micros_system_totals_file: str = "DLYSYSTT.TXT"
    micros_tender_detail_file: str = "TENDER_DETAIL.TXT"
    gift_card_payment_tenders: frozenset[str] = frozenset(
        {"G C Payment", "Gift Card Payment"}
    )

    @property
    def report_heading(self) -> str:
        return f"{self.location_name.upper()} - STORE {self.store}"

    @property
    def output_slug(self) -> str:
        return f"{self.location_name.replace(' ', '_')}_{self.store}"

    @property
    def micros_issue_column_number(self) -> int:
        """Human-facing one-based column number for the issue control."""

        return self.micros_issue_column_index + 1

    @property
    def micros_payment_column_number(self) -> int:
        """Human-facing one-based column number for the payment control."""

        return self.micros_payment_column_index + 1


_STORE_CONFIGS = {
    "9354": StoreConfig(
        store="9354",
        location_name="Richmond",
        scheduled_closed_weekdays=frozenset({0}),
        micros_default_path=Path("..") / "micros_data" / "RC-Richmond-current",
        micros_dropbox_relative_path=Path("micros_data") / "RC-Richmond-current",
        micros_source_label="Richmond Micros 3700 POS export",
        micros_issue_column_index=120,
        micros_payment_column_index=102,
    ),
    "9355": StoreConfig(
        store="9355",
        location_name="Virginia Beach",
        scheduled_closed_weekdays=frozenset({0}),
        micros_default_path=Path("..") / "GETLinkedData-VB",
        micros_dropbox_relative_path=Path("GETLinkedData-VB"),
        micros_source_label="Virginia Beach Micros 3700 POS export",
        micros_issue_column_index=120,
        micros_payment_column_index=102,
    ),
}

STORE_CONFIGS: Mapping[str, StoreConfig] = MappingProxyType(_STORE_CONFIGS)


def get_store_config(store: str | int) -> StoreConfig:
    store_number = str(store).strip()
    try:
        return STORE_CONFIGS[store_number]
    except KeyError as exc:
        supported = ", ".join(sorted(STORE_CONFIGS))
        raise ValueError(
            f"Unsupported store {store_number!r}; expected one of: {supported}."
        ) from exc


def get_operations_store_config(
    store: str | int,
    *,
    operations_root: Path,
    dropbox_root: Path | None = None,
) -> StoreConfig:
    """Return a store config whose Micros source is anchored to a known root.

    ``micros_default_path`` remains operations-relative for compatibility with
    direct development calls. Operator entrypoints pass ``dropbox_root`` so a
    workspace can live below ``Dropbox/Automations`` while Micros exports stay
    at the Dropbox root.
    """

    config = get_store_config(store)
    if dropbox_root is None:
        micros_path = Path(operations_root) / config.micros_default_path
    else:
        micros_path = Path(dropbox_root) / config.micros_dropbox_relative_path
    return replace(config, micros_default_path=micros_path.resolve(strict=False))
