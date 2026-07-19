from __future__ import annotations

from typing import Any

_FORMULA_PREFIXES = ("=", "+", "@", "\t", "\r")


def safe_excel_cell_value(value: Any) -> Any:
    """Return a cell value safe for writing untrusted text to Excel.

    Excel and openpyxl treat strings beginning with formula metacharacters as
    formulas. Prefixing such text with an apostrophe keeps the displayed text
    intact while forcing Excel to store it as a literal string.
    """

    if isinstance(value, str) and (value.startswith(_FORMULA_PREFIXES) or (len(value) > 1 and value.startswith("-"))):
        return f"'{value}"
    return value
