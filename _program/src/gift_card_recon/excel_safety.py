from __future__ import annotations

import unicodedata
from typing import Any


def safe_excel_cell_value(value: Any) -> Any:
    """Return a cell value safe for writing untrusted text to Excel.

    Ignore leading whitespace and control characters when identifying formula
    metacharacters. Prefixing unsafe text with an apostrophe keeps the displayed
    text intact while forcing Excel to store it as a literal string.
    """

    if not isinstance(value, str):
        return value
    for index, character in enumerate(value):
        if character.isspace() or unicodedata.category(character).startswith("C"):
            continue
        if character == "-" and not any(
            not (remaining.isspace() or unicodedata.category(remaining).startswith("C"))
            for remaining in value[index + 1 :]
        ):
            return value
        return f"'{value}" if character in "=+-@" else value
    return value
