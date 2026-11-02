"""A1-notation helpers.

The auditor never asks Excel for anything, so every address comes from these
pure-python conversions.
"""

from __future__ import annotations

import re

MAX_ROW = 1_048_576
MAX_COLUMN = 16_384

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$")


def column_to_index(letters: str) -> int:
    """``"A"`` -> 1, ``"AA"`` -> 27. Raises ``ValueError`` for junk."""
    if not letters or not letters.isalpha():
        raise ValueError(f"invalid column letters: {letters!r}")
    value = 0
    for char in letters.upper():
        value = value * 26 + (ord(char) - 64)
    if not 1 <= value <= MAX_COLUMN:
        raise ValueError(f"column out of range: {letters!r}")
    return value


def index_to_column(index: int) -> str:
    """1 -> ``"A"``, 27 -> ``"AA"``."""
    if index < 1:
        raise ValueError(f"column index must be >= 1, got {index}")
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def format_coordinate(row: int, column: int) -> str:
    return f"{index_to_column(column)}{row}"


def parse_coordinate(coordinate: str) -> tuple[int, int]:
    """``"C14"`` -> ``(14, 3)`` (row, column), 1-based."""
    match = _CELL_RE.match(coordinate.strip())
    if not match:
        raise ValueError(f"invalid cell coordinate: {coordinate!r}")
    row = int(match.group(2))
    column = column_to_index(match.group(1))
    if not 1 <= row <= MAX_ROW:
        raise ValueError(f"row out of range: {coordinate!r}")
    return row, column


def is_coordinate(text: str) -> bool:
    try:
        parse_coordinate(text)
    except ValueError:
        return False
    return True


def split_sheet_reference(text: str) -> tuple[str | None, str]:
    """Split ``"Sheet1!A1"`` / ``"'My Sheet'!A1"`` into its two halves.

    A quoted sheet name may contain ``!`` so the last unquoted ``!`` wins.
    """
    stripped = text.strip()
    if not stripped:
        return None, ""
    end = stripped.rfind("!")
    if end == -1:
        return None, stripped
    sheet = stripped[:end]
    rest = stripped[end + 1 :]
    if len(sheet) >= 2 and sheet[0] == "'" and sheet[-1] == "'":
        sheet = sheet[1:-1].replace("''", "'")
    if not rest:
        return sheet or None, ""
    return sheet or None, rest


def quote_sheet_name(name: str) -> str:
    """Quote a sheet name for use inside a formula when Excel would."""
    if re.match(r"^[A-Za-z_\\][A-Za-z0-9_.]*$", name):
        return name
    return "'" + name.replace("'", "''") + "'"
