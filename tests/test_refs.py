from __future__ import annotations

import pytest
from xlsx_audit.refs import (
    column_to_index,
    format_coordinate,
    index_to_column,
    is_coordinate,
    parse_coordinate,
    quote_sheet_name,
    split_sheet_reference,
)


@pytest.mark.parametrize(
    ("letters", "index"),
    [("A", 1), ("B", 2), ("Z", 26), ("AA", 27), ("AZ", 52), ("XFD", 16384)],
)
def test_column_conversion_roundtrip(letters: str, index: int) -> None:
    assert column_to_index(letters) == index
    assert index_to_column(index) == letters


def test_column_conversion_rejects_junk() -> None:
    with pytest.raises(ValueError):
        column_to_index("")
    with pytest.raises(ValueError):
        column_to_index("A1")
    with pytest.raises(ValueError):
        column_to_index("XFE")  # one past the last Excel column
    with pytest.raises(ValueError):
        index_to_column(0)


@pytest.mark.parametrize("coordinate", ["A1", "C14", "XFD1048576", "$B$7"])
def test_parse_coordinate_accepts_cells(coordinate: str) -> None:
    row, column = parse_coordinate(coordinate)
    assert row >= 1 and column >= 1
    assert format_coordinate(row, column) == coordinate.replace("$", "")


@pytest.mark.parametrize("coordinate", ["", "A0", "A", "1", "AAAA1", "A1048577", "A1:B2"])
def test_parse_coordinate_rejects_junk(coordinate: str) -> None:
    with pytest.raises(ValueError):
        parse_coordinate(coordinate)
    assert not is_coordinate(coordinate)


def test_split_sheet_reference() -> None:
    assert split_sheet_reference("Sheet1!A1") == ("Sheet1", "A1")
    assert split_sheet_reference("'My Sheet'!A1") == ("My Sheet", "A1")
    assert split_sheet_reference("'It''s'!B2") == ("It's", "B2")
    assert split_sheet_reference("A1") == (None, "A1")
    assert split_sheet_reference("Sheet1!") == ("Sheet1", "")
    assert split_sheet_reference("") == (None, "")


def test_split_sheet_reference_keeps_bang_inside_quotes() -> None:
    assert split_sheet_reference("'A!B'!C3") == ("A!B", "C3")


def test_quote_sheet_name() -> None:
    assert quote_sheet_name("Sheet1") == "Sheet1"
    assert quote_sheet_name("My Sheet") == "'My Sheet'"
    assert quote_sheet_name("It's") == "'It''s'"
