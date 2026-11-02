from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
import xlsx_audit.ooxml as ooxml
from xlsx_audit.ooxml import WorkbookReadError, formula_references, read_workbook

from factory import build_workbook, simple_sheet, write_bytes, zip_parts


def test_reads_values_strings_and_formulas(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "book.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "A1": {"v": 10},
                    "A2": {"v": 2.5},
                    "B1": {"t": "s", "v": "Merhaba"},
                    "B2": {"t": "b", "v": "1"},
                    "B3": {"t": "inlineStr", "v": "inline"},
                    "C1": {"f": "A1*2", "v": 20},
                    "C2": {"f": "A1/0"},
                    "C3": {"t": "e", "v": "#REF!"},
                },
            )
        ],
    )
    workbook = read_workbook(path)
    assert workbook.sheet_names() == ["Sheet1"]
    cells = workbook.sheet_cells("Sheet1")
    assert cells["A1"].value == "10"
    assert cells["A1"].is_numeric
    assert cells["B1"].value == "Merhaba"
    assert cells["B2"].value == "TRUE"
    assert cells["B3"].value == "inline"
    assert cells["C1"].formula == "=A1*2"
    assert cells["C1"].value == "20"
    assert cells["C2"].formula == "=A1/0"
    assert cells["C2"].value is None
    assert cells["C3"].value_type == "e"
    assert len(workbook.cells) == 8


def test_shared_formulas_are_translated(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "shared.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "A1": {"f": "B1*2", "v": 2, "shared": "0", "ref": "A1:A3"},
                    "A2": {"shared": "0", "v": 4},
                    "A3": {"shared": "0", "v": 6},
                    "B1": {"v": 1},
                    "B2": {"v": 2},
                    "B3": {"v": 3},
                },
            )
        ],
    )
    workbook = read_workbook(path)
    cells = workbook.sheet_cells("Sheet1")
    assert cells["A1"].formula == "=B1*2"
    assert cells["A2"].formula == "=B2*2"
    assert cells["A3"].formula == "=B3*2"


def test_shared_formula_without_master_is_dropped(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "orphan.xlsx",
        [simple_sheet("Sheet1", {"A1": {"shared": "7", "v": 1}})],
    )
    workbook = read_workbook(path)
    assert workbook.sheet_cells("Sheet1")["A1"].formula is None


def test_array_formula_is_kept(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "array.xlsx",
        [simple_sheet("Sheet1", {"A1": {"f": "SUM(B1:B2*2)", "v": 6, "array": "A1"}})],
    )
    workbook = read_workbook(path)
    assert workbook.sheet_cells("Sheet1")["A1"].formula == "=SUM(B1:B2*2)"


def test_external_links_are_collected(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "links.xlsx",
        [simple_sheet("Sheet1", {"A1": {"f": "[1]Sheet1!A1", "v": 1}})],
        external_targets=[
            r"file:///C:/Users/example/Budget.xlsx",
            "https://example.invalid/data.csv",
        ],
    )
    workbook = read_workbook(path)
    assert len(workbook.external_links) == 2
    assert workbook.external_links[0].is_file_path
    assert workbook.external_links[1].is_url
    assert not workbook.external_links[1].is_file_path


def test_defined_names_and_metadata(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "names.xlsx",
        [simple_sheet("Sheet1", {"A1": {"v": 1}})],
        defined_names=[
            {"name": "RATE", "formula": "Sheet1!$A$1"},
            {"name": "LocalName", "formula": "Sheet1!$A$1", "local_sheet_id": 0, "hidden": True},
        ],
        metadata={"creator": "example author", "title": "fake title"},
    )
    workbook = read_workbook(path)
    assert [n.name for n in workbook.defined_names] == ["RATE", "LocalName"]
    assert workbook.defined_names[1].hidden is True
    assert workbook.defined_names[1].local_sheet_id == 0
    assert workbook.properties["creator"] == "example author"


def test_hidden_sheet_state(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "hidden.xlsx",
        [
            simple_sheet("Visible", {}),
            simple_sheet("Shy", {}, state="hidden"),
            simple_sheet("Invisible", {}, state="veryHidden"),
        ],
    )
    workbook = read_workbook(path)
    assert workbook.sheets[1].is_hidden
    assert not workbook.sheets[1].is_very_hidden
    assert workbook.sheets[2].is_very_hidden


def test_formula_references_helper() -> None:
    refs = formula_references("=A1+Sheet2!C3")
    assert ("Sheet2", 3, 3) in refs
    assert (None, 1, 1) in refs
    assert formula_references("=@@@") == []


def test_missing_file(tmp_path: Path) -> None:
    with pytest.raises(WorkbookReadError, match="file not found"):
        read_workbook(tmp_path / "nope.xlsx")


def test_directory_is_rejected(tmp_path: Path) -> None:
    with pytest.raises(WorkbookReadError, match="directory"):
        read_workbook(tmp_path)


def test_non_zip_is_rejected(tmp_path: Path) -> None:
    path = write_bytes(tmp_path / "junk.xlsx", b"this is not a zip file")
    with pytest.raises(WorkbookReadError, match="not a readable"):
        read_workbook(path)


def test_package_without_workbook_part(tmp_path: Path) -> None:
    path = zip_parts(tmp_path / "empty.xlsx", [("hello.txt", "hi")])
    with pytest.raises(WorkbookReadError, match=r"xl/workbook\.xml"):
        read_workbook(path)


def test_invalid_workbook_xml(tmp_path: Path) -> None:
    path = zip_parts(tmp_path / "bad.xlsx", [("xl/workbook.xml", "<workbook>")])
    with pytest.raises(WorkbookReadError, match="invalid XML"):
        read_workbook(path)


def test_workbook_without_sheets(tmp_path: Path) -> None:
    path = zip_parts(
        tmp_path / "nosheets.xlsx",
        [
            (
                "xl/workbook.xml",
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
            )
        ],
    )
    with pytest.raises(WorkbookReadError, match="no <sheets>"):
        read_workbook(path)


def test_unknown_relationship(tmp_path: Path) -> None:
    workbook_xml = (
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="S" sheetId="1" r:id="rId404"/></sheets></workbook>'
    )
    path = zip_parts(
        tmp_path / "brokenrel.xlsx",
        [
            ("xl/workbook.xml", workbook_xml),
            ("xl/_rels/workbook.xml.rels", "<Relationships/>"),
        ],
    )
    with pytest.raises(WorkbookReadError, match="unknown relationship"):
        read_workbook(path)


def test_invalid_cell_coordinate(tmp_path: Path) -> None:
    sheet = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData><row r="1"><c r="ZZZZ1"><v>1</v></c></row></sheetData></worksheet>'
    )
    path = build_workbook(
        tmp_path / "badcoord.xlsx",
        [simple_sheet("Sheet1", {})],
        extra_parts={"xl/worksheets/sheet1.xml": sheet},
    )
    with pytest.raises(WorkbookReadError, match="invalid cell coordinate"):
        read_workbook(path)


def test_missing_shared_string(tmp_path: Path) -> None:
    sheet = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData><row r="1"><c r="A1" t="s"><v>9</v></c></row></sheetData></worksheet>'
    )
    path = build_workbook(
        tmp_path / "badstring.xlsx",
        [simple_sheet("Sheet1", {"A1": {"v": 1}})],
        extra_parts={"xl/worksheets/sheet1.xml": sheet},
    )
    with pytest.raises(WorkbookReadError, match="missing shared string"):
        read_workbook(path)


def test_too_many_entries(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(ooxml, "MAX_ENTRIES", 1)
    path = zip_parts(
        tmp_path / "many.xlsx",
        [("xl/workbook.xml", "<workbook/>"), ("a", "b")],
    )
    with pytest.raises(WorkbookReadError, match="too many parts"):
        read_workbook(path)


def test_part_too_large(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(ooxml, "MAX_PART_BYTES", 10)
    path = zip_parts(
        tmp_path / "big.xlsx",
        [("xl/workbook.xml", "<workbook>" + "x" * 100 + "</workbook>")],
    )
    with pytest.raises(WorkbookReadError, match="part too large"):
        read_workbook(path)


def test_total_size_limit(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(ooxml, "MAX_TOTAL_BYTES", 10)
    path = zip_parts(
        tmp_path / "total.xlsx",
        [("a.txt", "0123456789"), ("b.txt", "0123456789")],
    )
    with pytest.raises(WorkbookReadError, match="size limit"):
        read_workbook(path)


def test_sheet_without_sheet_data_has_no_cells(tmp_path: Path) -> None:
    sheet = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'
    path = build_workbook(
        tmp_path / "nodata.xlsx",
        [simple_sheet("Sheet1", {})],
        extra_parts={"xl/worksheets/sheet1.xml": sheet},
    )
    assert read_workbook(path).sheets[0].cells == ()


def test_zip_can_be_read_twice(tmp_path: Path) -> None:
    path = build_workbook(tmp_path / "twice.xlsx", [simple_sheet("S", {"A1": {"v": 1}})])
    assert len(read_workbook(path).cells) == 1
    assert len(read_workbook(path).cells) == 1
    with zipfile.ZipFile(path) as archive:  # the reader must close its handle
        assert archive.namelist()
