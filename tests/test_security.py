"""Untrusted input: entity expansion, path traversal, hostile nesting."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from xlsx_audit import audit_workbook
from xlsx_audit.formula_parser import FormulaSyntaxError, parse_formula
from xlsx_audit.ooxml import WorkbookReadError, read_workbook

from factory import build_workbook, simple_sheet, zip_parts


def test_external_entity_is_rejected(tmp_path: Path) -> None:
    payload = (
        '<?xml version="1.0"?>'
        '<!DOCTYPE workbook [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheets><sheet name="&xxe;" sheetId="1"/></sheets></workbook>'
    )
    path = zip_parts(tmp_path / "xxe.xlsx", [("xl/workbook.xml", payload)])
    with pytest.raises(WorkbookReadError, match="invalid XML"):
        read_workbook(path)


def test_entity_expansion_bomb_is_rejected(tmp_path: Path) -> None:
    payload = (
        '<?xml version="1.0"?>'
        "<!DOCTYPE lolz ["
        '<!ENTITY lol "lol">'
        '<!ENTITY lol1 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">'
        '<!ENTITY lol2 "&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;">'
        "]>"
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheets><sheet name="&lol2;" sheetId="1"/></sheets></workbook>'
    )
    path = zip_parts(tmp_path / "bomb.xlsx", [("xl/workbook.xml", payload)])
    with pytest.raises(WorkbookReadError, match="invalid XML"):
        read_workbook(path)


def test_zip_entries_are_never_extracted(tmp_path: Path) -> None:
    evil = tmp_path / "evil.xml"
    path = build_workbook(
        tmp_path / "slip.xlsx",
        [simple_sheet("S", {"A1": {"v": 1}})],
        extra_parts={"../../../evil.xml": "<evil/>"},
    )
    workbook = read_workbook(path)
    assert workbook.sheet_names() == ["S"]
    assert not evil.exists()
    assert not (tmp_path.parent / "evil.xml").exists()


def test_relationship_target_cannot_escape_the_package(tmp_path: Path) -> None:
    path = zip_parts(
        tmp_path / "escape.xlsx",
        [
            (
                "xl/workbook.xml",
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
            ),
            (
                "xl/_rels/workbook.xml.rels",
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                'Target="../../../../etc/passwd"/></Relationships>',
            ),
        ],
    )
    workbook = read_workbook(path)
    # The part resolves to a name that is not in the archive: no cells, no crash,
    # and nothing outside the package is opened.
    assert workbook.sheets[0].cells == ()


def test_hostile_nesting_is_a_parse_error_not_a_crash(tmp_path: Path) -> None:
    formula = "=" + "(" * 5000 + "1" + ")" * 5000
    with pytest.raises(FormulaSyntaxError, match="nested too deeply"):
        parse_formula(formula)


def test_long_unary_chain_does_not_recurse() -> None:
    tree = parse_formula("=" + "-" * 500 + "A1")
    assert tree is not None


def test_hostile_formula_in_a_workbook_is_reported(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "hostile.xlsx",
        [simple_sheet("S", {"A1": {"f": "(" * 3000 + "1" + ")" * 3000, "v": 1}})],
    )
    report = audit_workbook(path)
    assert "FORMULA_PARSE_ERROR" in [d.code for d in report.diagnostics]


def test_zip_without_worksheet_part_still_reads(tmp_path: Path) -> None:
    path = zip_parts(
        tmp_path / "missing.xlsx",
        [
            (
                "xl/workbook.xml",
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="Ghost" sheetId="1" r:id="rId1"/></sheets></workbook>',
            ),
            (
                "xl/_rels/workbook.xml.rels",
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                'Target="worksheets/sheet1.xml"/></Relationships>',
            ),
        ],
    )
    workbook = read_workbook(path)
    assert workbook.sheet_names() == ["Ghost"]
    assert workbook.sheets[0].cells == ()


def test_reading_does_not_modify_the_file(tmp_path: Path) -> None:
    path = build_workbook(tmp_path / "readonly.xlsx", [simple_sheet("S", {"A1": {"v": 1}})])
    before = path.read_bytes()
    audit_workbook(path)
    assert path.read_bytes() == before
    assert zipfile.is_zipfile(path)
