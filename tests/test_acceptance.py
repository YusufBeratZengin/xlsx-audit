"""The card's acceptance criteria, written as executable checks.

These tests mirror the public contract: a finding must carry the exact
``Sheet!Cell`` address, ``MAGIC_NUMBER_IN_FORMULA`` must fire for a rate typed
into a formula, and ``report.passed`` must be the single decision point for CI.
"""

from __future__ import annotations

import json
from pathlib import Path

from xlsx_audit import AuditReport, Severity, audit_workbook
from xlsx_audit.reporter import render_json, render_text

from factory import build_workbook, simple_sheet


def test_broken_reference_has_sheet_and_cell(tmp_path: Path) -> None:
    """AC: a cell holding #REF! or #VALUE! is reported as Sheet1!C14."""
    path = build_workbook(
        tmp_path / "ac1.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "C14": {"f": "A14+B14", "v": "#REF!", "t": "e"},
                    "D14": {"t": "e", "v": "#VALUE!"},
                },
            )
        ],
    )
    report = audit_workbook(path)
    locations = {d.location for d in report.diagnostics}
    assert "Sheet1!C14" in locations
    assert "Sheet1!D14" in locations
    assert report.passed is False
    assert all(d.severity is Severity.ERROR for d in report.diagnostics)


def test_magic_number_warning_code(tmp_path: Path) -> None:
    """AC: a numeric constant such as =B2*0.18 raises MAGIC_NUMBER_IN_FORMULA."""
    path = build_workbook(
        tmp_path / "ac2.xlsx",
        [simple_sheet("Sheet1", {"B2": {"v": 100}, "D2": {"f": "B2*0.18", "v": 18}})],
    )
    report = audit_workbook(path)
    codes = [d.code for d in report.diagnostics]
    assert "MAGIC_NUMBER_IN_FORMULA" in codes
    finding = next(d for d in report.diagnostics if d.code == "MAGIC_NUMBER_IN_FORMULA")
    assert finding.location == "Sheet1!D2"
    assert finding.value == "0.18"


def test_python_api_contract(tmp_path: Path) -> None:
    """The README snippet: audit_workbook(...).passed drives the build."""
    path = build_workbook(
        tmp_path / "ac3.xlsx",
        [simple_sheet("Sheet1", {"A1": {"v": 1}, "A2": {"f": "A1+1", "v": 2}})],
    )
    report: AuditReport = audit_workbook(str(path))
    assert report.passed is True
    assert report.to_dict()["passed"] is True
    assert json.loads(render_json(report))["status"] == "pass"


def test_acceptance_scenario_from_the_card(tmp_path: Path) -> None:
    """The Turkish usage example: two errors and a magic number, per cell."""
    path = build_workbook(
        tmp_path / "Butce_2026.xlsx",
        [
            simple_sheet("Gelirler", {"E15": {"t": "e", "v": "#REF!"}}),
            simple_sheet("Giderler", {"A8": {"v": 100}, "B8": {"f": "A8 * 1.20", "v": 120}}),
        ],
    )
    report = audit_workbook(path, lang="tr", strict=True)
    text = render_text(report)
    assert "Butce_2026.xlsx denetimi başarısız" in text
    assert "Gelirler!E15: Kırık formül referansı bulundu (#REF!)" in text
    assert "Giderler!B8: Formül içine gömülü sabit sayı: 1.2" in text
    assert report.passed is False


def test_committed_example_matches_documentation() -> None:
    """The README output is real: the committed example produces these codes."""
    example = Path(__file__).resolve().parent.parent / "examples" / "fake_financial_model.xlsx"
    report = audit_workbook(example)
    assert report.counts_by_code() == {
        "BROKEN_REFERENCE": 1,
        "EXTERNAL_LINK": 1,
        "MAGIC_NUMBER_IN_FORMULA": 1,
        "OMITTED_CELL_WARNING": 1,
        "VERY_HIDDEN_SHEET": 1,
    }
    assert report.passed is False
    assert report.stats.sheets == 3
    assert report.stats.cells == 15
    assert report.stats.formulas == 4


def test_strict_gate_is_the_release_switch(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "gate.xlsx", [simple_sheet("S", {"A1": {"f": "B1*1.07", "v": 1}})]
    )
    assert audit_workbook(path).passed is True
    assert audit_workbook(path, strict=True).passed is False
