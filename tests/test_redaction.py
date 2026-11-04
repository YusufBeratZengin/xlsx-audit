"""Reports must not leak paths, query strings, e-mail addresses or cell values.

The same masking has to apply no matter which rule produced the finding: a
magic-number warning on a formula that references ``C:\\Users\\...`` used to
serialise the raw path into JSON and SARIF.
"""

from __future__ import annotations

import json
from pathlib import Path

from xlsx_audit import Diagnostic, Severity, audit_workbook
from xlsx_audit.redact import mask_cell_value, mask_formula, mask_target, mask_text
from xlsx_audit.reporter import render_json, render_markdown, render_text
from xlsx_audit.sarif_reporter import render_sarif

from factory import build_workbook, simple_sheet

SECRET = "SecretAgent"
TOKEN = "secret123"


def _by_code(report, code: str):
    return [d for d in report.diagnostics if d.code == code]


def test_magic_number_on_external_path_formula_is_masked(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "leak.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "A1": {
                        "f": r"=[C:\Users\SecretAgent\Confidential.xlsx]Sheet1!A1*1.25",
                        "v": 10,
                    }
                },
            )
        ],
    )
    report = audit_workbook(path)
    magic = _by_code(report, "MAGIC_NUMBER_IN_FORMULA")
    assert len(magic) == 1
    assert magic[0].formula == "=[…/Confidential.xlsx]Sheet1!A1*1.25"
    assert SECRET not in (magic[0].formula or "")

    payload = json.dumps(report.to_dict())
    assert SECRET not in payload
    assert SECRET not in render_json(report)
    assert SECRET not in render_sarif(report)
    assert SECRET not in render_text(report)


def test_magic_number_on_url_formula_drops_the_query_string(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "url_leak.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "A1": {
                        "f": f'=1.25*LEN("https://api.example.invalid/data?token={TOKEN}")',
                        "v": 1,
                    }
                },
            )
        ],
    )
    report = audit_workbook(path)
    magic = _by_code(report, "MAGIC_NUMBER_IN_FORMULA")
    assert len(magic) == 1
    assert magic[0].formula == '=1.25*LEN("https://api.example.invalid/data")'
    assert TOKEN not in json.dumps(report.to_dict())
    assert TOKEN not in render_json(report)
    assert TOKEN not in render_sarif(report)


def test_every_report_format_uses_the_same_redaction(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "formats.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "A1": {
                        "f": r"=A2*1.5+[C:\Users\SecretAgent\Plan.xlsx]Sheet1!B1",
                        "v": 0,
                    },
                    "A2": {"v": 1},
                },
            )
        ],
    )
    report = audit_workbook(path)
    outputs = [
        render_text(report),
        render_json(report),
        render_markdown(report),
        render_sarif(report),
    ]
    for output in outputs:
        assert SECRET not in output
    assert "…/Plan.xlsx" in outputs[1]
    assert "…/Plan.xlsx" in outputs[3]


def test_range_warning_does_not_echo_the_cell_value(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "range_value.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "A1": {"v": 1},
                    "A2": {"v": 2},
                    "A3": {"v": 987654321},
                    "B1": {"f": "SUM(A1:A2)", "v": 3},
                },
            )
        ],
    )
    report = audit_workbook(path)
    finding = _by_code(report, "OMITTED_CELL_WARNING")[0]
    assert "A3" in finding.message
    assert "987654321" not in finding.message
    assert "<number>" in finding.message
    assert "987654321" not in json.dumps(report.to_dict())
    assert "987654321" not in render_sarif(report)


def test_diagnostic_redacts_directly_constructed_findings() -> None:
    finding = Diagnostic(
        code="X",
        severity=Severity.WARNING,
        formula=r"=A1+[C:\Users\SecretAgent\Book.xlsx]S!A1",
        value=r"C:\Users\SecretAgent\Book.xlsx",
        detail="path C:\\Users\\SecretAgent\\Book.xlsx",
        message="open C:\\Users\\SecretAgent\\Book.xlsx",
        params={"target": "https://example.invalid/x?token=abc"},
        rule="test",
    )
    assert SECRET not in (finding.formula or "")
    assert SECRET not in (finding.value or "")
    assert SECRET not in (finding.detail or "")
    assert SECRET not in finding.message
    assert SECRET not in finding.params["target"]
    assert finding.to_dict()["formula"] == "=A1+[…/Book.xlsx]S!A1"


def test_mask_target_keeps_the_tail_and_drops_the_directory() -> None:
    assert mask_target(r"C:\Users\example\Budget.xlsx") == "…/Budget.xlsx"
    assert mask_target("file:///C:/Users/example/Budget.xlsx") == "…/Budget.xlsx"
    assert mask_target("https://api.example/x?token=abc") == "https://api.example/x"
    assert mask_target("https://example.invalid") == "https://example.invalid/"
    assert mask_target("  ") == "<empty>"
    assert mask_target("/") == "…"


def test_mask_target_truncates_a_long_url_path() -> None:
    long_path = "/" + "a" * 80
    masked = mask_target("https://example.invalid" + long_path)
    assert masked.startswith("https://example.invalid")
    assert "…" in masked
    assert len(masked) < len("https://example.invalid" + long_path)
    assert long_path not in masked


def test_mask_formula_handles_file_and_unc_targets() -> None:
    assert mask_formula("[file:///tmp/example/Book.xlsx]") == "[…/Book.xlsx]"
    assert mask_formula(r"=\\server\share\Book.xlsx") == "=…/Book.xlsx"
    assert mask_formula("") == ""


def test_mask_formula_is_idempotent_and_leaves_plain_formulas_alone() -> None:
    plain = "=SUM(A1:A4)*1.2"
    assert mask_formula(plain) == plain
    assert mask_formula("[1]Sheet1!A1") == "[1]Sheet1!A1"
    assert mask_formula("[Book.xlsx]Sheet1!A1") == "[Book.xlsx]Sheet1!A1"
    once = mask_formula(r"=[C:\Users\example\Budget.xlsx]Sheet1!A1")
    assert once == "=[…/Budget.xlsx]Sheet1!A1"
    assert mask_formula(once) == once


def test_mask_text_and_cell_value_shapes() -> None:
    assert mask_text("mail jane.doe@example.invalid now") == "mail ***@example.invalid now"
    assert mask_text(r"see C:\Reports\secret.xlsx") == "see …/secret.xlsx"
    assert mask_text("see /home/example/secret.xlsx") == "see …/secret.xlsx"
    assert mask_text("=A1/B1") == "=A1/B1"
    assert mask_text("") == ""
    assert mask_cell_value(None) == ""
    assert mask_cell_value("  ") == ""
    assert mask_cell_value("42.5") == "<number>"
    assert mask_cell_value("TRUE") == "<boolean>"
    assert mask_cell_value("hello") == "<text:5>"
    assert mask_cell_value(r"C:\Users\example\Book.xlsx") == "…/Book.xlsx"
