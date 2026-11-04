from __future__ import annotations

import json
from pathlib import Path

import pytest
from xlsx_audit import Severity, audit_workbook
from xlsx_audit.sarif_reporter import sarif_document

from factory import build_workbook, simple_sheet


def codes(report) -> list[str]:
    return [d.code for d in report.diagnostics]


def by_code(report, code: str):
    return [d for d in report.diagnostics if d.code == code]


# ----------------------------------------------------------------- broken refs


def test_broken_reference_reports_sheet_and_cell(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "broken.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "C14": {"f": "#REF!+1", "v": "#REF!"},
                    "D1": {"t": "e", "v": "#VALUE!"},
                },
            )
        ],
    )
    report = audit_workbook(path)
    broken = by_code(report, "BROKEN_REFERENCE")
    assert len(broken) == 1
    assert broken[0].location == "Sheet1!C14"
    assert broken[0].severity is Severity.ERROR
    assert "#REF!" in broken[0].message
    errors = by_code(report, "FORMULA_ERROR_VALUE")
    assert [d.location for d in errors] == ["Sheet1!D1"]
    assert report.passed is False


def test_error_value_without_formula(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "value.xlsx",
        [simple_sheet("Data", {"B2": {"t": "e", "v": "#DIV/0!"}})],
    )
    report = audit_workbook(path)
    finding = by_code(report, "FORMULA_ERROR_VALUE")[0]
    assert finding.location == "Data!B2"
    assert finding.formula is None


def test_clean_workbook_passes(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "clean.xlsx",
        [simple_sheet("Sheet1", {"A1": {"v": 1}, "A2": {"f": "A1+1", "v": 2}})],
    )
    report = audit_workbook(path)
    assert report.passed
    assert report.status == "pass"
    assert report.diagnostics == []


# --------------------------------------------------------------- magic numbers


def test_magic_number_in_formula(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "magic.xlsx",
        [
            simple_sheet(
                "Giderler",
                {
                    "A8": {"v": 100},
                    "B8": {"f": "A8*1.20", "v": 120},
                    "B9": {"f": "B2*0.18", "v": 3},
                },
            )
        ],
    )
    report = audit_workbook(path)
    findings = by_code(report, "MAGIC_NUMBER_IN_FORMULA")
    assert {f.location for f in findings} == {"Giderler!B8", "Giderler!B9"}
    assert findings[0].severity is Severity.WARNING
    assert "1.2" in findings[0].message or "1.2" in (findings[0].value or "")
    assert report.passed  # warnings do not fail a normal run


def test_magic_numbers_allow_structural_constants(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "allowed.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "A1": {"f": "B1*1", "v": 1},
                    "A2": {"f": "B1+0", "v": 1},
                    "A3": {"f": "B1/2", "v": 1},
                    "A4": {"f": "B1*100", "v": 100},
                    "A5": {"f": "ROUND(B1,4)", "v": 1},
                    "A6": {"f": "INDEX(C1:C9,3)", "v": 1},
                    "B1": {"v": 1},
                },
            )
        ],
    )
    report = audit_workbook(path)
    assert by_code(report, "MAGIC_NUMBER_IN_FORMULA") == []


def test_magic_number_negative_and_percent(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "neg.xlsx",
        [simple_sheet("S", {"A1": {"f": "B1*-1.5", "v": 0}, "A2": {"f": "B1*18%", "v": 0}})],
    )
    findings = by_code(audit_workbook(path), "MAGIC_NUMBER_IN_FORMULA")
    values = sorted(f.value or "" for f in findings)
    assert values == ["-1.5", "18"]


def test_strict_mode_fails_on_warnings(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "strict.xlsx",
        [simple_sheet("S", {"A1": {"f": "B1*1.5", "v": 0}})],
    )
    assert audit_workbook(path).passed
    strict = audit_workbook(path, strict=True)
    assert strict.passed is False
    assert strict.status == "fail"


# -------------------------------------------------------------- external links


def test_external_link_and_url(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "links.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "A1": {"f": "[1]Sheet1!A1", "v": 1},
                    "A2": {"f": 'WEBSERVICE("https://example.invalid/api?token=fake")', "v": 0},
                },
            )
        ],
        external_targets=[r"file:///C:/Users/example/Budget.xlsx"],
    )
    report = audit_workbook(path)
    assert "EXTERNAL_LINK" in codes(report)
    urls = by_code(report, "EXTERNAL_URL")
    assert urls and all(u.severity is Severity.ERROR for u in urls)
    assert not any("token" in (u.message + (u.value or "") + (u.formula or "")) for u in urls)
    assert not any(
        "Users" in (d.message + (d.value or "") + (d.formula or "")) for d in report.diagnostics
    )
    dumped = json.dumps(report.to_dict())
    assert "token=fake" not in dumped
    assert "C:/Users" not in dumped


def test_external_formula_book_path_masked(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "masked_book.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "A1": {"f": r"[C:\Users\SecretAgent\Confidential.xlsx]Sheet1!A1", "v": 10},
                    "A2": {
                        "f": 'WEBSERVICE("https://api.internal.corp/query?token=secret123&user=admin")',
                        "v": 0,
                    },
                },
            )
        ],
    )
    report = audit_workbook(path)
    link_findings = by_code(report, "EXTERNAL_LINK")
    assert len(link_findings) == 1
    lf = link_findings[0]
    assert lf.location == "Sheet1!A1"
    assert lf.value == "…/Confidential.xlsx"
    assert "SecretAgent" not in lf.message
    assert "SecretAgent" not in (lf.formula or "")
    assert lf.formula == "=[…/Confidential.xlsx]Sheet1!A1"

    url_findings = by_code(report, "EXTERNAL_URL")
    assert len(url_findings) == 2
    for uf in url_findings:
        assert uf.location == "Sheet1!A2"
        assert uf.severity is Severity.ERROR
        assert "secret123" not in uf.message
        assert "secret123" not in (uf.value or "")
        assert "secret123" not in (uf.formula or "")
        assert uf.formula == '=WEBSERVICE("https://api.internal.corp/query")'

    # Verify JSON and SARIF do not leak sensitive paths or tokens
    json_text = json.dumps(report.to_dict())
    assert "SecretAgent" not in json_text
    assert "secret123" not in json_text

    sarif_text = json.dumps(sarif_document([report]))
    assert "SecretAgent" not in sarif_text
    assert "secret123" not in sarif_text


def test_worksheet_hyperlink_detected_and_masked(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "ws_links.xlsx",
        [
            simple_sheet(
                "Sheet1",
                cells={
                    "B2": {"v": "External Portal"},
                    "C5": {"v": "Private Budget"},
                },
                hyperlinks={"B2": "rId1", "C5": "rId2"},
                rels={
                    "rId1": "https://phishing.site/login?redirect=steal&session=abc123xyz",
                    "rId2": r"file:///C:/Users/victim/SecretPlan.xlsx",
                },
            )
        ],
    )
    report = audit_workbook(path)
    url_findings = by_code(report, "EXTERNAL_URL")
    assert len(url_findings) == 1
    uf = url_findings[0]
    assert uf.location == "Sheet1!B2"
    assert uf.sheet == "Sheet1"
    assert uf.cell == "B2"
    assert uf.severity is Severity.ERROR
    assert uf.value == "https://phishing.site/login"
    assert "session=" not in uf.message
    assert "session=" not in (uf.value or "")

    link_findings = by_code(report, "EXTERNAL_LINK")
    assert len(link_findings) == 1
    lf = link_findings[0]
    assert lf.location == "Sheet1!C5"
    assert lf.sheet == "Sheet1"
    assert lf.cell == "C5"
    assert lf.severity is Severity.WARNING
    assert lf.value == "…/SecretPlan.xlsx"
    assert "victim" not in lf.message
    assert "victim" not in (lf.value or "")


def test_worksheet_hyperlink_without_cell_ref(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "ws_rels_only.xlsx",
        [
            simple_sheet(
                "Sheet1",
                cells={"A1": {"v": 1}},
                rels={"rId1": "https://example.invalid/data?key=topsecret"},
            )
        ],
    )
    report = audit_workbook(path)
    url_findings = by_code(report, "EXTERNAL_URL")
    assert len(url_findings) == 1
    uf = url_findings[0]
    assert uf.location == "Sheet1"
    assert uf.sheet == "Sheet1"
    assert uf.cell is None
    assert uf.value == "https://example.invalid/data"
    assert "topsecret" not in uf.message


def test_external_http_link_is_an_error(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "http.xlsx",
        [simple_sheet("Sheet1", {"A1": {"v": 1}})],
        external_targets=["https://example.invalid/data.xlsx"],
    )
    report = audit_workbook(path)
    finding = by_code(report, "EXTERNAL_URL")[0]
    assert finding.severity is Severity.ERROR
    assert "example.invalid" in finding.message


def test_plain_formula_has_no_external_findings(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "plain.xlsx", [simple_sheet("S", {"A1": {"f": "SUM(B1:B3)", "v": 1}})]
    )
    assert by_code(audit_workbook(path), "EXTERNAL_LINK") == []


# --------------------------------------------------------------- hidden sheets


def test_hidden_and_very_hidden_sheets(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "hidden.xlsx",
        [
            simple_sheet("Visible", {}),
            simple_sheet("Shy", {}, state="hidden"),
            simple_sheet("TopSecret", {}, state="veryHidden"),
        ],
    )
    report = audit_workbook(path)
    very = by_code(report, "VERY_HIDDEN_SHEET")
    assert [d.sheet for d in very] == ["TopSecret"]
    assert very[0].severity is Severity.WARNING
    hidden = by_code(report, "HIDDEN_SHEET")
    assert [d.sheet for d in hidden] == ["Shy"]
    assert hidden[0].severity is Severity.INFO


def test_metadata_is_reported_without_values(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "meta.xlsx",
        [simple_sheet("S", {"A1": {"v": 1}})],
        metadata={"creator": "Jane Example", "title": "Fake Budget"},
    )
    report = audit_workbook(path)
    finding = by_code(report, "METADATA_PRESENT")[0]
    assert "Jane" not in finding.message
    assert "creator" in finding.message


def test_hidden_defined_name(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "name.xlsx",
        [simple_sheet("S", {"A1": {"v": 1}})],
        defined_names=[{"name": "SECRET_RATE", "formula": "S!$A$1", "hidden": True}],
    )
    finding = by_code(audit_workbook(path), "HIDDEN_DEFINED_NAME")[0]
    assert finding.value == "SECRET_RATE"


# ------------------------------------------------------------------- range gap


def test_omitted_cell_warning(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "range.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "A1": {"v": 1},
                    "A2": {"v": 2},
                    "A3": {"v": 3},
                    "B1": {"f": "SUM(A1:A2)", "v": 3},
                },
            )
        ],
    )
    findings = by_code(audit_workbook(path), "OMITTED_CELL_WARNING")
    assert len(findings) == 1
    assert findings[0].location == "Sheet1!B1"
    assert "A3" in findings[0].message


def test_no_warning_when_range_covers_data(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "covered.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "A1": {"v": 1},
                    "A2": {"v": 2},
                    "A3": {"v": 3},
                    "B1": {"f": "SUM(A1:A3)", "v": 6},
                },
            )
        ],
    )
    assert by_code(audit_workbook(path), "OMITTED_CELL_WARNING") == []


def test_no_warning_when_cell_already_referenced(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "referenced.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "A1": {"v": 1},
                    "A2": {"v": 2},
                    "A3": {"v": 3},
                    "B1": {"f": "SUM(A1:A2)+A3", "v": 6},
                },
            )
        ],
    )
    assert by_code(audit_workbook(path), "OMITTED_CELL_WARNING") == []


def test_horizontal_omitted_cell(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "horizontal.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "A1": {"v": 1},
                    "B1": {"v": 2},
                    "C1": {"v": 3},
                    "A2": {"f": "SUM(A1:B1)", "v": 3},
                },
            )
        ],
    )
    findings = by_code(audit_workbook(path), "OMITTED_CELL_WARNING")
    assert len(findings) == 1
    assert "C1" in findings[0].message


# ------------------------------------------------------------- circular refs


def test_three_level_cycle(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "cycle.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "A1": {"f": "B1+1", "v": 0},
                    "B1": {"f": "C1+1", "v": 0},
                    "C1": {"f": "A1+1", "v": 0},
                },
            )
        ],
    )
    report = audit_workbook(path)
    cycles = by_code(report, "CIRCULAR_REFERENCE")
    assert len(cycles) == 1
    assert cycles[0].severity is Severity.ERROR
    path_text = cycles[0].detail or ""
    assert "A1" in path_text and "B1" in path_text and "C1" in path_text
    assert path_text.count("->") == 3


def test_self_reference(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "self.xlsx", [simple_sheet("S", {"A1": {"f": "A1+1", "v": 0}})]
    )
    assert len(by_code(audit_workbook(path), "CIRCULAR_REFERENCE")) == 1


def test_multiple_cycles_in_same_component_reported(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "multi_cycle.xlsx",
        [
            simple_sheet(
                "Sheet1",
                {
                    "A1": {"f": "B1+C1", "v": 0},
                    "B1": {"f": "A1+1", "v": 0},
                    "C1": {"f": "A1+1", "v": 0},
                },
            )
        ],
    )
    report = audit_workbook(path)
    cycles = by_code(report, "CIRCULAR_REFERENCE")
    assert [c.detail for c in cycles] == [
        "Sheet1!A1 -> Sheet1!B1 -> Sheet1!A1",
        "Sheet1!A1 -> Sheet1!C1 -> Sheet1!A1",
    ]


def test_cross_sheet_cycle(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "cross.xlsx",
        [
            simple_sheet("One", {"A1": {"f": "Two!B1", "v": 0}}),
            simple_sheet("Two", {"B1": {"f": "One!A1", "v": 0}}),
        ],
    )
    cycling = by_code(audit_workbook(path), "CIRCULAR_REFERENCE")[0]
    assert "One!A1" in (cycling.detail or "")
    assert "Two!B1" in (cycling.detail or "")


def test_no_cycle_for_a_chain(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "chain.xlsx",
        [
            simple_sheet(
                "S",
                {
                    "A1": {"v": 1},
                    "A2": {"f": "A1+1", "v": 2},
                    "A3": {"f": "SUM(A1:A2)", "v": 3},
                },
            )
        ],
    )
    assert by_code(audit_workbook(path), "CIRCULAR_REFERENCE") == []


def test_cycle_through_named_range(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "named.xlsx",
        [simple_sheet("S", {"A1": {"f": "TOTAL+1", "v": 0}, "A2": {"v": 1}})],
        defined_names=[{"name": "TOTAL", "formula": "S!$A$1"}],
    )
    assert len(by_code(audit_workbook(path), "CIRCULAR_REFERENCE")) == 1


# ---------------------------------------------------------------- misc / i18n


def test_parse_error_is_a_warning_not_a_crash(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "parse.xlsx", [simple_sheet("S", {"A1": {"f": "SUM(", "v": 0}})]
    )
    report = audit_workbook(path)
    finding = by_code(report, "FORMULA_PARSE_ERROR")[0]
    assert finding.severity is Severity.WARNING
    assert report.passed


def test_ignore_suppresses_by_code(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "ignore.xlsx", [simple_sheet("S", {"A1": {"f": "B1*1.5", "v": 0}})]
    )
    report = audit_workbook(path, ignore=["MAGIC_NUMBER_IN_FORMULA"])
    assert report.diagnostics == []
    assert report.passed


def test_unknown_rule_is_rejected(tmp_path: Path) -> None:
    path = build_workbook(tmp_path / "r.xlsx", [simple_sheet("S", {})])
    with pytest.raises(ValueError, match="unknown rule"):
        audit_workbook(path, rules=["nope"])


def test_only_selected_rule_runs(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "only.xlsx",
        [simple_sheet("S", {"A1": {"f": "#REF!*2", "v": "#REF!"}})],
    )
    report = audit_workbook(path, rules=["magic_numbers"])
    assert report.rules == ("magic_numbers",)
    assert codes(report) == []


def test_turkish_messages(tmp_path: Path) -> None:
    from xlsx_audit.reporter import render_text

    path = build_workbook(
        tmp_path / "tr.xlsx",
        [simple_sheet("Gelirler", {"E15": {"t": "e", "v": "#REF!"}})],
    )
    report = audit_workbook(path, lang="tr")
    assert "Kırık formül referansı" in report.diagnostics[0].text("tr")
    assert "Kırık formül referansı" in render_text(report)
    assert report.lang == "tr"


def test_large_range_is_not_expanded(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "wide.xlsx",
        [simple_sheet("S", {"A1": {"f": "SUM(A1:XFD100)", "v": 0}})],
    )
    assert by_code(audit_workbook(path), "CIRCULAR_REFERENCE") == []
