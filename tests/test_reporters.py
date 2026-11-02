from __future__ import annotations

import json
from pathlib import Path

import pytest
from xlsx_audit import audit_workbook
from xlsx_audit.reporter import (
    render_json,
    render_markdown,
    render_report,
    render_text,
    use_color,
)
from xlsx_audit.sarif_reporter import render_sarif, sarif_document

from factory import build_workbook, simple_sheet


@pytest.fixture()
def bad_report(tmp_path: Path):
    path = build_workbook(
        tmp_path / "butce.xlsx",
        [
            simple_sheet(
                "Gelirler",
                {"E15": {"t": "e", "v": "#REF!"}},
            ),
            simple_sheet(
                "Giderler",
                {"A8": {"v": 100}, "B8": {"f": "A8 * 1.20", "v": 120}},
            ),
            simple_sheet("ÇokGizli", {}, state="veryHidden"),
        ],
        metadata={"creator": "example"},
    )
    return audit_workbook(path)


@pytest.fixture()
def good_report(tmp_path: Path):
    path = build_workbook(tmp_path / "clean.xlsx", [simple_sheet("S", {"A1": {"v": 1}})])
    return audit_workbook(path)


def test_text_report_lists_location_and_code(bad_report) -> None:
    text = render_text(bad_report)
    assert text.startswith("✖ [FAIL] butce.xlsx audit failed")
    assert "Gelirler!E15" in text
    assert "BROKEN_REFERENCE" in text
    assert "[WARNING] Giderler!B8" in text
    assert "findings ·" in text


def test_text_pass_report(good_report) -> None:
    text = render_text(good_report)
    assert text.startswith("✔ [PASS] clean.xlsx audit passed")
    assert "No findings." in text
    assert "(0 errors, 0 warnings)" in text.splitlines()[0]


def test_text_turkish_audits_in_turkish(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "butce_tr.xlsx",
        [simple_sheet("Gelirler", {"E15": {"t": "e", "v": "#REF!"}})],
    )
    report = audit_workbook(path, lang="tr")
    text = render_text(report, lang="tr")
    assert "denetimi başarısız" in text
    assert "[HATA] Gelirler!E15" in text
    assert "Kırık formül referansı" in text
    assert "bulgu ·" in text


def test_text_strict_marker(bad_report) -> None:
    bad_report.strict = True
    assert "[strict]" in render_text(bad_report)
    assert "[katı]" in render_text(bad_report, lang="tr")


def test_info_only_with_verbose(bad_report) -> None:
    plain = render_text(bad_report)
    verbose = render_text(bad_report, verbose=True)
    assert "METADATA_PRESENT" not in plain
    assert "METADATA_PRESENT" in verbose


def test_color_output_contains_ansi(bad_report) -> None:
    colored = render_text(bad_report, color=True)
    assert "\x1b[" in colored
    assert "\x1b[0m" in colored


def test_use_color_respects_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("NO_COLOR", raising=False)
    assert use_color(True) is True
    monkeypatch.setenv("NO_COLOR", "1")
    assert use_color(True) is False
    assert use_color() is False


def test_json_report(bad_report) -> None:
    payload = json.loads(render_json(bad_report))
    assert payload["status"] == "fail"
    assert payload["summary"]["errors"] >= 1
    assert payload["stats"]["sheets"] == 3
    assert set(payload["summary"]["by_code"]) >= {"BROKEN_REFERENCE", "MAGIC_NUMBER_IN_FORMULA"}
    assert payload["diagnostics"][0]["location"]


def test_markdown_report(bad_report) -> None:
    text = render_markdown(bad_report)
    assert text.startswith("# xlsx-audit report — butce.xlsx")
    assert "| Severity | Location | Code | Message |" in text
    assert "`Gelirler!E15`" in text
    assert "| HATA |" in render_markdown(bad_report, lang="tr")


def test_markdown_escapes_pipes() -> None:
    from xlsx_audit.models import AuditReport, Diagnostic, Severity

    report = AuditReport(
        path="pipe.xlsx",
        diagnostics=[
            Diagnostic(
                code="BROKEN_REFERENCE",
                severity=Severity.ERROR,
                message="a|b",
                sheet="S",
                cell="A1",
                rule="broken_refs",
            )
        ],
    )
    text = render_markdown(report)
    assert "a\\|b" in text


def test_sarif_report_shape(bad_report) -> None:
    document = json.loads(render_sarif(bad_report))
    assert document["version"] == "2.1.0"
    assert document["$schema"].endswith("sarif-2.1.0.json")
    run = document["runs"][0]
    driver = run["tool"]["driver"]
    assert driver["name"] == "xlsx-audit"
    assert driver["version"] == "0.1.0"
    rule_ids = [rule["id"] for rule in driver["rules"]]
    assert "BROKEN_REFERENCE" in rule_ids
    for result in run["results"]:
        assert result["ruleId"] in rule_ids
        assert result["ruleIndex"] == rule_ids.index(result["ruleId"])
        assert result["level"] in ("error", "warning", "note")
        location = result["locations"][0]
        assert location["physicalLocation"]["artifactLocation"]["uri"] == "butce.xlsx"
        assert location["physicalLocation"]["region"]["startLine"] == 1
        assert location["logicalLocations"][0]["fullyQualifiedName"]
        assert result["partialFingerprints"]["xlsxAudit/v1"]


def test_sarif_multiple_runs(bad_report, good_report) -> None:
    document = sarif_document([bad_report, good_report])
    assert len(document["runs"]) == 2
    assert document["runs"][1]["results"] == []
    assert document["runs"][1]["properties"]["status"] == "pass"


def test_sarif_dedupes_rule_list(bad_report, good_report) -> None:
    document = sarif_document([bad_report, good_report])
    rule_ids = [rule["id"] for rule in document["runs"][0]["tool"]["driver"]["rules"]]
    assert len(rule_ids) == len(set(rule_ids))


def test_render_report_dispatch(bad_report) -> None:
    assert "xlsx-audit report" in render_report(bad_report, "markdown")
    assert render_report(bad_report, "json").startswith("{")
    with pytest.raises(ValueError, match="unknown report format"):
        render_report(bad_report, "xml")


def test_report_counts(bad_report) -> None:
    assert len(bad_report.errors) == 1
    assert len(bad_report.warnings) >= 2
    assert bad_report.counts_by_code()["BROKEN_REFERENCE"] == 1
    assert bad_report.file_name == "butce.xlsx"
