"""CLI behaviour: exit codes, formats, output routing and error paths."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from xlsx_audit import __version__, cli
from xlsx_audit.cli import main

from factory import build_workbook, simple_sheet


@pytest.fixture()
def bad_book(tmp_path: Path) -> Path:
    return build_workbook(
        tmp_path / "butce.xlsx",
        [
            simple_sheet(
                "Gelirler",
                {"E15": {"t": "e", "v": "#REF!"}},
            ),
            simple_sheet(
                "Giderler",
                {"A8": {"v": 100}, "B8": {"f": "A8*1.20", "v": 120}},
            ),
        ],
    )


@pytest.fixture()
def good_book(tmp_path: Path) -> Path:
    return build_workbook(
        tmp_path / "clean.xlsx",
        [simple_sheet("Sheet1", {"A1": {"v": 1}, "A2": {"f": "A1+1", "v": 2}})],
    )


def test_clean_file_exits_zero(good_book: Path, capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["check", str(good_book)]) == 0
    out = capsys.readouterr().out
    assert "[PASS]" in out
    assert "No findings." in out


def test_errors_exit_one_and_show_location(
    bad_book: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    assert main(["check", str(bad_book)]) == 1
    out = capsys.readouterr().out
    assert "[FAIL]" in out
    assert "Gelirler!E15" in out
    assert "BROKEN_REFERENCE" in out
    assert "MAGIC_NUMBER_IN_FORMULA" in out


def test_strict_makes_warnings_fail(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "warn.xlsx", [simple_sheet("S", {"A1": {"f": "B1*1.5", "v": 1}})]
    )
    assert main(["check", str(path)]) == 0
    assert main(["check", str(path), "--strict"]) == 1


def test_exit_zero_flag(bad_book: Path) -> None:
    assert main(["check", str(bad_book), "--exit-zero"]) == 0


def test_missing_file_exits_two(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["check", str(tmp_path / "nope.xlsx")]) == 2
    captured = capsys.readouterr()
    assert "xlsx-audit:" in captured.err
    assert "file not found" in captured.err


def test_not_a_workbook_exits_two(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    path = tmp_path / "fake.xlsx"
    path.write_bytes(b"definitely not a zip")
    assert main(["check", str(path)]) == 2
    assert "not a readable" in capsys.readouterr().err


def test_output_write_error_exits_two(
    good_book: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    bad_output_path = tmp_path / "nonexistent_dir" / "out.txt"
    assert main(["check", str(good_book), "-o", str(bad_output_path)]) == 2
    captured = capsys.readouterr()
    assert "xlsx-audit: cannot write output to" in captured.err


def test_multiple_files_text(
    bad_book: Path, good_book: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    assert main(["check", str(bad_book), str(good_book)]) == 1
    out = capsys.readouterr().out
    assert "butce.xlsx" in out and "clean.xlsx" in out


def test_json_format_single(bad_book: Path, capsys: pytest.CaptureFixture[str]) -> None:
    main(["check", str(bad_book), "--format", "json"])
    payload = json.loads(capsys.readouterr().out)
    assert payload["status"] == "fail"
    assert payload["summary"]["errors"] == 1


def test_json_format_multiple(
    bad_book: Path, good_book: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    main(["check", str(bad_book), str(good_book), "--format", "json"])
    payload = json.loads(capsys.readouterr().out)
    assert isinstance(payload, list)
    assert [item["status"] for item in payload] == ["fail", "pass"]


def test_markdown_format(bad_book: Path, capsys: pytest.CaptureFixture[str]) -> None:
    main(["check", str(bad_book), "--format", "markdown"])
    out = capsys.readouterr().out
    assert out.startswith("# xlsx-audit report")
    assert "| Severity | Location | Code | Message |" in out


def test_sarif_format_single(bad_book: Path, capsys: pytest.CaptureFixture[str]) -> None:
    main(["check", str(bad_book), "--format", "sarif"])
    document = json.loads(capsys.readouterr().out)
    assert document["version"] == "2.1.0"
    assert len(document["runs"]) == 1


def test_sarif_format_multiple(
    bad_book: Path, good_book: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    main(["check", str(bad_book), str(good_book), "--format", "sarif"])
    document = json.loads(capsys.readouterr().out)
    assert len(document["runs"]) == 2


def test_output_writes_file(
    bad_book: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    target = tmp_path / "report.md"
    assert main(["check", str(bad_book), "--format", "markdown", "-o", str(target)]) == 1
    assert target.read_text(encoding="utf-8").startswith("# xlsx-audit report")
    captured = capsys.readouterr()
    assert captured.out == ""
    assert "report written to" in captured.err


def test_output_rejects_multiple_inputs(bad_book: Path, good_book: Path) -> None:
    assert main(["check", str(bad_book), str(good_book), "-o", "out.txt"]) == 2


def test_ignore_suppresses_code(bad_book: Path, capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["check", str(bad_book), "--ignore", "MAGIC_NUMBER_IN_FORMULA"]) == 1
    out = capsys.readouterr().out
    assert "MAGIC_NUMBER_IN_FORMULA" not in out
    assert "BROKEN_REFERENCE" in out


def test_ignore_accepts_comma_separated_list(
    bad_book: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    code = main(["check", str(bad_book), "--ignore", "BROKEN_REFERENCE,MAGIC_NUMBER_IN_FORMULA"])
    assert code == 0  # every finding in this workbook is suppressed
    out = capsys.readouterr().out
    assert "BROKEN_REFERENCE" not in out
    assert "[PASS]" in out


def test_ignore_everything_passes(bad_book: Path, capsys: pytest.CaptureFixture[str]) -> None:
    every_code = (
        "BROKEN_REFERENCE,FORMULA_ERROR_VALUE,MAGIC_NUMBER_IN_FORMULA,"
        "EXTERNAL_LINK,EXTERNAL_URL,HIDDEN_SHEET,VERY_HIDDEN_SHEET,METADATA_PRESENT,"
        "HIDDEN_DEFINED_NAME,OMITTED_CELL_WARNING,CIRCULAR_REFERENCE,FORMULA_PARSE_ERROR"
    )
    assert main(["check", str(bad_book), "--ignore", every_code]) == 0
    assert "[PASS]" in capsys.readouterr().out


def test_rules_subset(bad_book: Path, capsys: pytest.CaptureFixture[str]) -> None:
    main(["check", str(bad_book), "--rules", "magic_numbers"])
    out = capsys.readouterr().out
    assert "MAGIC_NUMBER_IN_FORMULA" in out
    assert "BROKEN_REFERENCE" not in out


def test_unknown_rule_exits_two(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    path = build_workbook(tmp_path / "book.xlsx", [simple_sheet("S", {})])
    assert main(["check", str(path), "--rules", "does_not_exist"]) == 2
    assert "unknown rule" in capsys.readouterr().err


def test_language_option(bad_book: Path, capsys: pytest.CaptureFixture[str]) -> None:
    main(["check", str(bad_book), "--lang", "tr"])
    out = capsys.readouterr().out
    assert "denetimi başarısız" in out
    assert "[HATA]" in out


def test_language_from_environment(
    bad_book: Path, capsys: pytest.CaptureFixture[str], monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("XLSX_AUDIT_LANG", "tr")
    main(["check", str(bad_book)])
    assert "[HATA]" in capsys.readouterr().out


def test_color_flags(bad_book: Path, capsys: pytest.CaptureFixture[str]) -> None:
    main(["check", str(bad_book), "--color"])
    assert "\x1b[" in capsys.readouterr().out
    main(["check", str(bad_book), "--no-color"])
    assert "\x1b[" not in capsys.readouterr().out


def test_verbose_shows_info(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    path = build_workbook(
        tmp_path / "meta.xlsx",
        [simple_sheet("S", {"A1": {"v": 1}})],
        metadata={"creator": "example"},
    )
    main(["check", str(path)])
    assert "METADATA_PRESENT" not in capsys.readouterr().out
    main(["check", str(path), "-v"])
    assert "METADATA_PRESENT" in capsys.readouterr().out


def test_rules_command(capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["rules"]) == 0
    out = capsys.readouterr().out
    for name in ("broken_refs", "magic_numbers", "external_links", "hidden_sheets"):
        assert name in out


def test_no_command_prints_help_and_returns_usage(
    capsys: pytest.CaptureFixture[str],
) -> None:
    assert main([]) == 2
    assert "usage:" in capsys.readouterr().out


def test_version_flag(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as excinfo:
        main(["--version"])
    assert excinfo.value.code == 0
    assert __version__ in capsys.readouterr().out


def test_invalid_format_is_rejected_by_argparse() -> None:
    with pytest.raises(SystemExit) as excinfo:
        main(["check", "x.xlsx", "--format", "xml"])
    assert excinfo.value.code == 2


def test_exit_codes_are_stable() -> None:
    assert (cli.EXIT_OK, cli.EXIT_FINDINGS, cli.EXIT_USAGE) == (0, 1, 2)
