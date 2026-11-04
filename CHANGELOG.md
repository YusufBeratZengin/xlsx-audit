# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Saved external hyperlink detection across worksheet relationships (`xl/worksheets/_rels/*.rels`) with cell coordinates.
- End-to-end regression tests for external target and query masking in formulas, JSON and SARIF reports.

### Fixed

- Validated cell text before conversion to `int` in shared string lookups, resolving strict mypy typing errors.
- Masked external workbook paths and sensitive URL query parameters within cell formulas and values across all report representations (JSON, SARIF, text, markdown).
- Handled filesystem errors when writing to `--output` in CLI with a friendly error message and exit code 2.

### Changed

- Circular-reference reporting now enumerates the elementary cycles of each strongly
  connected component (iterative Johnson's algorithm, at most 100 cycles per component)
  instead of reporting a single representative path. Several distinct loops sharing
  cells in one component are therefore all reported, each anchored at its smallest cell
  coordinate.
- Documented the formula constructs the static parser does not accept (whole-row
  references such as `1:1`, parenthesised unions and spill references), which are
  reported as `FORMULA_PARSE_ERROR` instead of being silently misread.

## [0.1.0] - 2026-11-03

First public release. Everything below is implemented and covered by the test suite.

### Added

- Bounded OOXML reader that works directly on the `.xlsx` zip: workbook part,
  relationships, shared strings, worksheets, defined names, document properties and
  external links. Zip-bomb limits and hardened XML parsing (`defusedxml`).
- Dependency-free Excel formula tokenizer and AST builder: A1 / `$A$1` references,
  quoted sheet names, external workbook prefixes (`[1]Sheet1!A1`,
  `[C:\Budget.xlsx]Sheet1!A1`), ranges, function calls, array constants, unary
  operators, percent and comparison operators.
- Shared-formula (`t="shared"`) translation onto follower cells, and `#REF!` for
  references shifted outside the sheet — the same behaviour as Excel.
- Rules: `broken_refs` (`#REF!` and every Excel error value), `magic_numbers`,
  `external_links` (saved links, URLs in formulas, `WEBSERVICE`/`RTD`/`FILTERXML`/
  `STOCKHISTORY`), `hidden_sheets` (`hidden`, `veryHidden`, hidden names, metadata),
  `circular_references` (iterative Tarjan SCC, full cycle path) and `range_checker`
  (`OMITTED_CELL_WARNING`).
- Reporters: colour terminal output, JSON, Markdown and SARIF 2.1.0 (including
  `partialFingerprints` for Code Scanning), with masked external targets.
- CLI: `check` and `rules` subcommands, `--strict`, `--format`, `--output`, `--lang en|tr`,
  `--ignore`, `--rules`, `--color/--no-color`, `--verbose`, `--exit-zero`; exit codes 0/1/2.
- Python API: `audit_workbook`, `audit_workbook_object`, `AuditReport`,
  `Workbook`, `Diagnostic`.
- pre-commit hook definition (`.pre-commit-hooks.yaml`) and a GitHub Actions example with
  SARIF upload.
- `examples/fake_financial_model.xlsx` plus `scripts/make_example.py` to rebuild it.
- `scripts/check.sh`: lint, format check, strict typing, tests and coverage gate.

### Known limitations

- Static analysis only; formulas are never evaluated.
- Whole-column references and ranges above 20 000 cells are not expanded for cycle
  detection.
- VBA macros are neither inspected nor executed.

[Unreleased]: https://github.com/YusufBeratZengin/xlsx-audit/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/YusufBeratZengin/xlsx-audit/releases/tag/v0.1.0
