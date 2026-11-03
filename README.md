# xlsx-audit

**A fast, Excel-free linter for `.xlsx` workbooks: broken formula references, circular
references, hard-coded constants, hidden sheets and external links, reported per cell.**

`xlsx-audit` reads the OOXML package directly (the zip that an `.xlsx` really is) and
never evaluates a formula. That means it runs on a Linux CI runner, in a pre-commit hook
or in a container with no Office, no COM automation and no LibreOffice installed.

```console
$ xlsx-audit check examples/fake_financial_model.xlsx
✖ [FAIL] fake_financial_model.xlsx audit failed (1 error, 4 warnings)

- [ERROR] Gelirler!E15: Broken formula reference found (#REF!). (BROKEN_REFERENCE)
- [WARNING] <workbook>: External workbook reference: …/Budget_2025.xlsx (EXTERNAL_LINK)
- [WARNING] Giderler!B8: Hard-coded numeric constant 1.2 in formula. (MAGIC_NUMBER_IN_FORMULA)
- [WARNING] Giderler!D11: Range Giderler!C8:C9 ends before the filled cell C10 (25). (OMITTED_CELL_WARNING)
- [WARNING] Notlar: Sheet is veryHidden: invisible in the Excel UI. (VERY_HIDDEN_SHEET)

5 findings · 3 sheets · 15 cells · 4 formulas
```

*(That output is from the committed `examples/fake_financial_model.xlsx`; every value in
it is synthetic.)*

## Why

Spreadsheets are business logic that nobody reviews. A deleted row leaves `#REF!`
behind, a VAT-style assumption is typed straight into a formula (`=A8*1.20`) and a
`veryHidden` sheet travels with the file when it is mailed to a regulator. The
European Spreadsheet Risks Group has collected such incidents for decades
([eusprig.org](http://www.eusprig.org)); the research prototype
[ExceLint](https://github.com/ExceLint/ExceLint) attacks the same problem but ships as a
Microsoft Excel add-in, so it cannot run on a server or in CI. `xlsx-audit` is the
headless counterpart: one command, deterministic output, machine-readable formats.

## Install

```bash
pip install xlsx-audit
```

Python 3.10+. One runtime dependency, `defusedxml`, which hardens XML parsing against
entity-expansion and external-entity attacks when the workbook comes from an untrusted
source. Everything else is the standard library.

For development:

```bash
git clone https://github.com/YusufBeratZengin/xlsx-audit
cd xlsx-audit
pip install -e ".[dev]"
bash scripts/check.sh
```

## 30-second example

```bash
xlsx-audit check budget.xlsx                # human report, exit 1 on errors
xlsx-audit check budget.xlsx --strict       # warnings fail too (release gate)
xlsx-audit check budget.xlsx --format json  # machine-readable
xlsx-audit check budget.xlsx --format sarif -o audit.sarif
xlsx-audit check budget.xlsx --lang tr      # Turkish messages
xlsx-audit rules                            # list rules
```

Try it without any file of your own:

```bash
python3 scripts/make_example.py
xlsx-audit check examples/fake_financial_model.xlsx
```

## Rules

| Code | Severity | What it means |
| --- | --- | --- |
| `BROKEN_REFERENCE` | error | `#REF!` — a reference was destroyed, usually by deleting a row/column |
| `FORMULA_ERROR_VALUE` | error | `#VALUE!`, `#DIV/0!`, `#NAME?`, `#NUM!`, `#N/A`, … in a cell or formula |
| `CIRCULAR_REFERENCE` | error | formulas depend on themselves, directly or through a chain (all cycles, with the full path) |
| `EXTERNAL_URL` | error | saved link to a URL, a URL string in a formula, or `WEBSERVICE`/`RTD`/`STOCKHISTORY`/`FILTERXML` |
| `EXTERNAL_LINK` | warning | link to another workbook (`[1]Sheet1!A1`, `[C:\…\Budget.xlsx]…`) |
| `MAGIC_NUMBER_IN_FORMULA` | warning | a business constant baked into a formula (`=B2*0.18`) |
| `OMITTED_CELL_WARNING` | warning | `=SUM(A1:A9)` while `A10` is already filled — the off-by-one total |
| `VERY_HIDDEN_SHEET` | warning | a sheet that cannot be unhidden from the Excel UI |
| `HIDDEN_SHEET` | info | a hidden sheet |
| `HIDDEN_DEFINED_NAME` | info | a hidden named range |
| `METADATA_PRESENT` | info | author/company/title metadata; values are never printed |
| `FORMULA_PARSE_ERROR` | warning | the static parser did not understand a construct; other rules skip that cell |
| `GRAPH_LIMIT_EXCEEDED` | info | the workbook is too large for cycle analysis; the rest still runs |

Structural constants (`0`, `1`, `2`, `100`) and index/precision arguments
(`ROUND(x, 2)`, `VLOOKUP(x, r, 3)`, …) are not magic numbers. A literal is reported
once per cell, not once per occurrence.

Findings are sorted by severity and cell, so two runs on the same file always produce
the same order. External targets are masked: the local directory and any query string
are removed before they reach the terminal, JSON or SARIF.

## Exit codes

| Code | Meaning |
| --- | --- |
| `0` | no findings (or only warnings without `--strict`) |
| `1` | findings that fail the run |
| `2` | usage error, or a file that could not be read as an `.xlsx` |

Use `--exit-zero` for reporting-only pipelines, `--ignore CODE[,CODE]` to suppress a
rule's findings, `--rules name[,name]` to run a subset, and `--no-color` (or `NO_COLOR`)
to turn ANSI colours off. `-v` also prints `info` findings.

## Python API

```python
from xlsx_audit import audit_workbook

report = audit_workbook("data.xlsx", strict=True)
if not report.passed:
    for finding in report.diagnostics:
        print(finding.location, finding.code, finding.message)
    raise SystemExit(1)
```

`report.to_dict()` is the same structure the JSON formatter emits. Individual rules can
be selected with `audit_workbook(path, rules=["broken_refs"])`, and `audit_workbook_object`
accepts an already-read `Workbook` (useful for tests and tooling).

## CI and pre-commit

GitHub Actions, with SARIF upload to the Security tab:

```yaml
- uses: actions/checkout@v4
- uses: actions/setup-python@v5
  with:
    python-version: "3.12"
- run: pip install xlsx-audit
- run: xlsx-audit check models/*.xlsx --format sarif -o xlsx-audit.sarif
  continue-on-error: true
- uses: github/codeql-action/upload-sarif@v3
  with:
    sarif_file: xlsx-audit.sarif
```

pre-commit — add this repository as a hook (`xlsx-audit` runs only on staged
`.xlsx`/`.xlsm` files):

```yaml
repos:
  - repo: https://github.com/YusufBeratZengin/xlsx-audit
    rev: v0.1.0
    hooks:
      - id: xlsx-audit
```

## How it works

```
xlsx_audit/
├── ooxml.py           # bounded zip reader: workbook.xml, rels, sharedStrings, sheets
├── formula_parser.py  # tokenizer + AST (refs, ranges, calls, arrays, operators)
├── refs.py            # A1 notation helpers
├── graph_audit.py     # dependency graph, iterative Tarjan SCC → cycles
├── rules/             # one module per rule, shared registry
├── reporter.py        # text / JSON / Markdown
├── sarif_reporter.py  # SARIF 2.1.0
└── cli.py             # argparse entry point
```

Shared formulas (`<f t="shared" si="0" ref="A1:A9"/>`) are translated onto every follower
cell with the same relative-reference shifting Excel performs, so a rule sees the formula
that actually runs in each cell. References shifted outside the sheet become `#REF!`,
exactly like Excel. Array (spill) formulas are read as written. Formula text is never
evaluated, so a workbook saved by a library that does not cache results is still audited.

## Limitations

- **Static only.** `xlsx-audit` does not compute values, so it cannot prove that a
  formula is *correct*, only that it contains known risk patterns. It is not a
  substitute for an audit firm's review or for a four-eyes check of a financial model.
- Whole-column references (`A:A`) and ranges larger than 20 000 cells are not expanded
  for cycle analysis; a cycle hidden exclusively behind such a reference is missed.
- Whole-row references (`1:1`), parenthesised unions (`SUM((A1,B2))`) and spill
  references (`A1#`) are not part of the parsed grammar; a cell using them is reported
  as `FORMULA_PARSE_ERROR` and the other rules skip it.
- Formulas nested deeper than 64 levels (Excel's own nesting limit) are not parsed; they
  are reported as `FORMULA_PARSE_ERROR` instead of being allowed to exhaust the stack.
- `IFERROR`/`IF` short-circuits and `INDIRECT`/`OFFSET` indirection are not followed.
- VBA macros are not inspected or executed (`.xlsm` is treated as static XML too).
- Text inside a `veryHidden` sheet is reported as a sheet-level finding; cell contents
  are not echoed, because they may be personal data.
- No official tax, VAT or legal rate is bundled anywhere in this project. If a rate is
  needed for a test, it is synthetic; anything else would be marked `NEEDS_SOURCE`.

## Licence

MIT — see [LICENSE](LICENSE). Copyright (c) 2026 YusufBeratZengin.
