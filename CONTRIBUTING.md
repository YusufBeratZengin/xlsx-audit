# Contributing

Thanks for taking the time to improve `xlsx-audit`. This project is small on purpose:
one dependency, no framework, deterministic output. Please keep it that way.

## Quick start

```bash
git clone https://github.com/YusufBeratZengin/xlsx-audit
cd xlsx-audit
python3 -m venv .venv
. .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -e ".[dev]"
bash scripts/check.sh
```

`scripts/check.sh` runs ruff (lint + format check), mypy in strict mode, the whole test
suite and a coverage gate. It must exit 0 before a pull request is opened. No network
access is needed: the tests build their own `.xlsx` fixtures with `zipfile`.

## What to work on

- Issues labelled [`good first issue`](https://github.com/YusufBeratZengin/xlsx-audit/labels/good%20first%20issue)
  are scoped for a first contribution.
- Bug reports are most useful with a *minimal* workbook that reproduces the finding.
  Never attach a real spreadsheet: remove names, numbers and links first, or rebuild the
  case with `tests/factory.py`.
- New rules are welcome, but a rule needs: a stable `CODE`, a severity, a test that is
  red without the rule and green with it, and a sentence in the README rule table.

## Adding a rule

1. Create `xlsx_audit/rules/<name>.py` with a `Rule` subclass (see
   `xlsx_audit/rules/base.py`). Use `self.diagnostic(..., message_key="...", params={...})`
   so the finding can be rendered in every language; add the template to
   `xlsx_audit/i18n.py` for both `en` and `tr`.
2. Register the class in `xlsx_audit/rules/__init__.py`.
3. Add tests in `tests/test_rules.py` covering the positive case, the negatives (near
   misses) and an unparseable formula.
4. Document the code in `README.md` and `README.tr.md`, and add a CHANGELOG entry.

## Style

- Python 3.10+, `from __future__ import annotations`, full type hints (`mypy --strict`).
- Line length 100, double quotes; `ruff format` decides the rest.
- No new runtime dependency without a written justification in the PR and in the README.
- Never print cell values that may be personal data; report the location and a masked
  form instead.
- Do not invent official rates, thresholds or legal values anywhere in code, docs or
  fixtures. If a number has no source, it does not belong in the repository
  (`NEEDS_SOURCE` in the issue instead).

## Tests

```bash
pytest -q                       # fast
pytest -q --cov=xlsx_audit      # with coverage
pytest -q -k magic              # one rule
```

Fixtures are synthetic. `tests/factory.py` can write a real OOXML package, which is the
supported way to reproduce an Excel quirk without shipping somebody's workbook.

## Commits and pull requests

- Short imperative commit subjects ("Add SARIF fingerprints"), one topic per commit.
- Describe *why* in the PR body, and paste the `scripts/check.sh` result.
- Update `CHANGELOG.md` under `## [Unreleased]`.
- By contributing you agree that your contribution is licensed under the MIT licence of
  this repository.

## Reporting security issues

Please do not open a public issue for a vulnerability. See [SECURITY.md](SECURITY.md).
