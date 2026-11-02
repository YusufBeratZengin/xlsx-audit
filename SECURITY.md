# Security policy

## Supported versions

| Version | Supported |
| --- | --- |
| 0.1.x | yes |

## Reporting a vulnerability

Please **do not** open a public issue for a security problem. Use GitHub's private
reporting form ("Security" tab → "Report a vulnerability") on
<https://github.com/YusufBeratZengin/xlsx-audit/security/advisories/new>. Include the
version, the operating system and a
minimal reproduction (a synthetic `.xlsx` built with `tests/factory.py` is ideal — never a
real workbook). You can expect an acknowledgement within a week.

## Threat model

`xlsx-audit` opens files that may come from untrusted sources (downloads, uploads,
email attachments). The relevant risks and the current mitigations:

- **Zip bombs / oversized packages.** Entry count, per-part size and total uncompressed
  size are bounded (`MAX_ENTRIES`, `MAX_PART_BYTES`, `MAX_TOTAL_BYTES` in
  `xlsx_audit/ooxml.py`). A package over the limit is rejected with
  `WorkbookReadError` instead of being expanded.
- **XML external entities and entity expansion.** All XML is parsed with `defusedxml`,
  which forbids DTDs, external entities and exponential entity expansion.
- **Path traversal.** Relationship targets are normalised with `posixpath` and only ever
  looked up inside the already-open zip; nothing is extracted to disk.
- **Read-only by design.** The tool never writes to the workbook it audits and never
  executes formulas or macros.
- **Data exposure in reports.** Findings carry locations, not cell contents. External
  targets are masked (directory and query string removed) before they are printed or
  written to JSON/SARIF. Author/company metadata is reported as "present" without its
  value.
- **Formula complexity.** Parsing is iterative where it matters (the dependency graph
  uses an iterative Tarjan SCC) and ranges larger than 20 000 cells are not expanded, so
  a hostile formula cannot force an unbounded allocation.

Out of scope: the accuracy of a spreadsheet's business logic. `xlsx-audit` is a static
linter, not an evaluator, and it cannot certify a financial model.

## Supply chain

- A single runtime dependency (`defusedxml`); Dependabot is enabled for pip and GitHub
  Actions.
- Releases are built from a tagged commit and published manually by the maintainer.
- No secrets, tokens or credentials are stored in this repository, and the test suite
  performs no network calls.
