"""Audit orchestration: read once, run rules, return a report."""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, List, Sequence

from .i18n import normalize_lang
from .models import AuditReport, Diagnostic, Severity, Workbook, WorkbookStats
from .ooxml import read_workbook
from .rules import Context, build_rules

__all__ = ["PARSE_ERROR_CODE", "audit_workbook", "audit_workbook_object"]

PARSE_ERROR_CODE = "FORMULA_PARSE_ERROR"


def audit_workbook(
    path: str | Path,
    *,
    strict: bool = False,
    lang: str = "en",
    ignore: Iterable[str] = (),
    rules: Sequence[str] | None = None,
) -> AuditReport:
    """Audit one ``.xlsx`` file and return an :class:`AuditReport`.

    ``strict`` makes warnings fail the run as well.  ``ignore`` drops findings by
    code (for example ``"MAGIC_NUMBER_IN_FORMULA"``).  Raises
    :class:`~xlsx_audit.ooxml.WorkbookReadError` when the file cannot be read.
    """
    workbook = read_workbook(path)
    return audit_workbook_object(workbook, strict=strict, lang=lang, ignore=ignore, rules=rules)


def audit_workbook_object(
    workbook: Workbook,
    *,
    strict: bool = False,
    lang: str = "en",
    ignore: Iterable[str] = (),
    rules: Sequence[str] | None = None,
) -> AuditReport:
    """Audit an already-parsed :class:`Workbook` (useful for tests and tooling)."""
    selected = build_rules(rules)
    language = normalize_lang(lang)
    ctx = Context(workbook=workbook, lang=language)
    findings: List[Diagnostic] = []
    for rule in selected:
        findings.extend(rule.check(ctx))
    findings.extend(_parse_failures(ctx))

    ignored = {code.strip().upper() for code in ignore if code.strip()}
    if ignored:
        findings = [d for d in findings if d.code.upper() not in ignored]

    diagnostics = sorted(findings, key=lambda d: d.sort_key())
    stats = WorkbookStats(
        sheets=len(workbook.sheets),
        cells=len(workbook.cells),
        formulas=sum(1 for cell in workbook.cells.values() if cell.formula),
        hidden_sheets=sum(1 for sheet in workbook.sheets if sheet.is_hidden),
        external_links=len(workbook.external_links),
    )
    return AuditReport(
        path=workbook.path,
        diagnostics=diagnostics,
        stats=stats,
        rules=tuple(rule.name for rule in selected),
        strict=strict,
        lang=language,
    )


def _parse_failures(ctx: Context) -> List[Diagnostic]:
    seen: set[tuple[str, str, str]] = set()
    out: List[Diagnostic] = []
    for failure in ctx.failures:
        key = (failure.sheet, failure.cell, failure.formula)
        if key in seen:
            continue
        seen.add(key)
        out.append(
            Diagnostic(
                code=PARSE_ERROR_CODE,
                severity=Severity.WARNING,
                message_key="formula_parse_error",
                sheet=failure.sheet,
                cell=failure.cell,
                formula=failure.formula,
                rule="formula_parser",
                detail="static parser could not tokenize this formula; other rules skipped it",
            )
        )
    return out
