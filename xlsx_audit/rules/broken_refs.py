"""Detect cells whose formulas are (or produce) Excel error values.

``#REF!`` is special: it means a reference was destroyed, usually by deleting a
row/column, and the formula silently keeps returning an error until somebody
notices.  The rule reports the exact ``Sheet!Cell`` so a reviewer can jump
straight there - the same contract Excel's own error checker gives.
"""

from __future__ import annotations

from typing import List

from ..formula_parser import ErrorLit, iter_nodes
from ..models import Cell, Diagnostic, Severity
from .base import Context, Rule, dedupe

__all__ = ["ERROR_CODES", "BrokenRefsRule"]

ERROR_CODES = {
    "#REF!",
    "#VALUE!",
    "#DIV/0!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#NULL!",
    "#SPILL!",
    "#CALC!",
    "#FIELD!",
    "#BLOCKED!",
    "#CONNECT!",
    "#UNKNOWN!",
    "#GETTING_DATA",
}


class BrokenRefsRule(Rule):
    name = "broken_refs"
    description = "Broken references (#REF!) and Excel error values in cells"

    def check(self, ctx: Context) -> List[Diagnostic]:
        findings: List[Diagnostic] = []
        for cell in sorted(ctx.workbook.cells.values(), key=lambda c: (c.sheet, c.row, c.column)):
            findings.extend(self._cached_error(ctx, cell))
            if cell.formula:
                findings.extend(self._formula_errors(ctx, cell))
        return dedupe(findings)

    def _cached_error(self, ctx: Context, cell: Cell) -> List[Diagnostic]:
        value = (cell.value or "").strip().upper()
        if value not in ERROR_CODES:
            return []
        code = "BROKEN_REFERENCE" if value == "#REF!" else "FORMULA_ERROR_VALUE"
        severity = Severity.ERROR
        return [
            self.diagnostic(
                code=code,
                severity=severity,
                message_key="broken_reference" if code == "BROKEN_REFERENCE" else "error_value",
                params={"code": value},
                sheet=cell.sheet,
                cell=cell.coordinate,
                formula=cell.formula,
                value=value,
            )
        ]

    def _formula_errors(self, ctx: Context, cell: Cell) -> List[Diagnostic]:
        tree = ctx.parse(cell)
        if tree is None:
            return []
        findings: List[Diagnostic] = []
        seen: set[str] = set()
        for node in iter_nodes(tree):
            if not isinstance(node, ErrorLit):
                continue
            error_code = node.code.upper()
            if error_code in seen:
                continue
            seen.add(error_code)
            if error_code == "#REF!":
                findings.append(
                    self.diagnostic(
                        code="BROKEN_REFERENCE",
                        severity=Severity.ERROR,
                        message_key="broken_reference",
                        params={"code": error_code},
                        sheet=cell.sheet,
                        cell=cell.coordinate,
                        formula=cell.formula,
                        value=error_code,
                    )
                )
            elif error_code in ERROR_CODES:
                findings.append(
                    self.diagnostic(
                        code="FORMULA_ERROR_VALUE",
                        severity=Severity.ERROR,
                        message_key="error_value",
                        params={"code": error_code},
                        sheet=cell.sheet,
                        cell=cell.coordinate,
                        formula=cell.formula,
                        value=error_code,
                    )
                )
        return findings
