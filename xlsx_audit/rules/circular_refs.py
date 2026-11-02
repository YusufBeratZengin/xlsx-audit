"""Circular reference rule built on the dependency graph."""

from __future__ import annotations

from typing import List

from ..graph_audit import find_cycles
from ..models import Diagnostic, Severity
from .base import Context, Rule

__all__ = ["CircularRefsRule"]


class CircularRefsRule(Rule):
    name = "circular_references"
    description = "Circular references across one or more sheets"

    def check(self, ctx: Context) -> List[Diagnostic]:
        try:
            cycles = find_cycles(ctx.workbook)
        except ValueError as exc:
            return [
                self.diagnostic(
                    code="GRAPH_LIMIT_EXCEEDED",
                    severity=Severity.INFO,
                    message=f"Circular-reference analysis skipped: {exc}",
                    detail=str(exc),
                )
            ]
        findings: List[Diagnostic] = []
        for cycle in cycles:
            sheet, cell = cycle.nodes[0]
            findings.append(
                self.diagnostic(
                    code="CIRCULAR_REFERENCE",
                    severity=Severity.ERROR,
                    message_key="circular_reference",
                    params={"path": cycle.path_text()},
                    sheet=sheet,
                    cell=cell,
                    detail=cycle.path_text(),
                )
            )
        return findings
