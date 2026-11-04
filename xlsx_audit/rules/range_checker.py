"""Spot ``=SUM(A1:A9)`` when ``A10`` is already filled.

This is the spreadsheet version of an off-by-one bug: a row is inserted at the
bottom of a block and the totals formula keeps its old extent, so the last row
silently drops out of every total.  The check is deliberately conservative: it
only fires when the cell immediately after a multi-cell range exists, is
numeric, and is not referenced anywhere else in the same formula.
"""

from __future__ import annotations

from typing import List, Optional, Set, Tuple

from ..formula_parser import FuncCall, Node, Range, Ref, iter_nodes
from ..models import Cell, Diagnostic, Severity
from ..redact import mask_cell_value
from ..refs import format_coordinate
from .base import Context, Rule

__all__ = ["RangeCheckerRule"]

AGGREGATE_FUNCTIONS = {
    "SUM",
    "SUMPRODUCT",
    "AVERAGE",
    "AVERAGEIF",
    "MIN",
    "MAX",
    "COUNT",
    "COUNTA",
    "PRODUCT",
    "MEDIAN",
    "STDEV",
    "STDEVP",
    "VAR",
    "VARP",
}


class RangeCheckerRule(Rule):
    name = "range_checker"
    description = "Aggregate ranges that stop before a filled neighbouring cell"

    def check(self, ctx: Context) -> List[Diagnostic]:
        findings: List[Diagnostic] = []
        for cell in ctx.formula_cells():
            tree = ctx.parse(cell)
            if tree is None:
                continue
            referenced = _direct_references(tree, cell.sheet)
            for call in iter_nodes(tree):
                if not isinstance(call, FuncCall) or call.name not in AGGREGATE_FUNCTIONS:
                    continue
                for arg in call.args:
                    if not isinstance(arg, Range):
                        continue
                    finding = self._check_range(ctx, cell, arg, referenced)
                    if finding is not None:
                        findings.append(finding)
        return findings

    def _check_range(
        self,
        ctx: Context,
        cell: Cell,
        node: Range,
        referenced: Set[Tuple[str, str]],
    ) -> Optional[Diagnostic]:
        start, end = node.start, node.end
        if start.row is None or start.column is None or end.row is None or end.column is None:
            return None
        if start.book is not None or end.book is not None:
            return None
        sheet_name = start.sheet or end.sheet or cell.sheet
        top, bottom = sorted((start.row, end.row))
        left, right = sorted((start.column, end.column))
        candidate: Optional[Tuple[int, int]] = None
        if left == right and bottom - top + 1 >= 2:
            candidate = (bottom + 1, left)
        elif top == bottom and right - left + 1 >= 2:
            candidate = (top, right + 1)
        if candidate is None:
            return None
        row, column = candidate
        coordinate = format_coordinate(row, column)
        if (sheet_name, coordinate) in referenced:
            return None
        target = ctx.workbook.sheet_cells(sheet_name).get(coordinate)
        if target is None or not target.is_numeric or target.formula is not None:
            return None
        if sheet_name == cell.sheet and coordinate == cell.coordinate:
            return None
        bounds = f"{format_coordinate(top, left)}:{format_coordinate(bottom, right)}"
        return self.diagnostic(
            code="OMITTED_CELL_WARNING",
            severity=Severity.WARNING,
            message_key="omitted_cell",
            params={
                "range": f"{sheet_name}!{bounds}",
                "cell": coordinate,
                "value": mask_cell_value(target.value),
            },
            sheet=cell.sheet,
            cell=cell.coordinate,
            formula=cell.formula,
            value=coordinate,
            detail=f"aggregate={bounds} candidate={sheet_name}!{coordinate}",
        )


def _direct_references(tree: Node, sheet: str) -> Set[Tuple[str, str]]:
    refs: Set[Tuple[str, str]] = set()
    for node in iter_nodes(tree):
        if isinstance(node, Ref) and node.row is not None and node.column is not None:
            refs.add((node.sheet or sheet, format_coordinate(node.row, node.column)))
        elif isinstance(node, Range):
            start, end = node.start, node.end
            if start.row is None or start.column is None or end.row is None or end.column is None:
                continue
            name = start.sheet or end.sheet or sheet
            top, bottom = sorted((start.row, end.row))
            left, right = sorted((start.column, end.column))
            if (bottom - top + 1) * (right - left + 1) > 10_000:
                continue
            for row in range(top, bottom + 1):
                for column in range(left, right + 1):
                    refs.add((name, format_coordinate(row, column)))
    return refs
