"""Find constants baked into formulas.

``=A1*1.20`` hides a business rule (a VAT rate, a commission, a growth
assumption) inside a formula.  When the rate changes nobody can find every
copy, and reviewers cannot see the assumption.  This rule flags literals that
are *not* structural (0, 1, 2 and 100 are allowed) and not a function's
positional/index argument.
"""

from __future__ import annotations

from typing import Dict, FrozenSet, Iterator, List, Tuple

from ..formula_parser import (
    Array,
    BinaryOp,
    FuncCall,
    Node,
    NumberLit,
    Percent,
    UnaryOp,
)
from ..models import Diagnostic, Severity
from .base import Context, Rule, dedupe

__all__ = ["DEFAULT_ALLOWED", "MagicNumbersRule"]

#: Literals that carry no business meaning.  0/1 are counters and flags, 2 is a
#: halving factor, 100 converts a fraction to percent.
DEFAULT_ALLOWED: FrozenSet[float] = frozenset({0.0, 1.0, 2.0, 100.0})

#: (function, zero-based argument index) pairs whose numbers are indices or
#: precisions rather than assumptions.
SKIP_ARGUMENTS: Dict[str, FrozenSet[int]] = {
    "CHOOSE": frozenset({0}),
    "HLOOKUP": frozenset({2}),
    "INDEX": frozenset({1, 2, 3}),
    "LARGE": frozenset({1}),
    "MATCH": frozenset({1, 2}),
    "RANK": frozenset({1}),
    "ROUND": frozenset({1}),
    "ROUNDDOWN": frozenset({1}),
    "ROUNDUP": frozenset({1}),
    "SMALL": frozenset({1}),
    "TEXT": frozenset({1}),
    "TRUNC": frozenset({1}),
    "VLOOKUP": frozenset({1, 2}),
}


def _format(value: float) -> str:
    if value == int(value) and abs(value) < 1e15:
        return str(int(value))
    return repr(value)


class MagicNumbersRule(Rule):
    name = "magic_numbers"
    description = "Numeric constants embedded inside formulas"

    def __init__(self, allowed: FrozenSet[float] = DEFAULT_ALLOWED) -> None:
        super().__init__()
        self.allowed = allowed

    def check(self, ctx: Context) -> List[Diagnostic]:
        findings: List[Diagnostic] = []
        for cell in ctx.formula_cells():
            tree = ctx.parse(cell)
            if tree is None:
                continue
            reported: set[float] = set()
            for literal, skipped in _literals(tree):
                if skipped or literal in self.allowed or literal in reported:
                    continue
                reported.add(literal)
                findings.append(
                    self.diagnostic(
                        code="MAGIC_NUMBER_IN_FORMULA",
                        severity=Severity.WARNING,
                        message_key="magic_number",
                        params={"value": _format(literal)},
                        sheet=cell.sheet,
                        cell=cell.coordinate,
                        formula=cell.formula,
                        value=_format(literal),
                    )
                )
        return dedupe(findings)


def _literals(node: Node, skip: bool = False) -> Iterator[Tuple[float, bool]]:
    if isinstance(node, NumberLit):
        yield node.value, skip
        return
    if isinstance(node, UnaryOp):
        if node.op == "-" and isinstance(node.operand, NumberLit):
            yield -node.operand.value, skip
            return
        yield from _literals(node.operand, skip)
        return
    if isinstance(node, FuncCall):
        skipped = SKIP_ARGUMENTS.get(node.name, frozenset())
        for index, arg in enumerate(node.args):
            yield from _literals(arg, skip or index in skipped)
        return
    if isinstance(node, BinaryOp):
        yield from _literals(node.left, skip)
        yield from _literals(node.right, skip)
        return
    if isinstance(node, Percent):
        yield from _literals(node.operand, skip)
        return
    if isinstance(node, Array):
        for row in node.rows:
            for item in row:
                yield from _literals(item, skip)
