"""Rule protocol and the per-run context shared by all rules."""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Callable, Dict, Iterable, List, Optional, Tuple

from ..formula_parser import FormulaSyntaxError, Node, parse_formula
from ..i18n import translator
from ..models import Cell, Diagnostic, Severity, Workbook

__all__ = ["Context", "ParseFailure", "Rule"]


@dataclass(frozen=True)
class ParseFailure:
    sheet: str
    cell: str
    formula: str


@dataclass
class Context:
    """Caches parsed formulas so several rules pay for parsing only once."""

    workbook: Workbook
    lang: str = "en"
    _cache: Dict[str, Optional[Node]] = field(default_factory=dict, repr=False)
    failures: List[ParseFailure] = field(default_factory=list, repr=False)
    t: Callable[..., str] = field(init=False)

    def __post_init__(self) -> None:
        self.t = translator(self.lang)

    def parse(self, cell: Cell) -> Optional[Node]:
        """Parsed AST for ``cell`` or ``None`` when the formula is unparseable."""
        formula = cell.formula
        if formula is None:
            return None
        if formula in self._cache:
            return self._cache[formula]
        try:
            tree: Optional[Node] = parse_formula(formula)
        except FormulaSyntaxError:
            tree = None
            self.failures.append(
                ParseFailure(sheet=cell.sheet, cell=cell.coordinate, formula=formula)
            )
        self._cache[formula] = tree
        return tree

    def formula_cells(self) -> Iterable[Cell]:
        return (cell for cell in self.workbook.cells.values() if cell.formula)


class Rule(ABC):
    """A single audit rule."""

    name: str = ""
    description: str = ""

    def __init__(self) -> None:
        if not self.name:  # pragma: no cover - guard for subclasses
            raise ValueError("rule must define a name")

    @abstractmethod
    def check(self, ctx: Context) -> List[Diagnostic]:
        """Return the findings for one workbook."""

    def diagnostic(
        self,
        code: str,
        severity: Severity,
        message: Optional[str] = None,
        sheet: Optional[str] = None,
        cell: Optional[str] = None,
        formula: Optional[str] = None,
        value: Optional[str] = None,
        detail: Optional[str] = None,
        message_key: Optional[str] = None,
        params: Optional[Dict[str, str]] = None,
    ) -> Diagnostic:
        return Diagnostic(
            code=code,
            severity=severity,
            message=message or "",
            sheet=sheet,
            cell=cell,
            formula=formula,
            value=value,
            rule=self.name,
            detail=detail,
            message_key=message_key or "",
            params=params or {},
        )


def dedupe(diagnostics: Iterable[Diagnostic]) -> List[Diagnostic]:
    """Drop exact duplicates while keeping the first occurrence order."""
    seen: set[Tuple[object, ...]] = set()
    out: List[Diagnostic] = []
    for item in diagnostics:
        key = (item.code, item.sheet, item.cell, item.message, item.value)
        if key in seen:
            continue
        seen.add(key)
        out.append(item)
    return out
