"""Core data model shared by the reader, the rules and the reporters."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, Iterable, Optional


class Severity(str, Enum):
    """How much a finding matters."""

    ERROR = "error"
    WARNING = "warning"
    INFO = "info"

    @property
    def rank(self) -> int:
        return _SEVERITY_RANK[self]

    @property
    def sarif_level(self) -> str:
        return {"error": "error", "warning": "warning", "info": "note"}[self.value]


_SEVERITY_RANK = {Severity.ERROR: 0, Severity.WARNING: 1, Severity.INFO: 2}


@dataclass(frozen=True)
class Diagnostic:
    """A single finding, always tied to a rule and (when applicable) a cell."""

    code: str
    severity: Severity
    message: str = ""
    sheet: Optional[str] = None
    cell: Optional[str] = None
    formula: Optional[str] = None
    value: Optional[str] = None
    rule: str = ""
    detail: Optional[str] = None
    message_key: str = ""
    params: Dict[str, str] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if not self.message and self.message_key:
            from .i18n import translator

            object.__setattr__(self, "message", translator("en")(self.message_key, **self.params))

    def text(self, lang: str = "en") -> str:
        """Localized message; falls back to the stored English text."""
        if self.message_key:
            from .i18n import translator

            return translator(lang)(self.message_key, **self.params)
        return self.message

    @property
    def location(self) -> str:
        if self.sheet and self.cell:
            return f"{self.sheet}!{self.cell}"
        if self.sheet:
            return self.sheet
        return "<workbook>"

    def sort_key(self) -> tuple[Any, ...]:
        return (
            self.severity.rank,
            self.sheet or "",
            _coord_sort_key(self.cell),
            self.code,
            self.message,
        )

    def to_dict(self, lang: str = "en") -> dict[str, Any]:
        out: dict[str, Any] = {
            "code": self.code,
            "severity": self.severity.value,
            "message": self.text(lang),
            "rule": self.rule,
            "location": self.location,
        }
        if self.sheet is not None:
            out["sheet"] = self.sheet
        if self.cell is not None:
            out["cell"] = self.cell
        if self.formula is not None:
            out["formula"] = self.formula
        if self.value is not None:
            out["value"] = self.value
        if self.detail is not None:
            out["detail"] = self.detail
        return out


def _coord_sort_key(cell: Optional[str]) -> tuple[int, int]:
    from .refs import parse_coordinate

    if not cell:
        return (0, 0)
    try:
        return parse_coordinate(cell)
    except ValueError:
        return (0, 0)


@dataclass(frozen=True)
class Cell:
    """A worksheet cell as it is stored in the OOXML part."""

    sheet: str
    coordinate: str
    row: int
    column: int
    formula: Optional[str] = None
    value: Optional[str] = None
    value_type: str = "n"
    style_id: int = 0

    @property
    def is_formula(self) -> bool:
        return self.formula is not None

    @property
    def is_numeric(self) -> bool:
        return self.value_type == "n" and self.value not in (None, "")


@dataclass(frozen=True)
class Sheet:
    """Worksheet metadata plus its cells, in document order."""

    name: str
    sheet_id: int
    state: str = "visible"  # visible | hidden | veryHidden
    cells: tuple[Cell, ...] = ()

    @property
    def is_hidden(self) -> bool:
        return self.state in ("hidden", "veryHidden")

    @property
    def is_very_hidden(self) -> bool:
        return self.state == "veryHidden"


@dataclass(frozen=True)
class DefinedName:
    name: str
    formula: str
    local_sheet_id: Optional[int] = None
    hidden: bool = False


@dataclass(frozen=True)
class ExternalLink:
    """A relationship target that points outside the workbook container."""

    part: str
    target: str
    kind: str = "externalLink"
    sheet: Optional[str] = None
    cell: Optional[str] = None

    @property
    def is_file_path(self) -> bool:
        lowered = self.target.lower()
        return lowered.startswith("file:") or (len(lowered) > 2 and lowered[1:3] in (":\\", ":/"))

    @property
    def is_url(self) -> bool:
        lowered = self.target.lower()
        return lowered.startswith(("http://", "https://", "ftp://"))


@dataclass
class Workbook:
    """Everything the auditor needs, read without Excel and without formulas being evaluated."""

    path: str
    sheets: list[Sheet] = field(default_factory=list)
    defined_names: list[DefinedName] = field(default_factory=list)
    external_links: list[ExternalLink] = field(default_factory=list)
    properties: dict[str, str] = field(default_factory=dict)

    _cells: Optional[dict[tuple[str, str], Cell]] = None
    _by_sheet: Optional[dict[str, dict[str, Cell]]] = None

    @property
    def cells(self) -> dict[tuple[str, str], Cell]:
        if self._cells is None:
            self._cells = {(c.sheet, c.coordinate): c for s in self.sheets for c in s.cells}
        return self._cells

    def sheet_cells(self, sheet_name: str) -> dict[str, Cell]:
        if self._by_sheet is None:
            self._by_sheet = {s.name: {c.coordinate: c for c in s.cells} for s in self.sheets}
        return self._by_sheet.get(sheet_name, {})

    @property
    def formulas(self) -> Iterable[Cell]:
        return (c for c in self.cells.values() if c.is_formula)

    def sheet_names(self) -> list[str]:
        return [s.name for s in self.sheets]


@dataclass(frozen=True)
class WorkbookStats:
    sheets: int = 0
    cells: int = 0
    formulas: int = 0
    hidden_sheets: int = 0
    external_links: int = 0

    def to_dict(self) -> dict[str, int]:
        return {
            "sheets": self.sheets,
            "cells": self.cells,
            "formulas": self.formulas,
            "hidden_sheets": self.hidden_sheets,
            "external_links": self.external_links,
        }


@dataclass
class AuditReport:
    """Result of auditing one workbook."""

    path: str
    diagnostics: list[Diagnostic] = field(default_factory=list)
    stats: WorkbookStats = field(default_factory=WorkbookStats)
    rules: tuple[str, ...] = ()
    strict: bool = False
    lang: str = "en"

    @property
    def errors(self) -> list[Diagnostic]:
        return [d for d in self.diagnostics if d.severity is Severity.ERROR]

    @property
    def warnings(self) -> list[Diagnostic]:
        return [d for d in self.diagnostics if d.severity is Severity.WARNING]

    @property
    def infos(self) -> list[Diagnostic]:
        return [d for d in self.diagnostics if d.severity is Severity.INFO]

    @property
    def passed(self) -> bool:
        """True when nothing blocks a release.

        In default mode warnings are advisory; with ``strict`` any finding fails.
        """
        if self.strict:
            return not self.diagnostics
        return not self.errors

    @property
    def status(self) -> str:
        return "pass" if self.passed else "fail"

    @property
    def file_name(self) -> str:
        import os

        return os.path.basename(self.path)

    def counts_by_code(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for diag in self.diagnostics:
            counts[diag.code] = counts.get(diag.code, 0) + 1
        return dict(sorted(counts.items()))

    def to_dict(self, lang: str = "en") -> dict[str, Any]:
        return {
            "tool": "xlsx-audit",
            "file": self.path,
            "status": self.status,
            "strict": self.strict,
            "passed": self.passed,
            "summary": {
                "errors": len(self.errors),
                "warnings": len(self.warnings),
                "infos": len(self.infos),
                "by_code": self.counts_by_code(),
            },
            "stats": self.stats.to_dict(),
            "rules": list(self.rules),
            "diagnostics": [d.to_dict(lang) for d in self.diagnostics],
        }
