"""Hidden sheets and document metadata.

A ``veryHidden`` sheet is not reachable from the Excel UI at all: whole models,
scratch calculations or (in badly sanitised exports) source data can travel with
a workbook that looks clean.  Hidden and very-hidden states are reported with
different severities, and author/company metadata is flagged without echoing the
values, because names are personal data.
"""

from __future__ import annotations

from typing import List

from ..models import Diagnostic, Severity
from .base import Context, Rule

__all__ = ["HiddenSheetsRule"]

_METADATA_KEYS = ("creator", "lastmodifiedby", "company", "manager", "title", "subject")


class HiddenSheetsRule(Rule):
    name = "hidden_sheets"
    description = "veryHidden worksheets, hidden sheets and document metadata"

    def check(self, ctx: Context) -> List[Diagnostic]:
        findings: List[Diagnostic] = []
        for sheet in ctx.workbook.sheets:
            if sheet.is_very_hidden:
                findings.append(
                    self.diagnostic(
                        code="VERY_HIDDEN_SHEET",
                        severity=Severity.WARNING,
                        message_key="very_hidden_sheet",
                        sheet=sheet.name,
                        detail="state=veryHidden",
                    )
                )
            elif sheet.state == "hidden":
                findings.append(
                    self.diagnostic(
                        code="HIDDEN_SHEET",
                        severity=Severity.INFO,
                        message_key="hidden_sheet",
                        sheet=sheet.name,
                        detail="state=hidden",
                    )
                )
        for defined in ctx.workbook.defined_names:
            if defined.hidden:
                findings.append(
                    self.diagnostic(
                        code="HIDDEN_DEFINED_NAME",
                        severity=Severity.INFO,
                        message_key="hidden_defined_name",
                        params={"name": defined.name},
                        sheet=None,
                        value=defined.name,
                    )
                )
        keys = sorted(k for k in ctx.workbook.properties if k.lower() in _METADATA_KEYS)
        if keys:
            findings.append(
                self.diagnostic(
                    code="METADATA_PRESENT",
                    severity=Severity.INFO,
                    message_key="metadata_present",
                    params={"keys": ", ".join(keys)},
                    detail="metadata values are not printed (possible personal data)",
                )
            )
        return findings
