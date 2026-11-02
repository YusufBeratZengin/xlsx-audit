"""xlsx-audit - static auditing for Excel workbooks.

The public API is one call::

    from xlsx_audit import audit_workbook

    report = audit_workbook("data.xlsx")
    assert report.passed, report.diagnostics
"""

from __future__ import annotations

from .audit import audit_workbook, audit_workbook_object
from .models import AuditReport, Diagnostic, Severity, Workbook
from .rules import rule_names

__version__ = "0.1.0"

__all__ = [
    "AuditReport",
    "Diagnostic",
    "Severity",
    "Workbook",
    "__version__",
    "audit_workbook",
    "audit_workbook_object",
    "rule_names",
]
