"""SARIF 2.1.0 output for GitHub Code Scanning and other SARIF consumers."""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from typing import Any, Dict, List, Tuple

from . import __version__
from .models import AuditReport, Diagnostic
from .redact import mask_formula, mask_text

__all__ = ["SARIF_RULES", "render_sarif", "sarif_document"]

SARIF_VERSION = "2.1.0"
SARIF_SCHEMA = "https://json.schemastore.org/sarif-2.1.0.json"
TOOL_URI = "https://github.com/YusufBeratZengin/xlsx-audit"

#: code -> (short description, default level)
SARIF_RULES: Dict[str, Tuple[str, str]] = {
    "BROKEN_REFERENCE": ("A formula or cell carries the #REF! error", "error"),
    "FORMULA_ERROR_VALUE": ("A cell carries an Excel error value", "error"),
    "FORMULA_PARSE_ERROR": ("A formula could not be parsed statically", "warning"),
    "MAGIC_NUMBER_IN_FORMULA": ("A numeric constant is embedded in a formula", "warning"),
    "EXTERNAL_LINK": ("The workbook links to an external workbook", "warning"),
    "EXTERNAL_URL": ("A formula reaches a network resource", "error"),
    "VERY_HIDDEN_SHEET": ("The workbook contains a veryHidden sheet", "warning"),
    "HIDDEN_SHEET": ("The workbook contains a hidden sheet", "note"),
    "HIDDEN_DEFINED_NAME": ("A defined name is hidden", "note"),
    "METADATA_PRESENT": ("The workbook carries document metadata", "note"),
    "CIRCULAR_REFERENCE": ("Formulas form a circular reference", "error"),
    "OMITTED_CELL_WARNING": ("An aggregate range stops before a filled cell", "warning"),
    "GRAPH_LIMIT_EXCEEDED": ("Circular-reference analysis was skipped", "note"),
}


def _artifact_uri(path: str) -> str:
    candidate = Path(path)
    try:
        relative = candidate.resolve().relative_to(Path.cwd().resolve())
        return relative.as_posix()
    except (ValueError, OSError):
        return candidate.name if candidate.name else os.path.basename(path)


def _fingerprint(diagnostic: Diagnostic) -> str:
    raw = "|".join(
        [
            diagnostic.code,
            diagnostic.location,
            diagnostic.message,
            diagnostic.value or "",
        ]
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:32]


def _result(diagnostic: Diagnostic, rule_index: Dict[str, int], uri: str) -> Dict[str, Any]:
    physical: Dict[str, Any] = {
        "artifactLocation": {"uri": uri},
        # A cell coordinate has no line in the XML part, so anchor the result at
        # the top of the file and expose the exact cell as a logical location.
        "region": {"startLine": 1, "startColumn": 1},
    }
    logical = {
        "name": diagnostic.cell or diagnostic.sheet or "workbook",
        "fullyQualifiedName": diagnostic.location,
        "kind": "member" if diagnostic.cell else "namespace",
    }
    properties: Dict[str, Any] = {"code": diagnostic.code, "rule": diagnostic.rule}
    if diagnostic.sheet:
        properties["sheet"] = diagnostic.sheet
    if diagnostic.cell:
        properties["cell"] = diagnostic.cell
    if diagnostic.value:
        properties["value"] = mask_text(diagnostic.value)
    if diagnostic.formula:
        properties["formula"] = mask_formula(diagnostic.formula)
    return {
        "ruleId": diagnostic.code,
        "ruleIndex": rule_index[diagnostic.code],
        "level": diagnostic.severity.sarif_level,
        "message": {"text": diagnostic.text("en")},
        "locations": [
            {
                "physicalLocation": physical,
                "logicalLocations": [logical],
            }
        ],
        "partialFingerprints": {"xlsxAudit/v1": _fingerprint(diagnostic)},
        "properties": properties,
    }


def sarif_document(reports: List[AuditReport]) -> Dict[str, Any]:
    all_codes: List[str] = []
    for report in reports:
        for diagnostic in report.diagnostics:
            if diagnostic.code not in all_codes:
                all_codes.append(diagnostic.code)
    rules = [
        {
            "id": code,
            "name": code.title().replace("_", ""),
            "shortDescription": {"text": SARIF_RULES.get(code, (code, "warning"))[0]},
            "helpUri": f"{TOOL_URI}#rule-{code.lower()}",
            "defaultConfiguration": {"level": SARIF_RULES.get(code, (code, "warning"))[1]},
        }
        for code in all_codes
    ]
    rule_index = {code: index for index, code in enumerate(all_codes)}
    runs: List[Dict[str, Any]] = []
    for report in reports:
        uri = _artifact_uri(report.path)
        runs.append(
            {
                "tool": {
                    "driver": {
                        "name": "xlsx-audit",
                        "version": __version__,
                        "informationUri": TOOL_URI,
                        "rules": rules,
                    }
                },
                "automationDetails": {"id": f"xlsx-audit/{report.file_name}"},
                "artifactLocation": {"uri": uri},
                "results": [
                    _result(diagnostic, rule_index, uri) for diagnostic in report.diagnostics
                ],
                "properties": {
                    "status": report.status,
                    "strict": report.strict,
                    "errors": len(report.errors),
                    "warnings": len(report.warnings),
                },
            }
        )
    return {
        "$schema": SARIF_SCHEMA,
        "version": SARIF_VERSION,
        "runs": runs,
    }


def render_sarif(report: AuditReport) -> str:
    return json.dumps(sarif_document([report]), ensure_ascii=False, indent=2) + "\n"
