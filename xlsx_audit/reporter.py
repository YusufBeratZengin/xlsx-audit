"""Human, JSON and Markdown renderers.

The terminal renderer uses plain ANSI escapes instead of a colour library: the
package intentionally has a single runtime dependency.  ``NO_COLOR`` and
``--no-color`` both disable it.
"""

from __future__ import annotations

import json
import os
from typing import Dict, List, Optional

from .i18n import normalize_lang
from .models import AuditReport, Diagnostic, Severity

__all__ = [
    "SEVERITY_LABELS",
    "render_json",
    "render_markdown",
    "render_report",
    "render_text",
    "use_color",
]

_COLORS: Dict[str, str] = {
    "reset": "\x1b[0m",
    "bold": "\x1b[1m",
    "dim": "\x1b[2m",
    "red": "\x1b[31m",
    "yellow": "\x1b[33m",
    "blue": "\x1b[34m",
    "green": "\x1b[32m",
}

SEVERITY_LABELS = {
    "en": {
        Severity.ERROR: "ERROR",
        Severity.WARNING: "WARNING",
        Severity.INFO: "INFO",
    },
    "tr": {
        Severity.ERROR: "HATA",
        Severity.WARNING: "UYARI",
        Severity.INFO: "BİLGİ",
    },
}

_SEVERITY_COLORS = {
    Severity.ERROR: "red",
    Severity.WARNING: "yellow",
    Severity.INFO: "blue",
}

_HEADLINES = {
    "en": {
        "fail": "✖ [FAIL] {name} audit failed ({errors}, {warnings})",
        "pass": "✔ [PASS] {name} audit passed ({errors}, {warnings})",
        "strict": " [strict]",
        "summary": "{count} findings · {sheets} sheets · {cells} cells · {formulas} formulas",
        "no_findings": "No findings.",
        "errors": "{n} error{s}",
        "warnings": "{n} warning{s}",
    },
    "tr": {
        "fail": "✖ [FAIL] {name} denetimi başarısız ({errors}, {warnings})",
        "pass": "✔ [PASS] {name} denetimi başarılı ({errors}, {warnings})",
        "strict": " [katı]",
        "summary": "{count} bulgu · {sheets} sayfa · {cells} hücre · {formulas} formül",
        "no_findings": "Bulgu yok.",
        "errors": "{n} Hata",
        "warnings": "{n} Uyarı",
    },
}


def use_color(color: bool | None = None) -> bool:
    """Resolve ``--color/--no-color`` against ``NO_COLOR`` and TTY state."""
    if color is not None:
        return color and not os.environ.get("NO_COLOR")
    if os.environ.get("NO_COLOR"):
        return False
    try:
        return os.isatty(1)
    except OSError:  # pragma: no cover - exotic streams
        return False


def _paint(text: str, color: bool, *styles: str) -> str:
    if not color:
        return text
    prefix = "".join(_COLORS[style] for style in styles)
    return f"{prefix}{text}{_COLORS['reset']}"


def _count_word(count: int, lang: str, kind: str) -> str:
    words = _HEADLINES[lang]
    if lang == "tr":
        return words[kind].format(n=count)
    suffix = "" if count == 1 else "s"
    return words[kind].format(n=count, s=suffix)


def render_text(
    report: AuditReport,
    *,
    lang: Optional[str] = None,
    color: bool = False,
    verbose: bool = False,
) -> str:
    """Render one report for a terminal."""
    language = normalize_lang(lang or report.lang)
    words = _HEADLINES[language]
    labels = SEVERITY_LABELS[language]
    name = report.file_name
    headline = words["pass" if report.passed else "fail"].format(
        name=name,
        errors=_count_word(len(report.errors), language, "errors"),
        warnings=_count_word(len(report.warnings), language, "warnings"),
    )
    if report.strict:
        headline += words["strict"]
    lines: List[str] = [
        _paint(headline, color, "bold", "green" if report.passed else "red"),
        "",
    ]
    shown = [d for d in report.diagnostics if verbose or d.severity is not Severity.INFO]
    if not shown:
        lines.append(_paint(words["no_findings"], color, "dim"))
    for diagnostic in shown:
        lines.append(_format_diagnostic(diagnostic, labels, color, language))
    lines.append("")
    lines.append(
        _paint(
            words["summary"].format(
                count=len(report.diagnostics),
                sheets=report.stats.sheets,
                cells=report.stats.cells,
                formulas=report.stats.formulas,
            ),
            color,
            "dim",
        )
    )
    return "\n".join(lines)


def _format_diagnostic(
    diagnostic: Diagnostic, labels: Dict[Severity, str], color: bool, lang: str = "en"
) -> str:
    label = labels[diagnostic.severity]
    painted_label = _paint(f"[{label}]", color, "bold", _SEVERITY_COLORS[diagnostic.severity])
    location = _paint(diagnostic.location, color, "bold")
    code = _paint(f"({diagnostic.code})", color, "dim")
    return f"- {painted_label} {location}: {diagnostic.text(lang)} {code}".rstrip()


def render_json(report: AuditReport, *, lang: Optional[str] = None) -> str:
    language = normalize_lang(lang or report.lang)
    return json.dumps(report.to_dict(language), ensure_ascii=False, indent=2) + "\n"


def render_markdown(
    report: AuditReport, *, lang: Optional[str] = None, verbose: bool = True
) -> str:
    language = normalize_lang(lang or report.lang)
    status = "PASS" if report.passed else "FAIL"
    strict = " (strict)" if report.strict else ""
    lines = [
        f"# xlsx-audit report — {report.file_name}",
        "",
        f"**Status:** {status}{strict}  ",
        f"**Errors:** {len(report.errors)} · **Warnings:** {len(report.warnings)} "
        f"· **Info:** {len(report.infos)}  ",
        f"**Sheets:** {report.stats.sheets} · **Cells:** {report.stats.cells} "
        f"· **Formulas:** {report.stats.formulas} · **External links:** "
        f"{report.stats.external_links}",
        "",
        "| Severity | Location | Code | Message |",
        "| --- | --- | --- | --- |",
    ]
    shown = [d for d in report.diagnostics if verbose or d.severity is not Severity.INFO]
    if not shown:
        lines.append("| - | - | - | No findings. |")
    for diagnostic in shown:
        label = SEVERITY_LABELS[language][diagnostic.severity]
        message = diagnostic.text(language).replace("|", "\\|").replace("\n", " ")
        location = diagnostic.location.replace("|", "\\|")
        lines.append(f"| {label} | `{location}` | `{diagnostic.code}` | {message} |")
    lines.append("")
    lines.append(
        "_Generated by [xlsx-audit](https://github.com/YusufBeratZengin/xlsx-audit); "
        "static analysis only, not a substitute for professional review._"
    )
    lines.append("")
    return "\n".join(lines)


def render_report(
    report: AuditReport,
    fmt: str = "text",
    *,
    lang: Optional[str] = None,
    color: bool = False,
    verbose: bool = False,
) -> str:
    """Dispatch to the renderer named by ``fmt`` (text, json, markdown, sarif)."""
    if fmt == "text":
        return render_text(report, lang=lang, color=color, verbose=verbose)
    if fmt == "json":
        return render_json(report, lang=lang)
    if fmt == "markdown":
        return render_markdown(report, lang=lang, verbose=True)
    if fmt == "sarif":
        from .sarif_reporter import render_sarif

        return render_sarif(report)
    raise ValueError(f"unknown report format: {fmt}")
