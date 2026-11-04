"""Redaction of sensitive text before a finding leaves the process.

Findings are written to terminals, JSON files and SARIF logs that are often
uploaded to CI artefact stores or GitHub Code Scanning.  The audited workbook
itself is local, but the *report* travels, so it must not carry:

* local directory paths (``C:\\Users\\someone\\Budget.xlsx``),
* URL query strings (they routinely hold session tokens and API keys),
* e-mail addresses found in formulas,
* raw cell values picked up by a rule (for example the value of the cell just
  outside an aggregate range).

Masking happens in one place - :class:`xlsx_audit.models.Diagnostic` - so it
applies to *every* rule and every serialiser, not just the external-link rule
that used to mask its own targets.
"""

from __future__ import annotations

import posixpath
import re
from typing import Match, Optional
from urllib.parse import urlsplit

__all__ = ["mask_cell_value", "mask_formula", "mask_target", "mask_text"]

#: URLs and file URIs.  Stops at whitespace, quotes and formula separators.
_URL_RE = re.compile(r"(?:https?|ftp|file)://[^\s\"'<>]+", re.IGNORECASE)

#: ``[C:\dir\Book.xlsx]`` external-workbook references inside a formula.
_BRACKET_RE = re.compile(r"\[([^\[\]]+)\]")

#: Windows drive paths, absolute or quoted.
_DRIVE_RE = re.compile(r"(?<![\w])[A-Za-z]:[\\/][^\s\"'<>]+")

#: UNC paths (``\\server\share\book.xlsx``).
_UNC_RE = re.compile(r"\\\\[^\s\"'<>]+")

#: POSIX absolute paths with a file extension (``/home/user/Book.xlsx``).  The
#: extension is required so ordinary division (``A1/B1``) is never touched.
_POSIX_PATH_RE = re.compile(r"(?<![\w:/])/(?:[\w.@+-]+/)+[\w.@+-]+\.[A-Za-z0-9]{1,8}")

_EMAIL_RE = re.compile(r"([A-Za-z0-9._%+-]+)@([A-Za-z0-9.-]+\.[A-Za-z]{2,})")

_NUMBER_RE = re.compile(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?")


def mask_target(target: str) -> str:
    """Keep the useful tail of a target, drop local directories and query strings.

    ``C:\\Users\\example\\Budget.xlsx`` becomes ``…/Budget.xlsx`` and
    ``https://api.example/query?token=...`` becomes ``https://api.example/query``.
    """
    cleaned = target.strip()
    if not cleaned:
        return "<empty>"
    lowered = cleaned.lower()
    if lowered.startswith(("http://", "https://", "ftp://")):
        parts = urlsplit(cleaned)
        path = parts.path or "/"
        if len(path) > 60:
            tail = path[-59:]
            path = "…" + (tail if tail.startswith("/") else "/" + tail)
        return f"{parts.scheme}://{parts.hostname or '?'}{path}"
    if lowered.startswith("file:"):
        cleaned = cleaned[5:]
        cleaned = cleaned.lstrip("/")
        cleaned = cleaned.replace("\\", "/")
    normalised = cleaned.replace("\\", "/")
    name = posixpath.basename(normalised.rstrip("/"))
    if not name:
        return "…"
    directory = posixpath.dirname(normalised.rstrip("/"))
    if directory:
        return "…/" + name
    return name


def _looks_like_path(text: str) -> bool:
    lowered = text.lower()
    if lowered.startswith(("file:", "\\\\", "//")):
        return True
    if re.match(r"[A-Za-z]:[\\/]", text):
        return True
    return "\\" in text or "/" in text


def mask_text(text: str) -> str:
    """Mask paths, URLs, query strings and e-mail addresses inside free text."""
    if not text:
        return text
    masked = _URL_RE.sub(lambda match: mask_target(match.group(0)), text)
    masked = _DRIVE_RE.sub(lambda match: mask_target(match.group(0)), masked)
    masked = _UNC_RE.sub(lambda match: mask_target(match.group(0)), masked)
    masked = _POSIX_PATH_RE.sub(lambda match: mask_target(match.group(0)), masked)
    return _EMAIL_RE.sub(r"***@\2", masked)


def _mask_bracket(match: Match[str]) -> str:
    inner = match.group(1)
    if not _looks_like_path(inner):
        return match.group(0)
    return f"[{mask_target(inner)}]"


def mask_formula(formula: str) -> str:
    """Return ``formula`` with paths, URLs and e-mail addresses masked.

    Idempotent: masking an already masked formula leaves it unchanged, so the
    model and the serialisers can both call it without double-encoding.
    """
    if not formula:
        return formula
    masked = _BRACKET_RE.sub(_mask_bracket, formula)
    masked = _URL_RE.sub(lambda match: mask_target(match.group(0)), masked)
    masked = _DRIVE_RE.sub(lambda match: mask_target(match.group(0)), masked)
    masked = _UNC_RE.sub(lambda match: mask_target(match.group(0)), masked)
    masked = _POSIX_PATH_RE.sub(lambda match: mask_target(match.group(0)), masked)
    return _EMAIL_RE.sub(r"***@\2", masked)


def mask_cell_value(value: Optional[str]) -> str:
    """Describe a cell value without echoing it.

    Numbers, text, booleans and paths are reduced to a shape hint; the cell
    coordinate in the finding is enough to inspect the real value in Excel.
    """
    text = (value or "").strip()
    if not text:
        return ""
    if _looks_like_path(text) or _URL_RE.fullmatch(text):
        return mask_target(text)
    if _NUMBER_RE.fullmatch(text):
        return "<number>"
    if text.lower() in ("true", "false"):
        return "<boolean>"
    return f"<text:{len(text)}>"
