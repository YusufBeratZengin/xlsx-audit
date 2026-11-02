"""Surface external workbook links and network egress hidden in formulas.

Two different leaks matter:

* a saved *link* (``xl/externalLinks/*``) that keeps pointing at somebody's
  ``C:\\Users\\...\\Budget.xlsx`` after the file is mailed around;
* a *formula* that reaches out at calculation time (``[1]Sheet1!A1``,
  ``WEBSERVICE("https://...")``), which can exfiltrate cell values.

Messages deliberately show a masked target: the local directory and any URL
query string are removed before the finding is printed or written to JSON.
"""

from __future__ import annotations

import posixpath
from typing import List
from urllib.parse import urlsplit

from ..formula_parser import FuncCall, Node, Ref, StringLit, iter_nodes
from ..models import Cell, Diagnostic, Severity
from .base import Context, Rule, dedupe

__all__ = ["ExternalLinksRule", "mask_target"]

NETWORK_FUNCTIONS = {"WEBSERVICE", "FILTERXML", "RTD", "STOCKHISTORY"}


def mask_target(target: str) -> str:
    """Keep the useful tail of a target, drop local directories and query strings."""
    cleaned = target.strip()
    if not cleaned:
        return "<empty>"
    lowered = cleaned.lower()
    if lowered.startswith(("http://", "https://", "ftp://")):
        parts = urlsplit(cleaned)
        path = parts.path or "/"
        if len(path) > 60:
            path = "…" + path[-59:]
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


class ExternalLinksRule(Rule):
    name = "external_links"
    description = "External workbook links, local paths and network calls"

    def check(self, ctx: Context) -> List[Diagnostic]:
        findings: List[Diagnostic] = []
        findings.extend(self._saved_links(ctx))
        for cell in ctx.formula_cells():
            tree = ctx.parse(cell)
            if tree is None:
                continue
            findings.extend(self._formula_links(ctx, cell, tree))
        return dedupe(findings)

    def _saved_links(self, ctx: Context) -> List[Diagnostic]:
        findings: List[Diagnostic] = []
        for link in ctx.workbook.external_links:
            target = mask_target(link.target)
            is_url = link.is_url
            findings.append(
                self.diagnostic(
                    code="EXTERNAL_URL" if is_url else "EXTERNAL_LINK",
                    severity=Severity.ERROR if is_url else Severity.WARNING,
                    message_key="external_url" if is_url else "external_link",
                    params={"target": target},
                    detail=f"part={link.part}",
                )
            )
        return findings

    def _formula_links(self, ctx: Context, cell: Cell, tree: Node) -> List[Diagnostic]:
        findings: List[Diagnostic] = []
        books: set[str] = set()
        urls: set[str] = set()
        network_calls: set[str] = set()
        for node in iter_nodes(tree):
            if isinstance(node, Ref) and node.book:
                books.add(node.book)
            elif isinstance(node, StringLit) and node.value.lower().startswith(
                ("http://", "https://", "ftp://")
            ):
                urls.add(node.value)
            elif isinstance(node, FuncCall) and node.name in NETWORK_FUNCTIONS:
                network_calls.add(node.name)
        for book in sorted(books):
            findings.append(
                self.diagnostic(
                    code="EXTERNAL_LINK",
                    severity=Severity.WARNING,
                    message_key="external_link_cell",
                    params={"book": book},
                    sheet=cell.sheet,
                    cell=cell.coordinate,
                    formula=cell.formula,
                    value=book,
                )
            )
        for url in sorted(urls):
            findings.append(
                self.diagnostic(
                    code="EXTERNAL_URL",
                    severity=Severity.ERROR,
                    message_key="external_url",
                    params={"target": mask_target(url)},
                    sheet=cell.sheet,
                    cell=cell.coordinate,
                    formula=cell.formula,
                    value=mask_target(url),
                )
            )
        for func in sorted(network_calls):
            findings.append(
                self.diagnostic(
                    code="EXTERNAL_URL",
                    severity=Severity.ERROR,
                    message_key="external_url",
                    params={"target": f"{func}()"},
                    sheet=cell.sheet,
                    cell=cell.coordinate,
                    formula=cell.formula,
                    value=func,
                )
            )
        return findings
