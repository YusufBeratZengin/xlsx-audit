"""Command line interface."""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from typing import List, Optional, Sequence

from . import __version__
from .audit import audit_workbook
from .i18n import normalize_lang
from .models import AuditReport
from .ooxml import WorkbookReadError
from .reporter import render_report, use_color
from .rules import RULE_CLASSES

__all__ = ["EXIT_FINDINGS", "EXIT_OK", "EXIT_USAGE", "build_parser", "main"]

EXIT_OK = 0
EXIT_FINDINGS = 1
EXIT_USAGE = 2

FORMATS = ("text", "json", "markdown", "sarif")

EPILOG = """examples:
  xlsx-audit check model.xlsx                 # audit one workbook
  xlsx-audit check model.xlsx --strict        # warnings fail the run too
  xlsx-audit check *.xlsx --format sarif -o audit.sarif
  xlsx-audit check book.xlsx --lang tr        # Turkish findings
  xlsx-audit rules                            # list every rule
"""


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="xlsx-audit",
        description="Audit Excel workbooks for broken formulas, circular references, "
        "magic numbers, hidden sheets and external links.",
        epilog=EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    subparsers = parser.add_subparsers(dest="command")

    check = subparsers.add_parser(
        "check",
        help="audit one or more .xlsx files",
        description="Audit one or more .xlsx files. Exit code 1 means findings, 2 means "
        "the file could not be read.",
    )
    check.add_argument("files", nargs="+", metavar="FILE", help="workbook(s) to audit")
    check.add_argument(
        "--strict",
        action="store_true",
        help="treat warnings as failures (recommended for CI release gates)",
    )
    check.add_argument(
        "--format",
        dest="fmt",
        choices=FORMATS,
        default="text",
        help="report format (default: text)",
    )
    check.add_argument(
        "-o",
        "--output",
        metavar="PATH",
        help="write the report to PATH instead of stdout (single input only; "
        "must not name the input workbook)",
    )
    check.add_argument(
        "--lang",
        default=os.environ.get("XLSX_AUDIT_LANG", "en"),
        choices=("en", "tr"),
        help="message language (default: en, env XLSX_AUDIT_LANG)",
    )
    check.add_argument(
        "--ignore",
        action="append",
        default=[],
        metavar="CODE[,CODE]",
        help="suppress findings by code; repeatable",
    )
    check.add_argument(
        "--rules",
        action="append",
        default=[],
        metavar="NAME[,NAME]",
        help="run only the named rules; repeatable",
    )
    check.add_argument(
        "--color",
        dest="color",
        action="store_true",
        default=None,
        help="force ANSI colours on a non-TTY stream",
    )
    check.add_argument(
        "--no-color",
        dest="color",
        action="store_false",
        help="disable ANSI colours (also honours NO_COLOR)",
    )
    check.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="also print informational findings",
    )
    check.add_argument(
        "--exit-zero",
        action="store_true",
        help="always exit 0 when the files could be read (reporting only)",
    )

    subparsers.add_parser("rules", help="list available rules")
    return parser


def _split_many(values: Sequence[str]) -> List[str]:
    out: List[str] = []
    for value in values:
        for part in value.split(","):
            cleaned = part.strip()
            if cleaned and cleaned not in out:
                out.append(cleaned)
    return out


def _same_file(first: Path, second: Path) -> bool:
    """True when two paths name the same filesystem object.

    ``resolve()`` catches a not-yet-existing output path and relative spellings;
    ``os.path.samefile`` additionally catches hard links and symlinks.
    """
    try:
        if first.resolve() == second.resolve():
            return True
    except OSError:
        pass
    try:
        return first.exists() and second.exists() and os.path.samefile(first, second)
    except OSError:
        return False


def _run_check(args: argparse.Namespace) -> int:
    lang = normalize_lang(args.lang)
    ignore = _split_many(args.ignore)
    selected = _split_many(args.rules) or None
    if args.output and len(args.files) > 1:
        print("xlsx-audit: --output accepts a single input file", file=sys.stderr)
        return EXIT_USAGE
    if args.output:
        output_path = Path(args.output)
        for raw_path in args.files:
            if _same_file(output_path, Path(raw_path)):
                print(
                    "xlsx-audit: refusing to write the report over the input workbook "
                    f"{raw_path}; choose a different --output path",
                    file=sys.stderr,
                )
                return EXIT_USAGE

    reports: List[AuditReport] = []
    read_errors = 0
    for raw_path in args.files:
        path = Path(raw_path)
        try:
            reports.append(
                audit_workbook(
                    path,
                    strict=args.strict,
                    lang=lang,
                    ignore=ignore,
                    rules=selected,
                )
            )
        except WorkbookReadError as exc:
            print(f"xlsx-audit: {path}: {exc}", file=sys.stderr)
            read_errors += 1
        except ValueError as exc:
            print(f"xlsx-audit: {exc}", file=sys.stderr)
            return EXIT_USAGE
    if read_errors:
        return EXIT_USAGE
    if not reports:
        return EXIT_USAGE

    color = use_color(args.color) and args.fmt == "text"
    if args.fmt == "text":
        body = (
            "\n".join(
                render_report(report, "text", lang=lang, color=color, verbose=args.verbose).rstrip()
                for report in reports
            )
            + "\n"
        )
    elif len(reports) == 1:
        body = render_report(reports[0], args.fmt, lang=lang, color=color, verbose=args.verbose)
    else:
        body = _render_many(reports, args.fmt)

    if args.output:
        try:
            Path(args.output).write_text(body, encoding="utf-8")
        except OSError as exc:
            print(f"xlsx-audit: cannot write output to {args.output}: {exc}", file=sys.stderr)
            return EXIT_USAGE
        print(f"xlsx-audit: report written to {args.output}", file=sys.stderr)
    else:
        sys.stdout.write(body)
        sys.stdout.flush()

    if args.exit_zero:
        return EXIT_OK
    return EXIT_OK if all(report.passed for report in reports) else EXIT_FINDINGS


def _render_many(reports: List[AuditReport], fmt: str) -> str:
    if fmt == "sarif":
        import json

        from .sarif_reporter import sarif_document

        return json.dumps(sarif_document(reports), ensure_ascii=False, indent=2) + "\n"
    if fmt == "json":
        import json

        return (
            json.dumps([report.to_dict() for report in reports], ensure_ascii=False, indent=2)
            + "\n"
        )
    return "\n".join(render_report(report, fmt) for report in reports)


def _run_rules() -> int:
    print("Available rules:\n")
    for rule_class in RULE_CLASSES:
        print(f"  {rule_class.name:<20} {rule_class.description}")
    print("\nUse --rules NAME[,NAME] to run a subset; --ignore CODE to hide findings.")
    return EXIT_OK


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)
    if args.command == "rules":
        return _run_rules()
    if args.command == "check":
        return _run_check(args)
    parser.print_help()
    return EXIT_USAGE


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
