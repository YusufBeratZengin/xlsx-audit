#!/usr/bin/env python3
"""Regenerate ``examples/fake_financial_model.xlsx``.

The workbook is deliberately wrong: it contains a destroyed reference, a
hard-coded multiplier, a veryHidden sheet and an external link to a made-up
local path.  Every value is synthetic ("example", xxxxxx); nothing here comes
from a real organisation, person or rate.

Only the standard library is used, so the example can be rebuilt in any
checkout:

    python3 scripts/make_example.py
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.sax.saxutils import quoteattr

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

CONTENT_TYPES = f"""{HEADER}
<Types xmlns="{CONTENT_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"""

ROOT_RELS = f"""{HEADER}
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="{REL_NS}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

SHEETS = {
    "Gelirler": {
        "A1": ("n", "100"),
        "A2": ("n", "250"),
        "A3": ("n", "175"),
        "B1": ("str", "Revenue line A (example)"),
        "B2": ("str", "Revenue line B (example)"),
        "B3": ("str", "Revenue line C (example)"),
        "C14": ("n", None, "SUM(A1:A3)"),
        "E15": ("err", "#REF!", "A15+B15"),
    },
    "Giderler": {
        "A8": ("n", "1000"),
        "B8": ("n", None, "A8*1.20"),
        "C8": ("n", "50"),
        "C9": ("n", "75"),
        "C10": ("n", "25"),
        "D11": ("n", None, "SUM(C8:C9)"),
    },
    "Notlar": {
        "A1": ("str", "Internal scratch sheet (example)"),
    },
}


def _cell(reference: str, value_type: str, value: str | None, formula: str | None) -> str:
    attribute = "" if value_type == "n" else f' t="{value_type}"'
    body = ""
    if formula is not None:
        body += f"<f>{formula}</f>"
    if value is not None:
        body += f"<v>{value}</v>"
    return f'<c r="{reference}"{attribute}>{body}</c>'


def _sheet_xml(cells: dict[str, tuple[str, str | None, str | None]]) -> str:
    rows: dict[int, list[tuple[str, str]]] = {}
    for reference, spec in cells.items():
        value_type, value = spec[0], spec[1]
        formula = spec[2] if len(spec) > 2 else None
        row = int("".join(ch for ch in reference if ch.isdigit()))
        rows.setdefault(row, []).append((reference, _cell(reference, value_type, value, formula)))
    body = "".join(
        f'<row r="{number}">{"".join(part for _ref, part in sorted(rows[number]))}</row>'
        for number in sorted(rows)
    )
    return f'{HEADER}<worksheet xmlns="{MAIN_NS}"><sheetData>{body}</sheetData></worksheet>'


def build(target: Path) -> Path:
    names = list(SHEETS)
    workbook_sheets = []
    for index, name in enumerate(names, start=1):
        state = ' state="veryHidden"' if name == "Notlar" else ""
        workbook_sheets.append(
            f'<sheet name={quoteattr(name)} sheetId="{index}" r:id="rIdS{index}"{state}/>'
        )
    workbook_xml = (
        f'{HEADER}<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}">'
        f"<sheets>{''.join(workbook_sheets)}</sheets></workbook>"
    )
    rels = [
        f'<Relationship Id="rIdS{index}" Type="{REL_NS}/worksheet" '
        f'Target="worksheets/sheet{index}.xml"/>'
        for index in range(1, len(names) + 1)
    ]
    rels.append(
        f'<Relationship Id="rIdEL1" Type="{REL_NS}/externalLink" '
        f'Target="externalLinks/externalLink1.xml"/>'
    )

    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "xl/workbook.xml": workbook_xml,
        "xl/_rels/workbook.xml.rels": (
            f'{HEADER}<Relationships xmlns="{PKG_REL_NS}">{"".join(rels)}</Relationships>'
        ),
        # A saved link to somebody else's desktop: the classic leak this tool
        # is built to find.  The path is obviously fake.
        "xl/externalLinks/externalLink1.xml": (
            f'{HEADER}<externalLink xmlns="{MAIN_NS}" xmlns:r="{REL_NS}">'
            '<externalBook r:id="rId1"/></externalLink>'
        ),
        "xl/externalLinks/_rels/externalLink1.xml.rels": (
            f'{HEADER}<Relationships xmlns="{PKG_REL_NS}">'
            f'<Relationship Id="rId1" Type="{REL_NS}/externalLinkPath" '
            'Target="file:///C:/Users/example/Budget_2025.xlsx" TargetMode="External"/>'
            "</Relationships>"
        ),
    }
    for index, name in enumerate(names, start=1):
        parts[f"xl/worksheets/sheet{index}.xml"] = _sheet_xml(SHEETS[name])

    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content)
    return target


if __name__ == "__main__":
    root = Path(__file__).resolve().parent.parent
    out = build(root / "examples" / "fake_financial_model.xlsx")
    print(f"wrote {out.relative_to(root)}")
