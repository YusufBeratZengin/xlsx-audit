"""Build real ``.xlsx`` packages for the test-suite.

openpyxl cannot produce the interesting cases (shared formulas, ``veryHidden``
sheets, external-link relationships), so the tests write the OOXML parts
directly.  Everything here is synthetic: names, numbers and paths are obviously
fake.
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
from xml.sax.saxutils import escape, quoteattr

from xlsx_audit.models import Cell, DefinedName, ExternalLink, Sheet, Workbook
from xlsx_audit.refs import parse_coordinate

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

CONTENT_TYPES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{CONTENT_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"""

ROOT_RELS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="{REL_NS}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""


def _cell_xml(reference: str, spec: Dict[str, object], shared_index: Optional[int]) -> str:
    attributes = [f'r="{reference}"']
    cell_type = spec.get("t")
    if cell_type:
        attributes.append(f't="{cell_type}"')
    if spec.get("s") is not None:
        attributes.append(f's="{spec["s"]}"')
    children = ""
    formula = spec.get("f")
    shared = spec.get("shared")
    if formula is not None or shared is not None:
        formula_attributes = ""
        if shared is not None:
            formula_attributes += f' t="shared" si="{shared}"'
        if spec.get("ref"):
            formula_attributes += f' ref="{spec["ref"]}"'
        if spec.get("array"):
            formula_attributes = f' t="array" ref="{spec["array"]}"'
        text = escape(str(formula)) if formula is not None else ""
        children += f"<f{formula_attributes}>{text}</f>"
    value = spec.get("v")
    if value is not None:
        if cell_type == "s" and shared_index is None:
            raise AssertionError("shared string index missing")
        rendered = str(shared_index) if cell_type == "s" else str(value)
        children += f"<v>{escape(rendered)}</v>"
    return f"<c {' '.join(attributes)}>{children}</c>"


def _sheet_xml(cells: Dict[str, Dict[str, object]], shared_strings: List[str]) -> str:
    rows: Dict[int, List[Tuple[int, str]]] = {}
    for reference, spec in cells.items():
        row = int("".join(ch for ch in reference if ch.isdigit()))
        column = 0
        for char in reference:
            if char.isalpha():
                column = column * 26 + (ord(char.upper()) - 64)
            else:
                break
        index = None
        if spec.get("t") == "s" and spec.get("v") is not None:
            text = str(spec["v"])
            if text not in shared_strings:
                shared_strings.append(text)
            index = shared_strings.index(text)
        rows.setdefault(row, []).append((column, _cell_xml(reference, spec, index)))

    body: List[str] = []
    for row_number in sorted(rows):
        entries = "".join(xml for _col, xml in sorted(rows[row_number]))
        body.append(f'<row r="{row_number}">{entries}</row>')
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{MAIN_NS}"><sheetData>{"".join(body)}</sheetData></worksheet>'
    )


def build_workbook(
    path: Path,
    sheets: Sequence[Dict[str, object]],
    *,
    external_targets: Sequence[str] = (),
    defined_names: Sequence[Dict[str, object]] = (),
    metadata: Optional[Dict[str, str]] = None,
    extra_parts: Optional[Dict[str, str]] = None,
    shared_strings_xml: Optional[str] = None,
) -> Path:
    """Write a workbook and return ``path``.

    ``sheets`` entries look like ``{"name": "Gelirler", "state": "visible",
    "cells": {"A1": {"v": 1}, "B1": {"f": "A1*1.2", "v": 1.2}}}``.
    """
    path = Path(path)
    shared_strings: List[str] = []
    sheet_parts: List[str] = []
    for _index, sheet in enumerate(sheets, start=1):
        cells = sheet.get("cells", {})
        assert isinstance(cells, dict)
        sheet_parts.append(_sheet_xml(cells, shared_strings))

    workbook_rels = [
        f'<Relationship Id="rIdS{index}" Type="{REL_NS}/worksheet" '
        f'Target="worksheets/sheet{index}.xml"/>'
        for index in range(1, len(sheets) + 1)
    ]
    if shared_strings:
        workbook_rels.append(
            f'<Relationship Id="rIdSS" Type="{REL_NS}/sharedStrings" Target="sharedStrings.xml"/>'
        )
    for index, _target in enumerate(external_targets, start=1):
        workbook_rels.append(
            f'<Relationship Id="rIdEL{index}" Type="{REL_NS}/externalLink" '
            f'Target="externalLinks/externalLink{index}.xml"/>'
        )

    workbook_sheets = []
    for index, sheet in enumerate(sheets, start=1):
        state = sheet.get("state")
        state_attr = f' state="{state}"' if state and state != "visible" else ""
        workbook_sheets.append(
            f'<sheet name={quoteattr(str(sheet["name"]))} sheetId="{index}" '
            f'r:id="rIdS{index}"{state_attr}/>'
        )

    names_xml = ""
    if defined_names:
        entries = []
        for defined in defined_names:
            attributes = f"name={quoteattr(str(defined['name']))}"
            if defined.get("local_sheet_id") is not None:
                attributes += f' localSheetId="{defined["local_sheet_id"]}"'
            if defined.get("hidden"):
                attributes += ' hidden="1"'
            entries.append(
                f"<definedName {attributes}>{escape(str(defined['formula']))}</definedName>"
            )
        names_xml = f"<definedNames>{''.join(entries)}</definedNames>"

    workbook_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}">'
        f"<sheets>{''.join(workbook_sheets)}</sheets>{names_xml}</workbook>"
    )

    parts: Dict[str, str] = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "xl/workbook.xml": workbook_xml,
        "xl/_rels/workbook.xml.rels": (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{PKG_REL_NS}">{"".join(workbook_rels)}</Relationships>'
        ),
    }
    for index, xml in enumerate(sheet_parts, start=1):
        parts[f"xl/worksheets/sheet{index}.xml"] = xml
    if shared_strings:
        parts["xl/sharedStrings.xml"] = shared_strings_xml or (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<sst xmlns="{MAIN_NS}" count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">'
            + "".join(f"<si><t>{escape(text)}</t></si>" for text in shared_strings)
            + "</sst>"
        )
    for index, target in enumerate(external_targets, start=1):
        parts[f"xl/externalLinks/externalLink{index}.xml"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<externalLink xmlns="{MAIN_NS}" xmlns:r="{REL_NS}">'
            f'<externalBook r:id="rId1"/></externalLink>'
        )
        parts[f"xl/externalLinks/_rels/externalLink{index}.xml.rels"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{PKG_REL_NS}">'
            f'<Relationship Id="rId1" Type="{REL_NS}/externalLinkPath" '
            f'Target={quoteattr(target)} TargetMode="External"/></Relationships>'
        )
    if metadata:
        core_entries = "".join(f"<{key}>{escape(value)}</{key}>" for key, value in metadata.items())
        parts["docProps/core.xml"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            f'xmlns:dc="http://purl.org/dc/elements/1.1/">{core_entries}</cp:coreProperties>'
        )
    if extra_parts:
        parts.update(extra_parts)

    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content)
    return path


def simple_sheet(
    name: str = "Sheet1",
    cells: Optional[Dict[str, Dict[str, object]]] = None,
    state: str = "visible",
) -> Dict[str, object]:
    return {"name": name, "state": state, "cells": cells or {}}


def write_bytes(path: Path, payload: bytes) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(payload)
    return path


def zip_parts(path: Path, parts: Iterable[Tuple[str, str]]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as archive:
        for name, content in parts:
            archive.writestr(name, content)
    return path


# --------------------------------------------------------------------------
# In-memory model helpers (unit tests that do not need a real package)
# --------------------------------------------------------------------------


def make_cell(
    sheet: str,
    coordinate: str,
    formula: Optional[str] = None,
    value: Optional[str] = None,
    value_type: str = "n",
) -> Cell:
    row, column = parse_coordinate(coordinate)
    return Cell(
        sheet=sheet,
        coordinate=coordinate,
        row=row,
        column=column,
        formula=formula,
        value=value,
        value_type=value_type,
    )


def make_workbook(
    sheets: Dict[str, Dict[str, Dict[str, Optional[str]]]],
    *,
    path: str = "fake_example.xlsx",
    external_links: Sequence[ExternalLink] = (),
    defined_names: Sequence[DefinedName] = (),
    properties: Optional[Dict[str, str]] = None,
    states: Optional[Dict[str, str]] = None,
) -> Workbook:
    """Build a :class:`Workbook` in memory.

    ``sheets`` maps a sheet name to ``{coordinate: {"formula": ..., "value": ...}}``.
    """
    built: List[Sheet] = []
    for index, (name, cells) in enumerate(sheets.items(), start=1):
        built.append(
            Sheet(
                name=name,
                sheet_id=index,
                state=(states or {}).get(name, "visible"),
                cells=tuple(
                    make_cell(
                        name,
                        coordinate,
                        spec.get("formula"),
                        spec.get("value"),
                        spec.get("type") or "n",
                    )
                    for coordinate, spec in cells.items()
                ),
            )
        )
    return Workbook(
        path=path,
        sheets=built,
        defined_names=list(defined_names),
        external_links=list(external_links),
        properties=dict(properties or {}),
    )
