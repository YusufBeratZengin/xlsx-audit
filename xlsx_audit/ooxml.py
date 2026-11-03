"""Read ``.xlsx`` packages straight from the OOXML zip container.

No Excel, no COM, no openpyxl: the reader walks ``xl/workbook.xml``, the
relationship parts, ``sharedStrings.xml`` and every worksheet part.  Malformed
or hostile input is rejected early (zip bomb limits, hardened XML parsing with
:mod:`defusedxml`).

Shared formulas (``<f t="shared" si="0" ref="A1:A9"/>``) are translated onto
each follower cell with the same relative-reference shifting Excel performs, so
rules see the formula that actually runs in every cell.
"""

from __future__ import annotations

import posixpath
import zipfile
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple, cast
from xml.etree.ElementTree import Element

from defusedxml.ElementTree import ParseError, fromstring

from .formula_parser import FormulaSyntaxError, parse_formula, shift_formula
from .models import Cell, DefinedName, ExternalLink, Sheet, Workbook
from .refs import format_coordinate, parse_coordinate

__all__ = ["WorkbookReadError", "read_workbook"]

MAX_ENTRIES = 20_000
MAX_PART_BYTES = 64 * 1024 * 1024
MAX_TOTAL_BYTES = 512 * 1024 * 1024

MAIN_NS_PREFIXES = ("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}",)
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


class WorkbookReadError(ValueError):
    """The file is not a readable, safe OOXML workbook."""


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _children(element: Element, name: str) -> Iterable[Element]:
    for child in element:
        if _local(child.tag) == name:
            yield child


def _first(element: Element, name: str) -> Optional[Element]:
    for child in _children(element, name):
        return child
    return None


def _attr(element: Element, name: str, default: Optional[str] = None) -> Optional[str]:
    if name in element.attrib:
        return element.attrib[name]
    for key, value in element.attrib.items():
        if _local(key) == name:
            return value
    return default


class _Package:
    """Bounded, read-only view over the zip entries."""

    def __init__(self, path: Path) -> None:
        try:
            self._zip = zipfile.ZipFile(path)
        except (zipfile.BadZipFile, OSError) as exc:
            raise WorkbookReadError(f"not a readable .xlsx package: {exc}") from exc
        self._names = set(self._zip.namelist())
        if len(self._names) > MAX_ENTRIES:
            raise WorkbookReadError("package has too many parts")
        total = sum(info.file_size for info in self._zip.infolist())
        if total > MAX_TOTAL_BYTES:
            raise WorkbookReadError("package expands beyond the size limit")
        self._total_read = 0

    @property
    def names(self) -> set[str]:
        return set(self._names)

    def has(self, part: str) -> bool:
        return part in self._names

    def read(self, part: str) -> bytes:
        info = self._zip.getinfo(part)
        if info.file_size > MAX_PART_BYTES:
            raise WorkbookReadError(f"part too large: {part}")
        self._total_read += info.file_size
        if self._total_read > MAX_TOTAL_BYTES:
            raise WorkbookReadError("package expands beyond the size limit")
        try:
            return self._zip.read(part)
        except (zipfile.BadZipFile, OSError, RuntimeError) as exc:
            raise WorkbookReadError(f"cannot read part {part}: {exc}") from exc

    def xml(self, part: str) -> Optional[Element]:
        if not self.has(part):
            return None
        try:
            return cast(Element, fromstring(self.read(part)))
        except (ParseError, ValueError) as exc:
            raise WorkbookReadError(f"invalid XML in {part}: {exc}") from exc

    def close(self) -> None:
        self._zip.close()


def _resolve_part(base_dir: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(base_dir, target))


def _rels(package: _Package, part: str) -> Dict[str, Tuple[str, str]]:
    """Map relationship id -> (resolved part or external target, target mode)."""
    directory, filename = posixpath.split(part)
    rels_part = posixpath.join(directory, "_rels", filename + ".rels")
    root = package.xml(rels_part)
    if root is None:
        return {}
    base_dir = directory
    mapping: Dict[str, Tuple[str, str]] = {}
    for rel in root:
        if _local(rel.tag) != "Relationship":
            continue
        rel_id = _attr(rel, "Id")
        target = _attr(rel, "Target")
        mode = _attr(rel, "TargetMode", "Internal") or "Internal"
        if not rel_id or target is None:
            continue
        if mode == "External":
            mapping[rel_id] = (target, "External")
        else:
            mapping[rel_id] = (_resolve_part(base_dir, target), "Internal")
    return mapping


def _read_shared_strings(package: _Package) -> List[str]:
    root = package.xml("xl/sharedStrings.xml")
    if root is None:
        return []
    strings: List[str] = []
    for si in root:
        if _local(si.tag) != "si":
            continue
        parts: List[str] = []
        for node in si.iter():
            if _local(node.tag) == "t" and node.text:
                parts.append(node.text)
        strings.append("".join(parts))
    return strings


def _cell_coordinate(raw_coordinate: Optional[str], row_number: int, column_number: int) -> str:
    if raw_coordinate:
        try:
            parse_coordinate(raw_coordinate)
        except ValueError as exc:
            raise WorkbookReadError(f"invalid cell coordinate {raw_coordinate!r}") from exc
        return raw_coordinate.upper()
    return format_coordinate(row_number, column_number)


def _cell_text(element: Element) -> Optional[str]:
    value = _first(element, "v")
    if value is not None:
        return value.text or ""
    inline = _first(element, "is")
    if inline is not None:
        return "".join(node.text or "" for node in inline.iter() if _local(node.tag) == "t")
    return None


class _RawFormula:
    __slots__ = ("kind", "ref", "shared_index", "text")

    def __init__(
        self,
        text: Optional[str],
        kind: Optional[str],
        shared_index: Optional[int],
        ref: Optional[str],
    ) -> None:
        self.text = text
        self.kind = kind
        self.shared_index = shared_index
        self.ref = ref


def _read_worksheet(
    package: _Package,
    part: str,
    sheet_name: str,
    shared_strings: List[str],
) -> Tuple[List[Cell], Dict[str, str]]:
    root = package.xml(part)
    if root is None:
        return [], {}
    sheet_data = _first(root, "sheetData")
    if sheet_data is None:
        return [], {}

    raw_cells: List[Tuple[str, int, int, Optional[str], str, int, Optional[_RawFormula]]] = []
    for row in _children(sheet_data, "row"):
        row_number = int(_attr(row, "r", "0") or 0)
        for column_number, element in enumerate(_children(row, "c"), start=1):
            raw_coordinate = _attr(element, "r")
            if raw_coordinate:
                try:
                    row_index, col_index = parse_coordinate(raw_coordinate)
                except ValueError as exc:
                    raise WorkbookReadError(
                        f"invalid cell coordinate {raw_coordinate!r} in {sheet_name!r}"
                    ) from exc
            else:
                row_index, col_index = row_number or 1, column_number
            coordinate = _cell_coordinate(raw_coordinate, row_index, col_index)
            cell_type = _attr(element, "t", "n") or "n"
            style_id = int(_attr(element, "s", "0") or 0)
            formula_element = _first(element, "f")
            formula = None
            if formula_element is not None:
                formula_type = _attr(formula_element, "t")
                shared_index = _attr(formula_element, "si")
                formula = _RawFormula(
                    text=formula_element.text,
                    kind=formula_type,
                    shared_index=int(shared_index) if shared_index is not None else None,
                    ref=_attr(formula_element, "ref"),
                )
            value = _cell_text(element)
            if cell_type == "s" and value is not None and value != "":
                try:
                    value = shared_strings[int(value)]
                except (ValueError, IndexError) as exc:
                    raise WorkbookReadError(
                        f"cell {sheet_name}!{coordinate} points at a missing shared string"
                    ) from exc
            elif cell_type == "b" and value in ("0", "1"):
                value = "TRUE" if value == "1" else "FALSE"
            raw_cells.append(
                (coordinate, row_index, col_index, value, cell_type, style_id, formula)
            )

    masters: Dict[int, Tuple[str, str]] = {}
    for coordinate, _row, _col, _value, _type, _style, formula in raw_cells:
        if (
            formula is not None
            and formula.kind == "shared"
            and formula.text
            and formula.shared_index is not None
        ):
            masters[formula.shared_index] = (coordinate, formula.text)

    cells: List[Cell] = []
    for coordinate, row_index, col_index, value, cell_type, style_id, formula in raw_cells:
        formula_text: Optional[str] = None
        if formula is not None:
            if formula.kind == "shared" and not formula.text:
                master = (
                    masters.get(formula.shared_index) if formula.shared_index is not None else None
                )
                if master is not None:
                    master_coordinate, master_text = master
                    master_row, master_column = parse_coordinate(master_coordinate)
                    try:
                        formula_text = shift_formula(
                            master_text, row_index - master_row, col_index - master_column
                        )
                    except FormulaSyntaxError:
                        formula_text = master_text
                else:
                    formula_text = None
            else:
                formula_text = formula.text
            if formula_text is not None and not formula_text.lstrip().startswith("="):
                formula_text = "=" + formula_text
        if formula_text is None and value in (None, ""):
            continue
        cells.append(
            Cell(
                sheet=sheet_name,
                coordinate=coordinate,
                row=row_index,
                column=col_index,
                formula=formula_text,
                value=value,
                value_type=cell_type,
                style_id=style_id,
            )
        )
    hyperlinks: Dict[str, str] = {}
    hyperlinks_elem = _first(root, "hyperlinks")
    if hyperlinks_elem is not None:
        for hl in _children(hyperlinks_elem, "hyperlink"):
            ref = _attr(hl, "ref")
            rel_id = _attr(hl, "id")
            if ref and rel_id:
                hyperlinks[rel_id] = ref
    return cells, hyperlinks


def _read_defined_names(root: Element) -> List[DefinedName]:
    container = _first(root, "definedNames")
    if container is None:
        return []
    names: List[DefinedName] = []
    for entry in _children(container, "definedName"):
        name = _attr(entry, "name")
        if not name:
            continue
        local = _attr(entry, "localSheetId")
        names.append(
            DefinedName(
                name=name,
                formula=(entry.text or "").strip(),
                local_sheet_id=int(local) if local is not None else None,
                hidden=(_attr(entry, "hidden", "0") or "0") in ("1", "true"),
            )
        )
    return names


def _read_properties(package: _Package) -> Dict[str, str]:
    props: Dict[str, str] = {}
    core = package.xml("docProps/core.xml")
    if core is not None:
        for node in core.iter():
            name = _local(node.tag)
            if name in ("creator", "lastModifiedBy", "title", "subject") and node.text:
                props[name] = node.text
    app = package.xml("docProps/app.xml")
    if app is not None:
        for node in app.iter():
            if _local(node.tag) in ("Company", "Manager") and node.text:
                props[_local(node.tag).lower()] = node.text
    return props


def _collect_external_links(
    package: _Package,
    workbook_part: str,
    sheet_parts: Optional[Dict[str, str]] = None,
    sheet_hyperlinks: Optional[Dict[str, Dict[str, str]]] = None,
) -> List[ExternalLink]:
    """Relationships that leave the package.

    A workbook relationship of type ``externalLink`` points at an *internal*
    part (``xl/externalLinks/externalLink1.xml``); the real destination lives in
    that part's own relationships with ``TargetMode="External"``.
    Worksheet parts also define relationships for external hyperlinks.
    """
    links: List[ExternalLink] = []
    seen: set[Tuple[str, str, Optional[str]]] = set()

    candidate_parts = [workbook_part]
    for entry in sorted(package.names):
        is_external = entry.startswith("xl/externalLinks/externalLink") and entry.endswith(".xml")
        is_worksheet = entry.startswith("xl/worksheets/") and entry.endswith(".xml")
        if (is_external or is_worksheet) and entry not in candidate_parts:
            candidate_parts.append(entry)

    if sheet_parts:
        for p in sheet_parts:
            if p not in candidate_parts:
                candidate_parts.append(p)

    for part in candidate_parts:
        sheet_name = sheet_parts.get(part) if sheet_parts else None
        hl_map = (sheet_hyperlinks.get(part) if sheet_hyperlinks else None) or {}

        for rel_id, (target, mode) in _rels(package, part).items():
            if mode != "External":
                continue
            cell_ref = hl_map.get(rel_id)
            key = (part, target, cell_ref)
            if key in seen:
                continue
            seen.add(key)
            is_sheet = part.startswith("xl/worksheets/") or sheet_name is not None
            kind = "hyperlink" if is_sheet else "externalLink"
            links.append(
                ExternalLink(
                    part=part,
                    target=target,
                    kind=kind,
                    sheet=sheet_name,
                    cell=cell_ref,
                )
            )
    return links


def read_workbook(source: str | Path) -> Workbook:
    """Read ``source`` and return a :class:`~xlsx_audit.models.Workbook`."""
    path = Path(source)
    if not path.exists():
        raise WorkbookReadError(f"file not found: {path}")
    if path.is_dir():
        raise WorkbookReadError(f"expected a file, got a directory: {path}")
    package = _Package(path)
    try:
        workbook_part = "xl/workbook.xml"
        if not package.has(workbook_part):
            raise WorkbookReadError("package has no xl/workbook.xml")
        root = package.xml(workbook_part)
        if root is None:  # pragma: no cover - has() already checked
            raise WorkbookReadError("package has no xl/workbook.xml")

        shared_strings = _read_shared_strings(package)
        rels = _rels(package, workbook_part)

        sheets_container = _first(root, "sheets")
        if sheets_container is None:
            raise WorkbookReadError("workbook.xml has no <sheets>")
        sheets: List[Sheet] = []
        sheet_parts_map: Dict[str, str] = {}
        sheet_hyperlinks_map: Dict[str, Dict[str, str]] = {}
        for entry in _children(sheets_container, "sheet"):
            name = _attr(entry, "name")
            if not name:
                raise WorkbookReadError("worksheet without a name")
            rel_id = _attr(entry, "id")
            part = None
            if rel_id is not None:
                if rel_id not in rels:
                    raise WorkbookReadError(
                        f"worksheet {name!r} references unknown relationship {rel_id!r}"
                    )
                part = rels[rel_id][0]
            if part is None:
                part = f"xl/worksheets/sheet{len(sheets) + 1}.xml"
            sheet_parts_map[part] = name
            cells, hyperlinks = _read_worksheet(package, part, name, shared_strings)
            sheet_hyperlinks_map[part] = hyperlinks
            sheets.append(
                Sheet(
                    name=name,
                    sheet_id=int(_attr(entry, "sheetId", "0") or 0),
                    state=_attr(entry, "state", "visible") or "visible",
                    cells=tuple(cells),
                )
            )
        workbook = Workbook(
            path=str(path),
            sheets=sheets,
            defined_names=_read_defined_names(root),
            external_links=_collect_external_links(
                package,
                workbook_part,
                sheet_parts=sheet_parts_map,
                sheet_hyperlinks=sheet_hyperlinks_map,
            ),
            properties=_read_properties(package),
        )
        return workbook
    finally:
        package.close()


def formula_references(formula: str) -> List[Tuple[Optional[str], Optional[int], Optional[int]]]:
    """Convenience helper: ``(sheet, row, column)`` for every direct reference."""
    try:
        tree = parse_formula(formula)
    except FormulaSyntaxError:
        return []
    from .formula_parser import Ref, iter_nodes

    out: List[Tuple[Optional[str], Optional[int], Optional[int]]] = []
    for node in iter_nodes(tree):
        if isinstance(node, Ref) and node.row is not None and node.column is not None:
            out.append((node.sheet, node.row, node.column))
    return out
