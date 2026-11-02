"""Formula dependency graph and circular-reference detection.

Every formula cell is a node; an edge points at each formula cell it reads.
Strongly connected components (Tarjan, iterative so deep models cannot blow the
Python stack) that contain more than one node - or a self-reference - are
circular references.  Excel only reports the first one it happens to evaluate;
this reports all of them with a readable path.
"""

from __future__ import annotations

from collections import deque
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

from .formula_parser import FormulaSyntaxError, NameRef, Node, Range, Ref, iter_nodes, parse_formula
from .models import DefinedName, Workbook
from .refs import format_coordinate

__all__ = ["Cycle", "build_dependency_graph", "find_cycles"]

NodeKey = Tuple[str, str]

#: Refuse to materialise ranges larger than this; such formulas are usually
#: whole-column references and expanding them would gain nothing.
MAX_RANGE_CELLS = 20_000
MAX_GRAPH_NODES = 200_000


@dataclass(frozen=True)
class Cycle:
    nodes: Tuple[NodeKey, ...]

    def path(self) -> List[NodeKey]:
        if not self.nodes:
            return []
        return [*list(self.nodes), self.nodes[0]]

    def path_text(self) -> str:
        return " -> ".join(f"{sheet}!{cell}" for sheet, cell in self.path())


def _defined_name_targets(defined: DefinedName, workbook: Workbook) -> List[NodeKey]:
    """Cells a defined name points at (best effort, same-workbook only)."""
    if not defined.formula:
        return []
    try:
        tree = parse_formula(defined.formula)
    except FormulaSyntaxError:
        return []
    default_sheet: Optional[str] = None
    if defined.local_sheet_id is not None:
        index = defined.local_sheet_id
        if 0 <= index < len(workbook.sheets):
            default_sheet = workbook.sheets[index].name
    targets: List[NodeKey] = []
    for node in iter_nodes(tree):
        if isinstance(node, Range):
            targets.extend(_expand_range(node, default_sheet))
        elif isinstance(node, Ref) and node.book is None:
            sheet = node.sheet or default_sheet
            if sheet and node.row and node.column:
                targets.append((sheet, format_coordinate(node.row, node.column)))
    return targets


def _expand_range(node: Range, default_sheet: Optional[str]) -> List[NodeKey]:
    start, end = node.start, node.end
    if start.book is not None or end.book is not None:
        return []
    sheet = start.sheet or end.sheet or default_sheet
    if not sheet:
        return []
    if start.row is None or start.column is None or end.row is None or end.column is None:
        return []
    top, bottom = sorted((start.row, end.row))
    left, right = sorted((start.column, end.column))
    if (bottom - top + 1) * (right - left + 1) > MAX_RANGE_CELLS:
        return []
    return [
        (sheet, format_coordinate(row, column))
        for row in range(top, bottom + 1)
        for column in range(left, right + 1)
    ]


def build_dependency_graph(workbook: Workbook) -> Dict[NodeKey, Set[NodeKey]]:
    """Adjacency map restricted to formula cells of this workbook."""
    formula_nodes: Set[NodeKey] = {
        (cell.sheet, cell.coordinate) for cell in workbook.cells.values() if cell.formula
    }
    if len(formula_nodes) > MAX_GRAPH_NODES:
        raise ValueError(
            f"workbook has {len(formula_nodes)} formula cells, above the {MAX_GRAPH_NODES} limit"
        )
    names: Dict[str, List[NodeKey]] = {}
    for defined in workbook.defined_names:
        names.setdefault(defined.name.upper(), []).extend(_defined_name_targets(defined, workbook))

    graph: Dict[NodeKey, Set[NodeKey]] = {node: set() for node in formula_nodes}
    for cell in workbook.cells.values():
        if not cell.formula:
            continue
        try:
            tree = parse_formula(cell.formula)
        except FormulaSyntaxError:
            continue
        source = (cell.sheet, cell.coordinate)
        edges = graph[source]
        for node in iter_nodes(tree):
            for target in _targets(node, cell.sheet, names):
                if target in formula_nodes and target != source:
                    edges.add(target)
                elif target == source:
                    edges.add(source)
    return graph


def _targets(node: Node, sheet: str, names: Dict[str, List[NodeKey]]) -> Iterable[NodeKey]:
    if isinstance(node, Range):
        yield from _expand_range(node, sheet)
    elif isinstance(node, Ref):
        if node.book is not None or node.row is None or node.column is None:
            return
        yield (node.sheet or sheet, format_coordinate(node.row, node.column))
    elif isinstance(node, NameRef):
        yield from names.get(node.name.upper(), [])


def _tarjan(graph: Dict[NodeKey, Set[NodeKey]]) -> List[List[NodeKey]]:
    """Iterative Tarjan SCC; returns components in deterministic order."""
    index_counter = 0
    indices: Dict[NodeKey, int] = {}
    low: Dict[NodeKey, int] = {}
    on_stack: Set[NodeKey] = set()
    stack: List[NodeKey] = []
    components: List[List[NodeKey]] = []

    for root in sorted(graph):
        if root in indices:
            continue
        work: List[Tuple[NodeKey, Iterable[NodeKey]]] = [(root, iter(sorted(graph[root])))]
        indices[root] = low[root] = index_counter
        index_counter += 1
        stack.append(root)
        on_stack.add(root)
        while work:
            node, children = work[-1]
            advanced = False
            for child in children:
                if child not in graph:
                    continue
                if child not in indices:
                    indices[child] = low[child] = index_counter
                    index_counter += 1
                    stack.append(child)
                    on_stack.add(child)
                    work.append((child, iter(sorted(graph[child]))))
                    advanced = True
                    break
                if child in on_stack:
                    low[node] = min(low[node], indices[child])
            if advanced:
                continue
            work.pop()
            if work:
                parent = work[-1][0]
                low[parent] = min(low[parent], low[node])
            if low[node] == indices[node]:
                component: List[NodeKey] = []
                while True:
                    member = stack.pop()
                    on_stack.discard(member)
                    component.append(member)
                    if member == node:
                        break
                components.append(sorted(component))
    return components


def _cycle_in_component(
    component: Sequence[NodeKey], graph: Dict[NodeKey, Set[NodeKey]]
) -> List[NodeKey]:
    """Shortest cycle through the smallest node of a strongly connected component."""
    members = set(component)
    start = min(component)
    for first in sorted(graph[start] & members):
        if first == start:
            return [start]
        parents: Dict[NodeKey, Optional[NodeKey]] = {first: None}
        queue: deque[NodeKey] = deque([first])
        while queue:
            current = queue.popleft()
            for nxt in sorted(graph[current] & members):
                if nxt == start:
                    path = [current]
                    parent = parents[current]
                    while parent is not None:
                        path.append(parent)
                        parent = parents[parent]
                    path.reverse()
                    return [start, *path]
                if nxt not in parents:
                    parents[nxt] = current
                    queue.append(nxt)
    return list(component)


def find_cycles(workbook: Workbook) -> List[Cycle]:
    """All circular references in ``workbook``, deterministically ordered."""
    graph = build_dependency_graph(workbook)
    cycles: List[Cycle] = []
    for component in _tarjan(graph):
        if len(component) == 1:
            node = component[0]
            if node not in graph[node]:
                continue
            cycles.append(Cycle((node,)))
            continue
        ordered = _cycle_in_component(component, graph)
        cycles.append(Cycle(tuple(ordered)))
    cycles.sort(key=lambda c: [(sheet, cell) for sheet, cell in c.nodes])
    return dedupe_cycles(cycles)


def dedupe_cycles(cycles: List[Cycle]) -> List[Cycle]:
    seen: Set[Tuple[NodeKey, ...]] = set()
    out: List[Cycle] = []
    for cycle in cycles:
        key = tuple(sorted(cycle.nodes))
        if key in seen:
            continue
        seen.add(key)
        out.append(cycle)
    return out
