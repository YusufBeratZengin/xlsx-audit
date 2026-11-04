"""Formula dependency graph and circular-reference detection.

Every formula cell is a node; an edge points at each formula cell it reads.
Strongly connected components (Tarjan, iterative so deep models cannot blow the
Python stack) are decomposed into elementary cycles (iterative Johnson's
algorithm, bounded by ``MAX_CYCLES_PER_COMPONENT`` per component).  Excel only
reports the first circular reference it happens to evaluate; this reports the
elementary cycles with a readable path.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Iterable, Iterator, List, Optional, Sequence, Set, Tuple

from .formula_parser import FormulaSyntaxError, NameRef, Node, Range, Ref, iter_nodes, parse_formula
from .models import DefinedName, Workbook
from .refs import format_coordinate

__all__ = ["Cycle", "build_dependency_graph", "dedupe_cycles", "find_cycles"]

NodeKey = Tuple[str, str]

#: Refuse to materialise ranges larger than this; such formulas are usually
#: whole-column references and expanding them would gain nothing.
MAX_RANGE_CELLS = 20_000
MAX_GRAPH_NODES = 200_000
MAX_CYCLES_PER_COMPONENT = 100


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


@dataclass
class _CycleFrame:
    node: NodeKey
    neighbors: Iterator[NodeKey]
    closed: bool = False


def _unblock(
    node: NodeKey,
    blocked: Set[NodeKey],
    block_map: Dict[NodeKey, Set[NodeKey]],
) -> None:
    stack = [node]
    while stack:
        current = stack.pop()
        if current in blocked:
            blocked.remove(current)
            for prev in block_map[current]:
                if prev in blocked:
                    stack.append(prev)
            block_map[current].clear()


def _cycles_in_component(
    component: Sequence[NodeKey], graph: Dict[NodeKey, Set[NodeKey]]
) -> List[Cycle]:
    """Elementary cycles in a strongly connected component (bounded by ``MAX_CYCLES_PER_COMPONENT``)."""
    cycles: List[Cycle] = []
    ordered_members = sorted(component)
    for idx, start in enumerate(ordered_members):
        if start in graph[start]:
            cycles.append(Cycle((start,)))
            if len(cycles) >= MAX_CYCLES_PER_COMPONENT:
                return cycles
        allowed = set(ordered_members[idx:])
        if len(allowed) < 2:
            continue
        start_neighbors = sorted((graph[start] & allowed) - {start})
        if not start_neighbors:
            continue
        blocked: Set[NodeKey] = {start}
        block_map: Dict[NodeKey, Set[NodeKey]] = {u: set() for u in allowed}
        path: List[NodeKey] = [start]
        work: List[_CycleFrame] = [_CycleFrame(start, iter(start_neighbors))]
        while work:
            frame = work[-1]
            nxt = next(frame.neighbors, None)
            if nxt is not None:
                if nxt == start:
                    cycles.append(Cycle(tuple(path)))
                    frame.closed = True
                    if len(cycles) >= MAX_CYCLES_PER_COMPONENT:
                        return cycles
                elif nxt not in blocked:
                    path.append(nxt)
                    blocked.add(nxt)
                    nxt_neighbors = sorted((graph[nxt] & allowed) - {nxt})
                    work.append(_CycleFrame(nxt, iter(nxt_neighbors)))
                continue
            work.pop()
            path.pop()
            if frame.closed:
                _unblock(frame.node, blocked, block_map)
                if work:
                    work[-1].closed = True
            else:
                for neighbor in (graph[frame.node] & allowed) - {frame.node}:
                    block_map[neighbor].add(frame.node)
    return cycles


def find_cycles(workbook: Workbook) -> List[Cycle]:
    """Elementary circular references in ``workbook``, deterministically ordered."""
    graph = build_dependency_graph(workbook)
    cycles: List[Cycle] = []
    for component in _tarjan(graph):
        cycles.extend(_cycles_in_component(component, graph))
    cycles.sort(key=lambda c: [(sheet, cell) for sheet, cell in c.nodes])
    return dedupe_cycles(cycles)


def _canonical_nodes(nodes: Sequence[NodeKey]) -> Tuple[NodeKey, ...]:
    if not nodes:
        return ()
    min_index = min(range(len(nodes)), key=nodes.__getitem__)
    return (*nodes[min_index:], *nodes[:min_index])


def dedupe_cycles(cycles: List[Cycle]) -> List[Cycle]:
    seen: Set[Tuple[NodeKey, ...]] = set()
    out: List[Cycle] = []
    for cycle in cycles:
        key = _canonical_nodes(cycle.nodes)
        if key in seen:
            continue
        seen.add(key)
        out.append(cycle)
    return out
