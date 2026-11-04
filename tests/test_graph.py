"""Dependency-graph unit tests (the circular-reference rule uses this module)."""

from __future__ import annotations

from pathlib import Path

import pytest
import xlsx_audit.graph_audit as graph_audit
from xlsx_audit.graph_audit import Cycle, build_dependency_graph, dedupe_cycles, find_cycles
from xlsx_audit.ooxml import read_workbook

from factory import build_workbook, simple_sheet


def cycles_for(tmp_path: Path, cells, **kwargs) -> list[Cycle]:
    path = build_workbook(tmp_path / "graph.xlsx", [simple_sheet("S", cells)], **kwargs)
    return find_cycles(read_workbook(path))


def test_graph_edges_between_formulas(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "edges.xlsx",
        [
            simple_sheet(
                "S",
                {
                    "A1": {"v": 1},
                    "A2": {"f": "A1+1", "v": 2},
                    "A3": {"f": "A2+1", "v": 3},
                },
            )
        ],
    )
    graph = build_dependency_graph(read_workbook(path))
    assert graph[("S", "A3")] == {("S", "A2")}
    assert graph[("S", "A2")] == set()  # A1 is a constant, not a formula node


def test_range_edges_are_expanded(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "range.xlsx",
        [
            simple_sheet(
                "S",
                {
                    "A1": {"f": "B1+1", "v": 1},
                    "B1": {"f": "A1+1", "v": 1},
                    "C1": {"f": "SUM(A1:B1)", "v": 2},
                },
            )
        ],
    )
    graph = build_dependency_graph(read_workbook(path))
    assert graph[("S", "C1")] == {("S", "A1"), ("S", "B1")}


def test_cycle_path_is_readable(tmp_path: Path) -> None:
    cycles = cycles_for(
        tmp_path,
        {
            "A1": {"f": "C1+1", "v": 0},
            "B1": {"f": "A1+1", "v": 0},
            "C1": {"f": "B1+1", "v": 0},
        },
    )
    assert len(cycles) == 1
    assert cycles[0].path_text() == "S!A1 -> S!C1 -> S!B1 -> S!A1"
    assert cycles[0].path()[-1] == cycles[0].path()[0]


def test_two_independent_cycles_are_both_reported(tmp_path: Path) -> None:
    cycles = cycles_for(
        tmp_path,
        {
            "A1": {"f": "A2", "v": 0},
            "A2": {"f": "A1", "v": 0},
            "B1": {"f": "B2", "v": 0},
            "B2": {"f": "B1", "v": 0},
        },
    )
    assert len(cycles) == 2


def test_multiple_cycles_in_single_scc_are_all_reported(tmp_path: Path) -> None:
    # A1 <-> B1 and B1 <-> C1 form one strongly connected component {A1, B1, C1}
    # with two elementary cycles: A1 -> B1 -> A1 and B1 -> C1 -> B1.
    cycles = cycles_for(
        tmp_path,
        {
            "A1": {"f": "B1", "v": 0},
            "B1": {"f": "A1+C1", "v": 0},
            "C1": {"f": "B1", "v": 0},
        },
    )
    assert [c.path_text() for c in cycles] == [
        "S!A1 -> S!B1 -> S!A1",
        "S!B1 -> S!C1 -> S!B1",
    ]


def test_figure_eight_cycles_sharing_start_node_are_all_reported(tmp_path: Path) -> None:
    cycles = cycles_for(
        tmp_path,
        {
            "A1": {"f": "B1+C1", "v": 0},
            "B1": {"f": "A1", "v": 0},
            "C1": {"f": "A1", "v": 0},
        },
    )
    assert [c.path_text() for c in cycles] == [
        "S!A1 -> S!B1 -> S!A1",
        "S!A1 -> S!C1 -> S!A1",
    ]


def test_self_reference_and_cycle_in_same_scc_are_both_reported(tmp_path: Path) -> None:
    cycles = cycles_for(
        tmp_path,
        {
            "A1": {"f": "A1+B1", "v": 0},
            "B1": {"f": "A1", "v": 0},
        },
    )
    assert [c.path_text() for c in cycles] == [
        "S!A1 -> S!A1",
        "S!A1 -> S!B1 -> S!A1",
    ]


def test_cycle_limit_per_component(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(graph_audit, "MAX_CYCLES_PER_COMPONENT", 1)
    cycles = cycles_for(
        tmp_path,
        {
            "A1": {"f": "B1+C1", "v": 0},
            "B1": {"f": "A1", "v": 0},
            "C1": {"f": "A1", "v": 0},
        },
    )
    assert [c.path_text() for c in cycles] == ["S!A1 -> S!B1 -> S!A1"]


def test_dedupe_cycles_removes_rotations_and_keeps_distinct_directions() -> None:
    a1, b1, c1 = ("S", "A1"), ("S", "B1"), ("S", "C1")
    forward = Cycle((a1, b1, c1))
    rotated = Cycle((b1, c1, a1))
    reverse = Cycle((a1, c1, b1))
    assert dedupe_cycles([forward, rotated, reverse]) == [forward, reverse]


def test_self_reference_is_a_cycle(tmp_path: Path) -> None:
    cycles = cycles_for(tmp_path, {"A1": {"f": "A1+1", "v": 0}})
    assert len(cycles) == 1
    assert cycles[0].path_text() == "S!A1 -> S!A1"


def test_defined_name_edges_are_followed(tmp_path: Path) -> None:
    cycles = cycles_for(
        tmp_path,
        {"A1": {"f": "TOTAL+1", "v": 0}},
        defined_names=[{"name": "TOTAL", "formula": "S!$A$1"}],
    )
    assert len(cycles) == 1


def test_graph_node_limit(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(graph_audit, "MAX_GRAPH_NODES", 1)
    path = build_workbook(
        tmp_path / "big.xlsx",
        [simple_sheet("S", {"A1": {"f": "A2", "v": 1}, "A2": {"f": "A1", "v": 1}})],
    )
    with pytest.raises(ValueError, match="above the"):
        build_dependency_graph(read_workbook(path))


def test_unparseable_formula_is_skipped(tmp_path: Path) -> None:
    path = build_workbook(tmp_path / "bad.xlsx", [simple_sheet("S", {"A1": {"f": "SUM(", "v": 0}})])
    assert build_dependency_graph(read_workbook(path)) == {("S", "A1"): set()}


def test_external_reference_is_not_an_edge(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "ext.xlsx", [simple_sheet("S", {"A1": {"f": "[1]S!A1", "v": 0}})]
    )
    assert build_dependency_graph(read_workbook(path)) == {("S", "A1"): set()}
