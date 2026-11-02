from __future__ import annotations

import pytest
from xlsx_audit.formula_parser import (
    Array,
    BinaryOp,
    BoolLit,
    ErrorLit,
    FormulaSyntaxError,
    FuncCall,
    Missing,
    NameRef,
    Percent,
    Range,
    Ref,
    StringLit,
    UnaryOp,
    canonical_formula,
    collect_numbers,
    collect_refs,
    iter_nodes,
    parse_formula,
    shift_formula,
    to_formula,
)


def test_nested_if_is_split_into_nodes() -> None:
    tree = parse_formula("=IF(SUM(A1:B10)>0, C1*1.18, 0)")
    assert isinstance(tree, FuncCall)
    assert tree.name == "IF"
    comparison = tree.args[0]
    assert isinstance(comparison, BinaryOp) and comparison.op == ">"
    assert isinstance(comparison.left, FuncCall) and comparison.left.name == "SUM"
    assert isinstance(comparison.left.args[0], Range)
    assert [n.value for n in collect_numbers(tree)] == [0.0, 1.18, 0.0]


def test_operator_precedence_and_serialization() -> None:
    tree = parse_formula("=1+2*3")
    assert isinstance(tree, BinaryOp) and tree.op == "+"
    assert to_formula(tree) == "1+2*3"
    assert to_formula(parse_formula("=(1+2)*3")) == "(1+2)*3"


def test_unary_and_percent() -> None:
    tree = parse_formula("=-A1^2+50%")
    assert isinstance(tree, BinaryOp)
    assert isinstance(tree.left, BinaryOp) and isinstance(tree.left.left, UnaryOp)
    assert isinstance(tree.right, Percent)


def test_string_escapes_and_concatenation() -> None:
    tree = parse_formula('=A1&"x""y"')
    assert isinstance(tree, BinaryOp) and tree.op == "&"
    assert isinstance(tree.right, StringLit)
    assert tree.right.value == 'x"y'
    assert to_formula(tree) == 'A1&"x""y"'


def test_error_literals() -> None:
    tree = parse_formula("=#REF!+1")
    assert isinstance(tree, BinaryOp)
    assert isinstance(tree.left, ErrorLit) and tree.left.code == "#REF!"
    assert isinstance(parse_formula("=#DIV/0!"), ErrorLit)


def test_sheet_qualified_and_absolute_refs() -> None:
    tree = parse_formula("=Sheet1!C14+$A$1*A2")
    left = tree.left  # type: ignore[union-attr]
    assert isinstance(left, Ref)
    assert (left.sheet, left.row, left.column) == ("Sheet1", 14, 3)
    middle = tree.right.left  # type: ignore[union-attr]
    assert isinstance(middle, Ref)
    assert (middle.abs_row, middle.abs_column) == (True, True)


def test_quoted_sheet_and_range() -> None:
    tree = parse_formula("='My Sheet'!A1:B2")
    assert isinstance(tree, Range)
    assert tree.start.sheet == "My Sheet"
    assert tree.end.sheet == "My Sheet"
    assert to_formula(tree) == "'My Sheet'!A1:'My Sheet'!B2"


def test_whole_column_range() -> None:
    tree = parse_formula("=SUM(A:A)")
    assert isinstance(tree, FuncCall)
    argument = tree.args[0]
    assert isinstance(argument, Range)
    assert argument.start.row is None
    assert argument.start.column == 1


def test_external_workbook_paths() -> None:
    tree = parse_formula(r"=[C:\Budget.xlsx]Sheet1!A1*2")
    left = tree.left  # type: ignore[union-attr]
    assert isinstance(left, Ref)
    assert left.book == r"C:\Budget.xlsx"
    assert left.is_external
    numbered = parse_formula("=[1]Data!B2")
    assert isinstance(numbered, Ref) and numbered.book == "1"


def test_names_and_structured_references() -> None:
    assert isinstance(parse_formula("=My_Name*2"), BinaryOp)
    tree = parse_formula("=SUM(Table1[Amount])")
    assert isinstance(tree, FuncCall)
    assert isinstance(tree.args[0], NameRef)
    assert tree.args[0].name == "Table1[Amount]"


def test_booleans_and_array_constants() -> None:
    tree = parse_formula('={TRUE,"a",1;FALSE,#N/A,2}')
    assert isinstance(tree, Array)
    assert isinstance(tree.rows[0][0], BoolLit)
    assert tree.rows[0][0].value is True
    assert isinstance(tree.rows[1][1], ErrorLit)


def test_missing_arguments() -> None:
    tree = parse_formula("=IF(A1,,0)")
    assert isinstance(tree, FuncCall)
    assert isinstance(tree.args[1], Missing)
    assert to_formula(tree) == "IF(A1,,0)"


def test_function_without_arguments() -> None:
    tree = parse_formula("=Today()")
    assert isinstance(tree, FuncCall)
    assert tree.name == "TODAY"
    assert tree.args == ()


def test_collect_refs_walks_every_branch() -> None:
    tree = parse_formula("=IF(A1>0,SUM(B1:B3),'Other'!C1)")
    refs = collect_refs(tree)
    kinds = [type(node).__name__ for node in refs]
    assert kinds == ["Ref", "Range", "Ref"]
    assert len(list(iter_nodes(tree))) > 5


@pytest.mark.parametrize(
    "formula",
    [
        "=",
        "=1+",
        "=SUM(A1",
        "=A1(",
        '="unterminated',
        "=#BOGUS!",
        "=1 2",
        "=@nope",
        "='unterminated!A1",
        "=[1",
        "=A1:B2:C3",
        "=SUM(A1,)extra",
    ],
)
def test_syntax_errors_are_reported(formula: str) -> None:
    with pytest.raises(FormulaSyntaxError):
        parse_formula(formula)


def test_canonical_formula_falls_back_to_raw_text() -> None:
    assert canonical_formula("=SUM(A1:A2)") == "SUM(A1:A2)"
    assert canonical_formula("=@@@") == "=@@@"


def test_shift_formula_moves_relative_parts() -> None:
    assert shift_formula("=A1*$B$2+SUM(C1:C3)", 1, 0) == "A2*$B$2+SUM(C2:C4)"
    assert shift_formula("=A1", 0, 1) == "B1"
    assert shift_formula("=$A$1", 5, 5) == "$A$1"


def test_shift_formula_breaks_at_sheet_edge() -> None:
    assert shift_formula("=A1", 0, -1) == "#REF!"
    assert shift_formula("=A2", -5, 0) == "#REF!"


def test_shift_formula_supports_ranges_and_functions() -> None:
    assert shift_formula("=SUM(A1:A3)+ROUND(B1,2)", 2, 1) == "SUM(B3:B5)+ROUND(C3,2)"


def test_to_formula_keeps_arrays_and_missing() -> None:
    tree = parse_formula("={1,2;3,4}")
    assert to_formula(tree) == "{1,2;3,4}"
