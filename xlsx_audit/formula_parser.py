"""A small, dependency-free Excel formula tokenizer and AST builder.

The parser only needs to understand enough of the Excel grammar to reason about
references and literals statically; it never evaluates a formula.  It handles
the constructs that show up in real workbooks: A1 / ``$A$1`` references,
quoted sheet names, external workbook prefixes (``[1]Sheet1!A1``,
``[C:\\Budget.xlsx]Sheet1!A1``), ranges, function calls, array constants,
string escapes, unary operators, percent and comparison operators.

Unknown-but-legal syntax raises :class:`FormulaSyntaxError` instead of silently
returning a half-parsed tree, so callers can report it.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable, Iterator, List, Optional, Sequence, Tuple, cast

from .refs import MAX_COLUMN, MAX_ROW, format_coordinate, parse_coordinate, quote_sheet_name

__all__ = [
    "Array",
    "BinaryOp",
    "BoolLit",
    "ErrorLit",
    "FormulaSyntaxError",
    "FuncCall",
    "Missing",
    "NameRef",
    "Node",
    "NumberLit",
    "Percent",
    "Range",
    "Ref",
    "StringLit",
    "UnaryOp",
    "canonical_formula",
    "collect_numbers",
    "collect_refs",
    "iter_nodes",
    "parse_formula",
    "shift_formula",
    "to_formula",
]

ERROR_CODES = (
    "#GETTING_DATA",
    "#BLOCKED!",
    "#CONNECT!",
    "#UNKNOWN!",
    "#VALUE!",
    "#DIV/0!",
    "#SPILL!",
    "#FIELD!",
    "#CALC!",
    "#NULL!",
    "#NAME?",
    "#REF!",
    "#NUM!",
    "#N/A",
)

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$")
_COL_RE = re.compile(r"^\$?([A-Za-z]{1,3})$")
_NAME_RE = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.\\]*$")
_ADDRESS_CHARS = set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789$")
_NAME_TRAIL_CHARS = set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_.\\")
_WORD_CHARS = _ADDRESS_CHARS | {"_", "."}


class FormulaSyntaxError(ValueError):
    """Raised when a formula cannot be tokenized or parsed."""

    def __init__(self, message: str, position: int = 0, formula: str = "") -> None:
        self.position = position
        self.formula = formula
        location = f" at position {position}"
        if formula:
            location += f" in {formula!r}"
        super().__init__(message + location)


# --------------------------------------------------------------------------- AST


class Node:
    """Base class for formula AST nodes."""

    __slots__ = ()


@dataclass(frozen=True)
class NumberLit(Node):
    value: float
    text: str

    def __post_init__(self) -> None:
        object.__setattr__(self, "text", self.text or _format_number(self.value))


@dataclass(frozen=True)
class StringLit(Node):
    value: str


@dataclass(frozen=True)
class BoolLit(Node):
    value: bool


@dataclass(frozen=True)
class ErrorLit(Node):
    code: str


@dataclass(frozen=True)
class Missing(Node):
    """An omitted function argument, as in ``IF(A1,,0)``."""


@dataclass(frozen=True)
class Ref(Node):
    """A single-cell reference. ``row``/``column`` are ``None`` for whole
    columns/rows such as ``A:A``."""

    sheet: Optional[str] = None
    row: Optional[int] = None
    column: Optional[int] = None
    abs_row: bool = False
    abs_column: bool = False
    book: Optional[str] = None

    @property
    def is_external(self) -> bool:
        return self.book is not None

    @property
    def address(self) -> str:
        if self.row is None and self.column is None:
            return "#REF!"
        if self.row is None:
            assert self.column is not None
            return ("$" if self.abs_column else "") + _column_letters(self.column)
        if self.column is None:
            return ("$" if self.abs_row else "") + str(self.row)
        from .refs import index_to_column

        col = ("$" if self.abs_column else "") + index_to_column(self.column)
        row = ("$" if self.abs_row else "") + str(self.row)
        return col + row

    def with_sheet(self, sheet: Optional[str]) -> Ref:
        if self.sheet is not None:
            return self
        return Ref(
            sheet=sheet,
            row=self.row,
            column=self.column,
            abs_row=self.abs_row,
            abs_column=self.abs_column,
            book=self.book,
        )


@dataclass(frozen=True)
class Range(Node):
    start: Ref
    end: Ref

    @property
    def is_external(self) -> bool:
        return self.start.is_external or self.end.is_external


@dataclass(frozen=True)
class NameRef(Node):
    name: str


@dataclass(frozen=True)
class FuncCall(Node):
    name: str
    args: Tuple[Node, ...] = ()


@dataclass(frozen=True)
class BinaryOp(Node):
    op: str
    left: Node
    right: Node


@dataclass(frozen=True)
class UnaryOp(Node):
    op: str
    operand: Node


@dataclass(frozen=True)
class Percent(Node):
    operand: Node


@dataclass(frozen=True)
class Array(Node):
    rows: Tuple[Tuple[Node, ...], ...] = ()


def _column_letters(column: int) -> str:
    from .refs import index_to_column

    return index_to_column(column)


def _format_number(value: float) -> str:
    if value == int(value) and abs(value) < 1e15:
        return str(int(value))
    return repr(value)


# ---------------------------------------------------------------------- tokenizer


@dataclass(frozen=True)
class Token:
    kind: str
    text: str
    pos: int
    value: object = None


_OPS = ("<>", "<=", ">=", "+", "-", "*", "/", "^", "&", "=", "<", ">", "%", ":")


class Tokenizer:
    def __init__(self, text: str) -> None:
        self.text = text
        self.pos = 0
        self.formula = text

    def error(self, message: str) -> FormulaSyntaxError:
        return FormulaSyntaxError(message, self.pos, self.formula)

    def tokenize(self) -> List[Token]:
        tokens: List[Token] = []
        text = self.text
        length = len(text)
        while self.pos < length:
            char = text[self.pos]
            if char in " \t\r\n":
                self.pos += 1
                continue
            if char == '"':
                tokens.append(self._string())
                continue
            if char == "#":
                tokens.append(self._error_literal())
                continue
            if char.isdigit() or (
                char == "." and self.pos + 1 < length and text[self.pos + 1].isdigit()
            ):
                tokens.append(self._number())
                continue
            if char in "(){};,":
                tokens.append(Token(char, char, self.pos))
                self.pos += 1
                continue
            if char in "+-*/^&%=<>:":
                for op in _OPS:
                    if text.startswith(op, self.pos):
                        tokens.append(Token("OP", op, self.pos))
                        self.pos += len(op)
                        break
                else:  # pragma: no cover - every char above is in _OPS
                    raise self.error(f"unexpected character {char!r}")
                continue
            if char in "'[" or char in _NAME_TRAIL_CHARS or char == "$":
                tokens.append(self._reference())
                continue
            raise self.error(f"unexpected character {char!r}")
        tokens.append(Token("EOF", "", self.pos))
        return tokens

    # -- token helpers ---------------------------------------------------

    def _string(self) -> Token:
        start = self.pos
        self.pos += 1
        value: List[str] = []
        text = self.text
        while self.pos < len(text):
            char = text[self.pos]
            if char == '"':
                if self.pos + 1 < len(text) and text[self.pos + 1] == '"':
                    value.append('"')
                    self.pos += 2
                    continue
                self.pos += 1
                return Token("STRING", text[start : self.pos], start, "".join(value))
            value.append(char)
            self.pos += 1
        raise FormulaSyntaxError("unterminated string literal", start, self.formula)

    def _error_literal(self) -> Token:
        start = self.pos
        for code in ERROR_CODES:
            if self.text.startswith(code, self.pos):
                self.pos += len(code)
                return Token("ERROR", code, start, code)
        raise self.error("unknown error literal")

    def _number(self) -> Token:
        start = self.pos
        match = re.compile(r"(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?").match(self.text, self.pos)
        if not match:  # pragma: no cover - caller guarantees a digit
            raise self.error("invalid number")
        self.pos = match.end()
        raw = match.group(0)
        try:
            value = float(raw)
        except ValueError:  # pragma: no cover - regex only matches numbers
            raise self.error("invalid number") from None
        return Token("NUMBER", raw, start, value)

    def _quoted_sheet(self) -> str:
        start = self.pos
        self.pos += 1
        parts: List[str] = []
        while self.pos < len(self.text):
            char = self.text[self.pos]
            if char == "'":
                if self.pos + 1 < len(self.text) and self.text[self.pos + 1] == "'":
                    parts.append("'")
                    self.pos += 2
                    continue
                self.pos += 1
                return "".join(parts)
            parts.append(char)
            self.pos += 1
        raise FormulaSyntaxError("unterminated sheet name", start, self.formula)

    def _reference(self) -> Token:
        start = self.pos
        book: Optional[str] = None
        sheet: Optional[str] = None
        if self.text[self.pos] == "[":
            close = self.text.find("]", self.pos)
            if close == -1:
                raise self.error("unterminated external workbook reference")
            book = self.text[self.pos + 1 : close]
            self.pos = close + 1

        if self.pos < len(self.text) and self.text[self.pos] == "'":
            sheet = self._quoted_sheet()
            if self.pos >= len(self.text) or self.text[self.pos] != "!":
                raise self.error("expected '!' after sheet name")
            self.pos += 1
        elif self.pos < len(self.text) and self.text[self.pos] in _NAME_TRAIL_CHARS:
            ident_start = self.pos
            while self.pos < len(self.text) and self.text[self.pos] in _NAME_TRAIL_CHARS:
                self.pos += 1
            ident = self.text[ident_start : self.pos]
            if self.pos < len(self.text) and self.text[self.pos] == "!":
                sheet = ident
                self.pos += 1
            elif self.pos < len(self.text) and self.text[self.pos] == "[" and book is None:
                return self._structured(ident, start)
            elif self.pos < len(self.text) and self.text[self.pos] == "(" and book is None:
                self.pos += 1
                return Token("FUNC", ident.upper(), start, ident.upper())
            else:
                self.pos = ident_start
        if self.pos >= len(self.text):
            raise FormulaSyntaxError("incomplete reference", start, self.formula)

        if self.text[self.pos] == "#":
            token = self._error_literal()
            return Token("ERROR", token.text, start, token.value)

        addr_start = self.pos
        while self.pos < len(self.text) and self.text[self.pos] in _WORD_CHARS:
            self.pos += 1
        address = self.text[addr_start : self.pos]
        if not address:
            raise self.error("expected a cell reference")

        kind = _classify_address(address)
        if self.pos < len(self.text) and self.text[self.pos] == "(" and book is None:
            self.pos += 1
            return Token("FUNC", address.upper(), start, address.upper())
        if kind is None:
            if _NAME_RE.match(address) and sheet is None and book is None:
                return Token("NAME", address, start, address)
            raise FormulaSyntaxError(f"invalid reference {address!r}", addr_start, self.formula)
        return Token("REF", self.text[start : self.pos], start, (book, sheet, address, kind))

    def _structured(self, table: str, start: int) -> Token:
        depth = 0
        self.pos += 1
        depth = 1
        while self.pos < len(self.text) and depth:
            if self.text[self.pos] == "[":
                depth += 1
            elif self.text[self.pos] == "]":
                depth -= 1
            self.pos += 1
        if depth:
            raise FormulaSyntaxError("unterminated structured reference", start, self.formula)
        return Token("NAME", self.text[start : self.pos], start, self.text[start : self.pos])


def _classify_address(address: str) -> Optional[str]:
    cell = _CELL_RE.match(address)
    if cell:
        row_text = cell.group(2)
        column_text = cell.group(1)
        if not 1 <= int(row_text) <= MAX_ROW:
            return None
        from .refs import column_to_index

        try:
            column_to_index(column_text)
        except ValueError:
            return None
        return "cell"
    column = _COL_RE.match(address)
    if column:
        from .refs import column_to_index

        try:
            column_to_index(column.group(1))
        except ValueError:
            return None
        return "column"
    return None


# ------------------------------------------------------------------------- parser


#: Excel itself refuses to nest functions deeper than 64 levels, and every level
#: costs roughly ten Python frames here; a higher cap would hit Python's own
#: recursion limit before this check ever fires.
MAX_NESTING = 64


class Parser:
    def __init__(self, tokens: Sequence[Token], formula: str) -> None:
        self.tokens = list(tokens)
        self.index = 0
        self.formula = formula
        self.depth = 0

    # -- cursor ----------------------------------------------------------

    @property
    def current(self) -> Token:
        return self.tokens[self.index]

    def peek(self, offset: int = 1) -> Token:
        position = self.index + offset
        if position >= len(self.tokens):
            return self.tokens[-1]
        return self.tokens[position]

    def advance(self) -> Token:
        token = self.tokens[self.index]
        if token.kind != "EOF":
            self.index += 1
        return token

    def expect(self, kind: str, text: Optional[str] = None) -> Token:
        token = self.current
        if token.kind != kind or (text is not None and token.text != text):
            wanted = text if text is not None else kind
            raise FormulaSyntaxError(
                f"expected {wanted!r} but found {token.text!r}", token.pos, self.formula
            )
        return self.advance()

    def at_op(self, *ops: str) -> bool:
        return self.current.kind == "OP" and self.current.text in ops

    # -- grammar ---------------------------------------------------------

    def parse(self) -> Node:
        node = self.parse_expression()
        if self.current.kind != "EOF":
            token = self.current
            raise FormulaSyntaxError(f"unexpected token {token.text!r}", token.pos, self.formula)
        return node

    def parse_expression(self) -> Node:
        return self._parse_binary(0)

    _LEVELS = (("=", "<>", "<", ">", "<=", ">="), ("&",), ("+", "-"), ("*", "/"), ("^",))

    def _parse_binary(self, level: int) -> Node:
        if level >= len(self._LEVELS):
            return self._parse_unary()
        ops = self._LEVELS[level]
        node = self._parse_binary(level + 1)
        while self.at_op(*ops):
            op = self.advance().text
            right = self._parse_binary(level + 1)
            node = BinaryOp(op, node, right)
        return node

    def _parse_unary(self) -> Node:
        operators: List[str] = []
        while self.at_op("+", "-"):
            operators.append(self.advance().text)
        node = self._parse_postfix()
        for op in reversed(operators):
            node = UnaryOp(op, node)
        return node

    def _parse_postfix(self) -> Node:
        node = self._parse_primary()
        while self.at_op("%"):
            self.advance()
            node = Percent(node)
        return node

    def _parse_primary(self) -> Node:
        token = self.current
        if token.kind == "NUMBER":
            self.advance()
            return NumberLit(float(token.value), token.text)  # type: ignore[arg-type]
        if token.kind == "STRING":
            self.advance()
            return StringLit(str(token.value))
        if token.kind == "ERROR":
            self.advance()
            return ErrorLit(str(token.value))
        if token.kind == "FUNC":
            self.advance()
            return self._nested(lambda: self._parse_call(str(token.value)))
        if token.kind == "NAME":
            self.advance()
            value = str(token.value)
            if value.upper() == "TRUE":
                return BoolLit(True)
            if value.upper() == "FALSE":
                return BoolLit(False)
            return NameRef(value)
        if token.kind == "REF":
            node = self._parse_ref()
            if self.at_op(":"):
                self.advance()
                if self.current.kind != "REF":
                    raise FormulaSyntaxError(
                        "expected a reference after ':'", self.current.pos, self.formula
                    )
                end = self._parse_ref()
                start = node
                if isinstance(start, Ref) and isinstance(end, Ref):
                    end = end.with_sheet(start.sheet)
                    node = Range(start, end)
            return node
        if token.kind == "(":
            self.advance()
            inner = self._nested(self.parse_expression)
            self.expect(")")
            return inner
        if token.kind == "{":
            return self._nested(self._parse_array)
        raise FormulaSyntaxError(f"unexpected token {token.text!r}", token.pos, self.formula)

    def _nested(self, parse: Callable[[], Node]) -> Node:
        self.depth += 1
        if self.depth > MAX_NESTING:
            raise FormulaSyntaxError("formula is nested too deeply", self.current.pos, self.formula)
        try:
            return parse()
        finally:
            self.depth -= 1

    def _parse_ref(self) -> Node:
        token = self.expect("REF")
        book, sheet, address, kind = cast(
            Tuple[Optional[str], Optional[str], str, str], token.value
        )
        if kind == "cell":
            row, column = parse_coordinate(address)
            abs_row, abs_column = _abs_flags(address)
            return Ref(
                sheet=sheet,
                row=row,
                column=column,
                abs_row=abs_row,
                abs_column=abs_column,
                book=book,
            )
        if kind == "column":
            from .refs import column_to_index

            column = column_to_index(address.replace("$", ""))
            return Ref(sheet=sheet, row=None, column=column, abs_column="$" in address, book=book)
        raise FormulaSyntaxError("unsupported reference kind", token.pos, self.formula)

    def _parse_call(self, name: str) -> Node:
        args: List[Node] = []
        if self.current.kind == ")":
            self.advance()
            return FuncCall(name, ())
        while True:
            if self.current.kind in (",", ";", ")"):
                args.append(Missing())
            else:
                args.append(self.parse_expression())
            if self.current.kind in (",", ";"):
                self.advance()
                continue
            self.expect(")")
            break
        return FuncCall(name, tuple(args))

    def _parse_array(self) -> Node:
        self.expect("{")
        rows: List[Tuple[Node, ...]] = []
        current: List[Node] = []
        while True:
            current.append(self.parse_expression())
            if self.current.kind == ",":
                self.advance()
                continue
            if self.current.kind == ";":
                self.advance()
                rows.append(tuple(current))
                current = []
                continue
            self.expect("}")
            rows.append(tuple(current))
            break
        return Array(tuple(rows))


def _abs_flags(address: str) -> Tuple[bool, bool]:
    """Return ``(abs_row, abs_column)`` for an A1-style address."""
    match = re.match(r"^(\$?)[A-Za-z]{1,3}(\$?)[0-9]{1,7}$", address)
    if not match:  # pragma: no cover - validated before calling
        return (False, False)
    return (match.group(2) == "$", match.group(1) == "$")


# ---------------------------------------------------------------------- public API


def parse_formula(formula: str) -> Node:
    """Parse ``formula`` (with or without a leading ``=``) into an AST."""
    text = formula.strip()
    if text.startswith("="):
        text = text[1:]
    if not text:
        raise FormulaSyntaxError("empty formula", 0, formula)
    tokens = Tokenizer(text).tokenize()
    return Parser(tokens, text).parse()


def iter_nodes(node: Node) -> Iterator[Node]:
    """Depth-first walk over every node in the tree."""
    yield node
    if isinstance(node, FuncCall):
        for arg in node.args:
            yield from iter_nodes(arg)
    elif isinstance(node, BinaryOp):
        yield from iter_nodes(node.left)
        yield from iter_nodes(node.right)
    elif isinstance(node, (UnaryOp, Percent)):
        yield from iter_nodes(node.operand)
    elif isinstance(node, Array):
        for row in node.rows:
            for item in row:
                yield from iter_nodes(item)


def collect_numbers(node: Node) -> List[NumberLit]:
    return [n for n in iter_nodes(node) if isinstance(n, NumberLit)]


def collect_refs(node: Node) -> List[Node]:
    """Every :class:`Ref` / :class:`Range` / :class:`NameRef` in the tree."""
    return [n for n in iter_nodes(node) if isinstance(n, (Ref, Range, NameRef))]


def to_formula(node: Node) -> str:
    """Serialize an AST back to formula text (used for shared-formula translation)."""
    return _serialize(node, 0)


_PRECEDENCE = {
    ":": 9,
    "^": 7,
    "*": 6,
    "/": 6,
    "+": 5,
    "-": 5,
    "&": 4,
    "=": 3,
    "<>": 3,
    "<": 3,
    ">": 3,
    "<=": 3,
    ">=": 3,
}


def _precedence(node: Node) -> int:
    if isinstance(node, BinaryOp):
        return _PRECEDENCE[node.op]
    if isinstance(node, (UnaryOp, Percent)):
        return 8
    if isinstance(node, Range):
        return 9
    return 10


def _serialize(node: Node, parent_precedence: int) -> str:
    if isinstance(node, NumberLit):
        return node.text
    if isinstance(node, StringLit):
        return '"' + node.value.replace('"', '""') + '"'
    if isinstance(node, BoolLit):
        return "TRUE" if node.value else "FALSE"
    if isinstance(node, ErrorLit):
        return node.code
    if isinstance(node, Missing):
        return ""
    if isinstance(node, Ref):
        prefix = ""
        if node.book is not None:
            prefix += f"[{node.book}]"
        if node.sheet is not None:
            prefix += quote_sheet_name(node.sheet) + "!"
        return prefix + node.address
    if isinstance(node, Range):
        return _serialize(node.start, 9) + ":" + _serialize(node.end, 9)
    if isinstance(node, NameRef):
        return node.name
    if isinstance(node, FuncCall):
        return f"{node.name}({','.join(_serialize(a, 0) for a in node.args)})"
    if isinstance(node, Array):
        rows = ";".join(",".join(_serialize(cell, 0) for cell in row) for row in node.rows)
        return "{" + rows + "}"
    if isinstance(node, UnaryOp):
        inner = _serialize(node.operand, 8)
        return f"{node.op}{inner}"
    if isinstance(node, Percent):
        return _serialize(node.operand, 8) + "%"
    if isinstance(node, BinaryOp):
        precedence = _PRECEDENCE[node.op]
        left = _serialize(node.left, precedence)
        right = _serialize(node.right, precedence + (0 if node.op == "^" else 1))
        text = f"{left}{node.op}{right}"
        if precedence < parent_precedence:
            return f"({text})"
        return text
    raise TypeError(f"cannot serialize {node!r}")  # pragma: no cover


def shift_formula(formula: str, row_delta: int, column_delta: int) -> str:
    """Translate a shared formula by ``row_delta``/``column_delta`` cells.

    Relative parts move, ``$``-anchored parts stay.  References pushed outside
    the sheet become ``#REF!`` exactly like Excel does.
    """
    tree = parse_formula(formula)
    return to_formula(_shift(tree, row_delta, column_delta))


def shift_node(node: Node, row_delta: int, column_delta: int) -> Node:
    return _shift(node, row_delta, column_delta)


def _shift(node: Node, row_delta: int, column_delta: int) -> Node:
    if isinstance(node, Ref):
        return _shift_ref(node, row_delta, column_delta)
    if isinstance(node, Range):
        start = _shift_ref(node.start, row_delta, column_delta)
        end = _shift_ref(node.end, row_delta, column_delta)
        return Range(start, end)
    if isinstance(node, FuncCall):
        return FuncCall(node.name, tuple(_shift(a, row_delta, column_delta) for a in node.args))
    if isinstance(node, BinaryOp):
        return BinaryOp(
            node.op,
            _shift(node.left, row_delta, column_delta),
            _shift(node.right, row_delta, column_delta),
        )
    if isinstance(node, UnaryOp):
        return UnaryOp(node.op, _shift(node.operand, row_delta, column_delta))
    if isinstance(node, Percent):
        return Percent(_shift(node.operand, row_delta, column_delta))
    if isinstance(node, Array):
        return Array(
            tuple(tuple(_shift(c, row_delta, column_delta) for c in row) for row in node.rows)
        )
    return node


def _shift_ref(ref: Ref, row_delta: int, column_delta: int) -> Ref:
    row = ref.row
    column = ref.column
    if row is not None and not ref.abs_row:
        row += row_delta
    if column is not None and not ref.abs_column:
        column += column_delta
    if (row is not None and not 1 <= row <= MAX_ROW) or (
        column is not None and not 1 <= column <= MAX_COLUMN
    ):
        return Ref(
            sheet=ref.sheet,
            row=None,
            column=None,
            book=ref.book,
        )
    return Ref(
        sheet=ref.sheet,
        row=row,
        column=column,
        abs_row=ref.abs_row,
        abs_column=ref.abs_column,
        book=ref.book,
    )


def canonical_formula(formula: str) -> str:
    """Normalized rendering of a formula, or the original text if unparseable."""
    try:
        return to_formula(parse_formula(formula))
    except FormulaSyntaxError:
        return formula.strip()


def cell_ref_text(ref: Ref) -> str:
    """Human-readable absolute location for a single-cell ref."""
    if ref.row is None or ref.column is None:
        return ref.address
    return (
        f"{ref.sheet}!{format_coordinate(ref.row, ref.column)}"
        if ref.sheet
        else format_coordinate(ref.row, ref.column)
    )
