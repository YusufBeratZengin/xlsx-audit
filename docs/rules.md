# Rule reference

Every finding has a stable `code`. Codes are the contract used by `--ignore`, JSON and
SARIF; they never change between minor versions. Messages are presentation only and can
be rendered in English (`en`, default) or Turkish (`tr`, `--lang tr`).

Severity decides the exit code:

- `error` — fails the run, even without `--strict`.
- `warning` — advisory; fails only with `--strict`.
- `info` — shown with `-v`, never fails.

## BROKEN_REFERENCE (error)

A `#REF!` either as a cached cell value or as a literal inside a formula. It almost
always means a row, column or sheet was deleted after the formula was written.

```
Gelirler!E15  =A15+B15   → BROKEN_REFERENCE
Gelirler!C3   =#REF!*1.2 → BROKEN_REFERENCE
```

## FORMULA_ERROR_VALUE (error)

Any other Excel error value: `#VALUE!`, `#DIV/0!`, `#NAME?`, `#NUM!`, `#N/A`, `#NULL!`,
`#SPILL!`, `#CALC!`, `#FIELD!`, `#BLOCKED!`, `#CONNECT!`, `#UNKNOWN!`, `#GETTING_DATA`.
One finding per distinct code per cell.

## CIRCULAR_REFERENCE (error)

The dependency graph (formula cells as nodes, references as edges) contains a strongly
connected component. Reported once per cycle with the full path:

```
S!A1 -> S!B1 -> S!C1 -> S!A1
```

Cross-sheet cycles and cycles through defined names are followed. Whole-column
references and ranges above 20 000 cells are not expanded (see Limitations).

## EXTERNAL_URL (error)

- a saved relationship or worksheet hyperlink whose target is `http://`, `https://` or `ftp://`;
- a URL string literal inside a formula;
- a call to `WEBSERVICE`, `FILTERXML`, `RTD` or `STOCKHISTORY`.

These can exfiltrate cell values at calculation time. Query strings are stripped before
the finding is rendered, stored, or serialized.

## EXTERNAL_LINK (warning)

A link to another workbook, either saved (`xl/externalLinks/*`, worksheet hyperlinks) or written in a formula
(`[1]Sheet1!A1`, `[C:\Users\example\Budget.xlsx]Sheet1!A1`). The target is masked to its
file name; the local directory is never printed or exposed in serialized reports.

## MAGIC_NUMBER_IN_FORMULA (warning)

A numeric literal embedded in a formula that is not structural:

```excel
=B2*0.18        → warning ("0.18")
=A8*1.20        → warning ("1.2")
=SUM(A1:A3)     → no finding
=ROUND(B1,2)    → no finding (2 is a precision argument)
=VLOOKUP(B1,C:D,3,FALSE) → no finding (3 is a column index)
```

Allowed everywhere: `0`, `1`, `2`, `100`. One finding per distinct literal per cell.
The reported `value` is the literal, so an allow-list can be built from the JSON output.

## OMITTED_CELL_WARNING (warning)

An aggregate function (`SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `PRODUCT`,
`MEDIAN`, `STDEV`, `STDEVP`, `VAR`, `VARP`, `SUMPRODUCT`, `AVERAGEIF`) receives a
multi-cell vertical or horizontal range and the cell immediately after the range is
filled with a number that the same formula does not reference anywhere:

```excel
A1:A3 = 10, 20, 30
A4    = 40
B1    = SUM(A1:A3)   → OMITTED_CELL_WARNING, "candidate A4"
```

Conservative by design: it does not fire for text, formulas, empty cells, external
sheets or when the extra cell is referenced elsewhere in the formula.

## VERY_HIDDEN_SHEET (warning) / HIDDEN_SHEET (info)

`state="veryHidden"` cannot be reached from the Excel UI at all, so it is a warning;
`state="hidden"` is informational. Cell contents are not reported.

## HIDDEN_DEFINED_NAME (info)

A named range with `hidden="1"`.

## METADATA_PRESENT (info)

Author, company, manager, title or subject metadata exists. Only the *keys* are printed,
because names are personal data.

## FORMULA_PARSE_ERROR (warning)

The static parser met a construct it does not model. The cell is skipped by the other
rules rather than guessed at; the finding tells you which cell to inspect by hand.

## GRAPH_LIMIT_EXCEEDED (info)

The workbook has more than 200 000 formula cells, so cycle analysis is skipped. Every
other rule still runs.
