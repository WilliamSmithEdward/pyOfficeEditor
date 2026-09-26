"""Sorting a range's rows, as Excel's Sort does.

Measured in Excel through the Sort object its dialog drives:

- Values order by kind, then within it: numbers, text, ``FALSE`` then
  ``TRUE``, and errors, which are all alike; blanks come last. Descending
  turns all but the blanks around. Text follows Excel's collation, and
  with case matched a lowercase letter comes before its capital, from the
  left. Rows that tie keep their order, and a later key decides between
  them.
- A hidden row, whether a filter or a hand hid it, keeps its place, and
  the rows shown are sorted among themselves.
- A row's cells move whole: values, formulas, styles, notes, threads and
  links. A formula moves as a copy does, its relative references shifting
  with it, and off the sheet to ``#REF!``, but a reference that names a
  sheet, even its own, stays as written. Conditional formats, validation
  and row heights stay where they are, and so do references into the range
  from outside it.
- Excel refuses a range holding merged cells, or one whose sorting would
  split an array formula. It keeps a refused sort's settings all the
  same; here a refused sort changes nothing.
- The sort is recorded in the sheet's ``<sortState>``: over the range less
  its header, its right and bottom edges cut back to the sheet's last cell.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyofficeeditor._xml import Element
from pyofficeeditor.excel._collate import CaseCollationKey, CollationKey, case_sort_key, sort_key
from pyofficeeditor.excel._comments import RT_COMMENTS, RT_THREADED_COMMENTS, move_notes
from pyofficeeditor.excel._formulas import sorted_formula
from pyofficeeditor.excel._reference import CellRef, RangeRef, column_index
from pyofficeeditor.excel._rowcol import RT_VML, related_parts
from pyofficeeditor.excel._schema import WORKSHEET_CHILD_ORDER, insert_in_schema_order
from pyofficeeditor.excel._values import CellError, CellValue, read_value

if TYPE_CHECKING:
    from pyofficeeditor.excel._sharedstrings import SharedStrings
    from pyofficeeditor.excel.worksheet import Worksheet

#: The order of a value's kind, ascending: numbers, text, logicals, errors.
_NUMBER, _TEXT, _LOGICAL, _ERROR = range(4)

#: The namespace Excel declares on every ``<sortState>`` it writes.
_XLRD2 = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
#: The most keys Excel's Sort takes.
MAX_SORT_KEYS = 64


@dataclass(frozen=True)
class SortKey:
    """One key of a sort: a column of the range, named by its letter, and
    whether it runs down from the largest."""

    column: str
    descending: bool = False


#: Where a value sorts ascending: its kind's place, then its place among
#: values of that kind.
Rank = tuple[int, float | CollationKey | CaseCollationKey]


def rank(value: CellValue, *, match_case: bool = False) -> Rank:
    """Where ``value`` sorts ascending. It is a raw value, a date its serial
    number, and not blank; errors are all alike, and a logical is a number
    here, so ``FALSE`` comes before ``TRUE``."""
    if isinstance(value, bool):
        return (_LOGICAL, value)
    if isinstance(value, (int, float)):
        return (_NUMBER, value)
    if isinstance(value, CellError):
        return (_ERROR, 0)
    text = str(value)
    return (_TEXT, case_sort_key(text) if match_case else sort_key(text))


def order(rows: Sequence[Sequence[CellValue]], descending: Sequence[bool], *, match_case: bool = False) -> list[int]:
    """The positions of ``rows`` in sorted order: each row its keys' values,
    compared a key at a time, blanks last whichever way a key runs, and rows
    that tie in their order.

    The rows are sorted by the last key, then by each key before it. Each
    pass keeps the order of the rows it ties, so the first key decides and
    the later ones break its ties, as comparing the keys in turn would.
    """
    positions = list(range(len(rows)))
    for index in reversed(range(len(descending))):
        filled = [position for position in positions if rows[position][index] is not None]
        blank = [position for position in positions if rows[position][index] is None]
        ranks = {position: rank(rows[position][index], match_case=match_case) for position in filled}
        # Turned around, a sort still keeps the order of what it ties.
        filled.sort(key=ranks.__getitem__, reverse=descending[index])
        positions = filled + blank
    return positions


# ----------------------------------------------------------------------
# Sorting a sheet's rows
# ----------------------------------------------------------------------


def sort_rows(sheet: Worksheet, block: RangeRef, keys: Sequence[SortKey], *, header: bool, match_case: bool) -> None:
    """Sort the rows of ``block`` on ``sheet`` by ``keys``, the first row
    left where it is when it is a header, as Excel's Sort does."""
    top = block.top + (1 if header else 0)
    columns = [column_index(key.column) for key in keys]
    _refuse(sheet, block)
    last_row, last_column = sheet.max_row, sheet.max_column
    elements = sheet.rows_by_number()
    # A row past the last that holds a cell is blank in every key, so it
    # sorts after the rest in its place and stays there.
    rows = [row for row in range(top, min(block.bottom, last_row) + 1) if not sheet.row_hidden(row)]
    strings = sheet.workbook.shared_strings
    values: list[list[CellValue]] = []
    for row in rows:
        cells = _cells_by_column(elements.get(row))
        values.append([_raw(cells.get(column), strings) for column in columns])
    permutation = order(values, [key.descending for key in keys], match_case=match_case)
    moves = {rows[source]: rows[target] for target, source in enumerate(permutation) if source != target}
    if moves:
        _unshare(sheet, moves, block)
        _move_cells(sheet, moves, block)
        _move_attached(sheet, moves, block)
        sheet.workbook.mark_values_changed()
    _record(sheet, block, top, keys, columns, match_case, last_row, last_column)
    sheet.invalidate()


def _refuse(sheet: Worksheet, block: RangeRef) -> None:
    """Measured: Excel refuses to sort merged cells, or to split an array
    formula, and moves nothing."""
    for merged in sheet.merged_ranges:
        if merged.intersects(block):
            raise ValueError(f"{block.a1} holds the merged cells {merged.a1}, and Excel does not sort merged cells.")
    for _, cell in sheet.cell_elements():
        formula = cell.child("f")
        ref = None if formula is None else formula.get("ref")
        if formula is None or ref is None or formula.get("t") not in ("array", "dataTable"):
            continue
        try:
            spans = RangeRef.parse(ref).normalized
        except ValueError:
            continue
        inside = block.left <= spans.left and spans.right <= block.right
        if spans.intersects(block) and (spans.top != spans.bottom or not inside):
            raise ValueError(f"sorting {block.a1} would split the array formula over {spans.a1}, which Excel refuses.")


def _cells_by_column(row: Element | None) -> dict[int, Element]:
    found: dict[int, Element] = {}
    for cell in () if row is None else row.children_named("c"):
        column = _column(cell)
        if column is not None:
            found[column] = cell
    return found


def _column(cell: Element) -> int | None:
    raw = cell.get("r")
    if raw is None:
        return None
    try:
        return CellRef.parse(raw).column
    except ValueError:
        return None


def _raw(cell: Element | None, strings: SharedStrings | None) -> CellValue:
    """A cell's value as a sort compares it: raw, a date as its serial
    number and a formula as the value cached for it."""
    if cell is None:
        return None
    return read_value(cell, shared_strings=strings, styles=None)


def _moving(moves: dict[int, int], block: RangeRef, row: int, column: int) -> bool:
    return row in moves and block.left <= column <= block.right


def _unshare(sheet: Worksheet, moves: dict[int, int], block: RangeRef) -> None:
    """Give each cell of a shared formula that has a cell moving its own
    text: a shared formula lives once, on its first cell, and the rest are
    read from it by where they stand."""
    groups: set[str] = set()
    members: list[tuple[Element, CellRef]] = []
    for reference, cell in sheet.cell_elements():
        formula = cell.child("f")
        if formula is None or formula.get("t") != "shared" or formula.get("si") is None:
            continue
        members.append((formula, reference))
        if _moving(moves, block, reference.row, reference.column):
            groups.add(formula.get("si") or "")
    # Every derived text is read before any is written, since the followers
    # are read from their first cell's element.
    derived = [(formula, sheet.get_formula(reference)) for formula, reference in members if formula.get("si") in groups]
    for formula, text in derived:
        for marker in ("t", "si", "ref"):
            formula.unset(marker)
        if text:
            formula.set_text(text)
    if derived:
        sheet.invalidate()


def _move_cells(sheet: Worksheet, moves: dict[int, int], block: RangeRef) -> None:
    """Take each moving row's cells in the block out, then put them into
    the row each goes to, their formulas as the sort writes them."""
    elements = sheet.rows_by_number()
    taken: dict[int, list[tuple[int, Element]]] = {}
    for source in moves:
        cells = [
            (column, cell)
            for column, cell in sorted(_cells_by_column(elements.get(source)).items())
            if block.left <= column <= block.right
        ]
        for _, cell in cells:
            parent = cell.parent
            if parent is not None:
                parent.remove(cell)
        taken[source] = cells
    for source, target in moves.items():
        step = target - source
        for column, cell in taken[source]:
            formula = cell.child("f")
            if formula is not None:
                if formula.text:
                    formula.set_text(sorted_formula(formula.text, step))
                ref = formula.get("ref")
                if ref is not None and formula.get("t") in ("array", "dataTable"):
                    span = RangeRef.parse(ref).normalized
                    moved = RangeRef(CellRef(span.top + step, span.left), CellRef(span.bottom + step, span.right))
                    formula.set("ref", moved.a1)
            cell.set("r", CellRef(target, column).a1)
        sheet.place_cells(target, [cell for _, cell in taken[source]])
    # A row left with no cell and nothing else of its own goes, as the
    # worksheet drops one when its last cell is removed.
    elements = sheet.rows_by_number()
    for source in moves:
        row = elements.get(source)
        if (
            row is not None
            and row.parent is not None
            and next(row.children_named("c"), None) is None
            and set(row.attributes) <= {"r", "spans"}
        ):
            row.parent.remove(row)
    sheet.reindex_rows()


def _move_attached(sheet: Worksheet, moves: dict[int, int], block: RangeRef) -> None:
    """Notes, threads and links go with their cells; each part keeps its
    entries in reading order, as Excel writes them."""
    package = sheet.workbook.package
    steps: dict[CellRef, int] = {}
    for name in related_parts(sheet, RT_COMMENTS):
        container = package.xml(name).root.child("commentList")
        if container is not None:
            steps.update(_move_refs(container, "comment", moves, block))
    for name in related_parts(sheet, RT_THREADED_COMMENTS):
        _move_refs(package.xml(name).root, "threadedComment", moves, block)
    if steps:
        for name in related_parts(sheet, RT_VML):
            text = package.read(name).decode("utf-8")
            moved = move_notes(text, steps)
            if moved != text:
                package.write(name, moved.encode("utf-8"))
    links = sheet.document.root.child("hyperlinks")
    for link in () if links is None else list(links.children_named("hyperlink")):
        raw = link.get("ref")
        try:
            span = RangeRef.parse(raw or "").normalized
        except ValueError:
            continue
        if span.top == span.bottom and span.top in moves and block.left <= span.left and span.right <= block.right:
            step = moves[span.top] - span.top
            link.set("ref", RangeRef(CellRef(span.top + step, span.left), CellRef(span.bottom + step, span.right)).a1)


def _move_refs(container: Element, name: str, moves: dict[int, int], block: RangeRef) -> dict[CellRef, int]:
    """Readdress each entry on a moving cell, then keep the entries in
    reading order. Returns how far each moved cell went."""
    steps: dict[CellRef, int] = {}
    entries = list(container.children_named(name))
    for entry in entries:
        try:
            cell = CellRef.parse(entry.get("ref") or "")
        except ValueError:
            continue
        if _moving(moves, block, cell.row, cell.column):
            steps[cell] = moves[cell.row] - cell.row
            entry.set("ref", CellRef(moves[cell.row], cell.column).a1)
    if steps:
        ordered = sorted(entries, key=lambda entry: _reading_order(entry.get("ref")))
        if ordered != entries:
            for entry in entries:
                container.remove(entry)
            for entry in ordered:
                container.append(entry)
    return steps


def _reading_order(raw: str | None) -> tuple[int, int]:
    try:
        return CellRef.parse(raw or "").sort_key
    except ValueError:
        return (0, 0)


def _record(
    sheet: Worksheet,
    block: RangeRef,
    top: int,
    keys: Sequence[SortKey],
    columns: Sequence[int],
    match_case: bool,
    last_row: int,
    last_column: int,
) -> None:
    """The sheet's ``<sortState>`` as Excel's Sort writes it: the range less
    its header, its right and bottom edges cut back to the sheet's last
    cell, and each key over the range's rows as given."""
    root = sheet.document.root
    existing = root.child("sortState")
    if existing is not None:
        root.remove(existing)
    right = max(min(block.right, last_column), block.left)
    bottom = max(min(block.bottom, last_row), top)
    state = Element.create("sortState")
    if match_case:
        state.set("caseSensitive", "1")
    state.set("ref", RangeRef(CellRef(top, block.left), CellRef(bottom, right)).a1)
    state.set("xmlns:xlrd2", _XLRD2)
    for key, column in zip(keys, columns, strict=True):
        condition = Element.create("sortCondition")
        if key.descending:
            condition.set("descending", "1")
        condition.set("ref", RangeRef(CellRef(top, column), CellRef(block.bottom, column)).a1)
        state.append(condition)
    insert_in_schema_order(root, state, WORKSHEET_CHILD_ORDER)


__all__ = ["MAX_SORT_KEYS", "Rank", "SortKey", "order", "rank", "sort_rows"]
