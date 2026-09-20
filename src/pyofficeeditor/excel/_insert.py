"""Inserting rows and columns, and everything that has to move with them.

Inserting a row is not a local edit. A cell's address is written into the
file in a dozen places, and every one of them has to move or the workbook is
quietly wrong rather than broken:

- the ``r`` on each ``<row>`` and each ``<c>`` below the insertion
- every formula on the sheet, and every formula on *other* sheets that
  reads from it, which is why the shift goes through the tokenizer: a bare
  ``A5`` means the formula's own sheet and a qualified ``Data!A5`` does not
- the ``ref`` of each shared-formula group
- merged ranges, hyperlinks, the sheet's autofilter
- each table's ``ref`` and its own autofilter
- the sheet's ``dimension`` and its page breaks
- defined names, at both scopes

Anything on the sheet that addresses cells and is *not* in that list makes
the insertion refuse rather than proceed. Conditional formatting, data
validation, protected ranges and drawings all carry cell addresses this
library does not model yet, and shifting the rest while leaving those behind
would produce a file that opens cleanly and highlights the wrong cells. A
refusal naming the element is the honest answer until they are modelled.

Deletion is not here. It needs one thing insertion does not: a reference to
a row that no longer exists becomes ``#REF!``, and a range only partly
deleted shrinks instead. That is a different transform and belongs in its
own change rather than bolted onto this one.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pyofficeeditor._xml import Element
from pyofficeeditor.excel._formulas import Shift, shift_formula, shift_range
from pyofficeeditor.excel._reference import MAX_COLUMN, MAX_ROW, CellRef, RangeRef

if TYPE_CHECKING:
    from pyofficeeditor.excel.worksheet import Worksheet

#: Elements that carry cell addresses this library cannot shift yet. Finding
#: one makes an insertion refuse, because moving everything else and leaving
#: these behind gives a workbook that opens and is wrong.
UNSHIFTABLE_ELEMENTS: dict[str, str] = {
    "conditionalFormatting": "conditional formatting rules",
    "dataValidations": "data validation rules",
    "protectedRanges": "protected ranges",
    "ignoredErrors": "ignored-error markers",
    "customSheetViews": "custom sheet views",
    "sortState": "a saved sort",
    "dataConsolidate": "a data consolidation",
    "scenarios": "scenarios",
    "drawing": "a drawing, chart or image anchored to cells",
    "oleObjects": "embedded OLE objects",
    "controls": "form controls",
    "picture": "a background picture",
    "extLst": "extension content, which may hold newer rules that address cells",
}


class UnshiftableContentError(ValueError):
    """The sheet carries something addressing cells that cannot be moved."""


def check_shiftable(sheet: Worksheet) -> None:
    """Refuse an insertion the sheet's contents would survive incorrectly."""
    found = [
        description
        for name, description in UNSHIFTABLE_ELEMENTS.items()
        if sheet.document.root.child(name) is not None
    ]
    if found:
        raise UnshiftableContentError(
            f"{sheet.name!r} carries {', '.join(found)}, which address cells this library "
            f"cannot move yet. Shifting everything else and leaving those behind would "
            f"produce a workbook that opens cleanly and points at the wrong cells, so the "
            f"insertion is refused instead."
        )


def insert_rows(sheet: Worksheet, at: int, count: int) -> None:
    """Insert blank rows, pushing everything at or below ``at`` down."""
    _check_bounds(at, count, MAX_ROW, "row")
    check_shiftable(sheet)
    highest = sheet.max_row
    if highest + count > MAX_ROW:
        raise ValueError(
            f"inserting {count} rows at {at} would push row {highest} past {MAX_ROW}, "
            f"the last row Excel has."
        )
    shift = Shift.rows(at, count)
    _shift_cells(sheet, shift)
    _shift_everything_else(sheet, shift)


def insert_columns(sheet: Worksheet, at: int, count: int) -> None:
    """Insert blank columns, pushing everything at or right of ``at`` over."""
    _check_bounds(at, count, MAX_COLUMN, "column")
    check_shiftable(sheet)
    highest = sheet.max_column
    if highest + count > MAX_COLUMN:
        raise ValueError(
            f"inserting {count} columns at {at} would push column {highest} past "
            f"{MAX_COLUMN}, the last column Excel has."
        )
    shift = Shift.columns(at, count)
    _shift_cells(sheet, shift)
    _shift_column_entries(sheet, shift)
    _shift_everything_else(sheet, shift)


def _check_bounds(at: int, count: int, limit: int, what: str) -> None:
    if at < 1 or at > limit:
        raise ValueError(f"{what} {at} is outside 1..{limit}")
    if count < 1:
        raise ValueError(f"the number of {what}s to insert must be at least 1; got {count}")


def _shift_cells(sheet: Worksheet, shift: Shift) -> None:
    """Rewrite the addresses on every affected row and cell.

    Rows are walked from the bottom up, and a row's cells right to left, so
    nothing ever lands on something that has not moved yet.
    """
    if sheet.document.root.child("sheetData") is None:
        return

    moving_columns = shift.columns_at is not None and shift.column_count
    for number, element in sorted(sheet.rows_by_number().items(), reverse=True):
        new_number = number
        if shift.rows_at is not None and shift.row_count and number >= shift.rows_at:
            new_number = number + shift.row_count
            element.set("r", str(new_number))
        if new_number != number or moving_columns:
            _move_cells(element, new_number, shift)

    sheet.reindex_rows()


def _move_cells(row: Element, row_number: int, shift: Shift) -> None:
    """Rewrite each cell's ``r`` in one row."""
    moving_columns = shift.columns_at is not None and shift.column_count
    cells = list(row.children_named("c"))
    if moving_columns:
        cells.reverse()
    for cell in cells:
        raw = cell.get("r")
        if raw is None:
            continue
        try:
            reference = CellRef.parse(raw)
        except ValueError:
            continue
        column = reference.column
        if moving_columns and shift.columns_at is not None and column >= shift.columns_at:
            column += shift.column_count
        cell.set("r", CellRef(row_number, column).a1)

    if moving_columns:
        _widen_span(row, shift.column_count)


def _widen_span(row: Element, column_count: int) -> None:
    """Keep a row's ``spans`` hint covering its cells after a column move."""
    current = row.get("spans")
    if current is None:
        return
    try:
        first, _, last = current.partition(":")
        low, high = int(first), int(last)
    except ValueError:
        return
    row.set("spans", f"{low}:{min(MAX_COLUMN, high + column_count)}")


def _shift_column_entries(sheet: Worksheet, shift: Shift) -> None:
    """Move the ``<col>`` entries for the columns that shifted."""
    if shift.columns_at is None or not shift.column_count:
        return
    container = sheet.document.root.child("cols")
    if container is None:
        return
    for entry in container.children_named("col"):
        for name in ("min", "max"):
            raw = entry.get(name)
            if raw is None:
                continue
            try:
                value = int(raw)
            except ValueError:
                continue
            if value >= shift.columns_at:
                entry.set(name, str(min(MAX_COLUMN, value + shift.column_count)))


def _shift_everything_else(sheet: Worksheet, shift: Shift) -> None:
    """Move every other stored address: formulas anywhere in the workbook,
    ranges on this sheet, and the workbook's defined names."""
    for other in sheet.workbook.sheets:
        _shift_formulas(other, shift, target_sheet=sheet.name)

    _shift_shared_formula_refs(sheet, shift)
    _shift_sheet_ranges(sheet, shift)
    _shift_tables(sheet, shift)
    _shift_breaks(sheet, shift)
    _shift_defined_names(sheet, shift)
    sheet.invalidate()


def _shift_formulas(sheet: Worksheet, shift: Shift, *, target_sheet: str) -> None:
    """Rewrite every formula on one sheet.

    Only the master of a shared-formula group carries text, so only masters
    are rewritten; the followers derive from it and move with it.
    """
    data = sheet.document.root.child("sheetData")
    if data is None:
        return
    for row in data.children_named("row"):
        for cell in row.children_named("c"):
            formula = cell.child("f")
            if formula is None:
                continue
            text = formula.text
            if not text:
                continue
            moved = shift_formula(
                text, shift, formula_sheet=sheet.name, target_sheet=target_sheet
            )
            if moved != text:
                formula.set_text(moved)


def _shift_shared_formula_refs(sheet: Worksheet, shift: Shift) -> None:
    """A shared-formula group records the block it covers."""
    data = sheet.document.root.child("sheetData")
    if data is None:
        return
    for row in data.children_named("row"):
        for cell in row.children_named("c"):
            formula = cell.child("f")
            if formula is not None:
                _shift_attribute(formula, "ref", shift)


def _shift_sheet_ranges(sheet: Worksheet, shift: Shift) -> None:
    """The sheet's own stored ranges: dimension, merges, hyperlinks, filter."""
    root = sheet.document.root

    dimension = root.child("dimension")
    if dimension is not None:
        _shift_attribute(dimension, "ref", shift)

    merges = root.child("mergeCells")
    if merges is not None:
        for entry in merges.children_named("mergeCell"):
            _shift_attribute(entry, "ref", shift)

    hyperlinks = root.child("hyperlinks")
    if hyperlinks is not None:
        for entry in hyperlinks.children_named("hyperlink"):
            _shift_attribute(entry, "ref", shift)

    sheet_filter = root.child("autoFilter")
    if sheet_filter is not None:
        _shift_attribute(sheet_filter, "ref", shift)


def _shift_tables(sheet: Worksheet, shift: Shift) -> None:
    """A table's extent and its own filter, which are separate refs."""
    for table in sheet.tables:
        root = table.document.root
        _shift_attribute(root, "ref", shift)
        table_filter = root.child("autoFilter")
        if table_filter is not None:
            _shift_attribute(table_filter, "ref", shift)


def _shift_breaks(sheet: Worksheet, shift: Shift) -> None:
    """Page breaks address a row or column by number rather than by range."""
    for container_name, at, count in (
        ("rowBreaks", shift.rows_at, shift.row_count),
        ("colBreaks", shift.columns_at, shift.column_count),
    ):
        if at is None or not count:
            continue
        container = sheet.document.root.child(container_name)
        if container is None:
            continue
        for entry in container.children_named("brk"):
            raw = entry.get("id")
            if raw is None:
                continue
            try:
                value = int(raw)
            except ValueError:
                continue
            if value >= at:
                entry.set("id", str(value + count))


def _shift_defined_names(sheet: Worksheet, shift: Shift) -> None:
    """A defined name's target is a formula, so it shifts the same way.

    An unqualified name has no sheet of its own, so the edited sheet stands
    in: Excel resolves such a name against the active sheet, and treating it
    as local is the reading that moves it with the data.
    """
    workbook = sheet.workbook
    container = workbook.package.xml(workbook.workbook_part).root.child("definedNames")
    if container is None:
        return
    for element in container.children_named("definedName"):
        text = element.text
        if not text:
            continue
        moved = shift_formula(
            text, shift, formula_sheet=sheet.name, target_sheet=sheet.name
        )
        if moved != text:
            element.set_text(moved)


def _shift_attribute(element: Element, name: str, shift: Shift) -> None:
    """Move a range stored in one attribute, leaving it alone if unparsable."""
    raw = element.get(name)
    if raw is None:
        return
    try:
        block = RangeRef.parse(raw)
    except ValueError:
        return
    moved = shift_range(block, shift)
    if moved.a1 != block.a1:
        element.set(name, moved.a1)


__all__ = [
    "UNSHIFTABLE_ELEMENTS",
    "UnshiftableContentError",
    "check_shiftable",
    "insert_columns",
    "insert_rows",
]
