"""Inserting and deleting rows and columns, and everything that moves.

Neither is a local edit. A cell's address is written into the file in a
dozen places, and every one of them has to move or the workbook is quietly
wrong rather than broken:

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
the operation refuse rather than proceed. Conditional formatting, data
validation, protected ranges and drawings all carry cell addresses this
library does not model yet, and moving the rest while leaving those behind
would produce a file that opens cleanly and highlights the wrong cells. A
refusal naming the element is the honest answer until they are modelled.

Deletion is not insertion run backwards. Three things only it has to do:

- A reference to something that is gone becomes ``#REF!``, while a *range*
  only partly deleted shrinks instead. Both ends of a range are therefore
  decided together rather than one at a time.
- A shared formula lives once, on its group's first cell. Delete that cell
  and the rest point at a formula that is not there, so any group about to
  lose its master is given its own text first.
- A merge that survives as a single cell is not a merge, and a table with
  no columns left is not a table. Both are removed rather than left behind
  as degenerate.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pyofficeeditor._xml import Element
from pyofficeeditor.excel._formulas import (
    Deletion,
    Shift,
    delete_in_formula,
    shift_formula,
    shift_range,
)
from pyofficeeditor.excel._reference import MAX_COLUMN, MAX_ROW, CellRef, RangeRef

if TYPE_CHECKING:
    from pyofficeeditor.excel._tables import Table
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


def delete_rows(sheet: Worksheet, at: int, count: int) -> None:
    """Delete rows, closing the gap and repointing what referred to them."""
    _check_bounds(at, count, MAX_ROW, "row")
    check_shiftable(sheet)
    deletion = Deletion.rows(at, count)
    _check_table_headers(sheet, deletion)
    _expand_orphaned_shared_formulas(sheet, deletion)
    _delete_formulas(sheet, deletion)
    _remove_rows(sheet, at, count)
    _delete_sheet_ranges(sheet, deletion)
    _delete_tables(sheet, deletion)
    _delete_breaks(sheet, deletion)
    _delete_defined_names(sheet, deletion)
    sheet.invalidate()


def delete_columns(sheet: Worksheet, at: int, count: int) -> None:
    """Delete columns, closing the gap and repointing what referred to them."""
    _check_bounds(at, count, MAX_COLUMN, "column")
    check_shiftable(sheet)
    deletion = Deletion.columns(at, count)
    _check_table_headers(sheet, deletion)
    _expand_orphaned_shared_formulas(sheet, deletion)
    _delete_formulas(sheet, deletion)
    _remove_columns(sheet, at, count)
    _delete_column_entries(sheet, at, count)
    _delete_sheet_ranges(sheet, deletion)
    _delete_tables(sheet, deletion)
    _delete_breaks(sheet, deletion)
    _delete_defined_names(sheet, deletion)
    sheet.invalidate()


def _check_table_headers(sheet: Worksheet, deletion: Deletion) -> None:
    """Refuse a deletion that would take a table's header row away.

    A table's column names have to equal the text in its header cells, so
    deleting that row leaves the table naming cells that hold something
    else. Excel does not offer the operation either.
    """
    for table in sheet.tables:
        block = table.ref
        if table.header_row_count and deletion.covers_row(block.top):
            raise ValueError(
                f"deleting row {block.top} would remove the header row of the table "
                f"{table.name!r}, whose column names come from it. Remove the table "
                f"first if that is what you mean."
            )
        if deletion.columns_at is not None and deletion.column_count:
            remaining = [
                column
                for column in range(block.left, block.right + 1)
                if not deletion.covers_column(column)
            ]
            if not remaining:
                raise ValueError(
                    f"deleting these columns would remove every column of the table "
                    f"{table.name!r}. Remove the table first if that is what you mean."
                )


def _expand_orphaned_shared_formulas(sheet: Worksheet, deletion: Deletion) -> None:
    """Give a shared group its own text when its master is about to go.

    A shared formula lives once, on the group's first cell, and the rest
    point at it by index. Delete that cell and the others reference a
    formula that is no longer there. Writing each survivor its own copy
    first costs some size, which Excel reclaims when it next saves, and
    keeps the formulas correct.
    """
    data = sheet.document.root.child("sheetData")
    if data is None:
        return

    doomed: set[str] = set()
    for row in data.children_named("row"):
        for cell in row.children_named("c"):
            formula = cell.child("f")
            reference = cell.get("r")
            if formula is None or reference is None:
                continue
            if (formula.get("t") or "") != "shared" or not formula.text:
                continue
            index = formula.get("si")
            if index is None:
                continue
            try:
                master = CellRef.parse(reference)
            except ValueError:
                continue
            if deletion.moved(master) is None:
                doomed.add(index)

    if not doomed:
        return

    # Every derived formula is read before anything is written. A follower
    # derives from its master's element, so stripping the master's markers
    # while still walking would leave the rest of the group with nothing to
    # derive from, and they would come out blank.
    derived: list[tuple[Element, str | None]] = []
    for row in data.children_named("row"):
        for cell in row.children_named("c"):
            formula = cell.child("f")
            reference = cell.get("r")
            if formula is None or reference is None:
                continue
            if (formula.get("t") or "") != "shared":
                continue
            if (formula.get("si") or "") not in doomed:
                continue
            try:
                own = sheet.get_formula(CellRef.parse(reference))
            except ValueError:
                own = None
            derived.append((formula, own))

    for formula, own in derived:
        formula.unset("t")
        formula.unset("si")
        formula.unset("ref")
        if own:
            formula.set_text(own)
    sheet.invalidate()


def _delete_formulas(sheet: Worksheet, deletion: Deletion) -> None:
    """Rewrite every formula in the workbook for the deletion."""
    for other in sheet.workbook.sheets:
        data = other.document.root.child("sheetData")
        if data is None:
            continue
        for row in data.children_named("row"):
            for cell in row.children_named("c"):
                formula = cell.child("f")
                if formula is None:
                    continue
                text = formula.text
                if not text:
                    continue
                moved = delete_in_formula(
                    text, deletion, formula_sheet=other.name, target_sheet=sheet.name
                )
                if moved != text:
                    formula.set_text(moved)


def _remove_rows(sheet: Worksheet, at: int, count: int) -> None:
    """Drop the deleted rows and pull the ones below them up."""
    data = sheet.document.root.child("sheetData")
    if data is None:
        return
    for number, element in sorted(sheet.rows_by_number().items()):
        if at <= number < at + count:
            data.remove(element)
            continue
        if number >= at + count:
            new_number = number - count
            element.set("r", str(new_number))
            _renumber_cells(element, new_number)
    sheet.reindex_rows()


def _remove_columns(sheet: Worksheet, at: int, count: int) -> None:
    """Drop the deleted columns from every row and pull the rest left."""
    data = sheet.document.root.child("sheetData")
    if data is None:
        return
    for number, element in sheet.rows_by_number().items():
        for cell in list(element.children_named("c")):
            raw = cell.get("r")
            if raw is None:
                continue
            try:
                reference = CellRef.parse(raw)
            except ValueError:
                continue
            column = reference.column
            if at <= column < at + count:
                element.remove(cell)
                continue
            if column >= at + count:
                cell.set("r", CellRef(number, column - count).a1)
        _narrow_span(element, at, count)


def _renumber_cells(row: Element, row_number: int) -> None:
    for cell in row.children_named("c"):
        raw = cell.get("r")
        if raw is None:
            continue
        try:
            reference = CellRef.parse(raw)
        except ValueError:
            continue
        cell.set("r", CellRef(row_number, reference.column).a1)


def _narrow_span(row: Element, at: int, count: int) -> None:
    current = row.get("spans")
    if current is None:
        return
    try:
        first, _, last = current.partition(":")
        low, high = int(first), int(last)
    except ValueError:
        return
    if high >= at:
        row.set("spans", f"{low}:{max(low, high - count)}")


def _delete_column_entries(sheet: Worksheet, at: int, count: int) -> None:
    """Drop and pull the ``<col>`` entries the deletion touched."""
    container = sheet.document.root.child("cols")
    if container is None:
        return
    for entry in list(container.children_named("col")):
        try:
            low = int(entry.get("min") or "0")
            high = int(entry.get("max") or "0")
        except ValueError:
            continue
        if at <= low and high < at + count:
            container.remove(entry)
            continue
        entry.set("min", str(max(1, low - count if low >= at + count else min(low, at))))
        entry.set("max", str(max(1, high - count if high >= at + count else min(high, at - 1) if high >= at else high)))
        if int(entry.get("max") or "0") < int(entry.get("min") or "0"):
            container.remove(entry)
    if next(container.children_named("col"), None) is None:
        sheet.document.root.remove(container)


def _delete_sheet_ranges(sheet: Worksheet, deletion: Deletion) -> None:
    """Shrink or drop the sheet's stored ranges."""
    root = sheet.document.root

    dimension = root.child("dimension")
    if dimension is not None:
        _delete_attribute(dimension, "ref", deletion, drop_if_gone=False)

    merges = root.child("mergeCells")
    if merges is not None:
        for entry in list(merges.children_named("mergeCell")):
            raw = entry.get("ref")
            if raw is None:
                continue
            try:
                block = RangeRef.parse(raw)
            except ValueError:
                continue
            moved = deletion.moved_range(block)
            # A merge that survives as one cell is not a merge any more,
            # which is what Excel does with it too.
            if moved is None or moved.is_single_cell:
                merges.remove(entry)
                continue
            entry.set("ref", moved.a1)
        remaining = sum(1 for _ in merges.children_named("mergeCell"))
        if remaining:
            merges.set("count", str(remaining))
        else:
            root.remove(merges)

    hyperlinks = root.child("hyperlinks")
    if hyperlinks is not None:
        for entry in list(hyperlinks.children_named("hyperlink")):
            if _delete_attribute(entry, "ref", deletion, drop_if_gone=True):
                hyperlinks.remove(entry)
        if next(hyperlinks.children_named("hyperlink"), None) is None:
            root.remove(hyperlinks)

    sheet_filter = root.child("autoFilter")
    if sheet_filter is not None and _delete_attribute(
        sheet_filter, "ref", deletion, drop_if_gone=True
    ):
        root.remove(sheet_filter)

    data = root.child("sheetData")
    if data is not None:
        for row in data.children_named("row"):
            for cell in row.children_named("c"):
                formula = cell.child("f")
                if formula is not None:
                    _delete_attribute(formula, "ref", deletion, drop_if_gone=False)


def _delete_tables(sheet: Worksheet, deletion: Deletion) -> None:
    """Shrink each table, or remove one the deletion emptied."""
    for table in sheet.tables:
        root = table.document.root
        raw = root.get("ref")
        if raw is None:
            continue
        try:
            block = RangeRef.parse(raw)
        except ValueError:
            continue
        moved = deletion.moved_range(block)
        if moved is None:
            sheet.remove_table(table.name)
            continue
        root.set("ref", moved.a1)
        table_filter = root.child("autoFilter")
        if table_filter is not None:
            _delete_attribute(table_filter, "ref", deletion, drop_if_gone=False)
        _drop_deleted_table_columns(table, deletion, block)


def _drop_deleted_table_columns(table: Table, deletion: Deletion, before: RangeRef) -> None:
    """Remove the ``<tableColumn>`` entries whose columns are gone."""
    if deletion.columns_at is None or not deletion.column_count:
        return
    container = table.document.root.child("tableColumns")
    if container is None:
        return
    for offset, entry in enumerate(list(container.children_named("tableColumn"))):
        if deletion.covers_column(before.left + offset):
            container.remove(entry)
    remaining = sum(1 for _ in container.children_named("tableColumn"))
    container.set("count", str(remaining))


def _delete_breaks(sheet: Worksheet, deletion: Deletion) -> None:
    for container_name, at, count in (
        ("rowBreaks", deletion.rows_at, deletion.row_count),
        ("colBreaks", deletion.columns_at, deletion.column_count),
    ):
        if at is None or not count:
            continue
        container = sheet.document.root.child(container_name)
        if container is None:
            continue
        for entry in list(container.children_named("brk")):
            raw = entry.get("id")
            if raw is None:
                continue
            try:
                value = int(raw)
            except ValueError:
                continue
            if at <= value < at + count:
                container.remove(entry)
            elif value >= at + count:
                entry.set("id", str(value - count))


def _delete_defined_names(sheet: Worksheet, deletion: Deletion) -> None:
    workbook = sheet.workbook
    container = workbook.package.xml(workbook.workbook_part).root.child("definedNames")
    if container is None:
        return
    for element in container.children_named("definedName"):
        text = element.text
        if not text:
            continue
        moved = delete_in_formula(
            text, deletion, formula_sheet=sheet.name, target_sheet=sheet.name
        )
        if moved != text:
            element.set_text(moved)


def _delete_attribute(
    element: Element, name: str, deletion: Deletion, *, drop_if_gone: bool
) -> bool:
    """Shrink a stored range.  Returns whether it vanished entirely."""
    raw = element.get(name)
    if raw is None:
        return False
    try:
        block = RangeRef.parse(raw)
    except ValueError:
        return False
    moved = deletion.moved_range(block)
    if moved is None:
        return drop_if_gone
    if moved.a1 != block.a1:
        element.set(name, moved.a1)
    return False


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
    "delete_columns",
    "delete_rows",
    "insert_columns",
    "insert_rows",
]
