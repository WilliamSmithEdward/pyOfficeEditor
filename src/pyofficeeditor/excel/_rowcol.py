"""Inserting and deleting rows and columns, and everything that moves.

Neither is a local edit. A cell's address is written into the file in about
twenty places, and every one of them has to move or the workbook is quietly
wrong rather than broken. In the worksheet:

- the ``r`` on each ``<row>`` and each ``<c>`` below the insertion
- every formula on the sheet, and every formula on *other* sheets that
  reads from it, which is why the shift goes through the tokenizer: a bare
  ``A5`` means the formula's own sheet and a qualified ``Data!A5`` does not
- the ``ref`` of each shared-formula group
- merged ranges, hyperlinks, the sheet's autofilter
- each table's ``ref`` and its own autofilter, and the columns an insertion
  inside a table adds to it, which a table has to list
- the sheet's ``dimension`` and its page breaks
- defined names, at both scopes
- every conditional formatting block, and the compatibility formula that
  makes its rules fire
- data validation, including the formulas that hold its bounds
- protected ranges, ignored errors, a saved sort, a data consolidation
- each scenario's input cells
- every custom sheet view's own selection, pane, autofilter and breaks
- the inline anchors of form controls and embedded objects
- the ``xm:sqref`` and ``xm:f`` of anything in ``extLst``

And in parts the worksheet only points at:

- a drawing's ``<xdr:from>``/``<xdr:to>``, holding **zero-based** indices
- a VML ``<x:Anchor>``, and the ``<x:Row>``/``<x:Column>`` beside it that
  names a comment's own cell
- each comment's ``ref`` in the comments part

And in every chart in the workbook, since a chart on any sheet, or on a
chart sheet, may read from this one: each ``<c:f>`` a series or a title
reads from, moved as a cell's formula is, measured.

And in pivot tables: the location of each on the sheet, and the source of
every cache that reads from it. An edit that cuts through a pivot table is
refused, as Excel refuses it, and one that takes all of a pivot table
deletes it, with its cache if nothing else reads from that.

Nothing is refused. There used to be a list of elements that made the
operation raise rather than risk moving everything else and leaving them
behind, which was the honest answer while they were unmodelled. It is empty
now.

Two of those entries were wrong in opposite directions. A background
``<picture>`` was refused although it tiles the sheet and names no cell at
all. Comments and legacy drawings were never on the list, so a sheet with a
comment shifted its cells and left the comment where it was.

Deletion is not insertion run backwards. Three things only it has to do:

- A reference to something that is gone becomes ``#REF!``, while a *range*
  only partly deleted shrinks instead. Both ends of a range are therefore
  decided together rather than one at a time.
- A shared formula lives once, on its group's first cell. Delete that cell
  and the rest point at a formula that is not there, so any group about to
  lose its master is given its own text first.
- A merge that survives as a single cell is not a merge, a table with no
  columns left is not a table, and a wrapper whose last entry went is not
  something Excel accepts. All three are removed rather than left behind as
  degenerate.
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from dataclasses import replace
from typing import TYPE_CHECKING

from pyofficeeditor._xml import Element
from pyofficeeditor.excel._addresses import (
    collapse_index,
    delete_cell,
    delete_index,
    delete_ref,
    delete_sqref,
    delete_vml_anchor,
    shift_cell,
    shift_index,
    shift_ref,
    shift_sqref,
    shift_vml_anchor,
)
from pyofficeeditor.excel._comments import RT_COMMENTS, RT_THREADED_COMMENTS, note_shapes
from pyofficeeditor.excel._conditional import (
    CONDITION_FORMULA_TYPES,
    ConditionalFormatting,
    ConditionalRule,
)
from pyofficeeditor.excel._formulas import (
    Deletion,
    Shift,
    delete_in_formula,
    shift_formula,
    shift_range,
)
from pyofficeeditor.excel._pivots import (
    RT_PIVOT_CACHE,
    RT_PIVOT_RECORDS,
    RT_PIVOT_TABLE,
    PivotTable,
    cache_parts,
    read_pivot_tables,
    worksheet_source,
)
from pyofficeeditor.excel._reference import MAX_COLUMN, MAX_ROW, CellRef, RangeRef
from pyofficeeditor.excel._xstring import decode, encode_attribute, escape
from pyofficeeditor.exceptions import PackageError

if TYPE_CHECKING:
    from pyofficeeditor.excel._tables import Table
    from pyofficeeditor.excel.worksheet import Worksheet
    from pyofficeeditor.opc import OpcPackage

def insert_rows(sheet: Worksheet, at: int, count: int) -> None:
    """Insert blank rows, pushing everything at or below ``at`` down."""
    _check_bounds(at, count, MAX_ROW, "row")
    # Every row the sheet records moves, a hidden or sized one with no cell
    # as much as one with data, so the last of those is what must fit.
    highest = max(sheet.rows_by_number(), default=0)
    if highest + count > MAX_ROW:
        raise ValueError(
            f"inserting {count} rows at {at} would push row {highest} past {MAX_ROW}, "
            f"the last row Excel has."
        )
    shift = Shift.rows(at, count)
    _check_pivot_tables(sheet, shift=shift)
    _shift_cells(sheet, shift)
    _shift_everything_else(sheet, shift)


def insert_columns(sheet: Worksheet, at: int, count: int) -> None:
    """Insert blank columns, pushing everything at or right of ``at`` over."""
    _check_bounds(at, count, MAX_COLUMN, "column")
    highest = sheet.max_column
    if highest + count > MAX_COLUMN:
        raise ValueError(
            f"inserting {count} columns at {at} would push column {highest} past "
            f"{MAX_COLUMN}, the last column Excel has."
        )
    shift = Shift.columns(at, count)
    _check_pivot_tables(sheet, shift=shift)
    _shift_cells(sheet, shift)
    _shift_column_entries(sheet, shift)
    _shift_everything_else(sheet, shift)


def delete_rows(sheet: Worksheet, at: int, count: int) -> None:
    """Delete rows, closing the gap and repointing what referred to them."""
    _check_bounds(at, count, MAX_ROW, "row")
    deletion = Deletion.rows(at, count)
    _check_table_headers(sheet, deletion)
    _check_pivot_tables(sheet, deletion=deletion)
    _expand_orphaned_shared_formulas(sheet, deletion)
    _delete_formulas(sheet, deletion)
    _remove_rows(sheet, at, count)
    _delete_sheet_ranges(sheet, deletion)
    _delete_conditional_formats(sheet, deletion)
    _delete_sheet_addresses(sheet, deletion)
    _move_custom_views(sheet, shift=None, deletion=deletion)
    _move_extensions(sheet, shift=None, deletion=deletion)
    _move_related_parts(sheet, shift=None, deletion=deletion)
    _move_charts(sheet, shift=None, deletion=deletion)
    _move_pivot_tables(sheet, shift=None, deletion=deletion)
    _move_pivot_caches(sheet, shift=None, deletion=deletion)
    _delete_tables(sheet, deletion)
    _delete_breaks(sheet, deletion)
    _delete_defined_names(sheet, deletion)
    sheet.invalidate()


def delete_columns(sheet: Worksheet, at: int, count: int) -> None:
    """Delete columns, closing the gap and repointing what referred to them."""
    _check_bounds(at, count, MAX_COLUMN, "column")
    deletion = Deletion.columns(at, count)
    _check_table_headers(sheet, deletion)
    _check_pivot_tables(sheet, deletion=deletion)
    _expand_orphaned_shared_formulas(sheet, deletion)
    _delete_formulas(sheet, deletion)
    _remove_columns(sheet, at, count)
    _delete_column_entries(sheet, at, count)
    _delete_sheet_ranges(sheet, deletion)
    _delete_conditional_formats(sheet, deletion)
    _delete_sheet_addresses(sheet, deletion)
    _move_custom_views(sheet, shift=None, deletion=deletion)
    _move_extensions(sheet, shift=None, deletion=deletion)
    _move_related_parts(sheet, shift=None, deletion=deletion)
    _move_charts(sheet, shift=None, deletion=deletion)
    _move_pivot_tables(sheet, shift=None, deletion=deletion)
    _move_pivot_caches(sheet, shift=None, deletion=deletion)
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


def _swap(parent: Element, old: Element, new: Element) -> None:
    """Put ``new`` where ``old`` sits, keeping the position.

    Conditional formatting has to stay between ``mergeCells`` and
    ``dataValidations``, and appending a rebuilt block would move it.
    """
    parent.insert_before(old, new)
    parent.remove(old)


def _delete_conditional_formats(sheet: Worksheet, deletion: Deletion) -> None:
    """Shrink each block's ranges, and drop a block with nothing left.

    A range that only partly overlaps the deletion shrinks, exactly as a
    merge or a formula's range does. A block whose every range is gone is
    removed rather than left behind with an empty ``sqref``, which Excel
    treats as malformed.
    """
    root = sheet.document.root
    for element in list(root.children_named("conditionalFormatting")):
        block = ConditionalFormatting.read(element)
        if not block.ranges:
            continue
        moved = tuple(
            survivor
            for area in block.ranges
            if (survivor := deletion.moved_range(area)) is not None
        )
        if not moved:
            root.remove(element)
            continue
        rules = tuple(_delete_in_rule(rule, deletion, sheet=sheet) for rule in block.rules)
        rebuilt = ConditionalFormatting(ranges=moved, rules=rules)
        anchored = tuple(rule.anchored_at(rebuilt.anchor) for rule in rebuilt.rules)
        _swap(root, element, ConditionalFormatting(ranges=moved, rules=anchored).write())


def _delete_in_rule(
    rule: ConditionalRule, deletion: Deletion, *, sheet: Worksheet
) -> ConditionalRule:
    """Break or shrink the references inside a rule's own condition.

    A rule holds its formulas as they read, and the formula code works on
    them as a part stores them, so each is escaped for it and read back.
    """
    if rule.kind not in CONDITION_FORMULA_TYPES or not rule.formulas:
        return rule
    return replace(
        rule,
        formulas=tuple(
            decode(delete_in_formula(escape(text), deletion, formula_sheet=sheet.name, target_sheet=sheet.name))
            for text in rule.formulas
        ),
    )


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
    if sheet_filter is not None and _move_filter(sheet_filter, shift=None, deletion=deletion):
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
            _move_filter(table_filter, shift=None, deletion=deletion)
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


CT_PIVOT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"


def _check_pivot_tables(sheet: Worksheet, *, shift: Shift | None = None, deletion: Deletion | None = None) -> None:
    """Refuse an edit that cuts through a pivot table, as Excel refuses it,
    measured: inserting inside one, or deleting part of one. Inserting at
    its first row or column moves it whole, and deleting all of it deletes
    it."""
    for pivot in read_pivot_tables(sheet.workbook.package, sheet.part_name):
        block = pivot.location
        if shift is not None:
            rows = shift.rows_at is not None and shift.row_count and block.top < shift.rows_at <= block.bottom
            columns = (
                shift.columns_at is not None and shift.column_count and block.left < shift.columns_at <= block.right
            )
            if rows or columns:
                raise ValueError(
                    f"inserting there would cut through the pivot table {pivot.name!r} at {block.a1}, and "
                    f"Excel refuses that too. Insert at its first row or column, or past its last."
                )
        if deletion is not None and _cuts(block, deletion):
            raise ValueError(
                f"deleting that would take part of the pivot table {pivot.name!r} at {block.a1}, and Excel "
                f"refuses that too. Delete all of its rows or columns to delete it, or none of them."
            )


def _cuts(block: RangeRef, deletion: Deletion) -> bool:
    """Whether a deletion takes part of a block and not all of it."""
    if deletion.rows_at is not None and deletion.row_count:
        covered = sum(1 for row in range(block.top, block.bottom + 1) if deletion.covers_row(row))
        return 0 < covered < block.height
    if deletion.columns_at is not None and deletion.column_count:
        covered = sum(1 for column in range(block.left, block.right + 1) if deletion.covers_column(column))
        return 0 < covered < block.width
    return False


def _move_pivot_tables(sheet: Worksheet, *, shift: Shift | None, deletion: Deletion | None) -> None:
    """Move each pivot table on the sheet with its cells, or delete one whose
    every row or column went, as Excel does."""
    package = sheet.workbook.package
    for pivot in read_pivot_tables(package, sheet.part_name):
        location = package.xml(pivot.part_name).root.child("location")
        if location is None:
            continue
        if shift is not None:
            moved: RangeRef | None = shift_range(pivot.location, shift)
        elif deletion is not None:
            moved = deletion.moved_range(pivot.location)
        else:
            continue
        if moved is None:
            _remove_pivot_table(sheet, pivot)
        elif moved.a1 != location.get("ref"):
            location.set("ref", moved.a1)


def _move_pivot_caches(sheet: Worksheet, *, shift: Shift | None, deletion: Deletion | None) -> None:
    """Move the source of every cache that reads from this sheet, as a
    formula's range moves. One deleted outright keeps its address, as
    Excel keeps it, measured."""
    workbook = sheet.workbook
    package = workbook.package
    for part in cache_parts(package, workbook.workbook_part):
        source = worksheet_source(package.xml(part).root)
        if source is None or decode(source.get("sheet") or "") != sheet.name:
            continue
        raw = source.get("ref")
        try:
            block = RangeRef.parse(raw or "")
        except ValueError:
            continue
        if shift is not None:
            moved: RangeRef | None = shift_range(block, shift)
        elif deletion is not None:
            moved = deletion.moved_range(block)
        else:
            continue
        if moved is not None and moved.a1 != raw:
            source.set("ref", moved.a1)


def _remove_pivot_table(sheet: Worksheet, pivot: PivotTable) -> None:
    """Take a pivot table out, and its cache with it if no other pivot table
    reads from it, as Excel drops one, measured."""
    workbook = sheet.workbook
    package = workbook.package
    relationships = package.relationships(sheet.part_name)
    for relationship in list(relationships.by_type(RT_PIVOT_TABLE)):
        if relationship.target_part == pivot.part_name:
            relationships.remove(relationship.id)
    package.remove_part(pivot.part_name)

    cache = pivot.cache_part
    if cache is None or not package.has_part(cache) or _cache_in_use(package, cache):
        return
    root = package.xml(workbook.workbook_part).root
    listed = root.child("pivotCaches")
    book_relationships = package.relationships(workbook.workbook_part)
    for relationship in list(book_relationships.by_type(RT_PIVOT_CACHE)):
        if relationship.target_part != cache:
            continue
        if listed is not None:
            for entry in list(listed.children_named("pivotCache")):
                if entry.get("r:id") == relationship.id:
                    listed.remove(entry)
        book_relationships.remove(relationship.id)
    if listed is not None and next(listed.children_named("pivotCache"), None) is None:
        # CT_PivotCaches needs an entry; an empty list is not allowed.
        root.remove(listed)
    records = [
        relationship.target_part
        for relationship in package.relationships(cache).by_type(RT_PIVOT_RECORDS)
        if not relationship.is_external
    ]
    package.remove_part(cache)
    for part in records:
        if package.has_part(part):
            package.remove_part(part)


def _cache_in_use(package: OpcPackage, cache: str) -> bool:
    """Whether any pivot table in the workbook reads from a cache."""
    types = package.content_types
    for name in package.part_names():
        if types.of(name) != CT_PIVOT_TABLE:
            continue
        if any(relationship.target_part == cache for relationship in package.relationships(name).by_type(RT_PIVOT_CACHE)):
            return True
    return False


#: The content types of the parts a chart lives in: DrawingML charts, and
#: the newer kinds, such as a waterfall, which Excel keeps in chartex parts.
CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
CT_CHART_EX = "application/vnd.ms-office.chartex+xml"


def chart_parts(package: OpcPackage) -> list[str]:
    """Every chart part in the workbook, wherever it is shown: a chart on
    any sheet, or on a chart sheet, may read from any sheet."""
    types = package.content_types
    return [name for name in package.part_names() if types.of(name) in (CT_CHART, CT_CHART_EX)]


def _move_charts(sheet: Worksheet, *, shift: Shift | None, deletion: Deletion | None) -> None:
    """Move the references in every chart that reads from this sheet.

    Measured, a chart's references behave as a cell's: a series grows with
    rows inserted inside it, moves with rows inserted above it, shrinks as
    rows inside it go, and one wholly deleted is written ``Data!#REF!``, as
    is a title linked to a deleted cell. Excel's object model goes on
    reporting the old address for a deleted one; the file does not. Each
    reference names its sheet, so a chart has no sheet of its own here.
    """
    package = sheet.workbook.package
    for part in chart_parts(package):
        for element in package.xml(part).root.descendants("f"):
            text = element.text
            if not text or "!" not in text:
                continue
            if shift is not None:
                moved = shift_formula(text, shift, formula_sheet="", target_sheet=sheet.name)
            elif deletion is not None:
                moved = delete_in_formula(text, deletion, formula_sheet="", target_sheet=sheet.name)
            else:
                moved = text
            if moved != text:
                element.set_text(moved)


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
    _shift_conditional_formats(sheet, shift)
    _shift_sheet_addresses(sheet, shift)
    _move_custom_views(sheet, shift=shift, deletion=None)
    _move_extensions(sheet, shift=shift, deletion=None)
    _move_related_parts(sheet, shift=shift, deletion=None)
    _move_charts(sheet, shift=shift, deletion=None)
    _move_pivot_tables(sheet, shift=shift, deletion=None)
    _move_pivot_caches(sheet, shift=shift, deletion=None)
    _shift_tables(sheet, shift)
    _shift_breaks(sheet, shift)
    _shift_defined_names(sheet, shift)
    sheet.invalidate()


def _shift_conditional_formats(sheet: Worksheet, shift: Shift) -> None:
    """Move each block's ranges, and the formulas that follow them.

    Three things move together, and missing any one leaves rules that
    highlight the wrong cells while the file opens cleanly:

    - every range in the ``sqref``
    - the condition of a ``cellIs`` or ``expression`` rule, which is a real
      formula and contains real references
    - the compatibility formula of the attribute-driven rules, which is
      rebuilt from the block's new anchor rather than shifted, because it
      names the top-left of the range by construction
    """
    root = sheet.document.root
    for element in list(root.children_named("conditionalFormatting")):
        block = ConditionalFormatting.read(element)
        if not block.ranges:
            continue
        moved = tuple(shift_range(area, shift) for area in block.ranges)
        rules = tuple(_shift_rule(rule, shift, sheet=sheet) for rule in block.rules)
        rebuilt = ConditionalFormatting(ranges=moved, rules=rules)
        anchored = tuple(rule.anchored_at(rebuilt.anchor) for rule in rebuilt.rules)
        _swap(root, element, ConditionalFormatting(ranges=moved, rules=anchored).write())


def _shift_rule(rule: ConditionalRule, shift: Shift, *, sheet: Worksheet) -> ConditionalRule:
    """Move the references inside a rule's own condition, on the formula as
    a part stores it; see :func:`_delete_in_rule`."""
    if rule.kind not in CONDITION_FORMULA_TYPES or not rule.formulas:
        return rule
    return replace(
        rule,
        formulas=tuple(
            decode(shift_formula(escape(text), shift, formula_sheet=sheet.name, target_sheet=sheet.name))
            for text in rule.formulas
        ),
    )


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
        _move_filter(sheet_filter, shift=shift, deletion=None)


def _shift_tables(sheet: Worksheet, shift: Shift) -> None:
    """A table's extent and its own filter, which are separate refs, and
    the columns an insertion inside a table adds to it."""
    for table in sheet.tables:
        root = table.document.root
        try:
            before = table.ref
        except ValueError:
            before = None
        _shift_attribute(root, "ref", shift)
        table_filter = root.child("autoFilter")
        if table_filter is not None:
            _move_filter(table_filter, shift=shift, deletion=None)
        at = shift.columns_at
        if before is not None and at is not None and before.left < at <= before.right:
            _add_table_columns(sheet, table, before, at - before.left, shift.column_count)


def _add_table_columns(sheet: Worksheet, table: Table, before: RangeRef, offset: int, count: int) -> None:
    """Take columns inserted inside a table into it, as Excel does.

    Each gets a ``<tableColumn>`` of its own, named ``ColumnN`` with the
    smallest ``N`` free and numbered after the highest id, and its header
    cell says the name. Measured, all of it; a table wider than its list of
    columns is one Excel refuses to open.
    """
    from pyofficeeditor.excel._tables import unique_column_names

    container = table.document.root.child("tableColumns")
    if container is None:
        return
    entries = list(container.children_named("tableColumn"))
    names = [decode(entry.get("name") or "") for entry in entries]
    resolved = unique_column_names(names[:offset] + [""] * count + names[offset:])
    identifiers = [int(raw) for raw in (entry.get("id") for entry in entries) if raw is not None and raw.isdigit()]
    first_id = max(identifiers, default=0) + 1
    anchor = entries[offset] if offset < len(entries) else None
    for index in range(count):
        name = resolved[offset + index]
        created = Element.create("tableColumn", {"id": str(first_id + index), "name": encode_attribute(name)})
        if anchor is None:
            container.append(created)
        else:
            container.insert_before(anchor, created)
        if table.header_row_count:
            sheet.set_value(CellRef(before.top, before.left + offset + index), name)
    container.set("count", str(len(entries) + count))


def _move_filter(element: Element, *, shift: Shift | None, deletion: Deletion | None) -> bool:
    """Move an ``<autoFilter>``: its range, and each filter column.

    A filter column is an offset from the range's left edge, so columns
    coming or going inside the range renumber it, as Excel renumbers it; a
    column whose own column is deleted goes, criterion and all. Returns
    whether the whole range was deleted.
    """
    raw = element.get("ref")
    if raw is None:
        return False
    try:
        before = RangeRef.parse(raw).normalized
    except ValueError:
        return False
    if shift is not None:
        after = shift_range(before, shift)
        mover: Shift | Deletion = shift
    else:
        if deletion is None:
            return False
        moved = deletion.moved_range(before)
        if moved is None:
            return True
        after, mover = moved, deletion
    for entry in list(element.children_named("filterColumn")):
        try:
            offset = int(entry.get("colId") or "0")
        except ValueError:
            continue
        landed = mover.moved(CellRef(before.top, before.left + offset))
        if landed is None:
            element.remove(entry)
            continue
        if landed.column - after.left != offset:
            entry.set("colId", str(landed.column - after.left))
    if after.a1 != before.a1:
        element.set("ref", after.a1)
    return False


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
    "delete_columns",
    "delete_rows",
    "insert_columns",
    "insert_rows",
]


#: Where a worksheet records a cell address outside the cells themselves,
#: as (container, entry, attribute, notation). Driving the shifting from a
#: table rather than a function per feature is what keeps this list
#: reviewable: adding a feature means adding a row.
#:
#: ``container`` is the wrapper element, or ``None`` when the entries sit
#: directly under ``<worksheet>``. An entry whose address is wholly deleted
#: is removed, which is why the notation matters: an ``sqref`` naming
#: several ranges survives as long as one of them does.
SHEET_ADDRESSES: tuple[tuple[str | None, str, str, str], ...] = (
    ("dataValidations", "dataValidation", "sqref", "sqref"),
    ("protectedRanges", "protectedRange", "sqref", "sqref"),
    ("ignoredErrors", "ignoredError", "sqref", "sqref"),
    (None, "sortState", "ref", "ref"),
    ("sortState", "sortCondition", "ref", "ref"),
    ("dataConsolidate", "dataRef", "ref", "ref"),
    ("scenarios", "scenario", "sqref", "sqref"),
)

#: The same, for entries whose address is a single cell.
SHEET_CELL_ADDRESSES: tuple[tuple[str, str, str], ...] = (
    ("scenarios/scenario", "inputCells", "r"),
)


def _address_entries(
    root: Element, container: str | None, entry: str
) -> list[tuple[Element, Element]]:
    """Every element a row of :data:`SHEET_ADDRESSES` points at."""
    if container is None:
        found = root.child(entry)
        return [] if found is None else [(root, found)]
    holder = root.child(container)
    if holder is None:
        return []
    if entry == container:
        return [(root, holder)]
    return [(holder, node) for node in list(holder.children_named(entry))]


def _shift_sheet_addresses(sheet: Worksheet, shift: Shift) -> None:
    """Move every stored address the table names."""
    root = sheet.document.root
    for container, entry, attribute, notation in SHEET_ADDRESSES:
        for _parent, element in _address_entries(root, container, entry):
            raw = element.get(attribute)
            if raw is None:
                continue
            moved = shift_sqref(raw, shift) if notation == "sqref" else shift_ref(raw, shift)
            element.set(attribute, moved)

    for path, entry, attribute in SHEET_CELL_ADDRESSES:
        for element in _nested(root, path, entry):
            raw = element.get(attribute)
            if raw is not None:
                element.set(attribute, shift_cell(raw, shift))

    _shift_inline_anchors(root, shift)
    _move_validation_formulas(sheet, shift=shift, deletion=None)


def _delete_sheet_addresses(sheet: Worksheet, deletion: Deletion) -> None:
    """Shrink every stored address, and drop what is left with nothing."""
    root = sheet.document.root
    for container, entry, attribute, notation in SHEET_ADDRESSES:
        for parent, element in _address_entries(root, container, entry):
            raw = element.get(attribute)
            if raw is None:
                continue
            moved = (
                delete_sqref(raw, deletion)
                if notation == "sqref"
                else delete_ref(raw, deletion)
            )
            if moved is None:
                parent.remove(element)
                continue
            element.set(attribute, moved)
        _drop_if_empty(root, container, entry)

    for path, entry, attribute in SHEET_CELL_ADDRESSES:
        for element in _nested(root, path, entry):
            raw = element.get(attribute)
            if raw is None:
                continue
            moved = delete_cell(raw, deletion)
            parent = element.parent
            if moved is None:
                if parent is not None:
                    parent.remove(element)
                continue
            element.set(attribute, moved)

    _delete_inline_anchors(root, deletion)
    _move_validation_formulas(sheet, shift=None, deletion=deletion)


def _nested(root: Element, path: str, entry: str) -> list[Element]:
    """Every ``entry`` under a slash-separated path of single children."""
    current: list[Element] = [root]
    for name in path.split("/"):
        following: list[Element] = []
        for node in current:
            following.extend(node.children_named(name))
        current = following
    found: list[Element] = []
    for node in current:
        found.extend(node.children_named(entry))
    return found


def _drop_if_empty(root: Element, container: str | None, entry: str) -> None:
    """Remove a wrapper whose entries have all gone, and fix its count.

    ``<dataValidations count="0"/>`` with no children is not something Excel
    accepts, so the wrapper goes with its last entry.
    """
    if container is None or container == entry:
        return
    holder = root.child(container)
    if holder is None:
        return
    remaining = sum(1 for _ in holder.children_named(entry))
    if not remaining:
        root.remove(holder)
    elif holder.get("count") is not None:
        holder.set("count", str(remaining))


def _anchor_ends(root: Element) -> Iterator[Element]:
    """Every ``<from>``/``<to>`` that carries a cell index.

    The element wrapping them differs by feature and none of the names is
    worth enumerating: a drawing uses ``<xdr:twoCellAnchor>`` or
    ``<xdr:oneCellAnchor>``, a form control a bare ``<anchor>`` nested inside
    ``mc:AlternateContent``. Looking for the wrapper by name missed every
    drawing, so the ends are found directly and a ``<from>`` with no ``col``
    or ``row`` child is simply skipped.
    """
    for name in ("from", "to"):
        for end in root.descendants(name):
            if _anchor_index(end, "col") is not None or _anchor_index(end, "row") is not None:
                yield end


def _anchor_index(end: Element, axis: str) -> Element | None:
    for child in end.children:
        if isinstance(child, Element) and (
            child.name == axis or child.name.endswith(f":{axis}")
        ):
            return child
    return None


def _shift_inline_anchors(root: Element, shift: Shift) -> None:
    for end in _anchor_ends(root):
        for axis, is_row in (("col", False), ("row", True)):
            node = _anchor_index(end, axis)
            if node is None or not node.text:
                continue
            try:
                value = int(node.text)
            except ValueError:
                continue
            node.set_text(str(shift_index(value, shift, is_row=is_row)))


def _delete_inline_anchors(root: Element, deletion: Deletion) -> None:
    for end in _anchor_ends(root):
        for axis, is_row in (("col", False), ("row", True)):
            node = _anchor_index(end, axis)
            if node is None or not node.text:
                continue
            try:
                value = int(node.text)
            except ValueError:
                continue
            node.set_text(str(collapse_index(value, deletion, is_row=is_row)))


#: Relationship types of the sheet's own parts that record cell addresses.
RT_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
RT_VML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"


def related_parts(sheet: Worksheet, relationship_type: str) -> list[str]:
    """The part names a sheet relates to by type.

    Public because more than one feature needs it: the shifting reaches a
    sheet's drawing and comments this way, and so does reading its shapes.
    """
    package = sheet.workbook.package
    try:
        relationships = package.relationships(sheet.part_name)
    except PackageError:
        return []
    found: list[str] = []
    for relationship in relationships.by_type(relationship_type):
        if relationship.is_external:
            continue
        name = relationship.target_part
        if package.has_part(name):
            found.append(name)
    return found


def _move_related_parts(sheet: Worksheet, *, shift: Shift | None, deletion: Deletion | None) -> None:
    """Move the addresses in a sheet's drawing, VML and comment parts.

    Three kinds live outside the worksheet, and none of them is reachable
    from the sheet's own XML:

    - a drawing's ``<xdr:from>``/``<xdr:to>`` hold **zero-based** column and
      row indices, so a shape sits one cell off if they are treated as
      one-based, and the file stays perfectly valid
    - a VML ``<x:Anchor>`` is eight comma-separated numbers, the first,
      third, fifth and seventh being zero-based column and row; this is how
      a comment and a form control say which cell they sit on
    - a comment part addresses its cell with an ordinary ``ref``

    Comments and legacy drawings were on neither the shifted list nor the
    refused one before this, so inserting a row on a sheet with a comment
    moved the cells and left the comment behind.
    """
    for name in related_parts(sheet, RT_DRAWING):
        document = sheet.workbook.package.xml(name)
        _move_anchor_ends(document.root, shift=shift, deletion=deletion)

    for name in related_parts(sheet, RT_COMMENTS):
        document = sheet.workbook.package.xml(name)
        _move_comments(document.root, shift=shift, deletion=deletion)

    for name in related_parts(sheet, RT_THREADED_COMMENTS):
        document = sheet.workbook.package.xml(name)
        _move_threads(document.root, shift=shift, deletion=deletion)

    for name in related_parts(sheet, RT_VML):
        _move_vml(sheet, name, shift=shift, deletion=deletion)


def _move_anchor_ends(
    root: Element, *, shift: Shift | None, deletion: Deletion | None
) -> bool:
    """Move every ``<from>``/``<to>`` in a drawing part."""
    changed = False
    for end in _anchor_ends(root):
        for axis, is_row in (("col", False), ("row", True)):
            node = _anchor_index(end, axis)
            if node is None or not node.text:
                continue
            try:
                value = int(node.text)
            except ValueError:
                continue
            moved = (
                shift_index(value, shift, is_row=is_row)
                if shift is not None
                else collapse_index(value, deletion, is_row=is_row)  # type: ignore[arg-type]
            )
            if moved != value:
                node.set_text(str(moved))
                changed = True
    return changed


def _move_comments(root: Element, *, shift: Shift | None, deletion: Deletion | None) -> bool:
    """Move each comment onto its cell's new address.

    A comment whose cell is deleted goes with it, which is what Excel does.
    """
    container = root.child("commentList")
    if container is None:
        return False
    changed = False
    for comment in list(container.children_named("comment")):
        raw = comment.get("ref")
        if raw is None:
            continue
        moved = (
            shift_cell(raw, shift) if shift is not None
            else delete_cell(raw, deletion)  # type: ignore[arg-type]
        )
        if moved is None:
            container.remove(comment)
            changed = True
            continue
        if moved != raw:
            comment.set("ref", moved)
            changed = True
    return changed


def _move_threads(root: Element, *, shift: Shift | None, deletion: Deletion | None) -> None:
    """Move each threaded comment, and each reply, onto its cell's new
    address. A thread whose cell is deleted goes with it, replies and all,
    as its placeholder note does; left behind, it would sit on a cell
    whose note had moved."""
    for element in list(root.children_named("threadedComment")):
        raw = element.get("ref")
        if raw is None:
            continue
        moved = (
            shift_cell(raw, shift) if shift is not None
            else delete_cell(raw, deletion)  # type: ignore[arg-type]
        )
        if moved is None:
            root.remove(element)
        elif moved != raw:
            element.set("ref", moved)


def _move_vml(
    sheet: Worksheet, name: str, *, shift: Shift | None, deletion: Deletion | None
) -> None:
    """Move every ``<x:Anchor>`` in a VML part.

    VML is not XML this library parses: it is an HTML-ish dialect with
    unquoted attributes and unclosed tags, and the ``<xml>`` root would be
    refused. The anchors are rewritten in the raw text instead, which is
    safe because an anchor is a self-contained run of eight numbers.
    """
    package = sheet.workbook.package
    raw = package.read(name)
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError:
        return
    original = text

    if deletion is not None:
        # A note whose cell is deleted goes, as its text in the comments
        # part does; collapsing its cell onto a neighbour would leave a box
        # with no text behind it.
        for cell, box in note_shapes(text).items():
            if deletion.covers_row(cell.row) or deletion.covers_column(cell.column):
                text = text.replace(box, "", 1)

    def move(match: re.Match[str]) -> str:
        body = match.group(1)
        moved = (
            shift_vml_anchor(body, shift)
            if shift is not None
            else delete_vml_anchor(body, deletion)  # type: ignore[arg-type]
        )
        return match.group(0).replace(body, moved, 1)

    rewritten = re.sub(r"<x:Anchor>(.*?)</x:Anchor>", move, text, flags=re.S)
    rewritten = _move_vml_owner(rewritten, shift=shift, deletion=deletion)
    if rewritten != original:
        package.write(name, rewritten.encode("utf-8"))


def _move_vml_owner(
    text: str, *, shift: Shift | None, deletion: Deletion | None
) -> str:
    """Move the ``<x:Row>``/``<x:Column>`` that names a note's own cell.

    This is not the ``<x:Anchor>``: the anchor is where the note's box is
    drawn, and this is the cell it belongs to, zero-based. Excel refuses to
    open a workbook whose VML and comments part disagree about that cell, so
    the two move together or neither does.
    """

    def rewrite(tag: str, is_row: bool) -> str:
        def move(match: re.Match[str]) -> str:
            try:
                value = int(match.group(1))
            except ValueError:
                return match.group(0)
            moved = (
                shift_index(value, shift, is_row=is_row)
                if shift is not None
                else collapse_index(value, deletion, is_row=is_row)  # type: ignore[arg-type]
            )
            return f"<x:{tag}>{moved}</x:{tag}>"

        return move  # type: ignore[return-value]

    text = re.sub(r"<x:Row>(\d+)</x:Row>", rewrite("Row", True), text)
    return re.sub(r"<x:Column>(\d+)</x:Column>", rewrite("Column", False), text)


def _move_validation_formulas(
    sheet: Worksheet, *, shift: Shift | None, deletion: Deletion | None
) -> None:
    """Move the references inside ``<formula1>`` and ``<formula2>``.

    A validation's bounds are formulas: ``$B$5`` for a whole-number limit,
    ``=$E$1:$E$2`` for a list's source, ``ISNUMBER($A2)`` for a custom rule.
    Moving the sqref and leaving these behind gives a range that validates
    against the wrong cells, which Excel reports no error for.
    """
    container = sheet.document.root.child("dataValidations")
    if container is None:
        return
    for entry in container.children_named("dataValidation"):
        for name in ("formula1", "formula2"):
            node = entry.child(name)
            if node is None or not node.text:
                continue
            moved = (
                shift_formula(
                    node.text, shift, formula_sheet=sheet.name, target_sheet=sheet.name
                )
                if shift is not None
                else delete_in_formula(
                    node.text,
                    deletion,  # type: ignore[arg-type]
                    formula_sheet=sheet.name,
                    target_sheet=sheet.name,
                )
            )
            if moved != node.text:
                node.set_text(moved)

#: Inside a ``<customSheetView>``, which repeats a slice of the sheet's own
#: structure per saved view: each attribute with the notation it uses.
CUSTOM_VIEW_ADDRESSES: tuple[tuple[str, str, str], ...] = (
    ("selection", "sqref", "sqref"),
    ("selection", "activeCell", "cell"),
    ("pane", "topLeftCell", "cell"),
)


def _move_custom_views(
    sheet: Worksheet, *, shift: Shift | None, deletion: Deletion | None
) -> None:
    """Move the addresses each saved view keeps.

    A custom sheet view stores its own selection, frozen pane, autofilter
    and page breaks, so every one of them is a second copy of an address the
    sheet also holds. Leaving them puts a user who switches to that view
    back on the cells that used to be there.
    """
    container = sheet.document.root.child("customSheetViews")
    if container is None:
        return
    for view in container.children_named("customSheetView"):
        for entry, attribute, notation in CUSTOM_VIEW_ADDRESSES:
            for node in view.children_named(entry):
                raw = node.get(attribute)
                if raw is None:
                    continue
                moved = _move_address(raw, notation, shift=shift, deletion=deletion)
                if moved is None:
                    node.unset(attribute)
                elif moved != raw:
                    node.set(attribute, moved)
        # A view's filter renumbers its columns like the sheet's own.
        for node in list(view.children_named("autoFilter")):
            if _move_filter(node, shift=shift, deletion=deletion):
                view.remove(node)
        for axis, is_row in (("rowBreaks", True), ("colBreaks", False)):
            breaks = view.child(axis)
            if breaks is not None:
                _move_break_entries(breaks, is_row=is_row, shift=shift, deletion=deletion)


def _move_break_entries(
    breaks: Element, *, is_row: bool, shift: Shift | None, deletion: Deletion | None
) -> None:
    """Move a ``<brk id=...>``, whose id is a one-based row or column."""
    for entry in list(breaks.children_named("brk")):
        raw = entry.get("id")
        if raw is None:
            continue
        try:
            value = int(raw)
        except ValueError:
            continue
        if shift is not None:
            moved = shift_index(value - 1, shift, is_row=is_row) + 1
        else:
            survivor = delete_index(value - 1, deletion, is_row=is_row)  # type: ignore[arg-type]
            if survivor is None:
                breaks.remove(entry)
                continue
            moved = survivor + 1
        entry.set("id", str(moved))
    remaining = sum(1 for _ in breaks.children_named("brk"))
    if breaks.get("count") is not None:
        breaks.set("count", str(remaining))


def _move_address(
    raw: str, notation: str, *, shift: Shift | None, deletion: Deletion | None
) -> str | None:
    """One address in whichever notation it is written in."""
    if shift is not None:
        if notation == "sqref":
            return shift_sqref(raw, shift)
        if notation == "cell":
            return shift_cell(raw, shift)
        return shift_ref(raw, shift)
    if notation == "sqref":
        return delete_sqref(raw, deletion)  # type: ignore[arg-type]
    if notation == "cell":
        return delete_cell(raw, deletion)  # type: ignore[arg-type]
    return delete_ref(raw, deletion)  # type: ignore[arg-type]


def _move_extensions(
    sheet: Worksheet, *, shift: Shift | None, deletion: Deletion | None
) -> None:
    """Move the addresses inside the sheet's ``<extLst>``.

    Everything newer than the 2006 schema lives here, and it spells its
    addresses in one consistent way whatever the feature: an ``<xm:sqref>``
    holding ranges and an ``<xm:f>`` holding a formula. A modern data bar,
    an x14 data validation and a sparkline group all use that pair, so one
    rule covers them rather than one per feature.

    An element whose ranges are all deleted is removed together with the
    ``x14`` rule that owns it, because a rule with no ``xm:sqref`` is not
    something Excel accepts.
    """
    extensions = sheet.document.root.child("extLst")
    if extensions is None:
        return
    for node in list(extensions.descendants("sqref")):
        if not node.text:
            continue
        moved = (
            shift_sqref(node.text, shift)
            if shift is not None
            else delete_sqref(node.text, deletion)  # type: ignore[arg-type]
        )
        if moved is None:
            _drop_extension_owner(node)
            continue
        if moved != node.text:
            node.set_text(moved)

    for node in extensions.descendants("f"):
        if not node.text:
            continue
        moved = (
            shift_formula(node.text, shift, formula_sheet=sheet.name, target_sheet=sheet.name)
            if shift is not None
            else delete_in_formula(
                node.text,
                deletion,  # type: ignore[arg-type]
                formula_sheet=sheet.name,
                target_sheet=sheet.name,
            )
        )
        if moved != node.text:
            node.set_text(moved)


def _drop_extension_owner(node: Element) -> None:
    """Remove the element an emptied ``<xm:sqref>`` belongs to.

    The sqref sits beside the rule it applies to rather than inside it, so
    the parent goes: an ``x14:conditionalFormatting`` whose range is gone
    takes its ``x14:cfRule`` with it.
    """
    owner = node.parent
    if owner is None:
        return
    grandparent = owner.parent
    if grandparent is None:
        owner.remove(node)
        return
    grandparent.remove(owner)
