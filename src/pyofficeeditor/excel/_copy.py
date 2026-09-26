"""Copying cells and pasting them, as Excel's Copy and Paste does.

Measured in Excel through ``Range.Copy Destination:=``, the paste all of
Ctrl+V:

- Each cell comes as it was, its value, formula and format, and a blank
  one clears the cell it lands on, format and all. A formula moves as a
  copy moves it: every relative reference, whichever sheet it names, and
  one the copy pushes off the sheet is ``#REF!``; see
  :func:`~pyofficeeditor.excel._formulas.copied_formula`.
- A destination a whole number of copies of the source both ways takes
  that many, and any other one copy, from its top left.
- On a sheet whose filter hides rows by a criterion, a copy takes only the
  rows and columns showing, rows hidden by hand among those it leaves out,
  unless none of the copy shows. A table's filter makes no difference.
- Merged cells wholly inside the source are merged where they land, and
  merged cells the paste covers go; one it covers in part is refused. An
  array formula comes whole, or as an array over the part copied, moved as
  its own first cell would move. One the paste covers in part is refused.
  A spill's formula comes and spills again, and the cells it spilled into
  come empty. A paste over a spill's formula takes what it spilled with it.
- Notes, threads and links come with their cells, in place of the
  destination's. A note's box keeps its size, and goes where a new note's
  box goes.
- Validation and conditional formats at the destination are cut back from
  it. The source's validation takes the pasted cells in: the same
  validation on the same sheet, and on another an identical one or a new
  one. The source's conditional formats come as new blocks, their formulas
  moved as the cells are.
- Whole rows bring their heights, formats and whether they are hidden, and
  whole columns their widths and whether they are hidden.

Not yet: a copy of a whole table, which Excel makes a new table of, and
cells taken out of a table, which Excel gives the table style's look as
formats of their own; a paste over a table's header or totals row or
across its edge, and one just below or beside a table, which Excel makes
it take in; a paste over cells a formula spilled into, which Excel then
shows ``#SPILL!``; shapes inside the source, which Excel copies with the
cells; pivot tables; and what-if data tables. Each is refused before
anything changes.
"""

from __future__ import annotations

import bisect
from collections import defaultdict
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING

from pyofficeeditor._xml import Element, XmlDocument
from pyofficeeditor.excel._addresses import cut_conditional_formats, cut_validations
from pyofficeeditor.excel._comments import CopiedNote, ThreadedComment
from pyofficeeditor.excel._conditional import CONDITION_FORMULA_TYPES, ConditionalFormatting, ConditionalRule
from pyofficeeditor.excel._formulas import copied_formula, join_areas
from pyofficeeditor.excel._hyperlinks import Hyperlink
from pyofficeeditor.excel._reference import MAX_COLUMN, MAX_ROW, AxisRef, CellRef, RangeRef
from pyofficeeditor.excel._schema import WORKSHEET_CHILD_ORDER, insert_in_schema_order
from pyofficeeditor.excel._xstring import decode, escape

if TYPE_CHECKING:
    from pyofficeeditor.excel.worksheet import Worksheet

#: A block, as ``(top, bottom, left, right)``.
Block = tuple[int, int, int, int]


@dataclass(frozen=True)
class _Axis:
    """The rows or the columns a copy takes: ``low`` to ``high`` less those
    ``skipped``, which a copy on a filtering sheet leaves out."""

    low: int
    high: int
    skipped: tuple[int, ...] = ()

    @property
    def size(self) -> int:
        return self.high - self.low + 1 - len(self.skipped)

    def place(self, index: int) -> int | None:
        """How far into the pasted block the source's ``index`` lands, from
        0, or ``None`` when the copy leaves it out."""
        if not self.low <= index <= self.high:
            return None
        at = bisect.bisect_left(self.skipped, index)
        if at < len(self.skipped) and self.skipped[at] == index:
            return None
        return index - self.low - at

    def runs(self, low: int, high: int) -> list[tuple[int, int]]:
        """Where ``low`` to ``high`` of the source land, as runs of places,
        one for each stretch the copy takes without a gap."""
        low, high = max(low, self.low), min(high, self.high)
        found: list[tuple[int, int]] = []
        start = low
        for index in [*(skipped for skipped in self.skipped if low <= skipped <= high), high + 1]:
            first, last = self.place(start), self.place(index - 1)
            if start <= index - 1 and first is not None and last is not None:
                found.append((first, last))
            start = index + 1
        return found


@dataclass(frozen=True)
class _Cell:
    """A cell as the copy took it."""

    row: int
    column: int
    element: Element
    #: The cell its formula is read from: itself, or the first cell of the
    #: array it starts a copy of.
    origin: CellRef
    #: How many rows and columns the array this cell starts covers.
    array: tuple[int, int] | None = None
    #: A spill's block, which moves with its formula.
    spill: RangeRef | None = None


@dataclass
class _Copy:
    """Everything a copy takes from its source, read before anything is
    pasted, so a paste over its own source reads what was there."""

    sheet: Worksheet
    rows: _Axis
    columns: _Axis
    whole: str | None
    cells: list[_Cell]
    merges: list[Block]
    notes: list[tuple[int, int, CopiedNote]]
    threads: list[tuple[int, int, ThreadedComment]]
    links: list[tuple[Block, Hyperlink]]
    validations: list[tuple[Element, list[RangeRef]]]
    formats: list[tuple[ConditionalFormatting, list[RangeRef]]]
    row_settings: list[tuple[int, dict[str, str]]]
    column_settings: list[tuple[int, float | None, bool]]


def copy_cells(sheet: Worksheet, cells: str | RangeRef, target: Worksheet, destination: str | RangeRef) -> RangeRef:
    """Copy ``cells`` of ``sheet`` and paste them at ``destination`` on
    ``target``, as Excel's Copy and Paste does; the block pasted."""
    if target.workbook is not sheet.workbook:
        raise ValueError(
            f"{target.name!r} is in another workbook; a copy pastes within its own, where a cell's "
            "format and text mean the same."
        )
    block, whole = _parsed(cells)
    landing, landed_whole = _parsed(destination)
    rows, columns = _axes(sheet, block)
    if whole == "rows" and landing.left != 1:
        raise ValueError(f"whole rows paste at column A, not at {landing.a1}.")
    if whole == "columns" and landing.top != 1:
        raise ValueError(f"whole columns paste at row 1, not at {landing.a1}.")
    tiles = _tiles(landing, rows.size, columns.size, single=landed_whole is None and landing.is_single_cell)
    _refuse(sheet, block, target, tiles)
    taken = _take(sheet, block, rows, columns, whole)
    for tile in tiles:
        _paste(taken, target, tile)
    target.workbook.mark_values_changed()
    target.invalidate()
    return RangeRef(tiles[0].start, tiles[-1].end)


def _parsed(reference: str | RangeRef) -> tuple[RangeRef, str | None]:
    """A block, and whether it is whole rows or whole columns."""
    if isinstance(reference, RangeRef):
        return reference.normalized, None
    try:
        return RangeRef.parse(reference).normalized, None
    except ValueError:
        span = AxisRef.parse(reference)
    if span.is_row:
        return RangeRef(CellRef(span.low, 1), CellRef(span.high, MAX_COLUMN)), "rows"
    return RangeRef(CellRef(1, span.low), CellRef(MAX_ROW, span.high)), "columns"


def _axes(sheet: Worksheet, block: RangeRef) -> tuple[_Axis, _Axis]:
    """The rows and columns a copy takes. Measured: all of them, unless the
    sheet's filter hides rows by a criterion, when the hidden rows and
    columns are left out, rows hidden by hand among them; and all of them
    again when none of the copy shows."""
    every = _Axis(block.top, block.bottom), _Axis(block.left, block.right)
    current = sheet.auto_filter
    if current is None or not current.filtering:
        return every
    hidden_rows = tuple(
        number for number in sorted(sheet.rows_by_number()) if block.top <= number <= block.bottom and sheet.row_hidden(number)
    )
    hidden_columns = tuple(column for column in _hidden_columns(sheet) if block.left <= column <= block.right)
    rows, columns = _Axis(block.top, block.bottom, hidden_rows), _Axis(block.left, block.right, hidden_columns)
    return every if rows.size == 0 or columns.size == 0 else (rows, columns)


def _hidden_columns(sheet: Worksheet) -> list[int]:
    container = sheet.document.root.child("cols")
    found: set[int] = set()
    for entry in () if container is None else container.children_named("col"):
        if entry.get("hidden") in ("1", "true"):
            found.update(range(int(entry.get("min") or 1), int(entry.get("max") or 1) + 1))
    return sorted(found)


def _tiles(landing: RangeRef, height: int, width: int, *, single: bool) -> list[RangeRef]:
    """Where the copies go. Measured: a destination a whole number of
    copies both ways takes that many, and any other one, at its top left."""
    if single or landing.height % height or landing.width % width:
        corners = [(landing.top, landing.left)]
    else:
        corners = [
            (top, left)
            for top in range(landing.top, landing.bottom + 1, height)
            for left in range(landing.left, landing.right + 1, width)
        ]
    tiles: list[RangeRef] = []
    for top, left in corners:
        if top + height - 1 > MAX_ROW or left + width - 1 > MAX_COLUMN:
            raise ValueError(f"a copy {height} rows by {width} columns does not fit on the sheet at {CellRef(top, left).a1}.")
        tiles.append(RangeRef(CellRef(top, left), CellRef(top + height - 1, left + width - 1)))
    return tiles


# ----------------------------------------------------------------------
# What stops a copy
# ----------------------------------------------------------------------


def _refuse(sheet: Worksheet, block: RangeRef, target: Worksheet, tiles: list[RangeRef]) -> None:
    """Refuse, before anything changes, what Excel refuses and what this
    does not do yet."""
    for tile in tiles:
        for merged in target.merged_ranges:
            if merged.intersects(tile) and not tile.contains(merged):
                raise ValueError(
                    f"the paste over {tile.a1} would take part of the merged cells {merged.a1}, "
                    "which Excel refuses: \"We can't do that to a merged cell.\""
                )
        for array, dynamic in _arrays(target):
            if not dynamic and array.intersects(tile) and not tile.contains(array):
                raise ValueError(
                    f"the paste over {tile.a1} would take part of the array formula over {array.a1}, "
                    "which Excel refuses: \"You can't change part of an array.\""
                )
            if dynamic and array.intersects(tile) and array.start not in tile:
                raise ValueError(
                    f"the paste over {tile.a1} lands in cells the formula at {array.start.a1} spilled into, "
                    "which Excel then shows #SPILL!; that is not supported yet."
                )
        for table in target.tables:
            data = table.data_range
            if table.ref.intersects(tile) and (data is None or not data.contains(tile)):
                raise ValueError(
                    f"the paste over {tile.a1} reaches the header, the totals row or past the edge of the "
                    f"table {table.name!r}; pasting there is not supported yet."
                )
            if _grows(table.ref, table.has_totals_row, tile):
                raise ValueError(
                    f"the paste over {tile.a1} would make the table {table.name!r} take it in, as Excel's does; "
                    "that is not supported yet."
                )
        _refuse_pivots_and_data_tables(target, tile)
    for table in sheet.tables:
        if block.contains(table.ref):
            raise ValueError(
                f"{block.a1} holds all of the table {table.name!r}, which Excel copies as a new table; "
                "that is not supported yet."
            )
        data = table.data_range
        inside = (
            data is not None and target is sheet and data.contains(block) and all(data.contains(tile) for tile in tiles)
        )
        if table.ref.intersects(block) and not inside:
            raise ValueError(
                f"{block.a1} takes cells of the table {table.name!r} out of it, which Excel gives the table "
                "style's look as formats of their own; that is not supported yet."
            )
    for shape in sheet.shapes:
        if shape.cells and block.contains(RangeRef.parse(shape.cells)):
            raise ValueError(
                f"{block.a1} holds the shape {shape.name!r}, which Excel copies with the cells; "
                "that is not supported yet."
            )
    _refuse_pivots_and_data_tables(sheet, block)


def _grows(table: RangeRef, totals: bool, tile: RangeRef) -> bool:
    """Measured: a paste on the row just below a table, within its columns,
    makes a table with no totals row take it in, and so does one on the
    column just beside it, within its rows."""
    below = not totals and tile.top == table.bottom + 1 and table.left <= tile.left and tile.right <= table.right
    beside = tile.left == table.right + 1 and table.top <= tile.top and tile.bottom <= table.bottom
    return below or beside


def _refuse_pivots_and_data_tables(sheet: Worksheet, block: RangeRef) -> None:
    for pivot in sheet.pivot_tables:
        if pivot.location.intersects(block):
            raise ValueError(f"{block.a1} meets the pivot table {pivot.name!r}; copying it is not supported yet.")
    for number, row in sheet.rows_by_number().items():
        if not block.top <= number <= block.bottom:
            continue
        for cell in row.children_named("c"):
            formula = cell.child("f")
            if formula is not None and formula.get("t") == "dataTable":
                raise ValueError(f"{block.a1} meets a what-if data table; copying it is not supported yet.")


def _arrays(sheet: Worksheet) -> list[tuple[RangeRef, bool]]:
    """Every array formula's block, and whether it is a spill: an array
    formula whose cell carries cell metadata, as Excel marks one."""
    found: list[tuple[RangeRef, bool]] = []
    for _, cell in sheet.cell_elements():
        formula = cell.child("f")
        ref = None if formula is None else formula.get("ref")
        if formula is None or ref is None or formula.get("t") != "array":
            continue
        try:
            found.append((RangeRef.parse(ref).normalized, cell.get("cm") is not None))
        except ValueError:
            continue
    return found


# ----------------------------------------------------------------------
# Taking the source
# ----------------------------------------------------------------------


def _take(sheet: Worksheet, block: RangeRef, rows: _Axis, columns: _Axis, whole: str | None) -> _Copy:
    arrays = {
        span: (reference, cell, cell.get("cm") is not None)
        for reference, cell in sheet.cell_elements()
        if (formula := cell.child("f")) is not None
        and formula.get("t") == "array"
        and (span := _block_of(formula.get("ref"))) is not None
        and span.intersects(block)
    }
    cells: list[_Cell] = []
    for number, row in sorted(sheet.rows_by_number().items()):
        place_row = rows.place(number)
        if place_row is None:
            continue
        for cell in row.children_named("c"):
            reference = _address(cell)
            place_column = None if reference is None else columns.place(reference.column)
            if reference is None or place_column is None:
                continue
            taken = _taken_cell(sheet, block, arrays, reference, cell)
            if taken is not None:
                element, origin, array, spill = taken
                cells.append(_Cell(place_row, place_column, element, origin, array, spill))

    merges = [
        (top, bottom, left, right)
        for merged in sheet.merged_ranges
        if block.contains(merged)
        and (top := rows.place(merged.top)) is not None
        and (bottom := rows.place(merged.bottom)) is not None
        and (left := columns.place(merged.left)) is not None
        and (right := columns.place(merged.right)) is not None
        and bottom - top == merged.bottom - merged.top
        and right - left == merged.right - merged.left
    ]
    notes = [
        (place_row, place_column, copied)
        for note in sheet.comments
        if (cell := _parse_cell(note.ref)) is not None
        and (place_row := rows.place(cell.row)) is not None
        and (place_column := columns.place(cell.column)) is not None
        and (copied := sheet.copied_note(cell)) is not None
    ]
    threads = [
        (place_row, place_column, thread)
        for thread in sheet.threaded_comments
        if (cell := _parse_cell(thread.ref)) is not None
        and (place_row := rows.place(cell.row)) is not None
        and (place_column := columns.place(cell.column)) is not None
    ]
    links = [
        ((top, bottom, left, right), link)
        for link in sheet.hyperlinks
        if block.contains(link.ref)
        and (top := rows.place(link.ref.top)) is not None
        and (bottom := rows.place(link.ref.bottom)) is not None
        and (left := columns.place(link.ref.left)) is not None
        and (right := columns.place(link.ref.right)) is not None
    ]
    validations = [
        (element, touched)
        for element in _validation_elements(sheet)
        if (touched := [_overlap(area, block) for area in _areas(element.get("sqref")) if area.intersects(block)])
    ]
    formats = [
        (formatting, touched)
        for formatting in sheet.conditional_formats
        if (touched := [_overlap(area, block) for area in formatting.ranges if area.intersects(block)])
    ]
    row_settings: list[tuple[int, dict[str, str]]] = []
    column_settings: list[tuple[int, float | None, bool]] = []
    if whole == "rows":
        present = sheet.rows_by_number()
        for number in range(block.top, block.bottom + 1):
            place = rows.place(number)
            if place is not None:
                row = present.get(number)
                settings = {} if row is None else {k: v for k, v in row.attributes.items() if k not in ("r", "spans")}
                row_settings.append((place, settings))
    if whole == "columns":
        for column in range(block.left, block.right + 1):
            place = columns.place(column)
            if place is not None:
                column_settings.append((place, sheet.column_width(column), sheet.column_hidden(column)))
    return _Copy(
        sheet, rows, columns, whole, cells, merges, notes, threads, links, validations, formats,
        row_settings, column_settings,
    )  # fmt: skip


def _taken_cell(
    sheet: Worksheet,
    block: RangeRef,
    arrays: dict[RangeRef, tuple[CellRef, Element, bool]],
    reference: CellRef,
    cell: Element,
) -> tuple[Element, CellRef, tuple[int, int] | None, RangeRef | None] | None:
    """A cell as the copy takes it, with the cell its formula is read from,
    the size of the array it starts and the spill it anchors; ``None`` for a
    cell a spill filled, which comes empty."""
    for span, (master, master_cell, dynamic) in arrays.items():
        if reference not in span:
            continue
        if dynamic:
            if reference != master:
                return None
            return _detached(cell), reference, None, span
        part = _overlap(span, block)
        element = _detached(cell)
        for formula in list(element.children_named("f")):
            element.remove(formula)
        if reference != part.start:
            return element, reference, None, None
        text = sheet.formula_of(master_cell, master) or ""
        formula = Element.create("f", {"t": "array"})
        formula.set_text(escape(text))
        element.insert(0, formula)
        return element, master, (part.height, part.width), None
    element = _detached(cell)
    formula = element.child("f")
    if formula is not None:
        text = sheet.formula_of(cell, reference)
        for marker in ("t", "si", "ref"):
            formula.unset(marker)
        if text:
            formula.set_text(escape(text))
    return element, reference, None, None


def _validation_elements(sheet: Worksheet) -> list[Element]:
    container = sheet.document.root.child("dataValidations")
    return [] if container is None else list(container.children_named("dataValidation"))


# ----------------------------------------------------------------------
# Pasting
# ----------------------------------------------------------------------


def _paste(taken: _Copy, target: Worksheet, tile: RangeRef) -> None:
    """Clear ``tile`` of all a paste replaces, then put the copy in it."""
    _clear(target, tile)
    by_row: dict[int, list[Element]] = defaultdict(list)
    for cell in taken.cells:
        destination = CellRef(tile.top + cell.row, tile.left + cell.column)
        element = _detached(cell.element)
        element.set("r", destination.a1)
        formula = element.child("f")
        if formula is not None and formula.text:
            formula.set_text(
                copied_formula(formula.text, destination.row - cell.origin.row, destination.column - cell.origin.column)
            )
        if formula is not None and cell.array is not None:
            height, width = cell.array
            formula.set("ref", RangeRef(destination, CellRef(destination.row + height - 1, destination.column + width - 1)).a1)
        if formula is not None and cell.spill is not None:
            rows, columns = destination.row - cell.origin.row, destination.column - cell.origin.column
            spill = RangeRef(cell.spill.start.translated(rows, columns), cell.spill.end.translated(rows, columns))
            formula.set("ref", spill.a1)
        by_row[destination.row].append(element)
    for number in sorted(by_row):
        target.place_cells(number, by_row[number])
    for top, bottom, left, right in taken.merges:
        _add_merge(target, _moved((top, bottom, left, right), tile))
    for row, column, note in taken.notes:
        target.paste_note(CellRef(tile.top + row, tile.left + column), note)
    for row, column, thread in taken.threads:
        _paste_thread(target, CellRef(tile.top + row, tile.left + column), thread)
    for places, link in taken.links:
        target.add_hyperlink(
            _moved(places, tile), link.target, location=link.location, display=link.display, tooltip=link.tooltip
        )
    for element, touched in taken.validations:
        _paste_validation(taken, target, tile, element, touched)
    for formatting, touched in taken.formats:
        _paste_format(taken, target, tile, formatting, touched)
    present = target.rows_by_number()
    for place, settings in taken.row_settings:
        number = tile.top + place
        if not settings and number not in present:
            continue
        row = target.row_element(number)
        for name in [name for name in row.attributes if name not in ("r", "spans")]:
            row.unset(name)
        for name, value in settings.items():
            row.set(name, value)
    for place, width, hidden in taken.column_settings:
        target.set_column_width(tile.left + place, width)
        target.set_column_hidden(tile.left + place, hidden)


def _clear(target: Worksheet, tile: RangeRef) -> None:
    """Take out of ``tile`` what a paste replaces: cells, merges, notes,
    threads and links, and cut validation and conditional formats back.
    Measured, validation loses the tile columns first and conditional
    formats rows first, as they lose cells a removal clears."""
    _drop_spills(target, tile)
    target.remove_cells(tile)
    container = target.document.root.child("mergeCells")
    for entry in [] if container is None else list(container.children_named("mergeCell")):
        merged = _block_of(entry.get("ref"))
        if merged is not None and tile.contains(merged) and container is not None:
            container.remove(entry)
    if container is not None:
        _recount(target.document.root, container, "mergeCell")
    for cell in {note.ref for note in target.comments} | {thread.ref for thread in target.threaded_comments}:
        reference = _parse_cell(cell)
        if reference is not None and reference in tile:
            target.remove_comment(reference)
    target.remove_hyperlink(tile)
    cut_validations(target.document.root, tile)
    cut_conditional_formats(target.document.root, tile)


def _drop_spills(target: Worksheet, tile: RangeRef) -> None:
    """A spill whose formula the paste replaces goes, and, measured, the
    cells it spilled into with it."""
    for span, dynamic in _arrays(target):
        if dynamic and span.start in tile:
            for spilled in span.cells():
                if spilled not in tile:
                    target.remove_cells(RangeRef(spilled, spilled))


def _paste_thread(target: Worksheet, cell: CellRef, thread: ThreadedComment) -> None:
    """A thread copied, replies and all, as Excel's Paste copies one."""
    target.add_threaded_comment(cell, thread.text, author=thread.author, when=thread.when)
    for reply in thread.replies:
        target.add_threaded_reply(cell, reply.text, author=reply.author, when=reply.when)
    if thread.resolved:
        target.resolve_threaded_comment(cell)


def _paste_validation(taken: _Copy, target: Worksheet, tile: RangeRef, element: Element, touched: list[RangeRef]) -> None:
    """The source's validation takes the pasted cells in. Measured: on the
    same sheet the validation itself does; on another sheet one identical
    to it there does, or else a new one, its formulas moved as the cells
    are; either way its areas join where they meet."""
    landed = [area for area in touched for area in _landing(taken, tile, area)]
    if not landed:
        return
    root = target.document.root
    if target is taken.sheet and element.parent is not None:
        element.set("sqref", " ".join(area.a1 for area in join_areas([*_areas(element.get("sqref")), *landed])))
        return
    anchor = _areas(element.get("sqref"))[0].start
    rows, columns = landed[0].top - anchor.row, landed[0].left - anchor.column
    wanted = _detached(element)
    for name in ("formula1", "formula2"):
        bound = wanted.child(name)
        if bound is not None and bound.text:
            bound.set_text(copied_formula(bound.text, rows, columns))
    for existing in _validation_elements(target):
        if _same_rule(existing, wanted):
            existing.set("sqref", " ".join(area.a1 for area in join_areas([*_areas(existing.get("sqref")), *landed])))
            return
    wanted.set("sqref", " ".join(area.a1 for area in join_areas(landed)))
    wanted.unset("xr:uid")
    container = root.child("dataValidations")
    if container is None:
        container = Element.create("dataValidations")
        insert_in_schema_order(root, container, WORKSHEET_CHILD_ORDER)
    container.append(wanted)
    _recount(root, container, "dataValidation")


def _same_rule(one: Element, other: Element) -> bool:
    """Whether two validations say the same, wherever each applies."""

    def rule(element: Element) -> tuple[object, ...]:
        settings = {name: value for name, value in element.attributes.items() if name not in ("sqref", "xr:uid")}
        bounds = tuple(bound.text for name in ("formula1", "formula2") if (bound := element.child(name)) is not None)
        return tuple(sorted(settings.items())), bounds

    return rule(one) == rule(other)


def _paste_format(
    taken: _Copy, target: Worksheet, tile: RangeRef, formatting: ConditionalFormatting, touched: list[RangeRef]
) -> None:
    """The source's conditional format comes as a new block over the
    pasted cells, after the sheet's others. Measured, its formulas move as
    the cells do, from the block's own first cell to the new block's."""
    landed = [area for area in touched for area in _landing(taken, tile, area)]
    if not landed:
        return
    anchor = formatting.ranges[0].start
    rows, columns = landed[0].top - anchor.row, landed[0].left - anchor.column
    priority = max((rule.priority for block in target.conditional_formats for rule in block.rules), default=0)
    rules: list[ConditionalRule] = []
    for offset, rule in enumerate(formatting.rules, start=1):
        moved = rule
        if rule.kind in CONDITION_FORMULA_TYPES and rule.formulas:
            moved = replace(
                rule, formulas=tuple(decode(copied_formula(escape(text), rows, columns)) for text in rule.formulas)
            )
        rules.append(replace(moved, priority=priority + offset).anchored_at(landed[0].start))
    written = ConditionalFormatting(ranges=tuple(landed), rules=tuple(rules)).write()
    insert_in_schema_order(target.document.root, written, WORKSHEET_CHILD_ORDER)


# ----------------------------------------------------------------------
# Places and blocks
# ----------------------------------------------------------------------


def _landing(taken: _Copy, tile: RangeRef, area: RangeRef) -> list[RangeRef]:
    """Where a source area lands in ``tile``: a block for each stretch of
    its rows and columns the copy takes without a gap."""
    return [
        RangeRef(CellRef(tile.top + top, tile.left + left), CellRef(tile.top + bottom, tile.left + right))
        for top, bottom in taken.rows.runs(area.top, area.bottom)
        for left, right in taken.columns.runs(area.left, area.right)
    ]


def _moved(places: Block, tile: RangeRef) -> RangeRef:
    top, bottom, left, right = places
    return RangeRef(CellRef(tile.top + top, tile.left + left), CellRef(tile.top + bottom, tile.left + right))


def _add_merge(target: Worksheet, block: RangeRef) -> None:
    """Record a merge as the pasted cells already stand, rather than
    through :meth:`Worksheet.merge`, which clears all but the first."""
    root = target.document.root
    container = root.child("mergeCells")
    if container is None:
        container = Element.create("mergeCells")
        insert_in_schema_order(root, container, WORKSHEET_CHILD_ORDER)
    container.append(Element.create("mergeCell", {"ref": block.a1}))
    _recount(root, container, "mergeCell")


def _recount(root: Element, container: Element, entry: str) -> None:
    """Set a container's count, or drop it when it holds nothing."""
    count = sum(1 for _ in container.children_named(entry))
    if count:
        container.set("count", str(count))
    else:
        root.remove(container)


def _overlap(area: RangeRef, block: RangeRef) -> RangeRef:
    return RangeRef(
        CellRef(max(area.top, block.top), max(area.left, block.left)),
        CellRef(min(area.bottom, block.bottom), min(area.right, block.right)),
    )


def _areas(sqref: str | None) -> list[RangeRef]:
    found: list[RangeRef] = []
    for piece in (sqref or "").split():
        block = _block_of(piece)
        if block is not None:
            found.append(block)
    return found


def _block_of(raw: str | None) -> RangeRef | None:
    try:
        return RangeRef.parse(raw or "").normalized
    except ValueError:
        return None


def _parse_cell(raw: str) -> CellRef | None:
    try:
        return CellRef.parse(raw)
    except ValueError:
        return None


def _address(cell: Element) -> CellRef | None:
    return _parse_cell(cell.get("r") or "")


def _detached(element: Element) -> Element:
    """A copy of an element, belonging to nothing."""
    return XmlDocument.parse(element.to_xml().encode("utf-8")).root


__all__ = ["copy_cells"]
