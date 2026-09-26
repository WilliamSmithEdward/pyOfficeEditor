"""Copying cells and pasting them, as Excel's Copy and Paste and its Paste
Special do.

Measured in Excel through ``Range.Copy Destination:=``, the paste all of
Ctrl+V:

- Each cell comes as it was, its value, formula and format, and a blank
  one clears the cell it lands on, format and all. A blank cell in a row or
  column with a format of its own comes with that format. A formula moves
  as a copy moves it: every relative reference, whichever sheet it names,
  and one the copy pushes off the sheet is ``#REF!``; see
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
  whole columns their widths, formats and whether they are hidden.

Measured through ``Range.PasteSpecial``, Paste Special brings part of that:

- ``formulas`` brings the cells' formulas and values, and ``values`` their
  values alone, a formula's as last calculated and text without its runs;
  either leaves the destination's formats, merged cells and the rest, and
  a cell it makes where the sheet had none takes the format its row or
  column shows, as a typed one does. A blank cell clears the value it
  lands on. ``values`` refuses merged cells at the destination the source
  does not have, as Excel refuses them.
- ``formats`` brings formats, merged cells and conditional formats, and the
  heights and widths of whole rows and columns; ``comments`` notes and
  threads; ``validation`` validation; ``column_widths`` the columns' widths
  and whether they are hidden.
- ``all_except_borders`` brings all but borders, the destination's
  staying; ``formulas_and_number_formats`` and ``values_and_number_formats``
  formulas or values and number formats, the rest of the destination's
  formats staying; ``all_merging_conditional_formats`` brings all, the
  destination's conditional formats staying beside the new ones.
- ``link`` writes a formula reading each cell copied: ``=$A$1`` for a single
  cell, and for more ``=A1`` and on down and across the whole destination.
- An operation adds, subtracts, multiplies or divides what comes into what
  is there. Two numbers give a number, ``#DIV/0!`` or ``#NUM!``; a formula
  on either side gives the formula combining them, as ``=(1+1)+10`` or
  ``=3+(E1*2)``, a number written as
  :func:`~pyofficeeditor.excel._formulas.number_in_formula` writes one. A
  blank counts as 0, where the other side is not blank too, and text, a
  logical or an error on either side leaves the destination's value.
  An operation onto an array formula is refused, as Excel refuses it.
- Skipping blanks leaves the destination as it is where the source cell is
  blank, its format too, unless it is part of merged cells; notes, links,
  validation and conditional formats come only where the source has them,
  the destination keeping its own elsewhere.
- Transposing turns the copy's rows into columns, and its references as
  :func:`~pyofficeeditor.excel._formulas.transposed_formula` says. A
  conditional format's formulas move as the cells would under a plain
  copy. A transposed paste over its own source is refused, as Excel
  refuses it.

Not yet, each refused before anything changes: a copy of a whole table,
which Excel makes a new table of, and cells taken out of a table, which
Excel gives the table style's look as formats of their own, when the paste
brings formulas or formats; a paste over a table's header or totals row or
across its edge, and one just below or beside a table, which Excel makes
it take in; a paste over cells a formula spilled into, which Excel then
shows ``#SPILL!``; shapes inside the source, which Excel copies with the
cells; pivot tables; what-if data tables. Of Paste Special: merged cells
at the destination a paste covers in part, which Excel pastes into in
ways of its own; a spill transposed, whose size only calculating it
gives; whole rows and columns transposed or linked; a transpose or a link
on a sheet whose filter hides some of the copy; column widths transposed;
links onto merged cells; an operation onto a spill.
"""

from __future__ import annotations

import bisect
import datetime as dt
import math
from collections import defaultdict
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING, Literal

from pyofficeeditor._xml import Element, XmlDocument, local_name
from pyofficeeditor.excel._addresses import cut_conditional_formats, cut_validations
from pyofficeeditor.excel._comments import CopiedNote, ThreadedComment
from pyofficeeditor.excel._conditional import CONDITION_FORMULA_TYPES, ConditionalFormatting, ConditionalRule
from pyofficeeditor.excel._formulas import (
    REF_ERROR,
    copied_formula,
    join_areas,
    number_in_formula,
    quote_sheet_name,
    transposed_formula,
)
from pyofficeeditor.excel._hyperlinks import Hyperlink
from pyofficeeditor.excel._insertformat import FormatMatcher, column_styles
from pyofficeeditor.excel._reference import MAX_COLUMN, MAX_ROW, AxisRef, CellRef, RangeRef
from pyofficeeditor.excel._schema import WORKSHEET_CHILD_ORDER, insert_in_schema_order
from pyofficeeditor.excel._values import CellError, CellValue, datetime_to_serial, read_value, write_value
from pyofficeeditor.excel._xstring import decode, escape

if TYPE_CHECKING:
    from pyofficeeditor.excel.worksheet import Worksheet

#: What a paste brings, by the Paste Special choice it is.
PasteKind = Literal[
    "all",
    "formulas",
    "values",
    "formats",
    "comments",
    "validation",
    "all_except_borders",
    "column_widths",
    "formulas_and_number_formats",
    "values_and_number_formats",
    "all_merging_conditional_formats",
    "link",
]
#: How Paste Special combines what it brings with what is there.
PasteOperation = Literal["add", "subtract", "multiply", "divide"]

#: A block, as ``(top, bottom, left, right)``.
Block = tuple[int, int, int, int]

#: The operator an operation writes between what is there and what comes.
_OPERATORS: dict[str, str] = {"add": "+", "subtract": "-", "multiply": "*", "divide": "/"}
#: What a cell holds, apart from its format.
_CONTENT_ATTRIBUTES = ("t", "cm", "vm", "ph")
_CONTENT_CHILDREN = ("f", "v", "is")


@dataclass(frozen=True)
class _Layers:
    """What one kind of paste brings, and so replaces."""

    #: The cells' contents: formulas and values, values alone, or formulas
    #: linking to them.
    contents: Literal["formulas", "values", "link"] | None = None
    #: The cells' formats: whole, all but borders, or number formats alone.
    style: Literal["whole", "except_borders", "number_format"] | None = None
    merges: bool = False
    notes: bool = False
    links: bool = False
    validation: bool = False
    #: Conditional formats, in place of the destination's or beside them.
    conditional: Literal["replace", "beside"] | None = None
    widths: bool = False
    #: Whole rows' and columns' own settings.
    settings: bool = False
    #: Shapes that move with the cells, which a paste of everything copies.
    shapes: bool = False


_LAYERS: dict[str, _Layers] = {
    "all": _Layers(
        "formulas", "whole", merges=True, notes=True, links=True, validation=True, conditional="replace",
        settings=True, shapes=True,
    ),
    "formulas": _Layers("formulas"),
    "values": _Layers("values"),
    "formats": _Layers(style="whole", merges=True, conditional="replace", settings=True),
    "comments": _Layers(notes=True),
    "validation": _Layers(validation=True),
    "all_except_borders": _Layers(
        "formulas", "except_borders", merges=True, notes=True, links=True, validation=True, conditional="replace",
        settings=True, shapes=True,
    ),
    "column_widths": _Layers(widths=True),
    "formulas_and_number_formats": _Layers("formulas", "number_format"),
    "values_and_number_formats": _Layers("values", "number_format"),
    "all_merging_conditional_formats": _Layers(
        "formulas", "whole", merges=True, notes=True, links=True, validation=True, conditional="beside",
        settings=True, shapes=True,
    ),
    "link": _Layers("link"),
}  # fmt: skip


@dataclass(frozen=True)
class _How:
    """A paste as asked for."""

    kind: str
    layers: _Layers
    operation: str | None = None
    skip_blanks: bool = False
    transpose: bool = False

    @property
    def plain(self) -> bool:
        """Whether this is the plain paste, Ctrl+V's, rather than Paste
        Special's."""
        return self.kind == "all" and self.operation is None and not self.skip_blanks and not self.transpose


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

    def index(self, place: int) -> int:
        """The source's row or column that lands ``place`` into the pasted
        block."""
        index = self.low + place
        for skipped in self.skipped:
            if skipped > index:
                break
            index += 1
        return index

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
    #: A cell a spill filled, which a paste of formulas leaves empty.
    spilled: bool = False
    #: The formula the cell works out and the cell it is written for, which
    #: an operation combines: for a cell of an array or a spill, the one
    #: its first cell holds.
    computes: tuple[str, CellRef] | None = None


@dataclass
class _Copy:
    """Everything a copy takes from its source, read before anything is
    pasted, so a paste over its own source reads what was there."""

    sheet: Worksheet
    block: RangeRef
    rows: _Axis
    columns: _Axis
    whole: str | None
    cells: list[_Cell]
    merges: list[Block]
    #: The places merged cells cover, which skipping blanks does not skip.
    merged: frozenset[tuple[int, int]]
    notes: list[tuple[int, int, CopiedNote]]
    threads: list[tuple[int, int, ThreadedComment]]
    links: list[tuple[Block, Hyperlink]]
    validations: list[tuple[Element, list[RangeRef]]]
    formats: list[tuple[ConditionalFormatting, list[RangeRef]]]
    row_settings: list[tuple[int, dict[str, str]]]
    #: Each whole column's place, width, whether hidden, and style.
    column_settings: list[tuple[int, float | None, bool, int | None]]


def copy_cells(
    sheet: Worksheet,
    cells: str | RangeRef,
    target: Worksheet,
    destination: str | RangeRef,
    *,
    paste: str = "all",
    operation: str | None = None,
    skip_blanks: bool = False,
    transpose: bool = False,
) -> RangeRef:
    """Copy ``cells`` of ``sheet`` and paste them at ``destination`` on
    ``target``, as Excel's Copy and Paste, or its Paste Special, does; the
    block pasted."""
    how = _asked(paste, operation, skip_blanks=skip_blanks, transpose=transpose)
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
    if whole is not None and (how.transpose or how.layers.contents == "link"):
        raise ValueError(f"whole {whole} {'transposed' if how.transpose else 'linked'} are not supported yet.")
    if (how.transpose or how.layers.contents == "link") and (rows.skipped or columns.skipped):
        raise ValueError(
            f"the sheet's filter hides some of {block.a1}; a paste "
            f"{'transposed' if how.transpose else 'linked'} from it is not supported yet."
        )
    height, width = (columns.size, rows.size) if how.transpose else (rows.size, columns.size)
    tiles = _tiles(landing, height, width, single=landed_whole is None and landing.is_single_cell)
    taken = _take(sheet, block, rows, columns, whole)
    _refuse(taken, target, tiles, how)
    if how.layers.contents == "link":
        _link(taken, target, RangeRef(tiles[0].start, tiles[-1].end))
    else:
        for tile in tiles:
            _paste(taken, target, tile, how)
    target.workbook.mark_values_changed()
    target.invalidate()
    return RangeRef(tiles[0].start, tiles[-1].end)


def _asked(paste: str, operation: str | None, *, skip_blanks: bool, transpose: bool) -> _How:
    """The paste asked for, refusing what does not go together."""
    layers = _LAYERS.get(paste)
    if layers is None:
        raise ValueError(f"{paste!r} is not a kind of paste; it is one of: {', '.join(_LAYERS)}.")
    if operation is not None and operation not in _OPERATORS:
        raise ValueError(f"{operation!r} is not an operation; it is one of: {', '.join(_OPERATORS)}.")
    if operation is not None and layers.contents not in ("formulas", "values"):
        raise ValueError(f"an operation combines values, so it goes with a paste that brings them, not {paste!r}.")
    if layers.contents == "link" and (operation is not None or skip_blanks or transpose):
        raise ValueError("Paste Link links each cell as it is, with no operation, skipping or transposing.")
    if layers.widths and transpose:
        raise ValueError("column widths transposed are not supported yet.")
    return _How(paste, layers, operation, skip_blanks, transpose)


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


def _refuse(taken: _Copy, target: Worksheet, tiles: list[RangeRef], how: _How) -> None:
    """Refuse, before anything changes, what Excel refuses and what this
    does not do yet."""
    sheet, block, layers = taken.sheet, taken.block, how.layers
    for tile in tiles:
        if how.transpose and target is sheet and tile.intersects(block):
            raise ValueError(
                f"the transposed paste over {tile.a1} overlaps {block.a1}, which Excel refuses: \"This selection "
                "isn't valid. Make sure the copy and paste areas don't overlap unless they are the same size "
                'and shape."'
            )
        _refuse_merges(taken, target, tile, how)
        if layers.contents is not None:
            _refuse_arrays(target, tile, how)
        if layers.contents is not None or layers.style is not None:
            _refuse_tables(target, tile)
        _refuse_pivots_and_data_tables(target, tile)
    brings_tables = layers.contents in ("formulas", "link") or layers.style in ("whole", "except_borders")
    for table in sheet.tables if brings_tables else ():
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
    for shape in sheet.shapes if layers.shapes else ():
        if shape.cells and block.contains(RangeRef.parse(shape.cells)):
            raise ValueError(
                f"{block.a1} holds the shape {shape.name!r}, which Excel copies with the cells; "
                "that is not supported yet."
            )
    if how.transpose and any(cell.spill is not None for cell in taken.cells):
        raise ValueError(
            f"{block.a1} holds a formula that spills, which transposed spills as far as calculating it "
            "says; that is not supported yet."
        )
    _refuse_pivots_and_data_tables(sheet, block)


def _refuse_merges(taken: _Copy, target: Worksheet, tile: RangeRef, how: _How) -> None:
    """Merged cells in the way. Measured: a paste that brings merged cells
    takes those it covers whole and refuses one it covers in part; a paste
    of values refuses any the source does not have there."""
    layers = how.layers
    landed = {_moved(places, tile, how).a1 for places in taken.merges}
    for merged in target.merged_ranges:
        if not merged.intersects(tile):
            continue
        whole = tile.contains(merged)
        if layers.merges and whole:
            continue
        if layers.merges and how.plain:
            raise ValueError(
                f"the paste over {tile.a1} would take part of the merged cells {merged.a1}, "
                "which Excel refuses: \"We can't do that to a merged cell.\""
            )
        if how.kind == "values" and merged.a1 not in landed:
            raise ValueError(
                f"the paste of values over {tile.a1} meets the merged cells {merged.a1}, which the copy does not "
                'have there, and Excel refuses it: "To do this, all the merged cells need to be the same size."'
            )
        if layers.merges or not whole or layers.contents == "link":
            raise ValueError(
                f"the paste over {tile.a1} meets the merged cells {merged.a1}, which Paste Special pastes into in "
                "ways of its own; that is not supported yet."
            )


def _refuse_arrays(target: Worksheet, tile: RangeRef, how: _How) -> None:
    for array, dynamic in _arrays(target):
        if not array.intersects(tile):
            continue
        if how.operation is not None and not dynamic:
            raise ValueError(
                f"the paste over {tile.a1} would work an operation into the array formula over {array.a1}, "
                "which Excel refuses: \"You can't change part of an array.\""
            )
        if how.operation is not None:
            raise ValueError(
                f"the paste over {tile.a1} would work an operation into the formula spilling over {array.a1}; "
                "that is not supported yet."
            )
        if not dynamic and not tile.contains(array):
            raise ValueError(
                f"the paste over {tile.a1} would take part of the array formula over {array.a1}, "
                "which Excel refuses: \"You can't change part of an array.\""
            )
        if dynamic and array.start not in tile:
            raise ValueError(
                f"the paste over {tile.a1} lands in cells the formula at {array.start.a1} spilled into, "
                "which Excel then shows #SPILL!; that is not supported yet."
            )


def _refuse_tables(target: Worksheet, tile: RangeRef) -> None:
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
    present: set[tuple[int, int]] = set()
    for number, row in sorted(sheet.rows_by_number().items()):
        place_row = rows.place(number)
        if place_row is None:
            continue
        for cell in row.children_named("c"):
            reference = _address(cell)
            place_column = None if reference is None else columns.place(reference.column)
            if reference is None or place_column is None:
                continue
            present.add((number, reference.column))
            cells.append(_taken_cell(sheet, block, arrays, reference, cell, (place_row, place_column)))
    if whole is None:
        cells.extend(_styled_blanks(sheet, block, rows, columns, present))

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
    merged = frozenset(
        (row, column) for top, bottom, left, right in merges for row in range(top, bottom + 1) for column in range(left, right + 1)
    )
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
    column_settings: list[tuple[int, float | None, bool, int | None]] = []
    if whole == "rows":
        present_rows = sheet.rows_by_number()
        for number in range(block.top, block.bottom + 1):
            place = rows.place(number)
            if place is not None:
                row = present_rows.get(number)
                settings = {} if row is None else {k: v for k, v in row.attributes.items() if k not in ("r", "spans")}
                row_settings.append((place, settings))
    if whole == "columns":
        for column in range(block.left, block.right + 1):
            place = columns.place(column)
            if place is not None:
                column_settings.append(
                    (place, sheet.column_width(column), sheet.column_hidden(column), sheet.column_style_index(column))
                )
    return _Copy(
        sheet, block, rows, columns, whole, cells, merges, merged, notes, threads, links, validations, formats,
        row_settings, column_settings,
    )  # fmt: skip


def _taken_cell(
    sheet: Worksheet,
    block: RangeRef,
    arrays: dict[RangeRef, tuple[CellRef, Element, bool]],
    reference: CellRef,
    cell: Element,
    place: tuple[int, int],
) -> _Cell:
    """A cell as the copy takes it, with the cell its formula is read from,
    the size of the array it starts and the spill it anchors."""
    row, column = place
    for span, (master, master_cell, dynamic) in arrays.items():
        if reference not in span:
            continue
        text = sheet.formula_of(master_cell, master) or ""
        computes = (escape(text), master) if text else None
        if dynamic:
            if reference != master:
                return _Cell(row, column, _detached(cell), reference, spilled=True, computes=computes)
            return _Cell(row, column, _detached(cell), reference, spill=span, computes=computes)
        part = _overlap(span, block)
        element = _detached(cell)
        for formula in list(element.children_named("f")):
            element.remove(formula)
        if reference != part.start:
            return _Cell(row, column, element, reference, computes=computes)
        formula = Element.create("f", {"t": "array"})
        formula.set_text(escape(text))
        element.insert(0, formula)
        return _Cell(row, column, element, master, array=(part.height, part.width), computes=computes)
    element = _detached(cell)
    formula = element.child("f")
    computes = None
    if formula is not None:
        text = sheet.formula_of(cell, reference)
        for marker in ("t", "si", "ref"):
            formula.unset(marker)
        if text:
            formula.set_text(escape(text))
            computes = (escape(text), reference)
    return _Cell(row, column, element, reference, computes=computes)


def _styled_blanks(
    sheet: Worksheet, block: RangeRef, rows: _Axis, columns: _Axis, present: set[tuple[int, int]]
) -> list[_Cell]:
    """The blank cells of a block that show a format of their row's or
    their column's, which, measured, a copy brings as formats of their
    own: the row's where it has one, and the column's otherwise."""
    row_styles = {
        number: style
        for number, row in sheet.rows_by_number().items()
        if block.top <= number <= block.bottom
        and row.get("customFormat") in ("1", "true")
        and (style := row.get("s")) not in (None, "0")
    }
    by_column = {
        column: style
        for column, style in column_styles(sheet.document.root).items()
        if block.left <= column <= block.right and style != "0"
    }
    found: list[_Cell] = []
    for number in range(block.top, block.bottom + 1) if by_column else sorted(row_styles):
        place_row = rows.place(number)
        if place_row is None:
            continue
        own = row_styles.get(number)
        for column in range(block.left, block.right + 1) if own else sorted(by_column):
            style = own or by_column.get(column)
            place_column = columns.place(column)
            if style is None or place_column is None or (number, column) in present:
                continue
            reference = CellRef(number, column)
            found.append(_Cell(place_row, place_column, Element.create("c", {"r": reference.a1, "s": style}), reference))
    return found


def _validation_elements(sheet: Worksheet) -> list[Element]:
    container = sheet.document.root.child("dataValidations")
    return [] if container is None else list(container.children_named("dataValidation"))


# ----------------------------------------------------------------------
# Pasting
# ----------------------------------------------------------------------


def _paste(taken: _Copy, target: Worksheet, tile: RangeRef, how: _How) -> None:
    """Put what the paste brings of the copy in ``tile``, in place of what
    was there."""
    layers = how.layers
    if layers.merges:
        _unmerge(target, tile)
    if layers.contents is not None or layers.style is not None:
        _paste_cells(taken, target, tile, how)
    if layers.merges:
        for places in taken.merges:
            _add_merge(target, _moved(places, tile, how))
    if layers.notes:
        _paste_notes(taken, target, tile, how)
    if layers.links:
        landed = [(_moved(places, tile, how), link) for places, link in taken.links]
        for area in [area for area, _ in landed] if how.skip_blanks else [tile]:
            target.remove_hyperlink(area)
        for area, link in landed:
            target.add_hyperlink(area, link.target, location=link.location, display=link.display, tooltip=link.tooltip)
    if layers.validation:
        cut = [area for _, touched in taken.validations for piece in touched for area in _landing(taken, tile, piece, how)]
        for area in cut if how.skip_blanks else [tile]:
            cut_validations(target.document.root, area)
        for element, touched in taken.validations:
            _paste_validation(taken, target, tile, element, touched, how)
    if layers.conditional is not None:
        if layers.conditional == "replace":
            cut = [area for _, touched in taken.formats for piece in touched for area in _landing(taken, tile, piece, how)]
            for area in cut if how.skip_blanks else [tile]:
                cut_conditional_formats(target.document.root, area)
        for formatting, touched in taken.formats:
            _paste_format(taken, target, tile, formatting, touched, how)
    if layers.settings:
        _paste_settings(taken, target, tile)
    if layers.widths:
        for column in range(taken.block.left, taken.block.right + 1):
            place = taken.columns.place(column)
            if place is not None:
                target.set_column_width(tile.left + place, taken.sheet.column_width(column))
                target.set_column_hidden(tile.left + place, taken.sheet.column_hidden(column))


def _paste_cells(taken: _Copy, target: Worksheet, tile: RangeRef, how: _How) -> None:
    """Each cell of ``tile`` as the paste leaves it: what it brings of the
    source cell landing there, and what it leaves of the cell there."""
    sources = {_at(tile, cell.row, cell.column, how): cell for cell in taken.cells}
    held = {
        reference: cell
        for number, row in target.rows_by_number().items()
        if tile.top <= number <= tile.bottom
        for cell in row.children_named("c")
        if (reference := _address(cell)) is not None and reference in tile
    }
    if how.layers.contents is not None:
        _drop_spills(target, tile, sources, how)
    lands = _Landing(taken, target, tile, how)
    placed: dict[int, list[Element]] = defaultdict(list)
    for at in sorted(sources.keys() | held.keys(), key=lambda reference: (reference.row, reference.column)):
        element = lands.cell(sources.get(at), held.get(at), at)
        if element is not None:
            placed[at.row].append(element)
    target.remove_cells(tile)
    for number in sorted(placed):
        target.place_cells(number, placed[number])


class _Landing:
    """How one source cell lands on one cell of a tile."""

    def __init__(self, taken: _Copy, target: Worksheet, tile: RangeRef, how: _How) -> None:
        self.taken, self.target, self.tile, self.how = taken, target, tile, how
        self.book = target.workbook
        self.matcher = FormatMatcher(self.book.styles)

    def cell(self, source: _Cell | None, held: Element | None, at: CellRef) -> Element | None:
        """The cell at ``at`` once the paste is made, or ``None`` for none."""
        how, layers = self.how, self.how.layers
        if how.skip_blanks and _blank(source) and (source is None or (source.row, source.column) not in self.taken.merged):
            return held
        if source is not None and source.spilled and layers.contents == "formulas" and how.operation is None:
            source = None
        formats_of_whole = self.taken.whole is not None and layers.contents is None
        if formats_of_whole and source is None and held is not None:
            # Measured: a paste of formats of whole rows or columns gives a
            # cell there the format the source shows where it holds none.
            source = self._shown(at)
        elif (
            formats_of_whole
            and source is not None
            and held is None
            and self.matcher.same(source.element.get("s") or "0", self._shown(at).element.get("s") or "0")
        ):
            # Measured: nor does it make a cell to hold the format the row
            # or column it lands in shows anyway.
            return None
        element = _detached(held) if held is not None else _new_cell(self.target, at)
        element.set("r", at.a1)
        if layers.contents is not None:
            self._contents(source, held, element, at)
        if layers.style is not None:
            style = self._style(source, element)
            if style is None:
                element.unset("s")
            else:
                element.set("s", style)
        empty = all(element.child(name) is None for name in _CONTENT_CHILDREN)
        if empty and held is None and layers.style not in ("whole", "except_borders"):
            return None
        return None if empty and (element.get("s") or "0") == "0" else element

    def _shown(self, at: CellRef) -> _Cell:
        """For a paste of whole rows or columns, the source cell landing at
        ``at`` as its row or column shows it where it holds none."""
        taken, tile = self.taken, self.tile
        row = taken.rows.index(at.row - tile.top) if taken.whole == "rows" else at.row
        column = taken.columns.index(at.column - tile.left) if taken.whole == "columns" else at.column
        reference = CellRef(row, column)
        style = _shown_style(taken.sheet, reference)
        element = Element.create("c", {"r": reference.a1} if style is None else {"r": reference.a1, "s": style})
        return _Cell(at.row - tile.top, at.column - tile.left, element, reference)

    def _contents(self, source: _Cell | None, held: Element | None, element: Element, at: CellRef) -> None:
        """Put what the source cell holds in ``element``, as the paste
        brings it."""
        how = self.how
        if how.operation is not None:
            self._operate(source, held, element, at)
            return
        _clear_contents(element)
        if source is None:
            return
        if how.layers.contents == "values":
            value = self._raw(source.element)
            if value is not None:
                write_value(element, value, shared_strings=self.book.ensure_shared_strings(), styles=None)
            return
        taken = _detached(source.element)
        for name in _CONTENT_ATTRIBUTES:
            value = taken.get(name)
            if value is not None:
                element.set(name, value)
        for name in _CONTENT_CHILDREN:
            for child in list(taken.children_named(name)):
                _append_content(element, child)
        formula = element.child("f")
        if formula is None:
            return
        if formula.text:
            formula.set_text(self._moved(formula.text, source.origin, at))
        if source.array is not None:
            height, width = reversed(source.array) if how.transpose else source.array
            formula.set("ref", RangeRef(at, CellRef(at.row + height - 1, at.column + width - 1)).a1)
        if source.spill is not None:
            rows, columns = at.row - source.origin.row, at.column - source.origin.column
            formula.set("ref", RangeRef(source.spill.start.translated(rows, columns), source.spill.end.translated(rows, columns)).a1)

    def _moved(self, formula: str, origin: CellRef, at: CellRef) -> str:
        """A formula written for ``origin`` as it lands at ``at``."""
        if self.how.transpose:
            return transposed_formula(formula, self.taken.block, self.tile.start, self.taken.sheet.name)
        return copied_formula(formula, at.row - origin.row, at.column - origin.column)

    def _operate(self, source: _Cell | None, held: Element | None, element: Element, at: CellRef) -> None:
        """Combine what comes with what is there. Measured: two numbers give
        a number; a formula on either side gives a formula combining them,
        itself in brackets and a number as it is; a blank counts as 0 unless
        both sides are blank; and text, a logical or an error on either side
        leaves what is there."""
        how = self.how
        if how.layers.contents == "formulas" and source is not None and source.computes is not None:
            text, origin = source.computes
            coming: tuple[str, float | str] | None = ("formula", decode(self._moved(text, origin, at)))
        else:
            coming = self._operand(None if source is None else source.element)
        written = None if held is None else self.target.formula_of(held, at)
        there = ("formula", written) if written else self._operand(held)
        if coming is None or there is None or (coming[0] == "blank" and there[0] == "blank"):
            return
        _clear_contents(element)
        operator = _OPERATORS[how.operation or "add"]
        if coming[0] != "formula" and there[0] != "formula":
            left = there[1] if isinstance(there[1], float) else 0.0
            right = coming[1] if isinstance(coming[1], float) else 0.0
            write_value(element, _worked(left, right, how.operation or "add"), shared_strings=None, styles=None)
            return
        formula = Element.create("f")
        formula.set_text(escape(f"{_operand_text(there)}{operator}{_operand_text(coming)}"))
        _append_content(element, formula)

    def _operand(self, cell: Element | None) -> tuple[str, float | str] | None:
        """A side of an operation: a number, or a blank; ``None`` for text,
        a logical or an error, which leave what is there."""
        value = None if cell is None else self._raw(cell)
        if value is None:
            return ("blank", 0.0)
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return ("number", float(value))
        return None

    def _raw(self, cell: Element) -> CellValue:
        """What a cell holds as it is stored: a date as its serial number,
        whatever its format, and text without its runs."""
        value = read_value(cell, shared_strings=self.book.shared_strings, styles=None)
        if isinstance(value, (dt.date, dt.time)):
            return datetime_to_serial(value, epoch_1904=self.book.epoch_1904)
        return value

    def _style(self, source: _Cell | None, cell: Element) -> str | None:
        """The format ``cell`` ends with, as the paste brings it."""
        layers = self.how.layers
        brought = None if source is None else source.element.get("s")
        if layers.style == "whole":
            return brought
        styles = self.book.styles
        if styles is None:
            return cell.get("s")
        coming, there = styles.cell_format(_index(brought)), styles.cell_format(_index(cell.get("s")))
        # Measured: where nothing would change, the cell keeps the format it
        # has, its entry and all, though another entry says the same.
        if layers.style == "except_borders":
            if coming.border == there.border:
                return brought
            wanted = replace(coming, border=there.border)
        else:
            if coming.number_format == there.number_format:
                return cell.get("s")
            wanted = replace(there, number_format=coming.number_format)
        index = styles.ensure_cell_format(wanted)
        return str(index) if index else None


def _new_cell(sheet: Worksheet, at: CellRef) -> Element:
    """A cell a paste makes where the sheet had none, with the format it
    showed: measured, a value pasted there takes it, as one typed does."""
    style = _shown_style(sheet, at)
    return Element.create("c", {"r": at.a1} if style is None else {"r": at.a1, "s": style})


def _shown_style(sheet: Worksheet, reference: CellRef) -> str | None:
    """The format a cell with nothing of its own shows: its row's, where the
    row has one, and its column's otherwise."""
    row = sheet.rows_by_number().get(reference.row)
    style = row.get("s") if row is not None and row.get("customFormat") in ("1", "true") else None
    if style not in (None, "0"):
        return style
    index = sheet.column_style_index(reference.column)
    return None if index is None else str(index)


def _blank(source: _Cell | None) -> bool:
    """Whether a source cell holds nothing, which skipping blanks skips."""
    return source is None or all(source.element.child(name) is None for name in _CONTENT_CHILDREN)


def _clear_contents(element: Element) -> None:
    for name in _CONTENT_ATTRIBUTES:
        element.unset(name)
    for name in _CONTENT_CHILDREN:
        for child in list(element.children_named(name)):
            element.remove(child)


def _append_content(element: Element, child: Element) -> None:
    """Add a formula, value or inline text to a cell in the order the
    schema gives them, ahead of anything else it holds."""
    def rank(node: Element) -> int:
        name = local_name(node.name)
        return _CONTENT_CHILDREN.index(name) if name in _CONTENT_CHILDREN else len(_CONTENT_CHILDREN)

    for existing in element.elements():
        if rank(existing) > rank(child):
            element.insert_before(existing, child)
            return
    element.append(child)


def _operand_text(side: tuple[str, float | str]) -> str:
    kind, value = side
    if kind == "formula":
        return f"({value})"
    return number_in_formula(value if isinstance(value, float) else 0.0)


def _worked(left: float, right: float, operation: str) -> float | CellError:
    """Two numbers combined, as Excel's operations combine them: a division
    by 0 is ``#DIV/0!`` and a result too large for a number ``#NUM!``."""
    if operation == "divide":
        if right == 0:
            return CellError("#DIV/0!")
        result = left / right
    elif operation == "subtract":
        result = left - right
    elif operation == "multiply":
        result = left * right
    else:
        result = left + right
    return CellError("#NUM!") if math.isinf(result) or math.isnan(result) else result


def _index(style: str | None) -> int | None:
    return int(style) if style is not None and style.isdigit() else None


def _unmerge(target: Worksheet, tile: RangeRef) -> None:
    """Take out the merged cells a paste covers, which the copy's replace."""
    container = target.document.root.child("mergeCells")
    if container is None:
        return
    for entry in list(container.children_named("mergeCell")):
        merged = _block_of(entry.get("ref"))
        if merged is not None and tile.contains(merged):
            container.remove(entry)
    _recount(target.document.root, container, "mergeCell")


def _drop_spills(target: Worksheet, tile: RangeRef, sources: dict[CellRef, _Cell], how: _How) -> None:
    """A spill whose formula the paste replaces goes, and, measured, the
    cells it spilled into with it."""
    for span, dynamic in _arrays(target):
        if not dynamic or span.start not in tile or (how.skip_blanks and _blank(sources.get(span.start))):
            continue
        for spilled in span.cells():
            if spilled not in tile:
                target.remove_cells(RangeRef(spilled, spilled))


def _paste_notes(taken: _Copy, target: Worksheet, tile: RangeRef, how: _How) -> None:
    """The copy's notes and threads in place of the destination's, or,
    skipping blanks, of those where the copy has one."""
    notes = {_at(tile, row, column, how): note for row, column, note in taken.notes}
    threads = {_at(tile, row, column, how): thread for row, column, thread in taken.threads}
    for cell in {note.ref for note in target.comments} | {thread.ref for thread in target.threaded_comments}:
        reference = _parse_cell(cell)
        if reference is None or reference not in tile:
            continue
        if not how.skip_blanks or reference in notes or reference in threads:
            target.remove_comment(reference)
    for at, note in notes.items():
        target.paste_note(at, note)
    for at, thread in threads.items():
        _paste_thread(target, at, thread)


def _paste_thread(target: Worksheet, cell: CellRef, thread: ThreadedComment) -> None:
    """A thread copied, replies and all, as Excel's Paste copies one."""
    target.add_threaded_comment(cell, thread.text, author=thread.author, when=thread.when)
    for reply in thread.replies:
        target.add_threaded_reply(cell, reply.text, author=reply.author, when=reply.when)
    if thread.resolved:
        target.resolve_threaded_comment(cell)


def _paste_validation(
    taken: _Copy, target: Worksheet, tile: RangeRef, element: Element, touched: list[RangeRef], how: _How
) -> None:
    """The source's validation takes the pasted cells in. Measured: on the
    same sheet the validation itself does; on another sheet one identical
    to it there does, or else a new one, its formulas moved as the cells
    are; either way its areas join where they meet."""
    landed = [area for area in touched for area in _landing(taken, tile, area, how)]
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
    taken: _Copy,
    target: Worksheet,
    tile: RangeRef,
    formatting: ConditionalFormatting,
    touched: list[RangeRef],
    how: _How,
) -> None:
    """The source's conditional format comes as a new block over the
    pasted cells, after the sheet's others. Measured, its formulas move as
    the cells would under a plain copy, from the block's own first cell to
    the new block's, transposed or not."""
    landed = [area for area in touched for area in _landing(taken, tile, area, how)]
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


def _paste_settings(taken: _Copy, target: Worksheet, tile: RangeRef) -> None:
    """Whole rows' heights, formats and whether they are hidden, and whole
    columns' widths, formats and whether they are hidden."""
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
    for place, width, hidden, style in taken.column_settings:
        target.set_column_width(tile.left + place, width)
        target.set_column_hidden(tile.left + place, hidden)
        target.set_column_style_index(tile.left + place, style)


# ----------------------------------------------------------------------
# Paste Link
# ----------------------------------------------------------------------


def _link(taken: _Copy, target: Worksheet, area: RangeRef) -> None:
    """A formula in each cell of ``area`` reading the source cell as far
    into the copy, and on past its end the same way. Measured: a single
    cell is read as ``$A$1`` from every cell, and a block as ``A1``; a sheet
    other than the source's is named."""
    block, sheet = taken.block, taken.sheet
    prefix = "" if target is sheet else f"{quote_sheet_name(sheet.name)}!"
    placed: dict[int, list[Element]] = defaultdict(list)
    held = {
        reference: cell
        for number, row in target.rows_by_number().items()
        if area.top <= number <= area.bottom
        for cell in row.children_named("c")
        if (reference := _address(cell)) is not None and reference in area
    }
    for at in area.cells():
        row, column = block.top + at.row - area.top, block.left + at.column - area.left
        if block.is_single_cell:
            read = CellRef(block.top, block.left, True, True).a1
        elif 1 <= row <= MAX_ROW and 1 <= column <= MAX_COLUMN:
            read = CellRef(row, column).a1
        else:
            read = REF_ERROR
        element = _detached(held[at]) if at in held else _new_cell(target, at)
        _clear_contents(element)
        formula = Element.create("f")
        formula.set_text(escape(prefix + read))
        _append_content(element, formula)
        placed[at.row].append(element)
    target.remove_cells(area)
    for number in sorted(placed):
        target.place_cells(number, placed[number])


# ----------------------------------------------------------------------
# Places and blocks
# ----------------------------------------------------------------------


def _at(tile: RangeRef, row: int, column: int, how: _How) -> CellRef:
    """Where a cell ``row`` and ``column`` places into the copy lands."""
    if how.transpose:
        return CellRef(tile.top + column, tile.left + row)
    return CellRef(tile.top + row, tile.left + column)


def _landing(taken: _Copy, tile: RangeRef, area: RangeRef, how: _How) -> list[RangeRef]:
    """Where a source area lands in ``tile``: a block for each stretch of
    its rows and columns the copy takes without a gap."""
    return [
        _moved((top, bottom, left, right), tile, how)
        for top, bottom in taken.rows.runs(area.top, area.bottom)
        for left, right in taken.columns.runs(area.left, area.right)
    ]


def _moved(places: Block, tile: RangeRef, how: _How) -> RangeRef:
    top, bottom, left, right = places
    if how.transpose:
        top, bottom, left, right = left, right, top, bottom
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


__all__ = ["PasteKind", "PasteOperation", "copy_cells"]
