"""A table giving up the rows at its bottom, as Excel's Resize does.

Measured in Excel through ``ListObject.Resize``, and through
``Range.RemoveDuplicates``, which resizes a table the same way once it has
moved the rows it keeps up:

- The table ends on the last row kept, or under it on its totals row, and
  its filter on the last row kept. A sort the table records is over the
  rows kept, its conditions left as they were.
- A totals row moves up to sit under the rows kept, with its cells, notes
  and links, and in the table's columns the rows it passes move down a row
  to make way. Nothing outside the table's columns moves.
- References follow as :class:`~pyofficeeditor.excel._formulas.TableShrink`
  says, in formulas anywhere in the workbook, validation, conditional
  formats, charts and defined names.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pyofficeeditor.excel._formulas import TableShrink, delete_in_formula
from pyofficeeditor.excel._reference import CellRef, RangeRef
from pyofficeeditor.excel._rowcol import remap_references
from pyofficeeditor.excel._sort import move_rows

if TYPE_CHECKING:
    from pyofficeeditor._xml import Element
    from pyofficeeditor.excel._tables import Table
    from pyofficeeditor.excel.worksheet import Worksheet


def shrink_table(sheet: Worksheet, table: Table, kept: int) -> None:
    """Give up ``table``'s data rows below row ``kept``, as Excel's Resize
    does when it shrinks a table from the bottom."""
    data = table.data_range
    if data is None or not data.top <= kept < data.bottom:
        return
    header, totals = table.header_row, table.totals_row
    shrink = TableShrink(
        left=data.left,
        right=data.right,
        starts=(data.top,) if header is None else (header, data.top),
        last=data.bottom,
        kept=kept,
        totals=totals,
    )
    _unshare_split_groups(sheet, shrink)
    remap_references(sheet, shrink)
    for column in table.document.root.descendants("tableColumn"):
        for name in ("calculatedColumnFormula", "totalsRowFormula"):
            _remap_text(column.child(name), shrink, sheet)
    if totals is not None:
        band = RangeRef(CellRef(kept + 1, data.left), CellRef(totals, data.right))
        moves = {row: row + 1 for row in range(kept + 1, totals)} | {totals: kept + 1}
        move_rows(sheet, moves, band, copies=False)
    root = table.document.root
    for element in (root, root.child("autoFilter"), root.child("sortState")):
        if element is not None:
            _remap_ref(element, shrink)
    sheet.invalidate()


def _remap_text(element: Element | None, shrink: TableShrink, sheet: Worksheet) -> None:
    if element is None or not element.text:
        return
    moved = delete_in_formula(element.text, shrink, formula_sheet=sheet.name, target_sheet=sheet.name)
    if moved != element.text:
        element.set_text(moved)


def _remap_ref(element: Element, shrink: TableShrink) -> None:
    raw = element.get("ref")
    try:
        block = RangeRef.parse(raw or "")
    except ValueError:
        return
    moved = shrink.moved_range(block)
    if moved != block:
        element.set("ref", moved.a1)


def _unshare_split_groups(sheet: Worksheet, shrink: TableShrink) -> None:
    """Give each cell of a shared formula its own text where the shrink
    moves a reference in any of them: a shared formula lives once, on its
    first cell, and the rest are read from it by where they stand, which
    a reference that moves in one cell and not the next would break."""
    for other in sheet.workbook.sheets:
        groups: dict[str, list[tuple[Element, str | None]]] = {}
        for reference, cell in other.cell_elements():
            formula = cell.child("f")
            index = None if formula is None else formula.get("si")
            if formula is None or formula.get("t") != "shared" or index is None:
                continue
            groups.setdefault(index, []).append((formula, other.formula_of(cell, reference)))
        # Every derived text is read before any is written, since the
        # followers are read from their first cell's element.
        split = [
            member
            for members in groups.values()
            if any(
                text and delete_in_formula(text, shrink, formula_sheet=other.name, target_sheet=sheet.name) != text
                for _, text in members
            )
            for member in members
        ]
        for formula, text in split:
            for marker in ("t", "si", "ref"):
                formula.unset(marker)
            if text:
                formula.set_text(text)
        if split:
            other.invalidate()


__all__ = ["shrink_table"]
