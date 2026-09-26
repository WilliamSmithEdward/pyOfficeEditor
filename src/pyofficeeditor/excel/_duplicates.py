"""Removing duplicate rows, as Excel's Remove Duplicates does.

Measured in Excel through ``Range.RemoveDuplicates``:

- Values are duplicates as Excel's filter compares text: case aside, a
  ligature as its letters and ``ß`` as ``ss``, while accents, hyphens,
  spaces, full-width forms and zero-width spaces all count. A logical is
  its text, so ``TRUE`` is a duplicate of the text ``true``. A number is a
  duplicate only of a number of the same value shown the same: ``2`` shown
  by a format of whole numbers is one of a plain ``2``, but ``1`` is not one
  of ``1`` shown ``1.00``, nor ``0.1`` of ``0.3-0.2``, and never of text. An
  error is a duplicate of the same error, a blank only of a blank, and empty
  text only of empty text.
- A row is a duplicate when each column compared holds a duplicate of an
  earlier row's; the first of them stays. Hidden rows count.
- The rows kept move up in order, as a sort moves rows: their values,
  formulas, formats, notes and links. The rows removed go below them and
  are cleared of all but their notes: values, formulas, formats and links
  go, and validation and conditional formats over them are cut back, a
  conditional format's rows first and a validation's columns first. Row
  heights, hidden rows, and validation and conditional formats elsewhere
  stay where they are.
- The range stops at the sheet's last cell, and a last row holding a
  formula over a range is taken for a total and left alone.
- Merged cells, and an array formula the move would split, are refused as
  the sort refuses them.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from typing import TYPE_CHECKING

from pyofficeeditor._xml import Element
from pyofficeeditor.excel._collate import sort_key
from pyofficeeditor.excel._reference import CellRef, RangeRef
from pyofficeeditor.excel._sort import check_movable, move_rows, row_values
from pyofficeeditor.excel._tokens import TokenKind, tokenize
from pyofficeeditor.excel._values import CellError, CellValue

if TYPE_CHECKING:
    from pyofficeeditor.excel.worksheet import Worksheet

#: What a value compares as, before the value itself.
_BLANK, _TEXT, _NUMBER, _ERROR = range(4)


def duplicate_key(value: CellValue, shown: str) -> tuple[object, ...]:
    """What ``value`` compares as when duplicates are sought: equal keys are
    duplicates. ``shown`` is the text its number format shows, which only a
    number needs."""
    if value is None:
        return (_BLANK,)
    if isinstance(value, bool):
        return (_TEXT, sort_key("TRUE" if value else "FALSE"))
    if isinstance(value, (int, float)):
        return (_NUMBER, float(value), sort_key(shown))
    if isinstance(value, CellError):
        return (_ERROR, value.code)
    return (_TEXT, sort_key(str(value)))


def remove_duplicate_rows(sheet: Worksheet, block: RangeRef, columns: Sequence[int], *, header: bool) -> int:
    """Remove the rows of ``block`` that duplicate an earlier row in
    ``columns``, the first row left out when it is a header, as Excel's
    Remove Duplicates does. Returns how many rows went."""
    check_movable(sheet, block)
    top = block.top + (1 if header else 0)
    bottom = min(block.bottom, sheet.max_row)
    if top <= bottom and _is_total(sheet, bottom, block):
        bottom -= 1
    if top > bottom:
        return 0
    rows = range(top, bottom + 1)
    kept: list[int] = []
    removed: list[int] = []
    seen: set[tuple[tuple[object, ...], ...]] = set()
    for row, values in zip(rows, row_values(sheet, rows, columns), strict=True):
        key = tuple(
            duplicate_key(value, sheet.get_text(CellRef(row, column)) if _is_number(value) else "")
            for column, value in zip(columns, values, strict=True)
        )
        if key in seen:
            removed.append(row)
        else:
            seen.add(key)
            kept.append(row)
    if not removed:
        return 0
    data = RangeRef(CellRef(top, block.left), CellRef(bottom, block.right))
    move_rows(sheet, {row: top + place for place, row in enumerate(kept + removed) if row != top + place}, data)
    cleared = RangeRef(CellRef(top + len(kept), block.left), CellRef(bottom, block.right))
    sheet.remove_cells(cleared)
    sheet.remove_hyperlink(cleared)
    _cut_validations(sheet.document.root, cleared)
    _cut_conditional_formats(sheet.document.root, cleared)
    sheet.invalidate()
    return len(removed)


def _is_number(value: CellValue) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _is_total(sheet: Worksheet, row: int, block: RangeRef) -> bool:
    """Measured: a last row holding a formula over a range, even one away
    from the data, is left alone, while one reading single cells moves."""
    for column in range(block.left, block.right + 1):
        formula = sheet.get_formula(CellRef(row, column))
        if formula and _reads_a_range(formula):
            return True
    return False


def _reads_a_range(formula: str) -> bool:
    tokens = tokenize(formula)
    for index, token in enumerate(tokens):
        if token.kind is TokenKind.AXIS:
            return True
        if (
            token.kind is TokenKind.REFERENCE
            and index + 2 < len(tokens)
            and tokens[index + 1].kind is TokenKind.TEXT
            and tokens[index + 1].raw.strip() == ":"
            and tokens[index + 2].kind is TokenKind.REFERENCE
        ):
            return True
    return False


# ----------------------------------------------------------------------
# Cutting validation and conditional formats back
# ----------------------------------------------------------------------


def _cut_validations(root: Element, hole: RangeRef) -> None:
    """Measured: a validation loses the cleared cells, its columns cut away
    first: ``C2:D7`` less ``A6:C7`` is ``D2:D7 C2:C5``."""
    container = root.child("dataValidations")
    if container is None:
        return
    for element in list(container.children_named("dataValidation")):
        _cut(element, hole, _columns_first)
    remaining = list(container.children_named("dataValidation"))
    if remaining:
        container.set("count", str(len(remaining)))
    else:
        root.remove(container)


def _cut_conditional_formats(root: Element, hole: RangeRef) -> None:
    """Measured: a conditional format loses the cleared cells, its rows cut
    away first: ``A2:E7`` less ``A6:C7`` is ``A2:E5 D6:E7``."""
    for element in list(root.children_named("conditionalFormatting")):
        _cut(element, hole, _rows_first)


def _cut(element: Element, hole: RangeRef, pieces: Callable[[RangeRef, RangeRef], list[RangeRef]]) -> None:
    """Take ``hole`` out of ``element``'s ``sqref``, each area it meets
    giving way to its pieces in its place, and drop the element if nothing
    is left of it."""
    raw = element.get("sqref") or ""
    try:
        areas = [RangeRef.parse(part).normalized for part in raw.split()]
    except ValueError:
        return
    if not any(area.intersects(hole) for area in areas):
        return
    left: list[RangeRef] = []
    for area in areas:
        left.extend(pieces(area, hole) if area.intersects(hole) else [area])
    parent = element.parent
    if not left:
        if parent is not None:
            parent.remove(element)
        return
    element.set("sqref", " ".join(area.a1 for area in left))


def _overlap(area: RangeRef, hole: RangeRef) -> RangeRef:
    return RangeRef(
        CellRef(max(area.top, hole.top), max(area.left, hole.left)),
        CellRef(min(area.bottom, hole.bottom), min(area.right, hole.right)),
    )


def _rows_first(area: RangeRef, hole: RangeRef) -> list[RangeRef]:
    """``area`` less ``hole``: the rows above and below it across the whole
    area, then the columns beside it in its rows."""
    cut = _overlap(area, hole)
    pieces: list[RangeRef] = []
    if area.top < cut.top:
        pieces.append(RangeRef(CellRef(area.top, area.left), CellRef(cut.top - 1, area.right)))
    if cut.bottom < area.bottom:
        pieces.append(RangeRef(CellRef(cut.bottom + 1, area.left), CellRef(area.bottom, area.right)))
    if area.left < cut.left:
        pieces.append(RangeRef(CellRef(cut.top, area.left), CellRef(cut.bottom, cut.left - 1)))
    if cut.right < area.right:
        pieces.append(RangeRef(CellRef(cut.top, cut.right + 1), CellRef(cut.bottom, area.right)))
    return pieces


def _columns_first(area: RangeRef, hole: RangeRef) -> list[RangeRef]:
    """``area`` less ``hole``: the columns beside it down the whole area,
    then the rows above and below it in its columns."""
    cut = _overlap(area, hole)
    pieces: list[RangeRef] = []
    if area.left < cut.left:
        pieces.append(RangeRef(CellRef(area.top, area.left), CellRef(area.bottom, cut.left - 1)))
    if cut.right < area.right:
        pieces.append(RangeRef(CellRef(area.top, cut.right + 1), CellRef(area.bottom, area.right)))
    if area.top < cut.top:
        pieces.append(RangeRef(CellRef(area.top, cut.left), CellRef(cut.top - 1, cut.right)))
    if cut.bottom < area.bottom:
        pieces.append(RangeRef(CellRef(cut.bottom + 1, cut.left), CellRef(area.bottom, cut.right)))
    return pieces


__all__ = ["duplicate_key", "remove_duplicate_rows"]
