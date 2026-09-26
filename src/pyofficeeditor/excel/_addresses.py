"""Every way a part spells a cell address, and how each one moves.

A worksheet records addresses in five different notations, and a row
insertion has to move all of them. Collecting the notations here rather than
writing a bespoke shifter per feature is what makes the list of features
this library can shift a table rather than a pile of special cases.

``sqref``
    Space-separated ranges: ``"I2:I9 K2:K4"``. Data validation, protected
    ranges, ignored errors and conditional formatting all use it, and a
    single cell is spelled ``"A1"`` rather than ``"A1:A1"``.

``ref``
    One range. A sort state, a sort condition, a table, a merge.

``r``
    One cell. A scenario's input cells, a comment.

A drawing anchor
    ``<xdr:col>4</xdr:col><xdr:row>5</xdr:row>``, and the same pair again
    under ``<xdr:to>``. These are **zero-based**: column 4 is E and row 5 is
    the sixth row. Shifting them as if they were one-based puts every shape
    one row and one column off, which no byte comparison catches because the
    file is still valid.

A VML anchor
    ``<x:Anchor>4, 0, 9, 0, 5, 44, 10, 19</x:Anchor>``: left column, left
    offset, top row, top offset, right column, right offset, bottom row,
    bottom offset. Zero-based again, and this is how a comment and a form
    control say which cell they sit on.

The two anchors do not simply move with their cells: how each moves turns
on the placement of the object it places, which
:mod:`~pyofficeeditor.excel._placement` works out.

A validation's or a conditional format's ``sqref`` also loses cells that
are cleared without moving, as Remove Duplicates clears the rows it
removes and a paste clears where it lands: :func:`cut_validations` and
:func:`cut_conditional_formats` cut them out, each in the order Excel
measured for it.
"""

from __future__ import annotations

from collections.abc import Callable

from pyofficeeditor._xml import Element
from pyofficeeditor.excel._formulas import Deletion, Remap, Shift, shift_range
from pyofficeeditor.excel._reference import MAX_COLUMN, MAX_ROW, CellRef, RangeRef


def shift_sqref(raw: str, shift: Shift) -> str:
    """Move every range in an ``sqref``, keeping the spacing convention."""
    moved: list[str] = []
    for piece in raw.split():
        try:
            block = RangeRef.parse(piece)
        except ValueError:
            moved.append(piece)
            continue
        moved.append(shift_range(block, shift).a1)
    return " ".join(moved)


def delete_sqref(raw: str, deletion: Remap, *, stretched: bool = True) -> str | None:
    """Shrink every range in an ``sqref``, or move it as another remap
    does; ``stretched`` is for a validation whose formulas read no cell
    relative to where it applies (see :meth:`Remap.moved_areas`).

    ``None`` when nothing survives, which means the element itself should
    go: an empty ``sqref`` is not something Excel accepts.
    """
    kept: list[str] = []
    for piece in raw.split():
        try:
            block = RangeRef.parse(piece)
        except ValueError:
            kept.append(piece)
            continue
        kept.extend(survivor.a1 for survivor in deletion.moved_areas((block,), stretched=stretched))
    return " ".join(kept) if kept else None


def shift_ref(raw: str, shift: Shift) -> str:
    """Move one range."""
    try:
        block = RangeRef.parse(raw)
    except ValueError:
        return raw
    return shift_range(block, shift).a1


def delete_ref(raw: str, deletion: Remap) -> str | None:
    """Shrink one range, or ``None`` when it is wholly deleted."""
    try:
        block = RangeRef.parse(raw)
    except ValueError:
        return raw
    survivor = deletion.moved_range(block)
    return None if survivor is None else survivor.a1


def shift_cell(raw: str, shift: Shift) -> str:
    """Move one cell address."""
    try:
        cell = CellRef.parse(raw)
    except ValueError:
        return raw
    moved = shift.moved(cell)
    return raw if moved is None else moved.a1


def delete_cell(raw: str, deletion: Remap) -> str | None:
    """Move one cell address, or ``None`` when its cell is deleted."""
    try:
        cell = CellRef.parse(raw)
    except ValueError:
        return raw
    moved = deletion.moved(cell)
    return None if moved is None else moved.a1


def shift_index(value: int, shift: Shift, *, is_row: bool) -> int:
    """Move a zero-based row or column index, as a drawing anchor holds it.

    The caller states which axis it is, because a zero-based number carries
    no clue. Treating one of these as one-based moves every shape by one
    cell, and the file stays valid.
    """
    at = shift.rows_at if is_row else shift.columns_at
    count = shift.row_count if is_row else shift.column_count
    if at is None or not count:
        return value
    limit = (MAX_ROW if is_row else MAX_COLUMN) - 1
    # ``at`` is one-based, the value is not.
    return min(value + count, limit) if value >= at - 1 else value


def delete_index(value: int, deletion: Deletion, *, is_row: bool) -> int | None:
    """The same, for a deletion. ``None`` when that row or column is gone."""
    at = deletion.rows_at if is_row else deletion.columns_at
    count = deletion.row_count if is_row else deletion.column_count
    if at is None or not count:
        return value
    start = at - 1
    if value < start:
        return value
    if value < start + count:
        return None
    return value - count


def collapse_index(value: int, deletion: Deletion, *, is_row: bool) -> int:
    """Where a zero-based index lands when its own row or column is
    deleted: on the edge of what remains, rather than nowhere."""
    at = deletion.rows_at if is_row else deletion.columns_at
    count = deletion.row_count if is_row else deletion.column_count
    if at is None or not count:
        return value
    survivor = delete_index(value, deletion, is_row=is_row)
    return max(at - 1, 0) if survivor is None else survivor


def parse_sqref(raw: str) -> tuple[RangeRef, ...]:
    """The ranges in an ``sqref``, which separates them with spaces."""
    blocks: list[RangeRef] = []
    for piece in raw.split():
        try:
            blocks.append(RangeRef.parse(piece))
        except ValueError:
            continue
    return tuple(blocks)


# ----------------------------------------------------------------------
# Cutting validation and conditional formats back
# ----------------------------------------------------------------------


def cut_validations(root: Element, hole: RangeRef) -> None:
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


def cut_conditional_formats(root: Element, hole: RangeRef) -> None:
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


__all__ = [
    "collapse_index",
    "cut_conditional_formats",
    "cut_validations",
    "delete_cell",
    "delete_index",
    "delete_ref",
    "delete_sqref",
    "parse_sqref",
    "shift_cell",
    "shift_index",
    "shift_ref",
    "shift_sqref",
]
