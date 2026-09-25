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
"""

from __future__ import annotations

from pyofficeeditor.excel._formulas import Deletion, Shift, shift_range
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


def delete_sqref(raw: str, deletion: Deletion) -> str | None:
    """Shrink every range in an ``sqref``.

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
        survivor = deletion.moved_range(block)
        if survivor is not None:
            kept.append(survivor.a1)
    return " ".join(kept) if kept else None


def shift_ref(raw: str, shift: Shift) -> str:
    """Move one range."""
    try:
        block = RangeRef.parse(raw)
    except ValueError:
        return raw
    return shift_range(block, shift).a1


def delete_ref(raw: str, deletion: Deletion) -> str | None:
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


def delete_cell(raw: str, deletion: Deletion) -> str | None:
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


__all__ = [
    "collapse_index",
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
