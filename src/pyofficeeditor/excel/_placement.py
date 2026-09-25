"""How a drawn object follows rows and columns as they come and go.

Excel gives every shape, picture, chart, control and embedded object one of
three placements, and each takes an inserted or deleted row differently.
The file says which in three places, and all three agree:

=====================  =================  ==========================  =====================
Placement              drawing anchor     sheet record ``<anchor>``   VML ``x:ClientData``
=====================  =================  ==========================  =====================
move and size          no ``editAs``      ``moveWithCells`` and       neither flag
                                          ``sizeWithCells``
move but don't size    ``oneCell``        ``moveWithCells`` alone     ``<x:SizeWithCells/>``
don't move or size     ``absolute``       neither                     both flags
=====================  =================  ==========================  =====================

The VML flags read backwards: ``<x:SizeWithCells/>`` is there when the
object does *not* size with its cells. A picture Excel inserts moves but
does not size; a shape, a chart, a group and a Forms button move and size.

Measured against Excel, on one axis at a time:

- **Move and size.** Each edge moves with the cell it sits in, so a row
  inserted inside the object stretches it and one deleted inside it
  shrinks it. A bottom or right edge lying exactly on the line where rows
  go in stays, because it belongs to the row above. An edge whose row is
  deleted lands on the edge of what remains, at offset zero; left where it
  was, a control's anchor turns upside down and Excel refuses the file. An
  object whose rows are all deleted is deleted with them.
- **Move but don't size.** The top-left corner moves with its cell, landing
  on the boundary when its row goes, and the object keeps its size, so the
  far corner is worked out again.
- **Don't move or size.** The object stays where it is on the sheet, and
  both corners are worked out again on the new grid.

A note is none of these. Its box follows the cell it annotates, by as many
rows or columns as the cell moves, keeping its size, whatever the rows
inside the box do, and the note goes when its cell does.

A VML anchor gives its offsets in the pixels of the screen that wrote it,
which is not always 96 to the inch. Where a far corner is worked out again,
it is converted at 96; Excel places a control by its record on the sheet,
measured, not by its VML, and writes the VML again when it saves.

Excel places an anchored shape by its anchor and never updates the
``a:xfrm`` when rows move, but this library reads a shape's box from its
transform, which is exact where turning an anchor into points is not. So
the transform moves with the anchor, by the distance the anchor moved on
the sheet's own grid, and a group's child coordinates stay as Excel leaves
them.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from functools import cached_property
from typing import Literal

from pyofficeeditor._xml import Element, local_name
from pyofficeeditor.excel._formulas import Deletion, Shift
from pyofficeeditor.excel._shapes import (
    MAX_ANCHOR_COLUMN,
    MAX_ANCHOR_ROW,
    SheetGrid,
    anchors_in,
    emu,
)

Placement = Literal["moveAndSize", "moveOnly", "free"]

#: EMU in a pixel, the unit a VML anchor gives its offsets in.
EMU_PER_PIXEL = 9525

#: The elements whose own ``a:xfrm`` says where they are. A chart's frame
#: is left out: Excel writes its transform as zero and the anchor alone
#: places it.
_TRANSFORMED = ("sp", "cxnSp", "pic", "grpSp")

_CLIENT_DATA = re.compile(r"<x:ClientData\b[^>]*>.*?</x:ClientData>", re.DOTALL)
_VML_ANCHOR = re.compile(r"(<x:Anchor>)(.*?)(</x:Anchor>)", re.DOTALL)


@dataclass(frozen=True)
class Corner:
    """One end of an anchor on one axis: a zero-based row or column, and
    how far into it, in EMU."""

    index: int
    offset: int


def _axis(edit: Shift | Deletion) -> tuple[bool, int, int] | None:
    """Which axis an edit touches, where it starts, zero-based, and how
    many rows or columns it takes; ``None`` for one that moves nothing."""
    if edit.rows_at is not None and edit.row_count:
        return True, edit.rows_at - 1, edit.row_count
    if edit.columns_at is not None and edit.column_count:
        return False, edit.columns_at - 1, edit.column_count
    return None


@dataclass(frozen=True)
class AxisEdit:
    """An insertion or a deletion, as it lands on the one axis it touches.

    ``before`` and ``after`` are the sheet's grid either side of it, which
    is what turns a corner into a distance: an inserted row is as tall as
    the sheet now says, and a deleted one as tall as it was.
    """

    is_row: bool
    #: The first row or column inserted or deleted, zero-based.
    start: int
    count: int
    deleting: bool
    before: SheetGrid
    after: SheetGrid

    @classmethod
    def of(
        cls, shift: Shift | None, deletion: Deletion | None, before: SheetGrid, after: SheetGrid
    ) -> AxisEdit | None:
        """The edit a shift or a deletion describes, or ``None`` for one
        that moves nothing."""
        edit = shift if shift is not None else deletion
        axis = None if edit is None else _axis(edit)
        if axis is None:
            return None
        is_row, start, count = axis
        return cls(is_row, start, count, shift is None, before, after)

    @property
    def end(self) -> int:
        """One past the last row or column inserted or deleted."""
        return self.start + self.count

    @property
    def limit(self) -> int:
        """The last index an anchor can name on this axis."""
        return MAX_ANCHOR_ROW if self.is_row else MAX_ANCHOR_COLUMN

    def size(self, index: int, *, after: bool) -> int:
        """How tall a row or how wide a column is, in EMU."""
        grid = self.after if after else self.before
        if self.is_row:
            return emu(grid.row_heights.get(index, grid.default_row))
        return emu(grid.column_widths.get(index, grid.default_column))

    def span(self, first: int, last: int, *, after: bool) -> int:
        """The EMU from the start of ``first`` to the start of ``last``."""
        return sum(self.size(index, after=after) for index in range(first, last))

    @cached_property
    def block(self) -> int:
        """How far the rows or columns inserted or deleted reach, in EMU:
        the distance everything past them moves. Worked out once, since
        every corner past the edit needs it."""
        return self.span(self.start, self.end, after=not self.deleting)


@dataclass(frozen=True)
class Moved:
    """Where an edge went, and how far its point moved on the sheet, in EMU:
    down or right is positive."""

    corner: Corner
    distance: int = 0
    #: Its own row or column was deleted.
    collapsed: bool = False


@dataclass(frozen=True)
class Box:
    """An object's two edges on one axis after an edit."""

    first: Moved
    last: Moved | None
    #: A deletion took every row, or every column, of a move-and-size
    #: object, which Excel deletes with them.
    swallowed: bool = False


def with_cells(corner: Corner, edit: AxisEdit, *, far: bool = False) -> Moved:
    """Where an edge goes when it moves with the cell it sits in.

    ``far`` for a bottom or right edge, which stays when it lies exactly on
    the line where rows or columns go in: measured, it belongs to the row
    above.
    """
    index, offset = corner.index, corner.offset
    if not edit.deleting:
        if index < edit.start or (far and index == edit.start and offset == 0):
            return Moved(corner)
        moved = min(index + edit.count, edit.limit)
        return Moved(Corner(moved, offset), edit.block)
    if index < edit.start:
        return Moved(corner)
    if index >= edit.end:
        return Moved(Corner(index - edit.count, offset), -edit.block)
    distance = edit.span(edit.start, index, after=False) + offset
    return Moved(Corner(edit.start, 0), -distance, collapsed=True)


def walk(edit: AxisEdit, corner: Corner, distance: int) -> Corner:
    """The corner ``distance`` EMU on from another, on the grid after the
    edit. A row of no height holds no corner, and one landing exactly on a
    line belongs to the row below it, as Excel writes a corner."""
    index, offset = corner.index, corner.offset + distance
    while offset < 0 and index > 0:
        index -= 1
        offset += edit.size(index, after=True)
    offset = max(offset, 0)
    while index < edit.limit:
        size = edit.size(index, after=True)
        if offset < size:
            break
        offset -= size
        index += 1
    return Corner(index, offset)


def swallows(first: Corner, last: Corner, start: int, end: int) -> bool:
    """Whether deleting the rows or columns from ``start`` up to ``end``
    takes all of a box on that axis: its first edge is in the deleted block
    and its last is too, or on the block's far line."""
    inside = start <= first.index < end
    return inside and (start <= last.index < end or last == Corner(end, 0))


def move_box(first: Corner, last: Corner | None, placement: Placement, edit: AxisEdit) -> Box:
    """Where an object's edges go under an edit, by its placement.

    ``last`` is ``None`` for a one-cell anchor, which has a corner and a
    size and no far corner to move.
    """
    if placement == "free":
        return Box(Moved(_stay(first, edit)), None if last is None else Moved(_stay(last, edit, far=True)))
    lead = with_cells(first, edit)
    if last is None:
        return Box(lead, None)
    if placement == "moveOnly":
        size = edit.span(first.index, last.index, after=False) + last.offset - first.offset
        return Box(lead, Moved(walk(edit, lead.corner, size), lead.distance))
    tail = with_cells(last, edit, far=True)
    return Box(lead, tail, edit.deleting and swallows(first, last, edit.start, edit.end))


def _stay(corner: Corner, edit: AxisEdit, *, far: bool = False) -> Corner:
    """The corner that keeps a point where it is on the sheet: the one it
    would move to with its cell, walked back by as far as that moved it."""
    moved = with_cells(corner, edit, far=far)
    return walk(edit, moved.corner, -moved.distance)


# ----------------------------------------------------------------------
# Placements, as each part records them
# ----------------------------------------------------------------------


def anchor_placement(anchor: Element) -> Placement:
    """A drawing anchor's placement. A one-cell anchor has a corner and a
    size, which is moving without sizing; an absolute one names no cell."""
    kind = local_name(anchor.name)
    if kind == "oneCellAnchor":
        return "moveOnly"
    if kind == "absoluteAnchor":
        return "free"
    edit_as = anchor.get("editAs")
    if edit_as == "oneCell":
        return "moveOnly"
    if edit_as == "absolute":
        return "free"
    return "moveAndSize"


def record_placement(anchor: Element) -> Placement:
    """A sheet record's ``<anchor>`` placement. Both attributes default to
    false, so an anchor with neither stays where it is."""
    moves = anchor.get("moveWithCells") in ("1", "true")
    sizes = anchor.get("sizeWithCells") in ("1", "true")
    if moves and sizes:
        return "moveAndSize"
    return "moveOnly" if moves else "free"


def vml_placement(client_data: str) -> Placement:
    """A VML object's placement, from flags that say what it does not do."""
    if _vml_flag(client_data, "MoveWithCells"):
        return "free"
    return "moveOnly" if _vml_flag(client_data, "SizeWithCells") else "moveAndSize"


def _vml_flag(client_data: str, name: str) -> bool:
    found = re.search(rf"<x:{name}\s*/>|<x:{name}>(.*?)</x:{name}>", client_data, re.DOTALL)
    return found is not None and (found.group(1) or "").strip().lower() not in ("false", "f", "0")


# ----------------------------------------------------------------------
# Drawing parts and the sheet's own records
# ----------------------------------------------------------------------


def swallowed_nodes(root: Element, deletion: Deletion) -> list[Element]:
    """The top-level nodes of a drawing whose object a deletion takes whole,
    for the caller to remove with every part only they use. Measured: Excel
    deletes a move-and-size shape, chart, group or control with its cells."""
    axis = _axis(deletion)
    if axis is None:
        return []
    is_row, start, count = axis
    found: list[Element] = []
    for node in root.elements():
        anchors = anchors_in(node)
        if not anchors or local_name(anchors[0].name) != "twoCellAnchor":
            continue
        if anchor_placement(anchors[0]) != "moveAndSize":
            continue
        first = _read(anchors[0].child("from"), is_row)
        last = _read(anchors[0].child("to"), is_row)
        if first is not None and last is not None and swallows(first, last, start, start + count):
            found.append(node)
    return found


def move_drawing(root: Element, edit: AxisEdit) -> None:
    """Move every anchor in a drawing part by its placement, and each
    shape's transform with it."""
    for node in root.elements():
        for anchor in anchors_in(node):
            box = _move_ends(anchor, anchor_placement(anchor), edit)
            if box is not None:
                _move_transform(anchor, box, edit)


def move_record_anchors(root: Element, edit: AxisEdit) -> None:
    """Move the ``<anchor>`` each control and embedded object records on the
    sheet, by the placement its own attributes give."""
    for child in root.elements():
        if local_name(child.name) == "sheetData":
            continue
        for anchor in child.descendants("anchor"):
            parent = anchor.parent
            if parent is not None and local_name(parent.name) in ("controlPr", "objectPr"):
                _move_ends(anchor, record_placement(anchor), edit)


def _move_ends(anchor: Element, placement: Placement, edit: AxisEdit) -> Box | None:
    start, end = anchor.child("from"), anchor.child("to")
    first = _read(start, edit.is_row)
    if start is None or first is None:
        return None
    last = _read(end, edit.is_row)
    box = move_box(first, last, placement, edit)
    _write(start, box.first.corner, edit.is_row)
    if end is not None and box.last is not None:
        _write(end, box.last.corner, edit.is_row)
    return box


def _names(is_row: bool) -> tuple[str, str]:
    return ("row", "rowOff") if is_row else ("col", "colOff")


def _read(end: Element | None, is_row: bool) -> Corner | None:
    """One end of an anchor on one axis, or ``None`` when it names no cell
    there."""
    if end is None:
        return None
    index_name, offset_name = _names(is_row)
    index = end.child(index_name)
    if index is None:
        return None
    offset = end.child(offset_name)
    try:
        return Corner(int(index.text), 0 if offset is None or not offset.text else int(offset.text))
    except ValueError:
        return None


def _write(end: Element, corner: Corner, is_row: bool) -> None:
    index_name, offset_name = _names(is_row)
    index = end.require(index_name)
    if index.text != str(corner.index):
        index.set_text(str(corner.index))
    offset = end.child(offset_name)
    if offset is None:
        if not corner.offset:
            return
        prefix = index.name.rpartition(":")[0]
        name = f"{prefix}:{offset_name}" if prefix else offset_name
        offset = Element.create(name)
        end.insert_after(index, offset)
    if offset.text != str(corner.offset):
        offset.set_text(str(corner.offset))


def _move_transform(anchor: Element, box: Box, edit: AxisEdit) -> None:
    """Move a shape's own transform by as far as its anchor moved. An
    all-zero transform, which is how Excel writes a control's drawing twin,
    says nothing and stays so."""
    top = box.first.distance
    bottom = top if box.last is None else box.last.distance
    if not top and not bottom:
        return
    body = next((child for child in anchor.elements() if local_name(child.name) in _TRANSFORMED), None)
    properties = None if body is None else (body.child("spPr") or body.child("grpSpPr"))
    transform = None if properties is None else properties.child("xfrm")
    offset = None if transform is None else transform.child("off")
    extent = None if transform is None else transform.child("ext")
    if offset is None or extent is None:
        return
    position, size = ("y", "cy") if edit.is_row else ("x", "cx")
    try:
        values = [int(offset.get(name) or 0) for name in ("x", "y")] + [
            int(extent.get(name) or 0) for name in ("cx", "cy")
        ]
        was, length = int(offset.get(position) or 0), int(extent.get(size) or 0)
    except ValueError:
        return
    if not any(values):
        return
    offset.set(position, str(was + top))
    extent.set(size, str(max(length + bottom - top, 0)))


# ----------------------------------------------------------------------
# VML
# ----------------------------------------------------------------------


def move_vml(text: str, edit: AxisEdit) -> str:
    """Move every ``<x:Anchor>`` in a VML part: a note's box with its cell,
    anything else by its placement. The offsets are in pixels here."""
    return _CLIENT_DATA.sub(lambda match: _move_client_data(match.group(0), edit), text)


def _move_client_data(block: str, edit: AxisEdit) -> str:
    anchor = _VML_ANCHOR.search(block)
    if anchor is None:
        return block
    numbers = [int(number) for number in re.findall(r"-?\d+", anchor.group(2))]
    if len(numbers) < 8:
        return block
    slots = (2, 3, 6, 7) if edit.is_row else (0, 1, 4, 5)
    first = Corner(numbers[slots[0]], numbers[slots[1]] * EMU_PER_PIXEL)
    last = Corner(numbers[slots[2]], numbers[slots[3]] * EMU_PER_PIXEL)
    owner = _note_owner(block, edit)
    if owner is not None:
        step = with_cells(Corner(owner, 0), edit).corner.index - owner
        first = Corner(min(max(first.index + step, 0), edit.limit), first.offset)
        last = Corner(min(max(last.index + step, 0), edit.limit), last.offset)
    else:
        box = move_box(first, last, vml_placement(block), edit)
        assert box.last is not None
        first, last = box.first.corner, box.last.corner
    moved = list(numbers)
    for slot, value in zip(
        slots,
        (first.index, round(first.offset / EMU_PER_PIXEL), last.index, round(last.offset / EMU_PER_PIXEL)),
        strict=True,
    ):
        moved[slot] = value
    if moved == numbers:
        return block
    rewritten = _renumber(anchor.group(2), moved)
    return block[: anchor.start(2)] + rewritten + block[anchor.end(2) :]


def _note_owner(block: str, edit: AxisEdit) -> int | None:
    """The zero-based row or column of the cell a note annotates, or
    ``None`` for anything that is not a note."""
    if re.search(r'\bObjectType="Note"', block) is None:
        return None
    tag = "Row" if edit.is_row else "Column"
    found = re.search(rf"<x:{tag}>(\d+)</x:{tag}>", block)
    return None if found is None else int(found.group(1))


def _renumber(raw: str, numbers: list[int]) -> str:
    """An anchor's text with its numbers replaced in place, keeping the
    spacing Excel lays it out with."""
    values = iter(numbers)
    return re.sub(r"-?\d+", lambda match: str(next(values, match.group(0))), raw)


__all__ = [
    "EMU_PER_PIXEL",
    "AxisEdit",
    "Box",
    "Corner",
    "Moved",
    "Placement",
    "anchor_placement",
    "move_box",
    "move_drawing",
    "move_record_anchors",
    "move_vml",
    "record_placement",
    "swallowed_nodes",
    "swallows",
    "vml_placement",
    "walk",
    "with_cells",
]
