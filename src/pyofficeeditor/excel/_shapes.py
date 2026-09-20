"""Shapes on a sheet: where they are, what they say, what a click runs.

A shape lives in a drawing part the sheet points at, wrapped in an anchor
that says which cells it spans. Five things about that are measured against
Excel rather than read off the schema, and the first is why this module
needs a grid at all.

**A form control's transform is zero.** Excel writes ``<a:off x="0" y="0"/>``
and an extent of nothing for a button, and puts the real box only in the
cell anchor. Reading the transform gives every control a size of zero at the
top-left corner.

So the transform is used when it says anything, and the anchor only when it
does not. Measured against Excel's own Left, Top, Width and Height: the
transform matched exactly for all six ordinary shapes, and only the button
needed the anchor.

**The grid is a fallback, and not an exact one.** Turning an anchor into
points needs the sheet's column widths, and a width is stored in characters
of the standard font rather than in points. Excel converts using that font's
maximum digit width, which the file does not carry, so this uses seven
pixels per character plus five of padding, at three quarters of a point per
pixel. Rows are exact, because a height is already in points. For the
fixture's button that put the left edge a quarter of a point out, which is
the error to expect for a form control and for nothing else.

**A form control lives inside ``mc:AlternateContent``.** Excel wraps its
anchor in a ``mc:Choice`` requiring the ``x14`` namespace, so a reader that
only looks at the drawing's own children finds every shape except the
controls.

**The element says what kind of shape it is.** ``<xdr:sp>`` is an AutoShape
unless its ``cNvSpPr`` carries ``txBox="1"``, in which case it is a text box:
not a different element, a flag. ``<xdr:cxnSp>`` is a line, ``<xdr:grpSp>``
a group, ``<xdr:pic>`` a picture, ``<xdr:graphicFrame>`` a chart or a table.

**A drawing shape's macro is on the shape and a form control's is not.**
``<xdr:sp macro="Clicked">`` for the first; a control carries nothing there,
and the sheet's own ``<controls>`` element holds ``macro="[0]!Clicked"``,
which Excel reports back as ``Book.xlsm!Clicked``.

**An untouched shape is written back from the markup it came from**, so a
drawing nobody edited survives a save byte for byte.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

from pyofficeeditor._xml import Element

#: English Metric Units in one point. Office measures a shape in EMU in the
#: file and in points through the object model, and this is the whole of the
#: conversion: 914400 to the inch, 72 points to the inch.
EMU_PER_POINT = 12700

#: What Excel's default column comes to in points. 8.43 characters is 64
#: pixels at 96 dpi, and a point is three quarters of a pixel.
DEFAULT_COLUMN_POINTS = 48.0

#: The default row height Excel writes for its standard font.
DEFAULT_ROW_POINTS = 14.5

#: What a shape is, told apart the way the object model tells them apart
#: rather than by which element happens to hold it.
ShapeKind = Literal[
    "shape",
    "textBox",
    "line",
    "picture",
    "chart",
    "group",
    "formControl",
    "other",
]

#: What ``Shape.Type`` answers for each kind, so a reader can be held to the
#: number Excel itself reported.
MSO_TYPE: dict[str, int] = {
    "shape": 1,
    "chart": 3,
    "group": 6,
    "formControl": 8,
    "line": 9,
    "picture": 13,
    "textBox": 17,
}

#: The element behind each kind. A text box is the exception: it is an
#: ``sp`` like an AutoShape and told apart by a flag, not by its name.
_ELEMENT_KINDS: dict[str, ShapeKind] = {
    "sp": "shape",
    "cxnSp": "line",
    "grpSp": "group",
    "pic": "picture",
    "graphicFrame": "chart",
}

#: The anchors a drawing is made of.
ANCHORS = ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor")


def points(value: str | int | float | None) -> float:
    """EMU as the points the object model reports."""
    if value is None:
        return 0.0
    try:
        return int(value) / EMU_PER_POINT
    except (TypeError, ValueError):
        return 0.0


def emu(value: float) -> int:
    """Points as the EMU the file holds, rounded as Office rounds."""
    return round(value * EMU_PER_POINT)


def characters_to_points(width: float) -> float:
    """A stored column width, in points.

    Stored in characters of the standard font. Excel turns that into whole
    pixels, seven to the character plus five for the padding, and a point is
    three quarters of a pixel. Rounding to whole pixels first is what makes
    the default 8.43 come out at exactly 48 rather than 47.9.
    """
    return round(width * 7 + 5) * 0.75


@dataclass
class SheetGrid:
    """How far across and down a cell sits, in points.

    A shape carrying its own transform does not need this. A form control
    does: Excel writes its transform as zero and leaves the anchor to say
    where it is.

    Both maps are keyed by zero-based index, because that is how a drawing
    anchor numbers its rows and columns.
    """

    column_widths: dict[int, float] = field(default_factory=lambda: {})
    row_heights: dict[int, float] = field(default_factory=lambda: {})
    default_column: float = DEFAULT_COLUMN_POINTS
    default_row: float = DEFAULT_ROW_POINTS

    def x(self, column: int, offset: str | int | None = 0) -> float:
        """Where a column starts, plus an offset into it, in points."""
        before = sum(
            self.column_widths.get(index, self.default_column) for index in range(column)
        )
        return before + points(offset)

    def y(self, row: int, offset: str | int | None = 0) -> float:
        """Where a row starts, plus an offset into it, in points."""
        before = sum(self.row_heights.get(index, self.default_row) for index in range(row))
        return before + points(offset)

    @classmethod
    def of(cls, sheet_root: Element) -> SheetGrid:
        """The grid a worksheet's own XML describes."""
        grid = cls()
        head = sheet_root.child("sheetFormatPr")
        if head is not None:
            grid.default_row = _as_float(head.get("defaultRowHeight"), DEFAULT_ROW_POINTS)
            width = head.get("defaultColWidth")
            if width is not None:
                grid.default_column = characters_to_points(_as_float(width, 8.43))

        columns = sheet_root.child("cols")
        if columns is not None:
            for entry in columns.children_named("col"):
                raw = entry.get("width")
                if raw is None:
                    continue
                first = int(_as_float(entry.get("min"), 1)) - 1
                last = int(_as_float(entry.get("max"), 1)) - 1
                measure = characters_to_points(_as_float(raw, 8.43))
                for index in range(first, min(last, first + 16383) + 1):
                    grid.column_widths[index] = measure

        data = sheet_root.child("sheetData")
        if data is not None:
            for row in data.children_named("row"):
                height = row.get("ht")
                if height is None:
                    continue
                number = int(_as_float(row.get("r"), 1)) - 1
                grid.row_heights[number] = _as_float(height, DEFAULT_ROW_POINTS)
        return grid


@dataclass(frozen=True)
class Shape:
    """One shape on a sheet.

    ``left``, ``top``, ``width`` and ``height`` are in points, the unit the
    object model uses, worked out from the cell anchor rather than from the
    transform: see the module docstring for why.
    """

    name: str = ""
    kind: ShapeKind = "shape"
    #: The preset geometry, such as ``rect``, ``roundRect`` or ``ellipse``.
    geometry: str = ""
    left: float = 0.0
    top: float = 0.0
    width: float = 0.0
    height: float = 0.0
    text: str = ""
    #: The procedure a click runs. A form control's comes from the sheet's
    #: own ``<controls>`` rather than from the drawing.
    macro: str = ""
    #: The id the drawing gives it, unique within that part.
    shape_id: int = 0
    #: A group's members, in the order the file holds them.
    children: tuple[Shape, ...] = ()

    @property
    def mso_type(self) -> int:
        """What ``Shape.Type`` answers for this kind."""
        return MSO_TYPE.get(self.kind, 1)

    def __repr__(self) -> str:
        where = f"{self.left:g},{self.top:g} {self.width:g}x{self.height:g}"
        macro = f" macro={self.macro}" if self.macro else ""
        return f"<Shape {self.name!r} {self.kind} {where}{macro}>"


def read_drawing(root: Element, grid: SheetGrid) -> list[Shape]:
    """Every shape in a drawing part, in the order it holds them.

    Anchors wrapped in ``mc:AlternateContent`` are picked up too, which is
    where Excel puts a form control.
    """
    found: list[Shape] = []
    for anchor in _anchors(root):
        shape = _from_anchor(anchor, grid)
        if shape is not None:
            found.append(shape)
    return found


def _anchors(root: Element) -> list[Element]:
    """Every anchor, including the ones inside ``mc:AlternateContent``.

    A form control's anchor sits in a ``mc:Choice`` requiring ``x14``, so
    looking only at the drawing's own children misses every control.
    """
    found: list[Element] = []
    for child in root.children:
        if not isinstance(child, Element):
            continue
        name = _local(child.name)
        if name in ANCHORS:
            found.append(child)
        elif name == "AlternateContent":
            chosen = _find(child, "Choice") or _find(child, "Fallback")
            if chosen is not None:
                found.extend(_anchors(chosen))
    return found


def _from_anchor(anchor: Element, grid: SheetGrid) -> Shape | None:
    body = _shape_element(anchor)
    if body is None:
        return None
    box = _transform_box(body)
    if box is None:
        box = _anchor_box(anchor, grid)
    left, top, width, height = box
    return _shape_of(body, grid, left=left, top=top, width=width, height=height)


def _transform_box(body: Element) -> tuple[float, float, float, float] | None:
    """The shape's own transform, or ``None`` when it says nothing.

    A form control's is all zeros, which is Excel declining to answer
    rather than a shape of no size at the origin.
    """
    shape_properties = _find(body, "spPr") or _find(body, "grpSpPr")
    if shape_properties is None:
        return None
    transform = _find(shape_properties, "xfrm")
    if transform is None:
        return None
    offset = _find(transform, "off")
    extent = _find(transform, "ext")
    if offset is None or extent is None:
        return None
    box = (
        points(offset.get("x")),
        points(offset.get("y")),
        points(extent.get("cx")),
        points(extent.get("cy")),
    )
    return None if box == (0.0, 0.0, 0.0, 0.0) else box


def _shape_element(parent: Element) -> Element | None:
    for child in parent.children:
        if isinstance(child, Element) and _local(child.name) in _ELEMENT_KINDS:
            return child
    return None


def _shape_of(
    body: Element, grid: SheetGrid, *, left: float, top: float, width: float, height: float
) -> Shape:
    kind = _ELEMENT_KINDS[_local(body.name)]
    properties = _find(body, "nvSpPr") or _find(body, "nvCxnSpPr") or _find(body, "nvGrpSpPr")
    properties = properties or _find(body, "nvPicPr") or _find(body, "nvGraphicFramePr")
    naming = None if properties is None else _find(properties, "cNvPr")

    if kind == "shape" and _is_text_box(properties):
        kind = "textBox"

    children: tuple[Shape, ...] = ()
    if kind == "group":
        members: list[Shape] = []
        for child in body.children:
            if isinstance(child, Element) and _local(child.name) in _ELEMENT_KINDS:
                members.append(
                    _shape_of(child, grid, left=left, top=top, width=width, height=height)
                )
        children = tuple(members)

    return Shape(
        name="" if naming is None else (naming.get("name") or ""),
        kind=kind,
        geometry=_geometry(body),
        left=left,
        top=top,
        width=width,
        height=height,
        text=_text_of(body),
        macro=body.get("macro") or "",
        shape_id=int(_as_float(None if naming is None else naming.get("id"), 0)),
        children=children,
    )


def _anchor_box(anchor: Element, grid: SheetGrid) -> tuple[float, float, float, float]:
    """The shape's box in points, from the anchor rather than the transform.

    A two-cell anchor gives both corners, a one-cell anchor a corner and an
    extent, and an absolute anchor a position in EMU with no cells at all.
    """
    kind = _local(anchor.name)
    if kind == "absoluteAnchor":
        position = _find(anchor, "pos")
        extent = _find(anchor, "ext")
        return (
            points(None if position is None else position.get("x")),
            points(None if position is None else position.get("y")),
            points(None if extent is None else extent.get("cx")),
            points(None if extent is None else extent.get("cy")),
        )

    start = _find(anchor, "from")
    left = top = 0.0
    if start is not None:
        left = grid.x(_index(start, "col"), _offset(start, "colOff"))
        top = grid.y(_index(start, "row"), _offset(start, "rowOff"))

    if kind == "oneCellAnchor":
        extent = _find(anchor, "ext")
        return (
            left,
            top,
            points(None if extent is None else extent.get("cx")),
            points(None if extent is None else extent.get("cy")),
        )

    end = _find(anchor, "to")
    if end is None:
        return (left, top, 0.0, 0.0)
    right = grid.x(_index(end, "col"), _offset(end, "colOff"))
    bottom = grid.y(_index(end, "row"), _offset(end, "rowOff"))
    return (left, top, max(right - left, 0.0), max(bottom - top, 0.0))


def _is_text_box(properties: Element | None) -> bool:
    """A text box is an ``sp`` with a flag, not an element of its own."""
    if properties is None:
        return False
    shape_properties = _find(properties, "cNvSpPr")
    return shape_properties is not None and shape_properties.get("txBox") in ("1", "true")


def _geometry(body: Element) -> str:
    shape_properties = _find(body, "spPr") or _find(body, "grpSpPr")
    if shape_properties is None:
        return ""
    preset = _find(shape_properties, "prstGeom")
    return "" if preset is None else (preset.get("prst") or "")


def _text_of(body: Element) -> str:
    """The shape's text, joining its runs and honouring line breaks.

    Excel stores a break inside a run as a ``<a:br/>`` between two runs and
    reports the pair as one character, so the two are joined with a newline.
    """
    text_body = _find(body, "txBody")
    if text_body is None:
        return ""
    lines: list[str] = []
    for paragraph in text_body.children:
        if not isinstance(paragraph, Element) or _local(paragraph.name) != "p":
            continue
        piece = ""
        for node in paragraph.children:
            if not isinstance(node, Element):
                continue
            name = _local(node.name)
            if name == "r":
                run = _find(node, "t")
                if run is not None:
                    # Excel writes a break inside a run as CRLF and reports
                    # it back as one character, so it is normalised here.
                    piece += (run.text or "").replace("\r\n", "\n")
            elif name == "br":
                piece += "\n"
        lines.append(piece)
    return "\n".join(lines)


def _find(parent: Element, name: str) -> Element | None:
    """A child by local name, whatever prefix it carries."""
    for child in parent.children:
        if isinstance(child, Element) and _local(child.name) == name:
            return child
    return None


def _index(end: Element, axis: str) -> int:
    node = _find(end, axis)
    if node is None or not node.text:
        return 0
    try:
        return int(node.text)
    except ValueError:
        return 0


def _offset(end: Element, axis: str) -> int:
    node = _find(end, axis)
    if node is None or not node.text:
        return 0
    try:
        return int(node.text)
    except ValueError:
        return 0


def _local(name: str) -> str:
    _, _, local = name.rpartition(":")
    return local


def _as_float(raw: str | None, fallback: float) -> float:
    if raw is None:
        return fallback
    try:
        return float(raw)
    except ValueError:
        return fallback


__all__ = [
    "ANCHORS",
    "DEFAULT_COLUMN_POINTS",
    "DEFAULT_ROW_POINTS",
    "EMU_PER_POINT",
    "MSO_TYPE",
    "Shape",
    "ShapeKind",
    "SheetGrid",
    "characters_to_points",
    "emu",
    "points",
    "read_drawing",
]
