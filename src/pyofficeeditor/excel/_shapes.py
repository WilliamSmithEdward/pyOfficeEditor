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

import re
from dataclasses import dataclass, field
from typing import Literal

from pyofficeeditor._xml import Element, escape_attribute, escape_text

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


#: The preset geometry Excel writes for each ``MsoAutoShapeType``, measured
#: by making one shape of each and reading the drawing back.
#:
#: Worth measuring rather than transcribing, because every plausible wrong
#: answer here is a real preset name belonging to some other shape. A table
#: that pairs 11 with ``cross``, 12 with ``star5``, 16 with ``can`` and 17
#: with ``cube`` looks right and is wrong on all four: Excel writes ``plus``,
#: ``pentagon``, ``foldedCorner`` and ``smileyFace``. Nothing complains. The
#: file is valid and the wrong shape appears.
#:
#: The five-pointed star is 92, not 12. The numbers are not contiguous:
#: ``MsoAutoShapeType`` leaves gaps, and only the ones Excel actually made
#: when asked are here.
PRESET_GEOMETRY: dict[int, str] = {
    1: "rect",  # msoShapeRectangle
    2: "parallelogram",
    3: "trapezoid",
    4: "diamond",
    5: "roundRect",  # msoShapeRoundedRectangle
    6: "octagon",
    7: "triangle",  # msoShapeIsoscelesTriangle
    8: "rtTriangle",
    9: "ellipse",  # msoShapeOval
    10: "hexagon",
    11: "plus",  # msoShapeCross
    12: "pentagon",  # msoShapeRegularPentagon
    13: "can",
    14: "cube",
    15: "bevel",
    16: "foldedCorner",
    17: "smileyFace",
    18: "donut",
    19: "noSmoking",
    20: "blockArc",
    21: "heart",
    22: "lightningBolt",
    23: "sun",
    24: "moon",
    25: "arc",
    26: "bracketPair",
    27: "bracePair",
    28: "plaque",
    92: "star5",  # msoShape5pointStar
    93: "star8",
    94: "star16",
    95: "star24",
    96: "star32",
}

#: The same the other way round, for reporting what a shape read from a
#: file would answer for ``AutoShapeType``.
AUTO_SHAPE_TYPES: dict[str, int] = {
    preset: number for number, preset in PRESET_GEOMETRY.items()
}


#: The anchors a drawing is made of.
ANCHORS = ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor")

#: The last column and row an anchor can name, zero-based. A drawing
#: counts from 0 where a cell reference counts from 1, so these are one
#: less than the sheet's 16384 by 1048576.
MAX_ANCHOR_COLUMN = 16383
MAX_ANCHOR_ROW = 1048575


#: What ``ControlFormat.Value`` answers for a check box or an option
#: button. Measured, because the off state is not zero: Excel answers
#: -4146, which is ``xlOff``, and a reader assuming 0 is wrong for every
#: unticked box.
XL_ON = 1
XL_OFF = -4146
XL_MIXED = 2

#: What a control's own part calls each kind, in the ``objectType``
#: attribute of ``formControlPr``. Excel's own spelling is kept rather
#: than translated, so what is read is what can be written back.
#:
#: ``Radio`` is an option button, ``Drop`` a drop down, ``GBox`` a group
#: box and ``Scroll`` a scroll bar.
CONTROL_KINDS = (
    "Button",
    "CheckBox",
    "Drop",
    "List",
    "Radio",
    "Spin",
    "Scroll",
    "GBox",
    "Label",
)

#: Which attribute holds a control's current value, by kind.
#:
#: This is the whole reason the value is read per kind rather than off one
#: attribute. A check box writes ``checked="Checked"`` and no ``val`` at
#: all; a drop down writes ``sel``, the 1-based selection, and leaves
#: ``val`` at 0; only a spinner and a scroll bar write ``val``. Reading
#: ``val`` alone answers 0 for a ticked box and 0 for a drop down with the
#: second item chosen. Both are plausible and both are wrong.
_VALUE_ATTRIBUTE: dict[str, str] = {
    "CheckBox": "checked",
    "Radio": "checked",
    "Drop": "sel",
    "List": "sel",
    "Spin": "val",
    "Scroll": "val",
}

#: What the ``checked`` attribute spells, and the number Excel answers for
#: it. Absent means off, which is how Excel writes an unticked box.
_CHECKED_STATES: dict[str, int] = {
    "Checked": XL_ON,
    "Mixed": XL_MIXED,
    "Unchecked": XL_OFF,
}

#: The bounds Excel gives a spinner and a scroll bar nobody configured, as
#: (minimum, maximum, increment, page). Measured by making one of each and
#: reading the part back, because the obvious guess of zero is wrong and
#: wrong in a way that hides: a spinner whose maximum is 0 is pinned at 0
#: whatever value it was given, and the file is perfectly valid.
RANGE_DEFAULTS: dict[str, tuple[int, int, int, int]] = {
    "Spin": (0, 30000, 1, 10),
    "Scroll": (0, 100, 1, 10),
}


@dataclass(frozen=True)
class FormControl:
    """What a Forms-toolbar control is wired to.

    A control is two halves. The drawing holds its box and its name, and
    everything about what it *does* lives in a part of its own that the
    sheet points at by relationship id. This is that part.

    ``value`` is the number Excel's own object model answers, so a reader
    can be held to it: 1 for a ticked box, -4146 for an unticked one, 2
    for the third state, and the 1-based selection for a list.
    """

    #: Excel's own ``objectType``: Button, CheckBox, Drop, List, Radio,
    #: Spin, Scroll, GBox or Label.
    kind: str = ""
    #: The cell the control writes to, such as ``$D$6``, as this control's
    #: own part holds it.
    #:
    #: An option button group stores the link once, on the button marked
    #: :attr:`first_button`, and the rest of the group carries nothing.
    #: Excel's object model resolves that and answers the group's cell for
    #: every member; this reports what the part says, which is empty for
    #: all but the first. The two disagree on purpose. Resolving it would
    #: mean deciding what makes a group, and that is not measured here:
    #: ``firstButton`` marks a start, but a group box holding the buttons
    #: is a second way to group them, and guessing between them would put
    #: an invented answer where a blank one is honest.
    linked_cell: str = ""
    #: The range a drop down or list box takes its items from.
    list_range: str = ""
    #: What the object model answers for ``Value``. Zero for a kind that
    #: has none, such as a button or a label.
    #:
    #: Measured: a control with a :attr:`linked_cell` takes its state from
    #: that cell on load and this is overwritten. A tick box stored ticked
    #: and linked to an empty cell opens unticked, and one stored unticked
    #: and linked to a cell holding TRUE opens ticked. So setting both is
    #: setting the cell, not the box.
    value: int = 0
    #: A spinner's or scroll bar's bounds and steps, and ``None`` where
    #: the part says nothing, which means Excel's own default for the
    #: kind: see :data:`RANGE_DEFAULTS`.
    minimum: int | None = None
    maximum: int | None = None
    increment: int | None = None
    page: int | None = None
    #: Whether this option button starts a group, which is where the
    #: group's :attr:`linked_cell` is kept. False for every other kind.
    first_button: bool = False
    #: The part this was read from, so an edit lands back in it.
    part_name: str = ""
    #: The relationship the sheet names that part by.
    relationship: str = ""

    @property
    def checked(self) -> bool | None:
        """A tick box or option button's state, or ``None`` for neither.

        ``None`` covers both the third state and every kind that has no
        tick at all, which is why the kind is checked rather than the
        value alone.
        """
        if self.kind not in ("CheckBox", "Radio"):
            return None
        if self.value == XL_ON:
            return True
        return False if self.value == XL_OFF else None


def read_control(root: Element, *, part_name: str = "", relationship: str = "") -> FormControl:
    """One control's own part, as the object model would report it."""
    kind = root.get("objectType") or ""
    attribute = _VALUE_ATTRIBUTE.get(kind)
    value = 0
    if attribute == "checked":
        value = _CHECKED_STATES.get(root.get("checked") or "", XL_OFF)
    elif attribute is not None:
        value = int(_as_float(root.get(attribute), 0.0))
    return FormControl(
        kind=kind,
        linked_cell=root.get("fmlaLink") or "",
        list_range=root.get("fmlaRange") or "",
        value=value,
        minimum=_as_int(root.get("min")),
        maximum=_as_int(root.get("max")),
        increment=_as_int(root.get("inc")),
        page=_as_int(root.get("page")),
        first_button=(root.get("firstButton") in ("1", "true")),
        part_name=part_name,
        relationship=relationship,
    )


def _as_int(raw: str | None) -> int | None:
    """An attribute as a whole number, or ``None`` when it says nothing."""
    if raw is None:
        return None
    try:
        return int(float(raw))
    except ValueError:
        return None


def control_range(control: FormControl) -> tuple[int, int, int, int]:
    """A control's bounds, filling in Excel's defaults for the kind."""
    fallback = RANGE_DEFAULTS.get(control.kind, (0, 0, 1, 10))
    chosen = (control.minimum, control.maximum, control.increment, control.page)
    return tuple(  # type: ignore[return-value]
        given if given is not None else default
        for given, default in zip(chosen, fallback, strict=True)
    )


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

    def column_at(self, across: float) -> tuple[int, int]:
        """Which column a point falls in, and how far into it, in EMU.

        The inverse of :meth:`x`, and needed to write an anchor rather than
        read one. Excel will not accept an offset larger than the cell
        holding it: it clamps one to the column's width and the shape lands
        somewhere other than where it was put. So the corner is worked out
        against the sheet's own widths instead of being left to Excel.
        """
        column, seen = 0, 0.0
        while column <= MAX_ANCHOR_COLUMN:
            width = self.column_widths.get(column, self.default_column)
            if width <= 0:
                column += 1
                continue
            if seen + width > across:
                return column, emu(max(across - seen, 0.0))
            seen += width
            column += 1
        return MAX_ANCHOR_COLUMN, 0

    def row_at(self, down: float) -> tuple[int, int]:
        """Which row a point falls in, and how far into it, in EMU."""
        row, seen = 0, 0.0
        while row <= MAX_ANCHOR_ROW:
            height = self.row_heights.get(row, self.default_row)
            if height <= 0:
                row += 1
                continue
            if seen + height > down:
                return row, emu(max(down - seen, 0.0))
            seen += height
            row += 1
        return MAX_ANCHOR_ROW, 0

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
    #: What a Forms control is wired to, and ``None`` for every other
    #: kind. The drawing says nothing about this: it comes from the part
    #: the sheet's own ``<control>`` points at.
    control: FormControl | None = None

    @property
    def mso_type(self) -> int:
        """What ``Shape.Type`` answers for this kind."""
        return MSO_TYPE.get(self.kind, 1)

    @property
    def auto_shape_type(self) -> int | None:
        """What ``Shape.AutoShapeType`` answers, or ``None`` for a geometry
        this does not name.

        A text box answers 1 as well as an AutoShape does, which is what
        Excel reports, because both are drawn as a rectangle.
        """
        return AUTO_SHAPE_TYPES.get(self.geometry)

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
                    # it back as one character, which is how it reads here.
                    piece += run.text or ""
            elif name == "br":
                piece += "\n"
        lines.append(piece)
    return "\n".join(lines)


# --------------------------------------------------------------------------
# Writing: the markup Excel expects for a shape this library adds
# --------------------------------------------------------------------------

#: A drawing part with nothing in it, for a sheet that has never had one.
EMPTY_DRAWING = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
)

#: The VML a sheet needs before it can hold a control at all: the id map,
#: and the shape type every button is drawn from. Taken from a file Excel
#: wrote, CRLFs included, because VML is not parsed here and what goes in
#: is what comes out.
EMPTY_VML = (
    '<xml xmlns:v="urn:schemas-microsoft-com:vml"\r\n'
    ' xmlns:o="urn:schemas-microsoft-com:office:office"\r\n'
    ' xmlns:x="urn:schemas-microsoft-com:office:excel">\r\n'
    ' <o:shapelayout v:ext="edit">\r\n'
    '  <o:idmap v:ext="edit" data="1"/>\r\n'
    ' </o:shapelayout><v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201"\r\n'
    '  path="m,l,21600r21600,l21600,xe">\r\n'
    '  <v:stroke joinstyle="miter"/>\r\n'
    '  <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>\r\n'
    '  <o:lock v:ext="edit" shapetype="t"/>\r\n'
    " </v:shapetype></xml>\r\n"
)

#: Where a form control's shape id starts. Excel numbers controls from
#: 1025 and ordinary drawing shapes from 2, and keeps the two apart.
FIRST_CONTROL_ID = 1025

#: The style Excel gives a new AutoShape, taken from one it wrote. A text
#: box gets none, which is also what Excel does.
_NEW_SHAPE_STYLE = (
    "<xdr:style>"
    '<a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef>'
    '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>'
    '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
    '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>'
    "</xdr:style>"
)

#: And the one it gives a line, which is a different four references: no
#: shade on the outline, no fill, an effect rather than none, and the
#: text colour from ``tx1`` rather than ``lt1``.
_NEW_LINE_STYLE = (
    "<xdr:style>"
    '<a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>'
    '<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>'
    '<a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef>'
    '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef>'
    "</xdr:style>"
)

#: Which style block each drawing kind gets. A text box gets none, which
#: is what Excel does: its text is the shape.
_SHAPE_STYLES: dict[str, str] = {
    "shape": _NEW_SHAPE_STYLE,
    "textBox": "",
    "line": _NEW_LINE_STYLE,
}

#: The preset each kind is drawn with when the caller names none. A
#: connector is not a rectangle: Excel writes ``prst="line"`` for one, and
#: a ``cxnSp`` carrying ``prst="rect"`` makes it refuse the workbook.
DEFAULT_GEOMETRY: dict[str, str] = {
    "shape": "rect",
    "textBox": "rect",
    "line": "line",
}


def corner_markup(
    which: str, grid: SheetGrid, across: float, down: float, *, wrapper: str = "xdr:"
) -> str:
    """One end of an anchor, as the cell it lands in plus an offset.

    ``wrapper`` is the prefix on the ``<from>``/``<to>`` element itself,
    and it is not always the same as the one on its children. A drawing
    writes ``<xdr:from><xdr:col>``; the sheet's own ``<controlPr>`` anchor
    writes ``<from><xdr:col>``, because the anchor is in the spreadsheetml
    namespace and only the coordinates come from the drawing one.

    Stripping the prefix off all five is the obvious simplification and
    makes Excel refuse the workbook rather than repair it, since ``<col>``
    then reads as a spreadsheetml element the schema has no place for.
    """
    column, column_offset = grid.column_at(across)
    row, row_offset = grid.row_at(down)
    return (
        f"<{wrapper}{which}>"
        f"<xdr:col>{column}</xdr:col><xdr:colOff>{column_offset}</xdr:colOff>"
        f"<xdr:row>{row}</xdr:row><xdr:rowOff>{row_offset}</xdr:rowOff>"
        f"</{wrapper}{which}>"
    )


def text_body(text: str, tag: str = "xdr:txBody") -> str:
    """A shape's text, as the paragraphs and runs DrawingML wants.

    An empty shape still needs the element: Excel writes one with a single
    empty paragraph rather than leaving it out.
    """
    empty = '<a:p><a:endParaRPr lang="en-US"/></a:p>'
    paragraphs = "".join(
        f'<a:p><a:r><a:rPr lang="en-US"/><a:t>{escape_text(line)}</a:t></a:r></a:p>'
        if line
        else empty
        for line in (text.split("\n") if text else [""])
    )
    return f'<{tag}><a:bodyPr vertOverflow="clip"/><a:lstStyle/>{paragraphs}</{tag}>'


def new_anchor(shape: Shape, grid: SheetGrid) -> str:
    """A drawing shape's anchor, holding both the cells and the transform.

    Excel reads the cells, and clamps an offset to the cell that holds it,
    so each corner is worked out against the sheet's own columns and rows
    rather than left to be clamped. The transform is written as well
    because that is what the object model reports back.
    """
    line = shape.kind == "line"
    element = "xdr:cxnSp" if line else "xdr:sp"
    identity = (
        f'<xdr:cNvPr id="{shape.shape_id}" name="{escape_attribute(shape.name)}"/>'
    )
    # Built outside the f-strings: a backslash inside one is 3.12 and up,
    # and this library runs on 3.10.
    box_flag = ' txBox="1"' if shape.kind == "textBox" else ""
    textlink = "" if line else ' textlink=""'
    # A line gets its own four references, a text box none at all, and an
    # AutoShape the accent-coloured set Excel gives a fresh one.
    style = _NEW_LINE_STYLE if line else _SHAPE_STYLES.get(shape.kind, _NEW_SHAPE_STYLE)
    contents = "" if line else text_body(shape.text)
    geometry = escape_attribute(
        shape.geometry or DEFAULT_GEOMETRY.get(shape.kind, "rect")
    )
    properties = (
        f"<xdr:nvCxnSpPr>{identity}<xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>"
        if line
        else f"<xdr:nvSpPr>{identity}<xdr:cNvSpPr{box_flag}/></xdr:nvSpPr>"
    )
    body = (
        f'<{element} macro="{escape_attribute(shape.macro)}"{textlink}>'
        f"{properties}"
        "<xdr:spPr><a:xfrm>"
        f'<a:off x="{emu(shape.left)}" y="{emu(shape.top)}"/>'
        f'<a:ext cx="{emu(shape.width)}" cy="{emu(shape.height)}"/>'
        f'</a:xfrm><a:prstGeom prst="{geometry}"><a:avLst/></a:prstGeom></xdr:spPr>'
        f"{style}{contents}"
        f"</{element}>"
    )
    return (
        "<xdr:twoCellAnchor>"
        f"{corner_markup('from', grid, shape.left, shape.top)}"
        f"{corner_markup('to', grid, shape.left + shape.width, shape.top + shape.height)}"
        f"{body}<xdr:clientData/></xdr:twoCellAnchor>"
    )


def control_drawing(shape: Shape, grid: SheetGrid) -> str:
    """The drawing half of a form control, as Excel writes it.

    Wrapped in an ``mc:AlternateContent``, with the shape hidden and its
    transform empty: where a control is comes from the anchor, which is
    the measured behaviour the reader already relies on.
    """
    return (
        '<mc:AlternateContent xmlns:mc='
        '"http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"'
        ' Requires="a14">'
        '<xdr:twoCellAnchor editAs="oneCell">'
        f"{corner_markup('from', grid, shape.left, shape.top)}"
        f"{corner_markup('to', grid, shape.left + shape.width, shape.top + shape.height)}"
        '<xdr:sp macro="" textlink="">'
        "<xdr:nvSpPr>"
        f'<xdr:cNvPr id="{shape.shape_id}" name="{escape_attribute(shape.name)}" hidden="1">'
        "<a:extLst>"
        '<a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}">'
        f'<a14:compatExt spid="_x0000_s{shape.shape_id}"/>'
        "</a:ext>"
        "</a:extLst>"
        "</xdr:cNvPr>"
        "<xdr:cNvSpPr/>"
        "</xdr:nvSpPr>"
        '<xdr:spPr bwMode="auto">'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/><a:ln w="9525"><a:miter lim="800000"/><a:headEnd/><a:tailEnd/></a:ln>'
        "</xdr:spPr>"
        "</xdr:sp>"
        '<xdr:clientData fPrintsWithSheet="0"/>'
        "</xdr:twoCellAnchor>"
        "</mc:Choice>"
        "<mc:Fallback/>"
        "</mc:AlternateContent>"
    )


def control_entry(shape: Shape, relationship: str, grid: SheetGrid) -> str:
    """The sheet's own record of a control: what it is and what it runs.

    This is the copy Excel reads for the macro; see
    :func:`set_vml_macro` for the one it does not.
    """
    macro = f' macro="{escape_attribute(shape.macro)}"' if shape.macro else ""
    corners = corner_markup(
        "from", grid, shape.left, shape.top, wrapper=""
    ) + corner_markup(
        "to", grid, shape.left + shape.width, shape.top + shape.height, wrapper=""
    )
    return (
        '<mc:AlternateContent xmlns:mc='
        '"http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice Requires="x14">'
        f'<control shapeId="{shape.shape_id}" r:id="{relationship}"'
        f' name="{escape_attribute(shape.name)}">'
        f'<controlPr defaultSize="0" autoFill="0" autoPict="0"{macro}>'
        f'<anchor moveWithCells="1" sizeWithCells="1">{corners}</anchor>'
        "</controlPr>"
        "</control>"
        "</mc:Choice>"
        "</mc:AlternateContent>"
    )


def check_control_kind(kind: str) -> str:
    """The kind, or raise before anything has been written.

    Called first rather than where the kind is used, because building a
    control writes four things and the last of them is the one that knows
    the kind is wrong. Discovering it there leaves a part, a relationship
    and an anchor behind for a control that was never made.

    The names are Excel's own ``objectType``, which is also what the
    reader reports, so a control read from a file can be written straight
    back. Keying this on anything else is how a Radio silently becomes a
    Button.
    """
    if kind not in _VML_STYLES:
        raise ValueError(
            f"{kind!r} is not a form control this can write. Excel's own kinds "
            f"are: {', '.join(sorted(_VML_STYLES))}."
        )
    return kind


def control_properties(control: FormControl) -> str:
    """A control's own part: which control it is and what it is wired to.

    The value is written back the way the kind spells it, which is the
    same three-way split the reader has to undo: ``checked`` for a tick,
    ``sel`` for a list, ``val`` for a spinner.
    """
    kind = control.kind or "Button"
    attributes = [f'objectType="{escape_attribute(kind)}"']
    if kind == "Drop":
        attributes.append('dropStyle="combo" dropLines="8"')
    if control.first_button:
        attributes.append('firstButton="1"')

    spelling = _VALUE_ATTRIBUTE.get(kind)
    if spelling == "checked" and control.value in (XL_ON, XL_MIXED):
        attributes.append(f'checked="{"Checked" if control.value == XL_ON else "Mixed"}"')
    elif spelling in ("sel", "val") and control.value:
        attributes.append(f'{spelling}="{control.value}"')

    if kind in RANGE_DEFAULTS:
        # A spinner or scroll bar without a maximum is pinned at zero, so
        # the kind's own default is written rather than left out.
        minimum, maximum, increment, page = control_range(control)
        if minimum:
            attributes.append(f'min="{minimum}"')
        attributes.append(f'max="{maximum}"')
        if increment != 1:
            attributes.append(f'inc="{increment}"')
        attributes.append(f'page="{page}"')
        attributes.append('dx="31"')

    if control.linked_cell:
        attributes.append(f'fmlaLink="{escape_attribute(control.linked_cell)}"')
    if control.list_range:
        attributes.append(f'fmlaRange="{escape_attribute(control.list_range)}"')
    attributes.append('lockText="1"')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"'
        f" {' '.join(attributes)}/>"
    )


def vml_anchor(shape: Shape, grid: SheetGrid) -> str:
    """The eight numbers a control's VML places it by.

    Column, offset, row, offset for each corner, zero-based, and the
    offsets are in half-points rather than EMU: measured against a file
    Excel wrote, where a button 12 points into its column says 24.
    """
    left_column, left_offset = grid.column_at(shape.left)
    top_row, top_offset = grid.row_at(shape.top)
    right_column, right_offset = grid.column_at(shape.left + shape.width)
    bottom_row, bottom_offset = grid.row_at(shape.top + shape.height)
    numbers = (
        left_column,
        _half_points(left_offset),
        top_row,
        _half_points(top_offset),
        right_column,
        _half_points(right_offset),
        bottom_row,
        _half_points(bottom_offset),
    )
    return ", ".join(str(one) for one in numbers)


def _half_points(offset: int) -> int:
    """An EMU offset as the half-points a VML anchor counts in."""
    return round(points(offset) * 2)


@dataclass(frozen=True)
class _VmlStyle:
    """How one kind of control is drawn, measured from a file Excel wrote.

    Every kind uses the same ``_x0000_t201`` shape type, and differs in
    its ``v:shape`` attributes, the fill and lock block inside it, and
    which ``x:ClientData`` elements it carries in which order.
    """

    #: The ``ObjectType`` the VML spells, which is not always the
    #: ``objectType`` the control part spells. A tick box is ``CheckBox``
    #: in the part and ``Checkbox`` in the VML, and writing the part's
    #: spelling into the VML makes Excel refuse the workbook.
    object_type: str
    attributes: str
    body: str
    #: The ClientData children, in the order Excel writes them. The
    #: schema is a sequence, so the order is not cosmetic.
    slots: tuple[str, ...]
    align: str = "left"


_FILL_WINDOW = '<v:fill color2="window [65]"/>'
_LOCK = '<o:lock v:ext="edit" rotation="t"/>'
_LOCK_TEXT = '<o:lock v:ext="edit" rotation="t" text="t"/>'
_TICK_SLOTS = (
    "SizeWithCells", "Anchor", "AutoFill", "AutoLine", "TextVAlign",
    "Checked", "FmlaLink", "NoThreeD",
)
_LIST_SLOTS = (
    "FmlaLink", "Val", "Min", "Max", "Inc", "Page", "Dx", "FmlaRange",
    "Sel", "NoThreeD2", "SelType", "LCT",
)

#: What each kind looks like in the VML, taken from ``controls.xlsm``,
#: which Excel authored with one of each.
_VML_STYLES: dict[str, _VmlStyle] = {
    "Button": _VmlStyle(
        object_type="Button",
        attributes=' o:button="t" fillcolor="buttonFace [67]"',
        body='<v:fill color2="buttonFace [67]" o:detectmouseclick="t"/>' + _LOCK,
        slots=("Anchor", "PrintObject", "AutoFill", "FmlaMacro", "TextHAlign", "TextVAlign"),
        align="center",
    ),
    "CheckBox": _VmlStyle(
        object_type="Checkbox",
        attributes=(
            ' filled="f" fillcolor="windowText [64]" stroked="f"'
            ' strokecolor="window [65]" strokeweight="3e-5mm"'
        ),
        body=_FILL_WINDOW + '<v:path shadowok="t" strokeok="t" fillok="t"/>' + _LOCK,
        slots=_TICK_SLOTS,
    ),
    "Radio": _VmlStyle(
        object_type="Radio",
        attributes=(
            ' filled="f" fillcolor="windowText [64]" stroked="f"'
            ' strokecolor="window [65]" strokeweight="3e-5mm"'
        ),
        body=_FILL_WINDOW + '<v:path shadowok="t" strokeok="t" fillok="t"/>' + _LOCK,
        slots=(*_TICK_SLOTS, "FirstButton"),
    ),
    "Drop": _VmlStyle(
        object_type="Drop",
        attributes=' filled="f" fillcolor="windowText [64]" strokecolor="windowText [64]"',
        body=_FILL_WINDOW + _LOCK_TEXT,
        slots=("SizeWithCells", "Anchor", "AutoFill", *_LIST_SLOTS, "DropStyle", "DropLines"),
    ),
    "List": _VmlStyle(
        object_type="List",
        attributes=' fillcolor="window [65]" strokecolor="windowText [64]"',
        body=_FILL_WINDOW + _LOCK_TEXT,
        slots=("SizeWithCells", "Anchor", *_LIST_SLOTS),
    ),
    "Spin": _VmlStyle(
        object_type="Spin",
        attributes="",
        body=_LOCK_TEXT,
        slots=("Anchor", "FmlaLink", "Val", "Min", "Max", "Inc", "Page", "Dx"),
    ),
    "Scroll": _VmlStyle(
        object_type="Scroll",
        attributes="",
        body=_LOCK_TEXT,
        slots=(
            "SizeWithCells", "Anchor", "FmlaLink", "Val", "Min", "Max",
            "Inc", "Page", "Horiz", "Dx",
        ),
    ),
    "GBox": _VmlStyle(
        object_type="GBox",
        attributes=' fillcolor="window [65]" strokecolor="windowText [64]"',
        body=_FILL_WINDOW + _LOCK,
        slots=("SizeWithCells", "Anchor", "NoThreeD"),
    ),
    "Label": _VmlStyle(
        object_type="Label",
        attributes=' filled="f" fillcolor="windowText [64]" strokecolor="windowText [64]"',
        body=_FILL_WINDOW + _LOCK,
        slots=("Anchor", "AutoFill"),
    ),
}

#: The ClientData children that are the same whatever the control is.
_FIXED_SLOTS: dict[str, str] = {
    "SizeWithCells": "<x:SizeWithCells/>",
    "PrintObject": "<x:PrintObject>False</x:PrintObject>",
    "AutoFill": "<x:AutoFill>False</x:AutoFill>",
    "AutoLine": "<x:AutoLine>False</x:AutoLine>",
    "TextHAlign": "<x:TextHAlign>Center</x:TextHAlign>",
    "TextVAlign": "<x:TextVAlign>Center</x:TextVAlign>",
    "NoThreeD": "<x:NoThreeD/>",
    "NoThreeD2": "<x:NoThreeD2/>",
    "Horiz": "<x:Horiz/>",
    "Dx": "<x:Dx>31</x:Dx>",
    "SelType": "<x:SelType>Single</x:SelType>",
    "LCT": "<x:LCT>Normal</x:LCT>",
    "DropStyle": "<x:DropStyle>Combo</x:DropStyle>",
    "DropLines": "<x:DropLines>8</x:DropLines>",
}


def _client_data_slot(
    slot: str, shape: Shape, control: FormControl, grid: SheetGrid
) -> str:
    """One ClientData child, or nothing when it has nothing to say."""
    if slot == "Anchor":
        return f"<x:Anchor>{vml_anchor(shape, grid)}</x:Anchor>"
    if slot == "FmlaMacro":
        return f"<x:FmlaMacro>{escape_text(shape.macro)}</x:FmlaMacro>" if shape.macro else ""
    if slot == "FmlaLink":
        return (
            f"<x:FmlaLink>{escape_text(control.linked_cell)}</x:FmlaLink>"
            if control.linked_cell
            else ""
        )
    if slot == "FmlaRange":
        return (
            f"<x:FmlaRange>{escape_text(control.list_range)}</x:FmlaRange>"
            if control.list_range
            else ""
        )
    if slot == "Checked":
        # 1 ticked, 2 the third state, and nothing at all for off, which
        # is the same absence the reader turns back into xlOff.
        if control.value == XL_ON:
            return "<x:Checked>1</x:Checked>"
        return "<x:Checked>2</x:Checked>" if control.value == XL_MIXED else ""
    if slot == "FirstButton":
        return "<x:FirstButton/>" if control.first_button else ""
    if slot in ("Val", "Sel"):
        spelling = _VALUE_ATTRIBUTE.get(control.kind)
        value = control.value if spelling == slot.lower() else 0
        return f"<x:{slot}>{value}</x:{slot}>"
    if slot in ("Min", "Max", "Inc", "Page"):
        minimum, maximum, increment, page = control_range(control)
        chosen = {"Min": minimum, "Max": maximum, "Inc": increment, "Page": page}[slot]
        return f"<x:{slot}>{chosen}</x:{slot}>"
    return _FIXED_SLOTS.get(slot, "")


def control_vml(shape: Shape, grid: SheetGrid) -> str:
    """The VML Excel actually draws a control from.

    Every kind is drawn from the same ``_x0000_t201`` shape type and
    differs in everything else: the attributes on the shape, the fill
    inside it, and which ``x:ClientData`` children it carries in which
    order. All of that is measured rather than derived, because an
    ObjectType Excel does not recognise makes it refuse the workbook
    rather than draw the control differently.
    """
    control = shape.control or FormControl()
    style = _VML_STYLES[check_control_kind(control.kind or "Button")]

    text = ""
    if shape.text:
        rendered = "<br/>".join(escape_text(line) for line in shape.text.split("\n"))
        text = (
            "<v:textbox style='mso-direction-alt:auto' o:singleclick=\"f\">"
            f"<div style='text-align:{style.align}'>{rendered}</div></v:textbox>"
        )
    children = "".join(
        _client_data_slot(slot, shape, control, grid) for slot in style.slots
    )
    return (
        f'<v:shape id="{vml_id(shape.name)}" o:spid="_x0000_s{shape.shape_id}"'
        ' type="#_x0000_t201"'
        f" style='position:absolute;margin-left:{shape.left:g}pt;"
        f"margin-top:{shape.top:g}pt;width:{shape.width:g}pt;height:{shape.height:g}pt;"
        "z-index:1;mso-wrap-style:tight'"
        f'{style.attributes} o:insetmode="auto">'
        f"{style.body}{text}"
        f'<x:ClientData ObjectType="{style.object_type}">'
        f"{children}"
        "</x:ClientData>"
        "</v:shape>"
    )


def vml_id(name: str) -> str:
    """A shape's name as a VML id, which cannot hold a space."""
    cleaned = re.sub(r"[^A-Za-z0-9_.-]", "", name)
    return cleaned if cleaned and not cleaned[0].isdigit() else f"_{cleaned}"


#: One control's VML shape, found by the id the drawing gave it.
_VML_SHAPE = r'(<v:shape\b[^>]*?o:spid="_x0000_s{id}".*?</v:shape>)'

#: The macro inside it, which Excel writes but does not read back: the
#: sheet's own ``<controlPr macro=...>`` is what a click actually runs.
_VML_MACRO = re.compile(r"<x:FmlaMacro>.*?</x:FmlaMacro>", re.DOTALL)


def with_vml_shape(vml: str, markup: str) -> str:
    """A VML part with one more shape in it, before the closing tag."""
    at = vml.rfind("</xml>")
    return vml if at < 0 else vml[:at] + markup + vml[at:]


def without_vml_shape(vml: str, shape_id: int) -> str:
    """A VML part with one control's shape taken out."""
    pattern = re.compile(_VML_SHAPE.format(id=shape_id), re.DOTALL)
    return pattern.sub("", vml, count=1)


#: The content type of each part a shape needs. VML is a Default by
#: extension rather than an Override, which is how Excel writes it.
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing"
CT_CONTROL_PROPERTIES = "application/vnd.ms-excel.controlproperties+xml"

#: The relationship a sheet names a control's own part by.
RT_CONTROL_PROPERTIES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp"
)

#: The namespaces a control's markup uses from inside the worksheet part.
#: A sheet that has never carried one declares none of them.
NS_SPREADSHEET_DRAWING = (
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
)
NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
NS_MARKUP_COMPATIBILITY = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def anchor_holding(root: Element, name: str) -> Element | None:
    """The top-level node to remove to take a named shape off a sheet.

    Not the anchor itself where a control is concerned: Excel wraps a
    control's anchor in an ``mc:AlternateContent``, and removing the inner
    anchor would leave the empty wrapper behind.
    """
    for child in root.children:
        if not isinstance(child, Element):
            continue
        local = _local(child.name)
        if local in ANCHORS:
            body = _shape_element(child)
            if body is not None and _named(body, name) is not None:
                return child
        elif local == "AlternateContent":
            chosen = _find(child, "Choice") or _find(child, "Fallback")
            if chosen is not None and anchor_holding(chosen, name) is not None:
                return child
    return None


def find_shape_element(root: Element, name: str) -> Element | None:
    """The element behind a named shape, anywhere in a drawing.

    Anchors inside ``mc:AlternateContent`` are searched too, and so are a
    group's members, so a shape nested in a group is reachable.
    """
    for anchor in _anchors(root):
        body = _shape_element(anchor)
        found = None if body is None else _named(body, name)
        if found is not None:
            return found
    return None


def _named(body: Element, name: str) -> Element | None:
    if _shape_name(body) == name:
        return body
    for child in body.children:
        if isinstance(child, Element) and _local(child.name) in _ELEMENT_KINDS:
            found = _named(child, name)
            if found is not None:
                return found
    return None


def _shape_name(body: Element) -> str:
    properties = _find(body, "nvSpPr") or _find(body, "nvCxnSpPr") or _find(body, "nvGrpSpPr")
    properties = properties or _find(body, "nvPicPr") or _find(body, "nvGraphicFramePr")
    naming = None if properties is None else _find(properties, "cNvPr")
    return "" if naming is None else (naming.get("name") or "")


def qualified_macro(macro: str, existing: str = "") -> str:
    """A macro name as a control's own record spells it.

    Excel writes ``[0]!Clicked`` there rather than the bare name a drawing
    shape carries. The bracketed number indexes the workbook holding the
    procedure and is not always 0: ``shapes.xlsm`` says ``[1]``, and it
    carries an external link where ``controls.xlsm`` carries none.

    What decides the number is not measured here, so an existing prefix is
    kept rather than replaced, and a name given with a ``!`` already in it
    is written through untouched. Only a control that has never had a
    macro gets a guessed ``[0]``.
    """
    if not macro:
        return ""
    if "!" in macro:
        return macro
    prefix = existing.partition("!")[0] if "!" in existing else "[0]"
    return f"{prefix}!{macro}"


def set_vml_macro(text: str, shape_id: int, macro: str) -> str:
    """One control's VML with its macro set, cleared, or added.

    Measured: Excel reads the sheet's ``<controlPr macro=...>`` and ignores
    this copy. A file where the two disagree opens cleanly and runs what
    the sheet says. It is written anyway, because Excel writes both and a
    stale name left in the package is a trap for the next reader.

    ``<x:FmlaMacro>`` goes after ``<x:AutoFill>`` where there is one, which
    is where Excel puts it.
    """
    pattern = re.compile(_VML_SHAPE.format(id=shape_id), re.DOTALL)
    match = pattern.search(text)
    if match is None:
        return text

    block = match.group(1)
    replacement = f"<x:FmlaMacro>{escape_text(macro)}</x:FmlaMacro>" if macro else ""
    if _VML_MACRO.search(block) is not None:
        changed = _VML_MACRO.sub(lambda _: replacement, block, count=1)
    elif not macro:
        return text
    else:
        anchor = re.search(r"</x:AutoFill>|</x:Anchor>", block)
        if anchor is None:
            changed = block.replace("</x:ClientData>", f"{replacement}</x:ClientData>", 1)
        else:
            at = anchor.end()
            changed = block[:at] + replacement + block[at:]
    return text[: match.start(1)] + changed + text[match.end(1) :]


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
    "AUTO_SHAPE_TYPES",
    "CONTROL_KINDS",
    "CT_CONTROL_PROPERTIES",
    "CT_DRAWING",
    "CT_VML",
    "DEFAULT_COLUMN_POINTS",
    "DEFAULT_GEOMETRY",
    "DEFAULT_ROW_POINTS",
    "EMPTY_DRAWING",
    "EMPTY_VML",
    "EMU_PER_POINT",
    "FIRST_CONTROL_ID",
    "MAX_ANCHOR_COLUMN",
    "MAX_ANCHOR_ROW",
    "MSO_TYPE",
    "NS_MARKUP_COMPATIBILITY",
    "NS_SPREADSHEET_DRAWING",
    "NS_X14",
    "PRESET_GEOMETRY",
    "RANGE_DEFAULTS",
    "RT_CONTROL_PROPERTIES",
    "XL_MIXED",
    "XL_OFF",
    "XL_ON",
    "FormControl",
    "Shape",
    "ShapeKind",
    "SheetGrid",
    "anchor_holding",
    "characters_to_points",
    "check_control_kind",
    "control_drawing",
    "control_entry",
    "control_properties",
    "control_range",
    "control_vml",
    "corner_markup",
    "emu",
    "find_shape_element",
    "new_anchor",
    "points",
    "qualified_macro",
    "read_control",
    "read_drawing",
    "set_vml_macro",
    "text_body",
    "vml_anchor",
    "vml_id",
    "with_vml_shape",
    "without_vml_shape",
]
