"""Column widths, row heights, hiding, and frozen panes.

Rows and columns are stored in completely different places, and only one of
them is simple.

A row's settings live on its own ``<row>`` element, which the sheet already
has whenever the row holds cells::

    <row r="1" ht="24" customHeight="1">...</row>

A column's live in a separate ``<cols>`` block before ``sheetData``, and each
entry covers a *range* of columns rather than one::

    <cols><col min="1" max="1" width="18.6328125" customWidth="1"/>
          <col min="3" max="3" width="10.08984375" customWidth="1"/></cols>

So setting one column's width inside a span means splitting that span into
up to three entries, copying every attribute to each piece. Editing the
entry in place instead would resize every column it covers, which is the
same shared-state trap that cell formatting has.

**A width is not the number you type into Excel.** Setting ``ColumnWidth =
18`` in VBA stores ``18.6328125``; the offset was ``+0.6328125`` for every
integer width measured, and ``8.43``, Excel's default, broke even that
pattern by landing on ``9.08984375``. The unit is a count of ``0`` glyphs in
the workbook's default font plus padding, so converting needs that font's
maximum digit width in pixels, which is not in the file. This module
therefore exposes the stored number and says so, rather than applying a
conversion that would be right for one font and quietly wrong for the rest.
Row heights carry no such trap: they are points, and ``24`` stores as ``24``.

Both settings need a companion flag. A ``width`` without ``customWidth="1"``
and an ``ht`` without ``customHeight="1"`` are ignored by Excel, so a value
set without them looks like it simply did not take.
"""

from __future__ import annotations

from dataclasses import dataclass

from pyofficeeditor._xml import Element
from pyofficeeditor.excel._reference import MAX_COLUMN, CellRef

#: Every attribute a ``<col>`` entry can carry besides ``min`` and ``max``.
#: Splitting a span copies all of them to each piece, so a column that was
#: hidden and styled stays hidden and styled.
COLUMN_ATTRIBUTES = (
    "width",
    "style",
    "hidden",
    "bestFit",
    "customWidth",
    "phonetic",
    "outlineLevel",
    "collapsed",
)

#: Which pane Excel activates after a freeze, decided by which splits exist.
#: Getting it wrong leaves the cursor in a pane the user cannot see.
_ACTIVE_PANE = {
    (True, True): "bottomRight",
    (False, True): "bottomLeft",
    (True, False): "topRight",
}


@dataclass(frozen=True)
class Freeze:
    """Where a sheet's panes are frozen."""

    #: The number of columns frozen on the left.
    columns: int = 0
    #: The number of rows frozen at the top.
    rows: int = 0

    @property
    def is_frozen(self) -> bool:
        return self.columns > 0 or self.rows > 0

    @property
    def top_left(self) -> CellRef:
        """The first unfrozen cell, which is what ``topLeftCell`` names."""
        return CellRef(self.rows + 1, self.columns + 1)

    @classmethod
    def at(cls, reference: CellRef) -> Freeze:
        """The freeze that puts ``reference`` at the top left of the
        scrolling area, which is how Excel's own command is described:
        freezing at B2 pins row 1 and column A.
        """
        return cls(columns=reference.column - 1, rows=reference.row - 1)

    def write(self) -> Element:
        element = Element.create("pane")
        if self.columns:
            element.set("xSplit", str(self.columns))
        if self.rows:
            element.set("ySplit", str(self.rows))
        element.set("topLeftCell", self.top_left.a1)
        element.set("activePane", _ACTIVE_PANE[(self.columns > 0, self.rows > 0)])
        element.set("state", "frozen")
        return element

    @classmethod
    def read(cls, element: Element | None) -> Freeze:
        if element is None or (element.get("state") or "split") not in ("frozen", "frozenSplit"):
            return cls()
        return cls(columns=_as_int(element.get("xSplit")), rows=_as_int(element.get("ySplit")))


def _as_int(raw: str | None) -> int:
    if raw is None:
        return 0
    try:
        return int(float(raw))
    except ValueError:
        return 0


def column_entry(container: Element, index: int) -> Element | None:
    """The ``<col>`` covering a column, without changing anything."""
    for entry in container.children_named("col"):
        low = _as_int(entry.get("min"))
        high = _as_int(entry.get("max"))
        if low <= index <= high:
            return entry
    return None


def isolate_column(container: Element, index: int) -> Element:
    """A ``<col>`` covering exactly one column, splitting a span if needed.

    A span covering 1 to 5 becomes up to three entries when column 3 is
    isolated, each carrying the original's attributes. Editing the span in
    place instead would resize all five.
    """
    if not 1 <= index <= MAX_COLUMN:
        raise ValueError(f"column {index} is outside 1..{MAX_COLUMN}")

    existing = column_entry(container, index)
    if existing is None:
        created = Element.create("col", {"min": str(index), "max": str(index)})
        _insert_ordered(container, created)
        return created

    low = _as_int(existing.get("min"))
    high = _as_int(existing.get("max"))
    if low == index and high == index:
        return existing

    attributes = {name: existing.get(name) for name in COLUMN_ATTRIBUTES}
    middle = Element.create("col", {"min": str(index), "max": str(index)})
    _copy_attributes(attributes, middle)

    if low < index:
        # The span keeps its head and the new entry follows it.
        existing.set("max", str(index - 1))
        container.insert_after(existing, middle)
        if high > index:
            tail = Element.create("col", {"min": str(index + 1), "max": str(high)})
            _copy_attributes(attributes, tail)
            container.insert_after(middle, tail)
    else:
        # The span starts at the wanted column, so it keeps the tail.
        existing.set("min", str(index + 1))
        container.insert_before(existing, middle)

    return middle


def _copy_attributes(attributes: dict[str, str | None], target: Element) -> None:
    for name, value in attributes.items():
        if value is not None:
            target.set(name, value)


def _insert_ordered(container: Element, entry: Element) -> None:
    """Keep ``<cols>`` in ascending order, which is how Excel writes it."""
    index = _as_int(entry.get("min"))
    for existing in container.children_named("col"):
        if _as_int(existing.get("min")) > index:
            container.insert_before(existing, entry)
            return
    container.append(entry)


__all__ = [
    "COLUMN_ATTRIBUTES",
    "Freeze",
    "column_entry",
    "isolate_column",
]
