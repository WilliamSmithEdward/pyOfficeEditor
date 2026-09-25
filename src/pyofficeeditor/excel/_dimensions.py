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

A height needs its companion flag: an ``ht`` without ``customHeight="1"``
is ignored, so a height set without it looks like it did not take. A width
is honoured with or without ``customWidth="1"``, measured, and the flag is
written anyway, as Excel writes it for a width set by hand.

**A ``<col>`` without a width is a column of width 0.** Measured: an entry
carrying only ``outlineLevel``, ``collapsed``, ``style`` or nothing at all
shows as a hidden column, and stays hidden when its outline is expanded. So
an entry made for a column at the standard width is given that width, as
Excel gives it: the sheet's ``defaultColWidth`` when it has one, and
otherwise a width fixed by the Normal font, measured per font in
:data:`STANDARD_WIDTHS`.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyofficeeditor._xml import Element
from pyofficeeditor.excel._reference import MAX_COLUMN, CellRef

if TYPE_CHECKING:
    from pyofficeeditor.excel.worksheet import Worksheet

#: The width Excel stores for a column at the standard width, by the
#: Normal style's font, measured on Excel 16 at 100% display scaling. It is
#: the standard eight digits plus padding, rounded up to a multiple of
#: eight pixels and stored in digits, so fonts whose widest digit has the
#: same pixel width store the same number. At other scalings the pixels,
#: and so the number, differ: a file Excel wrote at 150% stored 8.7265625
#: for Aptos Narrow 11.
STANDARD_WIDTHS: dict[tuple[str, float], float] = {
    ("Calibri", 11): 9.140625,
    ("Aptos Narrow", 11): 9.140625,
    ("Arial", 10): 9.140625,
    ("Calibri", 10): 9.140625,
    ("Tahoma", 10): 9.140625,
    ("Aptos", 11): 9.0,
    ("Arial", 11): 9.0,
    ("Calibri", 12): 9.0,
    ("Calibri Light", 11): 9.0,
    ("Cambria", 11): 9.0,
    ("Consolas", 11): 9.0,
    ("Courier New", 10): 9.0,
    ("Garamond", 12): 9.0,
    ("Times New Roman", 12): 9.0,
    ("Verdana", 10): 9.0,
    ("Arial Narrow", 10): 9.33203125,
    ("Segoe UI", 9): 9.33203125,
    ("Georgia", 11): 8.88671875,
}

#: Calibri 11's and Aptos Narrow 11's, the Normal fonts of Excel before
#: 2023 and since, for a font not measured.
FALLBACK_STANDARD_WIDTH = 9.140625


def standard_width(format_properties: Element | None, font_name: str | None, font_size: float | None) -> float:
    """What a column at a sheet's standard width stores as its width.

    ``format_properties`` is the sheet's ``<sheetFormatPr>``, whose
    ``defaultColWidth``, when a standard width has been set, is the answer
    for every font measured; the font is the Normal style's.
    """
    if format_properties is not None:
        raw = format_properties.get("defaultColWidth")
        if raw is not None:
            try:
                return float(raw)
            except ValueError:
                pass
    if font_name is None or font_size is None:
        return FALLBACK_STANDARD_WIDTH
    return STANDARD_WIDTHS.get((font_name, font_size), FALLBACK_STANDARD_WIDTH)


def sheet_standard_width(sheet: Worksheet) -> float:
    """What a column at a sheet's standard width stores, which an entry
    made for one has to carry: by the sheet's own default, else the
    workbook's Normal font."""
    styles = sheet.workbook.styles
    font = None if styles is None else styles.cell_format(None).font
    return standard_width(
        sheet.document.root.child("sheetFormatPr"),
        None if font is None else font.name,
        None if font is None else font.size,
    )

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


def isolate_column(container: Element, index: int, *, width: float) -> Element:
    """A ``<col>`` covering exactly one column, splitting a span if needed.

    A span covering 1 to 5 becomes up to three entries when column 3 is
    isolated, each carrying the original's attributes. Editing the span in
    place instead would resize all five. A column with no entry gets one at
    ``width``, the standard width, since an entry without one is a column
    of width 0.
    """
    if not 1 <= index <= MAX_COLUMN:
        raise ValueError(f"column {index} is outside 1..{MAX_COLUMN}")

    existing = column_entry(container, index)
    if existing is None:
        created = Element.create("col", {"min": str(index), "max": str(index), "width": format_width(width)})
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


def format_width(width: float) -> str:
    """A width as Excel writes one: ``9`` rather than ``9.0``."""
    return f"{width:.15g}"


def says_nothing(entry: Element, standard: float) -> bool:
    """Whether a ``<col>`` says no more than having no entry would: nothing
    besides its span, or only the standard width."""
    attributes = {name for name in COLUMN_ATTRIBUTES if entry.get(name) is not None}
    if not attributes:
        return True
    if attributes != {"width"}:
        return False
    try:
        return float(entry.get("width") or "") == standard
    except ValueError:
        return False


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
    "FALLBACK_STANDARD_WIDTH",
    "STANDARD_WIDTHS",
    "Freeze",
    "column_entry",
    "format_width",
    "isolate_column",
    "says_nothing",
    "sheet_standard_width",
    "standard_width",
]
