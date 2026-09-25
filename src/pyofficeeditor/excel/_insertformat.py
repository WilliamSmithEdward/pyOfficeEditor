"""What an inserted row or column takes from the one before it.

Excel's Insert formats a new row like the row above it, and a new column
like the column to its left, unless it is told not to. Measured against
Excel, one feature at a time.

Taken:

- a row's height (``ht``, ``customHeight``), its style (``s``,
  ``customFormat``), its outline level and its ``x14ac:dyDescent``; a
  column's ``<col>`` settings, width, style and outline level among them
- each cell's style, where it differs from what the new row or column would
  give that cell anyway: the row's style when the row has one, and the
  column's otherwise. The cell comes across empty. Styles are compared by
  what they hold, not by index: Excel keeps duplicates of its default
  format, a merge makes one, and a cell in one takes nothing across.
- a conditional format, a data validation or a protected range that ends on
  the row above: it grows over the new rows. Both the 2006 form and the
  ``x14`` one in ``extLst`` do.
- a sparkline on the row above, which is copied into each new row reading
  its data as far further down as the new row is, the way a relative
  formula copies

Left behind: hidden, even for a row inserted inside a collapsed group;
``collapsed``; values; merges; hyperlinks, though a hyperlink cell's style
comes across; tables; autofilters; ignored errors. Row 1 and column A have
nothing before them, and a row or column inserted there takes nothing.

A hidden column Excel stores at width 0 has no width of its own to give, and
a new column beside it gets the standard width.

A visible row or column that lands in a collapsed group opens it: Excel
clears ``collapsed`` on the summary row or column of every group the new
ones join, wherever the sheet puts its summaries.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pyofficeeditor._xml import Element, XmlDocument, local_name
from pyofficeeditor.excel._dimensions import COLUMN_ATTRIBUTES, column_entry, format_width, says_nothing
from pyofficeeditor.excel._formulas import translate_formula
from pyofficeeditor.excel._reference import MAX_COLUMN, MAX_ROW, CellRef, RangeRef

if TYPE_CHECKING:
    from pyofficeeditor.excel._formats import CellFormat
    from pyofficeeditor.excel._styles import Styles
    from pyofficeeditor.excel.worksheet import Worksheet

#: What a new row takes from the row above it, by local name.
ROW_FORMAT = ("ht", "customHeight", "s", "customFormat", "outlineLevel", "dyDescent")

#: What a new column does not take from the column to its left.
_NOT_TAKEN = ("hidden", "collapsed")


class _Formats:
    """Cell styles compared by what they hold, as Excel compares them."""

    def __init__(self, styles: Styles | None) -> None:
        self._styles = styles
        self._resolved: dict[str, CellFormat] = {}

    def same(self, first: str, second: str) -> bool:
        if first == second:
            return True
        if self._styles is None:
            return False
        return self._format(first) == self._format(second)

    def _format(self, style: str) -> CellFormat:
        found = self._resolved.get(style)
        if found is None:
            assert self._styles is not None
            found = self._styles.cell_format(int(style) if style.isdigit() else None)
            self._resolved[style] = found
        return found


def copy_row_format(sheet: Worksheet, at: int, count: int) -> RangeRef | None:
    """Format rows ``at`` to ``at + count - 1``, just inserted and empty,
    like the row above them. Returns the block of cells it made, for the
    sheet's ``dimension`` to cover once the rest has moved."""
    if at <= 1:
        return None
    root = sheet.document.root
    data = root.child("sheetData")
    above = sheet.rows_by_number().get(at - 1)
    if data is None or above is None:
        return None
    taken = [(name, value) for name, value in above.attributes.items() if local_name(name) in ROW_FORMAT]
    row_style = above.get("s") if _true(above.get("customFormat")) else None
    styles = _column_styles(root)
    formats = _Formats(sheet.workbook.styles)
    cells: list[tuple[int, str]] = []
    for cell in above.children_named("c"):
        column = _column_of(cell)
        style = cell.get("s") or "0"
        if column is not None and not formats.same(style, row_style or styles.get(column, "0")):
            cells.append((column, style))
    if not taken and not cells:
        return None

    previous = above
    for number in range(at, at + count):
        row = Element.create("row", {"r": str(number)})
        for name, value in taken:
            row.set(name, value)
        for column, style in cells:
            row.append(Element.create("c", {"r": CellRef(number, column).a1, "s": style}))
        data.insert_after(previous, row)
        previous = row
    sheet.reindex_rows()

    level = _level(above.get("outlineLevel"))
    if level:
        _open_row_groups(sheet, at, at + count - 1, level)
    if not cells:
        return None
    return RangeRef(CellRef(at, cells[0][0]), CellRef(at + count - 1, cells[-1][0]))


def place_column_entries(root: Element, at: int, count: int, *, copy_format: bool, standard: float) -> Element | None:
    """Move the ``<col>`` entries for columns ``at`` onward ``count`` to the
    right, and give the new columns the left column's settings, or none.

    An entry spanning the insertion is split there first, so a new column
    takes nothing it was not given. Then the left column's settings, less
    ``hidden`` and ``collapsed``, go to the new ones: by stretching its entry
    when that says the same, as Excel does, or in an entry of their own.
    Returns the entry the new columns take their settings from, or ``None``.
    """
    container = root.child("cols")
    if container is None:
        return None
    for entry in list(container.children_named("col")):
        low, high = _span(entry)
        if low < at <= high:
            tail = _copied(entry)
            tail.set("min", str(at))
            entry.set("max", str(at - 1))
            container.insert_after(entry, tail)
    for entry in container.children_named("col"):
        low, high = _span(entry)
        if low >= at:
            entry.set("min", str(min(low + count, MAX_COLUMN)))
            entry.set("max", str(min(high + count, MAX_COLUMN)))
    if not copy_format or at <= 1:
        return None

    left = column_entry(container, at - 1)
    if left is None:
        return None
    settings = {name: value for name in COLUMN_ATTRIBUTES if name not in _NOT_TAKEN and (value := left.get(name)) is not None}
    if not _positive(settings.get("width")):
        settings["width"] = format_width(standard)
        settings.pop("customWidth", None)
    new = Element.create("col", {"min": str(at), "max": str(at + count - 1), **settings})
    if says_nothing(new, standard):
        return None
    if _settings(left) == settings and _span(left)[1] == at - 1:
        left.set("max", str(at + count - 1))
        placed = left
    else:
        container.insert_after(left, new)
        placed = new
    following = _next_col(placed)
    if following is not None and _span(following)[0] == at + count and _settings(following) == _settings(placed):
        placed.set("max", str(_span(following)[1]))
        container.remove(following)
    return placed


def copy_column_format(
    sheet: Worksheet, at: int, count: int, source: Element | None, *, standard: float
) -> RangeRef | None:
    """Format columns ``at`` to ``at + count - 1``, just inserted and
    empty, like the column to their left: its cells' styles, and its
    groups opened. ``source`` is the ``<col>`` the new columns took their
    settings from, as :func:`place_column_entries` returned it. Returns the
    block of cells it made, as :func:`copy_row_format` does."""
    if at <= 1:
        return None
    root = sheet.document.root
    column_style = "0" if source is None else (source.get("style") or "0")
    formats = _Formats(sheet.workbook.styles)
    added: list[int] = []
    for number, row in sorted(sheet.rows_by_number().items()):
        default = row.get("s") if _true(row.get("customFormat")) else column_style
        for cell in row.children_named("c"):
            column = _column_of(cell)
            if column is None or column < at - 1:
                continue
            style = cell.get("s") or "0"
            if column == at - 1 and not formats.same(style, default or "0"):
                previous = cell
                for new_column in range(at, at + count):
                    created = Element.create("c", {"r": CellRef(number, new_column).a1, "s": style})
                    row.insert_after(previous, created)
                    previous = created
                added.append(number)
            break
    level = 0 if source is None else _level(source.get("outlineLevel"))
    if level:
        _open_column_groups(root, at, at + count - 1, level, standard=standard)
    if not added:
        return None
    return RangeRef(CellRef(added[0], at), CellRef(added[-1], at + count - 1))


def grow_ranges(root: Element, at: int, count: int, *, is_row: bool) -> None:
    """Stretch each conditional format, data validation and protected range
    ending on the row above, or the column to the left, over the new ones."""
    for element in root.children_named("conditionalFormatting"):
        _grow_attribute(element, at, count, is_row=is_row)
    for container, entry in (("dataValidations", "dataValidation"), ("protectedRanges", "protectedRange")):
        holder = root.child(container)
        if holder is not None:
            for element in holder.children_named(entry):
                _grow_attribute(element, at, count, is_row=is_row)
    extensions = root.child("extLst")
    if extensions is None:
        return
    for sqref in extensions.descendants("sqref"):
        owner = sqref.parent
        if owner is None or local_name(owner.name) not in ("conditionalFormatting", "dataValidation"):
            continue
        grown = _grown(sqref.text, at, count, is_row=is_row)
        if grown != sqref.text:
            sqref.set_text(grown)


def copy_sparklines(root: Element, at: int, count: int, *, is_row: bool) -> None:
    """Copy each sparkline on the row above, or in the column to the left,
    into each new row or column, reading its data that much further on."""
    extensions = root.child("extLst")
    if extensions is None:
        return
    for sparkline in list(extensions.descendants("sparkline")):
        location = sparkline.child("sqref")
        parent = sparkline.parent
        try:
            cell = CellRef.parse(location.text if location is not None else "")
        except ValueError:
            continue
        if parent is None or (cell.row if is_row else cell.column) != at - 1:
            continue
        previous = sparkline
        for step in range(1, count + 1):
            rows, columns = (step, 0) if is_row else (0, step)
            copy = XmlDocument.parse(sparkline.to_xml().encode("utf-8")).root
            copied_location = copy.child("sqref")
            if copied_location is not None:
                copied_location.set_text(cell.translated(rows, columns).a1)
            data = copy.child("f")
            if data is not None and data.text:
                data.set_text(translate_formula(data.text, rows, columns))
            parent.insert_after(previous, copy)
            previous = copy


# ----------------------------------------------------------------------
# Groups
# ----------------------------------------------------------------------


def _open_row_groups(sheet: Worksheet, first: int, last: int, level: int) -> None:
    """Clear ``collapsed`` on the summary row of each group rows ``first``
    to ``last``, visible and at ``level``, have joined."""
    rows = sheet.rows_by_number()
    below = _summary_after(sheet.document.root, "summaryBelow")
    number, step = (last + 1, 1) if below else (first - 1, -1)
    threshold = level
    while 1 <= number <= MAX_ROW and threshold > 0:
        row = rows.get(number)
        found = 0 if row is None else _level(row.get("outlineLevel"))
        if found < threshold:
            if row is not None:
                row.unset("collapsed")
            threshold = found
        number += step


def _open_column_groups(root: Element, first: int, last: int, level: int, *, standard: float) -> None:
    """The same for columns, whose summary sits to the right unless the
    sheet says otherwise. An entry left saying nothing goes, as Excel drops
    it."""
    container = root.child("cols")
    if container is None:
        return
    right = _summary_after(root, "summaryRight")
    number, step = (last + 1, 1) if right else (first - 1, -1)
    threshold = level
    while 1 <= number <= MAX_COLUMN and threshold > 0:
        entry = column_entry(container, number)
        found = 0 if entry is None else _level(entry.get("outlineLevel"))
        if found < threshold:
            if entry is not None and entry.unset("collapsed") and says_nothing(entry, standard):
                container.remove(entry)
            threshold = found
        number += step


def _summary_after(root: Element, attribute: str) -> bool:
    """Whether a sheet's outline puts each summary after its group, which
    is the default for rows and columns alike."""
    properties = root.child("sheetPr")
    outline = None if properties is None else properties.child("outlinePr")
    value = None if outline is None else outline.get(attribute)
    return value is None or _true(value)


# ----------------------------------------------------------------------
# Ranges
# ----------------------------------------------------------------------


def _grow_attribute(element: Element, at: int, count: int, *, is_row: bool) -> None:
    raw = element.get("sqref")
    if raw is None:
        return
    grown = _grown(raw, at, count, is_row=is_row)
    if grown != raw:
        element.set("sqref", grown)


def _grown(raw: str, at: int, count: int, *, is_row: bool) -> str:
    """An ``sqref`` with each range that ends on the row, or column, before
    ``at`` stretched ``count`` further."""
    pieces: list[str] = []
    for piece in raw.split():
        try:
            block = RangeRef.parse(piece)
        except ValueError:
            pieces.append(piece)
            continue
        end = block.end
        if is_row and end.row == at - 1:
            block = RangeRef(block.start, CellRef(min(end.row + count, MAX_ROW), end.column))
        elif not is_row and end.column == at - 1:
            block = RangeRef(block.start, CellRef(end.row, min(end.column + count, MAX_COLUMN)))
        else:
            pieces.append(piece)
            continue
        pieces.append(block.a1)
    return " ".join(pieces)


# ----------------------------------------------------------------------
# Small things
# ----------------------------------------------------------------------


def _true(value: str | None) -> bool:
    return value in ("1", "true")


def _level(value: str | None) -> int:
    try:
        return int(value or 0)
    except ValueError:
        return 0


def _positive(value: str | None) -> bool:
    try:
        return float(value or 0) > 0
    except ValueError:
        return False


def _column_of(cell: Element) -> int | None:
    try:
        return CellRef.parse(cell.get("r") or "").column
    except ValueError:
        return None


def _column_styles(root: Element) -> dict[int, str]:
    """Each styled column's style, by column number."""
    styles: dict[int, str] = {}
    container = root.child("cols")
    if container is None:
        return styles
    for entry in container.children_named("col"):
        style = entry.get("style")
        if style is None:
            continue
        low, high = _span(entry)
        for column in range(low, min(high, MAX_COLUMN) + 1):
            styles[column] = style
    return styles


def _span(entry: Element) -> tuple[int, int]:
    try:
        return int(entry.get("min") or 0), int(entry.get("max") or 0)
    except ValueError:
        return 0, 0


def _settings(entry: Element) -> dict[str, str]:
    return {name: value for name in COLUMN_ATTRIBUTES if (value := entry.get(name)) is not None}


def _copied(entry: Element) -> Element:
    return Element.create("col", {"min": entry.get("min") or "", "max": entry.get("max") or "", **_settings(entry)})


def _next_col(entry: Element) -> Element | None:
    parent = entry.parent
    if parent is None:
        return None
    siblings = list(parent.children_named("col"))
    index = next(position for position, sibling in enumerate(siblings) if sibling is entry)
    return siblings[index + 1] if index + 1 < len(siblings) else None


def widen_dimension(root: Element, block: RangeRef) -> None:
    """Keep the sheet's ``dimension`` covering cells just added, once
    everything the insertion moves has moved."""
    element = root.child("dimension")
    raw = None if element is None else element.get("ref")
    if element is None or raw is None:
        return
    try:
        current = RangeRef.parse(raw)
    except ValueError:
        return
    widened = current.expanded(block.start).expanded(block.end)
    if widened.a1 != raw:
        element.set("ref", widened.a1)


__all__ = [
    "ROW_FORMAT",
    "copy_column_format",
    "copy_row_format",
    "copy_sparklines",
    "grow_ranges",
    "place_column_entries",
    "widen_dimension",
]
