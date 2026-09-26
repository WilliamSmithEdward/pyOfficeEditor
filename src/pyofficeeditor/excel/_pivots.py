"""Pivot tables: where each one is, and what its cache summarises.

A pivot table is a part the sheet points at, whose ``<location ref>`` is
the block of the sheet's grid it fills. It points at a cache definition,
which the workbook lists in ``<pivotCaches>`` and whose
``<worksheetSource ref="A1:C11" sheet="Data"/>`` names the cells the cache
was read from; a cache may serve more than one pivot table.

Measured, with Excel making each edit on a fresh copy and saving it:

- Excel refuses to insert or delete rows or columns that cut through a
  pivot table: inserting inside it, or deleting part of it. It allows
  inserting at its first row or column, which moves it whole.
- Deleting every row or column of a pivot table deletes the table, and a
  cache no pivot table uses any more goes with it.
- A pivot table's location moves with its cells, and a cache's source
  moves, grows and shrinks as a formula's range does. A source deleted
  outright keeps its address, since the attribute has no way to write
  ``#REF!``, and a renamed sheet is renamed there too.

What a pivot table shows, for GETPIVOTDATA, is read from its definition:
each ``<i>`` of ``<rowItems>`` and ``<colItems>`` is a row or column of
its data area, naming the item each field of that axis stands at, and the
values themselves are the cells Excel wrote there.
"""

from __future__ import annotations

import datetime as dt
import math
from collections.abc import Callable
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyofficeeditor.excel import _collate
from pyofficeeditor.excel._numfmt import format_value
from pyofficeeditor.excel._reference import CellRef, RangeRef
from pyofficeeditor.excel._values import CellError, datetime_to_serial
from pyofficeeditor.excel._xstring import decode

if TYPE_CHECKING:
    from pyofficeeditor._xml import Element
    from pyofficeeditor.opc import OpcPackage

RT_PIVOT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
RT_PIVOT_CACHE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
RT_PIVOT_RECORDS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords"


@dataclass(frozen=True)
class PivotTable:
    """A pivot table on a sheet, and the cells its cache was read from."""

    name: str
    #: The block of the sheet the pivot table fills.
    location: RangeRef
    part_name: str
    #: The cache definition it reads from.
    cache_part: str | None = None
    #: The sheet and range the cache was read from, when it was read from a
    #: range.
    source_sheet: str | None = None
    source_range: RangeRef | None = None
    #: The table or defined name the cache was read from, when it was.
    source_name: str | None = None


def read_pivot_tables(package: OpcPackage, sheet_part: str) -> list[PivotTable]:
    """The pivot tables a sheet holds, in the order it points at them."""
    found: list[PivotTable] = []
    for relationship in package.relationships(sheet_part).by_type(RT_PIVOT_TABLE):
        part = relationship.target_part
        if relationship.is_external or not package.has_part(part):
            continue
        root = package.xml(part).root
        location = root.child("location")
        raw = None if location is None else location.get("ref")
        try:
            block = RangeRef.parse(raw or "")
        except ValueError:
            continue
        cache = next(
            (
                entry.target_part
                for entry in package.relationships(part).by_type(RT_PIVOT_CACHE)
                if not entry.is_external and package.has_part(entry.target_part)
            ),
            None,
        )
        source = None if cache is None else worksheet_source(package.xml(cache).root)
        sheet, reference, name = _source_of(source)
        found.append(
            PivotTable(
                name=decode(root.get("name") or ""),
                location=block.normalized,
                part_name=part,
                cache_part=cache,
                source_sheet=sheet,
                source_range=reference,
                source_name=name,
            )
        )
    return found


def worksheet_source(cache: Element) -> Element | None:
    """A cache definition's ``<worksheetSource>``, if it reads from cells."""
    source = cache.child("cacheSource")
    return None if source is None else source.child("worksheetSource")


def cache_parts(package: OpcPackage, workbook_part: str) -> list[str]:
    """Every cache definition the workbook lists."""
    return [
        relationship.target_part
        for relationship in package.relationships(workbook_part).by_type(RT_PIVOT_CACHE)
        if not relationship.is_external and package.has_part(relationship.target_part)
    ]


# ----------------------------------------------------------------------
# The report, as GETPIVOTDATA reads it
# ----------------------------------------------------------------------

#: The field a report's rows or columns list its values under, when it
#: summarizes more than one.
VALUES = -2

#: What turns on a field's subtotals besides its default one.
_SUBTOTAL_FLAGS = (
    "sumSubtotal", "countASubtotal", "avgSubtotal", "maxSubtotal", "minSubtotal", "productSubtotal",
    "countSubtotal", "stdDevSubtotal", "stdDevPSubtotal", "varSubtotal", "varPSubtotal",
)  # fmt: skip

Item = str | float | bool | CellError | None


@dataclass(frozen=True)
class PivotItem:
    """An item of a field: its value in the cache, a date as its serial
    number, and where the cache lists it; the text the report shows it as;
    and the name it was given in the report, if it was."""

    value: Item
    index: int
    text: str
    caption: str | None


@dataclass(frozen=True)
class PivotField:
    """A field of a report: the name the report gives it, its items in
    the report's order, ``None`` for an entry that is a subtotal, and
    whether it shows subtotals and whether at the top of each group.

    A grouped field says how: ``years``, ``quarters`` or ``months`` of a
    date, or ``range`` for numbers in bins of ``interval`` from ``start``.
    """

    name: str
    items: tuple[PivotItem | None, ...]
    subtotals: bool
    subtotals_top: bool
    grouping: str | None = None
    start: float = 0.0
    interval: float = 0.0

    def number_index(self, wanted: float) -> int | None:
        """The item a number picks. Measured: a date grouped by years,
        quarters or months takes the whole number part as the year, the
        quarter or the month; a bin is found by whole steps from its start,
        toward zero, and takes a number only at or below the bin's start;
        anything else has to be equal."""
        target: int | None = None
        if self.grouping == "range":
            if self.interval <= 0:
                return None
            steps = math.trunc((wanted - self.start) / self.interval)
            if wanted > self.start + steps * self.interval:
                return None
            target = steps + 1
        elif self.grouping in ("years", "quarters", "months"):
            target = math.trunc(wanted)
        for index, item in enumerate(self.items):
            if item is None:
                continue
            if target is None:
                if isinstance(item.value, float) and item.value == wanted:
                    return index
            elif self.grouping == "years":
                if item.text.isdigit() and int(item.text) == target:
                    return index
            elif item.index == target:
                return index
        return None


@dataclass(frozen=True)
class PivotLine:
    """A row or column of a report's data area: its kind (``data``,
    ``grand``, ``blank`` or a subtotal's function), the item each field of
    its axis stands at, the outer fields first, and the value it shows
    when its axis lists the values."""

    kind: str
    items: tuple[int, ...]
    data: int


@dataclass(frozen=True)
class PivotReport:
    """What a pivot table shows and where: the fields, the lines of its
    data area as the file lists them, its filters, and its values, each
    by its name and the name of the field it summarizes."""

    name: str
    location: RangeRef
    #: The location with the filters above it, which name the report too.
    extent: RangeRef
    first_data_row: int
    first_data_column: int
    fields: tuple[PivotField, ...]
    rows: tuple[int, ...]
    columns: tuple[int, ...]
    row_lines: tuple[PivotLine, ...]
    column_lines: tuple[PivotLine, ...]
    #: Each filter's field, with the item it is set to, or ``None`` when it
    #: shows every item.
    pages: tuple[tuple[int, int | None], ...]
    values: tuple[tuple[str, str], ...]
    #: Whether the report is one this reads: not one from a cube, whose
    #: fields GETPIVOTDATA names another way, and with its lines written.
    readable: bool = True

    def value_index(self, name: str) -> int | None:
        """The value a name picks: its own name, or else the field it
        summarizes, the first such."""
        for index, (own, _) in enumerate(self.values):
            if _collate.equal(own, name):
                return index
        for index, (_, source) in enumerate(self.values):
            if _collate.equal(source, name):
                return index
        return None

    def field_index(self, name: str) -> int | None:
        for index, field in enumerate(self.fields):
            if _collate.equal(field.name, name):
                return index
        return None

    def item_index(self, field: int, wanted: str | float | bool) -> int | None:
        """The item a value picks: text by the name it was given or the
        text it is shown as, a number by its value."""
        found = self.fields[field]
        if isinstance(wanted, float):
            return found.number_index(wanted)
        for index, item in enumerate(found.items):
            if item is None:
                continue
            if isinstance(wanted, str):
                if (item.caption is not None and _collate.equal(item.caption, wanted)) or _collate.equal(item.text, wanted):
                    return index
            elif isinstance(item.value, bool) and item.value == wanted:
                return index
        return None

    def cell(self, value: int, chosen: dict[int, int]) -> CellRef | None:
        """The cell of the data area showing a value for the items chosen,
        a field left out meaning its total, or ``None`` where the report
        shows no such cell. A filter chosen has to be set to that item."""
        remaining = dict(chosen)
        for field, item in self.pages:
            if field in remaining and (item is None or remaining.pop(field) != item):
                return None
        if any(field not in self.rows and field not in self.columns for field in remaining):
            return None
        row = self._line(self.row_lines, self.rows, remaining, value)
        column = self._line(self.column_lines, self.columns, remaining, value)
        if row is None or column is None:
            return None
        top = self.location.top + self.first_data_row
        left = self.location.left + self.first_data_column
        return CellRef(top + row, left + column)

    def _line(self, lines: tuple[PivotLine, ...], axis: tuple[int, ...], chosen: dict[int, int], value: int) -> int | None:
        """The line of an axis that stands at the items chosen of its
        fields and shows the value wanted, when the axis lists the values.

        Measured: the line with no other field at an item, the total of the
        fields left out, is the one; failing that, a line whose other
        fields are at items, if it is the only one. A group's own line,
        which leaves the fields within it out, shows its subtotal only when
        the field has subtotals at the top."""
        values_at = axis.index(VALUES) if VALUES in axis else None
        wanted = {position: chosen[field] for position, field in enumerate(axis) if field in chosen}
        others: list[int] = []
        for index, line in enumerate(lines):
            if line.kind == "blank":
                continue
            assigned = 0 if line.kind == "grand" else len(line.items)
            if values_at is not None:
                if values_at < assigned:
                    shown = line.items[values_at]
                elif line.kind == "data":
                    continue
                else:
                    shown = line.data
                if shown != value:
                    continue
            if any(position >= assigned or line.items[position] != item for position, item in wanted.items()):
                continue
            if line.kind == "data" and assigned < len(axis):
                inner = axis[assigned - 1] if assigned else VALUES
                if inner == VALUES or not (self.fields[inner].subtotals and self.fields[inner].subtotals_top):
                    continue
            if all(position in wanted or position == values_at for position in range(assigned)):
                return index
            others.append(index)
        return others[0] if len(others) == 1 else None


def read_pivot_report(
    package: OpcPackage, part: str, formats: Callable[[int], str], *, epoch_1904: bool = False
) -> PivotReport | None:
    """A pivot table's report, or ``None`` when it has no location to be
    found by. One whose cache or layout this does not follow comes back
    not ``readable``. ``formats`` gives a number format's code by its id,
    for the text an item is shown as."""
    root = package.xml(part).root
    location = root.child("location")
    if location is None:
        return None
    try:
        block = RangeRef.parse(location.get("ref") or "").normalized
    except ValueError:
        return None
    pages_rows = _number(location.get("rowPageCount"))
    top = max(block.top - pages_rows - 1, 1) if pages_rows else block.top
    extent = RangeRef(CellRef(top, block.left), CellRef(block.bottom, block.right))
    unread = PivotReport(decode(root.get("name") or ""), block, extent, 0, 0, (), (), (), (), (), (), (), readable=False)
    cache = next(
        (
            entry.target_part
            for entry in package.relationships(part).by_type(RT_PIVOT_CACHE)
            if not entry.is_external and package.has_part(entry.target_part)
        ),
        None,
    )
    if cache is None:
        return unread
    definition = package.xml(cache).root
    container = definition.child("cacheFields")
    sources = [] if container is None else list(container.children_named("cacheField"))
    elements = _children(root, "pivotFields", "pivotField")
    rows = tuple(_number(entry.get("x")) for entry in _children(root, "rowFields", "field"))
    columns = tuple(_number(entry.get("x")) for entry in _children(root, "colFields", "field"))
    if (
        definition.child("cacheHierarchies") is not None
        or len(elements) != len(sources)
        or (rows and root.child("rowItems") is None)
        or (columns and root.child("colItems") is None)
    ):
        return unread
    fields = tuple(
        _field(element, source, formats, epoch_1904=epoch_1904) for element, source in zip(elements, sources, strict=True)
    )
    values = tuple(
        (decode(entry.get("name") or ""), decode(sources[_number(entry.get("fld"))].get("name") or ""))
        for entry in _children(root, "dataFields", "dataField")
        if 0 <= _number(entry.get("fld")) < len(sources)
    )
    return PivotReport(
        name=unread.name,
        location=block,
        extent=extent,
        first_data_row=_number(location.get("firstDataRow")),
        first_data_column=_number(location.get("firstDataCol")),
        fields=fields,
        rows=rows,
        columns=columns,
        row_lines=_lines(root.child("rowItems")),
        column_lines=_lines(root.child("colItems")),
        pages=tuple(
            (_number(entry.get("fld")), None if entry.get("item") is None else _number(entry.get("item")))
            for entry in _children(root, "pageFields", "pageField")
        ),
        values=values,
    )


def _children(root: Element, container: str, name: str) -> list[Element]:
    found = root.child(container)
    return [] if found is None else list(found.children_named(name))


def _number(raw: str | None) -> int:
    try:
        return int(raw or 0)
    except ValueError:
        return 0


def _lines(container: Element | None) -> tuple[PivotLine, ...]:
    """An axis's lines, each ``<i>`` repeating the first ``r`` items of the
    one before it; a grand total's placeholder item left off."""
    if container is None:
        return ()
    found: list[PivotLine] = []
    previous: list[int] = []
    for line in container.children_named("i"):
        kind = line.get("t") or "data"
        items = previous[: _number(line.get("r"))] + [_number(entry.get("v")) for entry in line.children_named("x")]
        found.append(PivotLine(kind, () if kind == "grand" else tuple(items), _number(line.get("i"))))
        previous = items
    return tuple(found)


def _field(element: Element, source: Element, formats: Callable[[int], str], *, epoch_1904: bool) -> PivotField:
    """A field as the report has it, its items read from the cache: from
    the groups a grouped field makes, or else from the values it holds."""
    group = source.child("fieldGroup")
    listed = None if group is None else group.child("groupItems")
    ranged = None if listed is None or group is None else group.child("rangePr")
    if listed is None:
        listed = source.child("sharedItems")
    cached = [] if listed is None else [_cached(entry, epoch_1904=epoch_1904) for entry in listed.elements()]
    code = formats(_number(element.get("numFmtId") or source.get("numFmtId")))
    items: list[PivotItem | None] = []
    for entry in _children(element, "items", "item"):
        index = _number(entry.get("x"))
        if (entry.get("t") or "data") != "data" or not 0 <= index < len(cached):
            items.append(None)
            continue
        value = cached[index]
        caption = entry.get("n")
        text = _shown(value, code, epoch_1904=epoch_1904)
        items.append(PivotItem(value, index, text, None if caption is None else decode(caption)))
    subtotals = element.get("defaultSubtotal") != "0" or any(element.get(flag) in ("1", "true") for flag in _SUBTOTAL_FLAGS)
    grouping = None if ranged is None else ranged.get("groupBy") or "range"
    return PivotField(
        name=decode(element.get("name") or source.get("name") or ""),
        items=tuple(items),
        subtotals=subtotals,
        subtotals_top=element.get("subtotalTop") not in ("0", "false"),
        grouping=grouping,
        start=0.0 if ranged is None else _decimal(ranged.get("startNum")),
        interval=0.0 if ranged is None else _decimal(ranged.get("groupInterval")),
    )


def _decimal(raw: str | None) -> float:
    try:
        return float(raw or 0)
    except ValueError:
        return 0.0


def _cached(entry: Element, *, epoch_1904: bool) -> Item:
    """A cached value: text, a number, a logical, a date as its serial, an
    error, or ``None`` for a blank."""
    raw = decode(entry.get("v") or "")
    kind = entry.name
    if kind == "n":
        try:
            return float(raw)
        except ValueError:
            return None
    if kind == "b":
        return raw in ("1", "true")
    if kind == "d":
        try:
            return datetime_to_serial(dt.datetime.fromisoformat(raw), epoch_1904=epoch_1904)
        except ValueError:
            return None
    if kind == "e":
        return CellError(raw)
    if kind == "s":
        return raw
    return None


def _shown(value: Item, code: str, *, epoch_1904: bool) -> str:
    """The text a report shows an item as."""
    if value is None:
        return "(blank)"
    if isinstance(value, CellError):
        return value.code
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return format_value(float(value), code, epoch_1904=epoch_1904)
    return value


def _source_of(source: Element | None) -> tuple[str | None, RangeRef | None, str | None]:
    if source is None:
        return None, None, None
    sheet = source.get("sheet")
    raw = source.get("ref")
    reference: RangeRef | None = None
    if raw is not None:
        try:
            reference = RangeRef.parse(raw).normalized
        except ValueError:
            reference = None
    name = source.get("name")
    return (
        None if sheet is None else decode(sheet),
        reference,
        None if name is None else decode(name),
    )


__all__ = [
    "RT_PIVOT_CACHE",
    "RT_PIVOT_RECORDS",
    "RT_PIVOT_TABLE",
    "VALUES",
    "PivotField",
    "PivotItem",
    "PivotLine",
    "PivotReport",
    "PivotTable",
    "cache_parts",
    "read_pivot_report",
    "read_pivot_tables",
    "worksheet_source",
]
