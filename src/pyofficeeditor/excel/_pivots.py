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
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyofficeeditor.excel._reference import RangeRef
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
    "PivotTable",
    "cache_parts",
    "read_pivot_tables",
    "worksheet_source",
]
