"""A sheet's state as the tests that hold an edit to a file Excel saved
after the same edit compare it: what each cell holds, each row's own
settings, and what hangs on the cells.

Attributes Excel works out again whenever it saves are left out, and a
row with nothing of its own is left out as Excel leaves it out of the
file, so what is compared is what the edit decides.
"""

from __future__ import annotations

import re

from pyofficeeditor.excel import CellFormat, CellValue, Table, Worksheet
from pyofficeeditor.excel._rowcol import RT_VML, related_parts

#: Row attributes Excel works out again whenever it saves, as hints for
#: drawing: which columns hold cells, and the descent of the row's font.
ROW_HINTS = frozenset({"r", "spans", "x14ac:dyDescent"})


def cells(sheet: Worksheet) -> dict[str, tuple[str | None, CellValue, CellFormat]]:
    """Each cell's formula, value and format, by address. The format is
    what the cell's style index resolves to rather than the index, which
    Excel renumbers as it saves when a style has fallen out of use."""
    return {
        reference.a1: (sheet.get_formula(reference), sheet.get_value(reference), sheet.get_format(reference))
        for reference, _ in sheet.cell_elements()
    }


def rows(sheet: Worksheet) -> dict[int, dict[str, str]]:
    """Each row's own attributes: its height, style, and whether it is
    hidden. A row with no cell and nothing of its own is left out, as Excel
    leaves it out of the file."""
    found: dict[int, dict[str, str]] = {}
    for number, row in sheet.rows_by_number().items():
        own = {name: value for name, value in row.attributes.items() if name not in ROW_HINTS}
        if own or next(row.children_named("c"), None) is not None:
            found[number] = own
    return found


def columns(sheet: Worksheet) -> dict[int, tuple[float | None, bool]]:
    """Each column the sheet sets a width or hides, with its width and
    whether it is hidden."""
    container = sheet.document.root.child("cols")
    numbers: set[int] = set()
    for entry in () if container is None else container.children_named("col"):
        numbers.update(range(int(entry.get("min") or 1), int(entry.get("max") or 1) + 1))
    return {
        number: (sheet.column_width(number), sheet.column_hidden(number))
        for number in sorted(numbers)
        if sheet.column_width(number) is not None or sheet.column_hidden(number)
    }


def attached(sheet: Worksheet) -> dict[str, object]:
    """Notes and their boxes, links, validation and conditional formats
    with their formulas, merged cells, the filter and the tables, each as
    its ranges."""
    package = sheet.workbook.package
    boxes = [
        re.findall(r"<x:(Row|Column|Anchor)>([^<]*)</x:", package.read(name).decode("utf-8"))
        for name in related_parts(sheet, RT_VML)
    ]
    filtered = sheet.document.root.child("autoFilter")
    return {
        "notes": {note.ref: note.text for note in sheet.comments},
        "note boxes": boxes,
        "links": {link.ref.a1: link.target for link in sheet.hyperlinks},
        "validations": [
            ([block.a1 for block in rule.ranges], rule.formula1, rule.formula2) for rule in sheet.data_validations
        ],
        "conditional formats": [
            ([area.a1 for area in formatting.ranges], [list(rule.formulas) for rule in formatting.rules])
            for formatting in sheet.conditional_formats
        ],
        "merged": [block.a1 for block in sheet.merged_ranges],
        "filter": None if filtered is None else filtered.to_xml(),
        "tables": {table.name: table_state(table) for table in sheet.tables},
    }


def table_state(table: Table) -> tuple[str, str | None, str | None]:
    """Where a table is, its filter and the sort it records."""
    root = table.document.root
    filtered, state = root.child("autoFilter"), root.child("sortState")
    return (
        table.ref.a1,
        None if filtered is None else filtered.to_xml(),
        None if state is None else state.to_xml(),
    )


def sort_state(sheet: Worksheet) -> str | None:
    """The sort the sheet records, as markup."""
    state = sheet.document.root.child("sortState")
    return None if state is None else state.to_xml()
