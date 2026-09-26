"""Sorting a range's rows, held to Excel's Sort.

``sorts.xlsx`` is authored by ``scripts/build_excel_fixtures.py``, one sort
to a sheet and saved unsorted. Excel then sorted every sheet through its
Sort object and saved the result as ``sorts_sorted.xlsx``, and
``sorts_answers.json`` records each sort and the error Excel refused it
with. The library sorts ``sorts.xlsx`` the same way, and each sheet has to
come out as Excel's did: every cell's formula, value and style, the rows,
notes, links, validation, conditional formats, the filter and the sort
recorded. Values are compared once the library has calculated, since a
sort changes what formulas read.
"""

from __future__ import annotations

import datetime as dt
import json
import re
from dataclasses import dataclass
from pathlib import Path

import pytest

from pyofficeeditor.excel import CellValue, SortKey, Workbook, Worksheet
from pyofficeeditor.excel._formulas import sorted_formula
from pyofficeeditor.excel._rowcol import RT_VML, related_parts

FIXTURES = Path(__file__).parent / "fixtures" / "excel"
WORKBOOK = FIXTURES / "sorts.xlsx"
SORTED = FIXTURES / "sorts_sorted.xlsx"
ANSWERS = FIXTURES / "sorts_answers.json"
#: The day sorts.xlsx was authored on.
AUTHORED = dt.date(2026, 9, 26)

needs_workbook = pytest.mark.skipif(
    not (WORKBOOK.exists() and SORTED.exists() and ANSWERS.exists()),
    reason="run scripts/build_excel_fixtures.py to author sorts.xlsx with real Excel",
)


@dataclass(frozen=True)
class _Sort:
    """One sort Excel made, as the answers record it."""

    cells: str
    keys: tuple[SortKey, ...]
    header: bool
    match_case: bool
    #: The error Excel refused the sort with, or empty.
    refused: str

    def apply(self, sheet: Worksheet) -> None:
        sheet.sort(self.cells, self.keys, header=self.header, match_case=self.match_case)


def _sorts() -> dict[str, _Sort]:
    if not ANSWERS.exists():
        return {}
    answers = json.loads(ANSWERS.read_text(encoding="utf-8"))
    return {
        name: _Sort(
            str(entry["range"]),
            tuple(SortKey(str(column), bool(descending)) for column, descending in entry["keys"]),
            bool(entry["header"]),
            bool(entry["matchCase"]),
            str(entry["refused"]),
        )
        for name, entry in answers.items()
    }


SORTS = _sorts()


@pytest.fixture(scope="module")
def by_library() -> Workbook:
    """sorts.xlsx sorted here, every sort Excel made, and calculated."""
    book = Workbook.from_bytes(WORKBOOK.read_bytes())
    for name, sort in SORTS.items():
        if not sort.refused:
            sort.apply(book[name])
    book.calculate(today=AUTHORED)
    return book


@pytest.fixture(scope="module")
def by_excel() -> Workbook:
    return Workbook.from_bytes(SORTED.read_bytes())


#: Row attributes Excel works out again whenever it saves, as hints for
#: drawing: which columns hold cells, and the descent of the row's font.
_ROW_HINTS = frozenset({"r", "spans", "x14ac:dyDescent"})


def _cells(sheet: Worksheet) -> dict[str, tuple[str | None, CellValue, str | None]]:
    return {
        reference.a1: (sheet.get_formula(reference), sheet.get_value(reference), element.get("s"))
        for reference, element in sheet.cell_elements()
    }


def _rows(sheet: Worksheet) -> dict[int, dict[str, str]]:
    """Each row's own attributes: its height, style, and whether it is
    hidden. A row with no cell and nothing of its own is left out, as Excel
    leaves it out of the file."""
    found: dict[int, dict[str, str]] = {}
    for number, row in sheet.rows_by_number().items():
        own = {name: value for name, value in row.attributes.items() if name not in _ROW_HINTS}
        if own or next(row.children_named("c"), None) is not None:
            found[number] = own
    return found


def _attached(sheet: Worksheet) -> dict[str, object]:
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
        "validations": [[block.a1 for block in rule.ranges] for rule in sheet.data_validations],
        "conditional formats": [[block.a1 for block in rule.ranges] for rule in sheet.conditional_formats],
        "merged": [block.a1 for block in sheet.merged_ranges],
        "filter": None if filtered is None else filtered.to_xml(),
    }


def _sort_state(sheet: Worksheet) -> str | None:
    state = sheet.document.root.child("sortState")
    return None if state is None else state.to_xml()


@needs_workbook
@pytest.mark.parametrize("name", sorted(SORTS))
def test_a_sort_leaves_the_sheet_as_excel_leaves_it(name: str, by_library: Workbook, by_excel: Workbook) -> None:
    mine, excel = by_library[name], by_excel[name]
    assert _cells(mine) == _cells(excel)
    assert _rows(mine) == _rows(excel)
    assert _attached(mine) == _attached(excel)
    # Excel keeps the settings of a sort it refused, as the dialog left
    # them; the library changes nothing when it refuses one.
    if not SORTS[name].refused:
        assert _sort_state(mine) == _sort_state(excel)


@needs_workbook
@pytest.mark.parametrize("name", sorted(name for name, sort in SORTS.items() if sort.refused))
def test_what_excel_refuses_to_sort_is_refused_and_left_alone(name: str) -> None:
    sort = SORTS[name]
    book = Workbook.from_bytes(WORKBOOK.read_bytes())
    sheet = book[name]
    before = sheet.document.root.to_xml()
    with pytest.raises(ValueError, match="merged cells" if "merged" in sort.refused else "array formula"):
        sort.apply(sheet)
    assert sheet.document.root.to_xml() == before
    assert not book.is_modified


@needs_workbook
def test_the_measured_sorts_cover_what_the_sort_has_to_get_right() -> None:
    """Refusals, both directions, several keys, case, and no header."""
    assert {name for name, sort in SORTS.items() if sort.refused} == {"Array", "Merged", "Spill"}
    assert any(key.descending for sort in SORTS.values() for key in sort.keys)
    assert any(len(sort.keys) > 1 for sort in SORTS.values())
    assert any(sort.match_case for sort in SORTS.values())
    assert any(not sort.header for sort in SORTS.values())


#: Measured through ``Range.Formula`` once Excel's Sort had moved each
#: formula's row, on sheet S beside sheets O and P and a name ``nm``: the
#: formula, how many rows down it went, and what Excel then wrote.
MOVED_FORMULAS = [
    ("B3", -1, "B2"),
    ("B2", 3, "B5"),
    ("B1", -1, "#REF!"),
    ("B1", 3, "B4"),
    ("OFFSET(B3,0,0)", -1, "OFFSET(B2,0,0)"),
    ("ROW(B2)", 3, "ROW(B5)"),
    ('INDIRECT("B3")', -1, 'INDIRECT("B3")'),
    ("nm", -1, "nm"),
    # A reference naming a sheet stays as written, its own sheet included.
    ("O!B3", -1, "O!B3"),
    ("O!B1", -1, "O!B1"),
    ("O!$B3", -1, "O!$B3"),
    ("O!B$2", 3, "O!B$2"),
    ("SUM(O!B3:B4)", -1, "SUM(O!B3:B4)"),
    ("SUM(O:P!B3)", -1, "SUM(O:P!B3)"),
    ("SUM(O!3:3)", -1, "SUM(O!3:3)"),
    ("SUM(O!B:B)", 3, "SUM(O!B:B)"),
    ("B2+O!B2", 3, "B5+O!B2"),
    ("S!B3", -1, "S!B3"),
    ("S!B1", 3, "S!B1"),
    ("SUM(S!B3,B3)", -1, "SUM(S!B3,B2)"),
    ("SUM(S!2:2)", 3, "SUM(S!2:2)"),
    ("S!$B$3", -1, "S!$B$3"),
    # From sorts.xlsx: a range that turns over keeps its markers in place.
    ("SUM(B$4:B5)+SUM($2:3)", -2, "SUM(B$3:B4)+SUM($1:2)"),
    ("B1+B2", -2, "#REF!+#REF!"),
    ("SUM(B$3:B6)+SUM($4:5)", -1, "SUM(B$3:B5)+SUM($4:4)"),
]


@pytest.mark.parametrize(("formula", "rows", "expected"), MOVED_FORMULAS)
def test_a_sorted_formula_moves_as_excel_moves_it(formula: str, rows: int, expected: str) -> None:
    assert sorted_formula(formula, rows) == expected


# ----------------------------------------------------------------------
# The method's own contract
# ----------------------------------------------------------------------


@pytest.fixture
def sheet(live_empty_xlsx: Path) -> Worksheet:
    book = Workbook.from_bytes(live_empty_xlsx.read_bytes())
    data = book.sheets[0]
    for row, (name, amount) in enumerate([("n", "amount"), ("b", 2), ("c", 3), ("a", 2)], start=1):
        data[f"A{row}"] = name
        data[f"B{row}"] = amount
    return data


def test_keys_are_letters_or_sort_keys(sheet: Worksheet) -> None:
    sheet.sort("A1:B4", [SortKey("B", descending=True), "A"], header=True)
    assert [sheet[f"A{row}"].value for row in range(2, 5)] == ["c", "a", "b"]
    sheet.sort("A1:B4", "A", header=True)
    assert [sheet[f"A{row}"].value for row in range(2, 5)] == ["a", "b", "c"]


def test_without_a_header_the_first_row_is_sorted_too(sheet: Worksheet) -> None:
    sheet.sort("A1:B4", "A")
    assert [sheet[f"A{row}"].value for row in range(1, 5)] == ["a", "b", "c", "n"]


def test_a_sort_is_saved(sheet: Worksheet) -> None:
    sheet.sort("A1:B4", SortKey("A", descending=True), header=True, match_case=True)
    reopened = Workbook.from_bytes(sheet.workbook.to_bytes()).sheets[0]
    assert [reopened[f"A{row}"].value for row in range(2, 5)] == ["c", "b", "a"]
    assert _sort_state(reopened) == (
        '<sortState caseSensitive="1" ref="A2:B4" '
        'xmlns:xlrd2="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2">'
        '<sortCondition descending="1" ref="A2:A4"/></sortState>'
    )
    assert sheet.workbook.values_changed


@pytest.mark.parametrize(
    ("cells", "by", "header", "message"),
    [
        ("A1:B4", [], False, "1 to 64 keys"),
        ("A1:B4", ["A"] * 65, False, "1 to 64 keys"),
        ("A1:B4", "C", False, "not a column of A1:B4"),
        ("A1:B1", "A", True, "nothing below its header"),
    ],
)
def test_a_sort_needs_keys_in_its_range_and_rows_to_sort(
    sheet: Worksheet, cells: str, by: str | list[str], header: bool, message: str
) -> None:
    before = sheet.document.root.to_xml()
    with pytest.raises(ValueError, match=message):
        sheet.sort(cells, by, header=header)
    assert sheet.document.root.to_xml() == before


def test_a_table_is_not_sorted_yet(sheet: Worksheet) -> None:
    sheet.add_table("Amounts", "A1:B4")
    with pytest.raises(ValueError, match="sorting a table is not supported yet"):
        sheet.sort("A1:B4", "B", header=True)
