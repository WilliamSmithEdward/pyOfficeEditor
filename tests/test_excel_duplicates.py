"""Removing duplicate rows, held to Excel's Remove Duplicates.

``duplicates.xlsx`` is authored by ``scripts/build_excel_fixtures.py``:
sheet Pairs holds 92 pairs of values, each in a range of its own with a
marker beside each value, so the second marker goes exactly when Excel
takes the pair for duplicates; each other sheet holds one more thing
measured. Excel then removed every range's duplicates and saved the
result as ``duplicates_removed.xlsx``, and ``duplicates_answers.json``
records each range and the error Excel refused it with. The library
removes the same duplicates from ``duplicates.xlsx``, and each sheet has to
come out as Excel's did: every cell's formula, value and style, the rows,
notes, links, validation and conditional formats.
"""

from __future__ import annotations

import datetime as dt
import json
from dataclasses import dataclass
from pathlib import Path

import pytest
import sheet_state

from pyofficeeditor.excel import FilterColumn, RangeRef, Workbook, Worksheet, column_letter, criteria

FIXTURES = Path(__file__).parent / "fixtures" / "excel"
WORKBOOK = FIXTURES / "duplicates.xlsx"
REMOVED = FIXTURES / "duplicates_removed.xlsx"
ANSWERS = FIXTURES / "duplicates_answers.json"
#: The day duplicates.xlsx was authored on.
AUTHORED = dt.date(2026, 9, 26)

needs_workbook = pytest.mark.skipif(
    not (WORKBOOK.exists() and REMOVED.exists() and ANSWERS.exists()),
    reason="run scripts/build_excel_fixtures.py to author duplicates.xlsx with real Excel",
)


@dataclass(frozen=True)
class _Removal:
    """One Remove Duplicates Excel ran, as the answers record it."""

    sheet: str
    cells: str
    #: The columns compared, as Excel was given them: counting from 1.
    offsets: tuple[int, ...]
    header: bool
    #: The error Excel refused it with, or empty.
    refused: str

    def columns(self, sheet: Worksheet) -> list[str]:
        """The columns compared, by letter. Measured, Excel counts from the
        range's first column, or from its table's when it is in one."""
        block = RangeRef.parse(self.cells).normalized
        left = next((table.ref.left for table in sheet.tables if table.ref.intersects(block)), block.left)
        return [column_letter(left + offset - 1) for offset in self.offsets]

    def apply(self, sheet: Worksheet) -> int:
        return sheet.remove_duplicates(self.cells, self.columns(sheet), header=self.header)


def _removals() -> list[_Removal]:
    if not ANSWERS.exists():
        return []
    answers = json.loads(ANSWERS.read_text(encoding="utf-8"))
    found: list[_Removal] = []
    for name, entry in answers.items():
        sheet, _, cells = str(name).partition("!")
        found.append(
            _Removal(
                sheet,
                cells,
                tuple(int(offset) for offset in entry["columns"]),
                bool(entry["header"]),
                str(entry["refused"]),
            )
        )
    return found


REMOVALS = _removals()
#: Measured: Excel refuses a table with a totals row while its filter hides
#: rows by a criterion, but only once it has moved the rows it keeps up and
#: cleared the rest, leaving the table as it was. The library refuses
#: before it changes anything, so these sheets are not held to Excel's.
REFUSED_PART_WAY = {"TableFilteredTotals"}
SHEETS = sorted({removal.sheet for removal in REMOVALS} - REFUSED_PART_WAY)


@pytest.fixture(scope="module")
def by_library() -> Workbook:
    """duplicates.xlsx with every range Excel cleaned cleaned here too, and
    calculated."""
    book = Workbook.from_bytes(WORKBOOK.read_bytes())
    for removal in REMOVALS:
        if not removal.refused:
            removal.apply(book[removal.sheet])
    book.calculate(today=AUTHORED)
    return book


@pytest.fixture(scope="module")
def by_excel() -> Workbook:
    return Workbook.from_bytes(REMOVED.read_bytes())


@needs_workbook
@pytest.mark.parametrize("name", SHEETS)
def test_duplicates_go_as_excel_removes_them(name: str, by_library: Workbook, by_excel: Workbook) -> None:
    mine, excel = by_library[name], by_excel[name]
    assert sheet_state.cells(mine) == sheet_state.cells(excel)
    assert sheet_state.rows(mine) == sheet_state.rows(excel)
    assert sheet_state.attached(mine) == sheet_state.attached(excel)


@needs_workbook
def test_each_pair_is_judged_as_excel_judges_it(by_library: Workbook, by_excel: Workbook) -> None:
    """The pairs one at a time, for a failure that names the pair: the
    second marker stays exactly when Excel kept the second value."""
    mine, excel = by_library["Pairs"], by_excel["Pairs"]
    differ = [
        removal.cells
        for removal in REMOVALS
        if removal.sheet == "Pairs"
        and (mine[f"B{RangeRef.parse(removal.cells).bottom}"].value is None)
        != (excel[f"B{RangeRef.parse(removal.cells).bottom}"].value is None)
    ]
    assert differ == []


#: What the library says when it refuses, by what Excel said.
_REFUSALS = {
    "You can't change part of an array.": "array formula",
    "To do this, all the merged cells need to be the same size.": "merged cells",
    "Application-defined or object-defined error": "runs past the table",
    "You can't rearrange cells within a table this way, because it might affect other table cells in an "
    "unexpected way.": "totals row and its filter",
}


@needs_workbook
@pytest.mark.parametrize("removal", [removal for removal in REMOVALS if removal.refused], ids=lambda removal: removal.sheet)
def test_what_excel_refuses_is_refused_and_left_alone(removal: _Removal) -> None:
    book = Workbook.from_bytes(WORKBOOK.read_bytes())
    sheet = book[removal.sheet]
    before = sheet.document.root.to_xml()
    tables = [table.document.root.to_xml() for table in sheet.tables]
    with pytest.raises(ValueError, match=_REFUSALS[removal.refused]):
        removal.apply(sheet)
    assert sheet.document.root.to_xml() == before
    assert [table.document.root.to_xml() for table in sheet.tables] == tables
    assert not book.is_modified


@needs_workbook
def test_references_into_a_table_follow_it_as_excel_moves_them(by_library: Workbook, by_excel: Workbook) -> None:
    """Formulas on another sheet, defined names and charts read the tables
    as Excel left them: a range shaped like a part of a table shrinks with
    it, and a reference to its totals row follows the row up."""
    assert sheet_state.cells(by_library["Reader"]) == sheet_state.cells(by_excel["Reader"])
    names = {name.name: name.refers_to for name in by_excel.defined_names if not name.is_builtin}
    assert {name.name: name.refers_to for name in by_library.defined_names if not name.is_builtin} == names
    assert names == {
        "TableGone": "Table!$B$7",
        "TableSpan": "Table!$B$2:$B$5",
        "TotalsCell": "TableTotals!$B$6",
        "TotalsSpan": "TableTotals!$B$2:$B$5",
    }
    for name in ("Table", "TableTotals"):
        assert [chart.references for chart in by_library[name].charts] == [
            chart.references for chart in by_excel[name].charts
        ]


@needs_workbook
def test_excel_refuses_a_filtered_table_with_a_totals_row_part_way(by_excel: Workbook) -> None:
    """What Excel leaves behind when it refuses: the rows kept moved up and
    the rest cleared, the rows its filter hid still hidden, and the table,
    its totals row and its filter as they were."""
    before = Workbook.from_bytes(WORKBOOK.read_bytes())["TableFilteredTotals"]
    after = by_excel["TableFilteredTotals"]
    assert [after[f"A{row}"].value for row in range(2, 10)] == ["a", "b", "c", "d", None, None, None, "Total"]
    assert [after.row_hidden(row) for row in range(2, 10)] == [before.row_hidden(row) for row in range(2, 10)]
    assert sheet_state.attached(after)["tables"] == sheet_state.attached(before)["tables"]


@needs_workbook
def test_the_measured_removals_cover_what_matters() -> None:
    """Refusals, a key over two columns, no header, pairs that are
    duplicates and pairs that are not, and tables."""
    assert {removal.sheet for removal in REMOVALS if removal.refused} == {
        "ArrayRows",
        "Merged",
        "TableFilteredTotals",
        "TableLonger",
        "TableWider",
    }
    assert any(len(removal.offsets) > 1 for removal in REMOVALS)
    assert any(not removal.header for removal in REMOVALS)
    assert len([removal for removal in REMOVALS if removal.sheet == "Pairs"]) == 92
    assert len([removal for removal in REMOVALS if removal.sheet.startswith("Table")]) == 18


# ----------------------------------------------------------------------
# The method's own contract
# ----------------------------------------------------------------------


@pytest.fixture
def sheet(live_empty_xlsx: Path) -> Worksheet:
    book = Workbook.from_bytes(live_empty_xlsx.read_bytes())
    data = book.sheets[0]
    for row, (name, amount) in enumerate([("n", "amount"), ("b", 2), ("B", 3), ("a", 2), ("b", 2)], start=1):
        data[f"A{row}"] = name
        data[f"B{row}"] = amount
    return data


def test_every_column_is_compared_unless_named(sheet: Worksheet) -> None:
    assert sheet.remove_duplicates("A1:B5", header=True) == 1
    assert [(sheet[f"A{row}"].value, sheet[f"B{row}"].value) for row in range(2, 6)] == [
        ("b", 2),
        ("B", 3),
        ("a", 2),
        (None, None),
    ]
    assert sheet.workbook.values_changed


def test_named_columns_alone_decide(sheet: Worksheet) -> None:
    assert sheet.remove_duplicates("A1:B5", "A", header=True) == 2
    assert [sheet[f"A{row}"].value for row in range(2, 6)] == ["b", "a", None, None]
    assert [sheet[f"B{row}"].value for row in range(2, 6)] == [2, 2, None, None]


def test_nothing_to_remove_changes_nothing(sheet: Worksheet) -> None:
    before = sheet.document.root.to_xml()
    assert sheet.remove_duplicates("B2:B3") == 0
    assert sheet.document.root.to_xml() == before


@pytest.mark.parametrize(
    ("cells", "columns", "header", "message"),
    [
        ("A1:B5", [], False, "at least one column"),
        ("A1:B5", "C", False, "not a column of A1:B5"),
        ("A1:B1", None, True, "nothing below its header"),
    ],
)
def test_a_removal_needs_columns_in_its_range_and_rows(
    sheet: Worksheet, cells: str, columns: str | list[str] | None, header: bool, message: str
) -> None:
    before = sheet.document.root.to_xml()
    with pytest.raises(ValueError, match=message):
        sheet.remove_duplicates(cells, columns, header=header)
    assert sheet.document.root.to_xml() == before


def test_a_range_in_a_table_stands_for_all_of_it(sheet: Worksheet) -> None:
    """Measured: Excel takes the whole table whatever part of it the range
    is, and the table gives up the row it no longer needs."""
    table = sheet.add_table("Amounts", "A1:B5")
    assert sheet.remove_duplicates("B3") == 1
    assert [(sheet[f"A{row}"].value, sheet[f"B{row}"].value) for row in range(2, 6)] == [
        ("b", 2),
        ("B", 3),
        ("a", 2),
        (None, None),
    ]
    assert table.ref.a1 == "A1:B4"


def test_a_tables_columns_are_named_from_the_table(sheet: Worksheet) -> None:
    sheet.add_table("Amounts", "A1:B5")
    assert sheet.remove_duplicates("B2:B5", "A") == 2
    assert [sheet[f"A{row}"].value for row in range(2, 6)] == ["b", "a", None, None]


@pytest.mark.parametrize(
    ("cells", "columns", "message"),
    [
        ("A1:C5", None, "runs past the table"),
        ("A2:B5", "C", "not a column of A1:B5"),
    ],
)
def test_a_removal_from_a_table_stays_in_it(sheet: Worksheet, cells: str, columns: str | None, message: str) -> None:
    table = sheet.add_table("Amounts", "A1:B5")
    before = sheet.document.root.to_xml(), table.document.root.to_xml()
    with pytest.raises(ValueError, match=message):
        sheet.remove_duplicates(cells, columns)
    assert (sheet.document.root.to_xml(), table.document.root.to_xml()) == before


def test_a_totals_row_moves_up_and_what_reads_it_follows(sheet: Worksheet) -> None:
    sheet["A6"] = "Total"
    sheet["B6"].formula = "=SUBTOTAL(109,Amounts[amount])"
    sheet["D1"].formula = "=B6"
    table = sheet.add_table("Amounts", "A1:B6", totals_row=True)
    assert sheet.remove_duplicates("A1:B5") == 1
    assert (sheet["A5"].value, sheet["B5"].formula) == ("Total", "SUBTOTAL(109,Amounts[amount])")
    assert sheet["A6"].value is None
    assert sheet["D1"].formula == "B5"
    assert table.ref.a1 == "A1:B5"


def test_a_tables_filter_is_applied_again(sheet: Worksheet) -> None:
    """Measured: the rows the filter hid are shown, then hidden again as
    they now fall, and those the table gave up stay shown."""
    sheet.add_table("Amounts", "A1:B5")
    sheet.set_table_filter("Amounts", [FilterColumn(1, criteria(">2"))])
    assert [sheet.row_hidden(row) for row in range(2, 6)] == [True, False, True, True]
    assert sheet.remove_duplicates("A1:B5", "A") == 2
    assert [sheet.row_hidden(row) for row in range(2, 6)] == [True, True, False, False]


def test_a_filtered_table_with_a_totals_row_is_refused(sheet: Worksheet) -> None:
    sheet["A6"] = "Total"
    table = sheet.add_table("Amounts", "A1:B6", totals_row=True)
    sheet.set_table_filter("Amounts", [FilterColumn(1, criteria(">1"))])
    before = sheet.document.root.to_xml(), table.document.root.to_xml()
    with pytest.raises(ValueError, match="totals row and its filter"):
        sheet.remove_duplicates("A1:B5")
    assert (sheet.document.root.to_xml(), table.document.root.to_xml()) == before
