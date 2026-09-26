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

from pyofficeeditor.excel import RangeRef, Workbook, Worksheet, column_letter

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
    #: The columns compared, by letter.
    columns: tuple[str, ...]
    header: bool
    #: The error Excel refused it with, or empty.
    refused: str

    def apply(self, sheet: Worksheet) -> int:
        return sheet.remove_duplicates(self.cells, self.columns, header=self.header)


def _removals() -> list[_Removal]:
    if not ANSWERS.exists():
        return []
    answers = json.loads(ANSWERS.read_text(encoding="utf-8"))
    found: list[_Removal] = []
    for name, entry in answers.items():
        sheet, _, cells = str(name).partition("!")
        left = RangeRef.parse(cells).normalized.left
        found.append(
            _Removal(
                sheet,
                cells,
                tuple(column_letter(left + int(offset) - 1) for offset in entry["columns"]),
                bool(entry["header"]),
                str(entry["refused"]),
            )
        )
    return found


REMOVALS = _removals()
SHEETS = sorted({removal.sheet for removal in REMOVALS})


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


@needs_workbook
@pytest.mark.parametrize("removal", [removal for removal in REMOVALS if removal.refused], ids=lambda removal: removal.sheet)
def test_what_excel_refuses_is_refused_and_left_alone(removal: _Removal) -> None:
    book = Workbook.from_bytes(WORKBOOK.read_bytes())
    sheet = book[removal.sheet]
    before = sheet.document.root.to_xml()
    with pytest.raises(ValueError, match="merged cells" if "merged" in removal.refused else "array formula"):
        removal.apply(sheet)
    assert sheet.document.root.to_xml() == before
    assert not book.is_modified


@needs_workbook
def test_the_measured_removals_cover_what_matters() -> None:
    """Refusals, a key over two columns, no header, and pairs that are
    duplicates and pairs that are not."""
    assert {removal.sheet for removal in REMOVALS if removal.refused} == {"ArrayRows", "Merged"}
    assert any(len(removal.columns) > 1 for removal in REMOVALS)
    assert any(not removal.header for removal in REMOVALS)
    assert len([removal for removal in REMOVALS if removal.sheet == "Pairs"]) == 92


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


def test_a_table_is_not_cleaned_yet(sheet: Worksheet) -> None:
    sheet.add_table("Amounts", "A1:B5")
    with pytest.raises(ValueError, match="removing duplicates from a table is not supported yet"):
        sheet.remove_duplicates("A1:B5", header=True)
