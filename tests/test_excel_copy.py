"""Copying cells and pasting them, held to Excel's Copy and Paste.

``copies.xlsx`` is authored by ``scripts/build_excel_fixtures.py``: each
sheet holds a source block. Excel then made every copy
``copies_answers.json`` records, in order, through ``Range.Copy
Destination:=``, and saved the result as ``copies_pasted.xlsx``; the
answers record each copy's source and destination and the error Excel
refused it with. The library makes the same copies of ``copies.xlsx``, and
each sheet has to come out as Excel's did: every cell's formula, value and
format once calculated, the rows and columns, and the notes and their
boxes, links, validation, conditional formats and merged cells.
"""

from __future__ import annotations

import datetime as dt
import json
from dataclasses import dataclass
from pathlib import Path

import pytest
import sheet_state

from pyofficeeditor.excel import RangeRef, Workbook, Worksheet

FIXTURES = Path(__file__).parent / "fixtures" / "excel"
WORKBOOK = FIXTURES / "copies.xlsx"
PASTED = FIXTURES / "copies_pasted.xlsx"
ANSWERS = FIXTURES / "copies_answers.json"
#: The day copies.xlsx was authored on.
AUTHORED = dt.date(2026, 9, 26)

needs_workbook = pytest.mark.skipif(
    not (WORKBOOK.exists() and PASTED.exists() and ANSWERS.exists()),
    reason="run scripts/build_excel_fixtures.py to author copies.xlsx with real Excel",
)


@dataclass(frozen=True)
class _Copy:
    """One copy Excel made, as the answers record it."""

    sheet: str
    source: str
    onto: str
    target: str
    #: The error Excel refused it with, or empty.
    refused: str

    def apply(self, book: Workbook) -> RangeRef:
        return book[self.sheet].copy_range(self.source, self.target, to=book[self.onto])


def _copies() -> list[_Copy]:
    if not ANSWERS.exists():
        return []
    answers = json.loads(ANSWERS.read_text(encoding="utf-8"))
    return [
        _Copy(entry["sheet"], entry["source"], entry["onto"], entry["target"], entry["refused"])
        for _, entry in sorted(answers.items())
    ]


COPIES = _copies()
#: Sheets whose copies the library refuses for now, though Excel makes
#: them, and what it says: cells taken out of a table, which Excel gives
#: the table style's look as formats of their own, and a paste onto cells
#: a formula spilled into, which Excel then shows #SPILL!.
NOT_YET = {
    "PartOut": "table style's look",
    "TableRows": "table style's look",
    "SpillOver": "spilled into",
}
SHEETS = sorted(({copy.sheet for copy in COPIES} | {copy.onto for copy in COPIES}) - set(NOT_YET))


@pytest.fixture(scope="module")
def by_library() -> Workbook:
    """copies.xlsx with every copy Excel made made here too, in the same
    order, and calculated."""
    book = Workbook.from_bytes(WORKBOOK.read_bytes())
    for copy in COPIES:
        if not copy.refused and copy.sheet not in NOT_YET:
            copy.apply(book)
    book.calculate(today=AUTHORED)
    return book


@pytest.fixture(scope="module")
def by_excel() -> Workbook:
    return Workbook.from_bytes(PASTED.read_bytes())


@needs_workbook
@pytest.mark.parametrize("name", SHEETS)
def test_a_paste_leaves_the_sheet_as_excel_leaves_it(name: str, by_library: Workbook, by_excel: Workbook) -> None:
    mine, excel = by_library[name], by_excel[name]
    assert sheet_state.cells(mine) == sheet_state.cells(excel)
    assert sheet_state.rows(mine) == sheet_state.rows(excel)
    assert sheet_state.columns(mine) == sheet_state.columns(excel)
    assert sheet_state.attached(mine) == sheet_state.attached(excel)


#: What the library says when it refuses, by what Excel said.
_REFUSALS = {
    "We can't do that to a merged cell.": "merged cells",
    "You can't change part of an array.": "part of the array formula",
}


@needs_workbook
@pytest.mark.parametrize("copy", [copy for copy in COPIES if copy.refused], ids=lambda copy: copy.sheet)
def test_what_excel_refuses_is_refused_and_left_alone(copy: _Copy) -> None:
    book = Workbook.from_bytes(WORKBOOK.read_bytes())
    target = book[copy.onto]
    before = target.document.root.to_xml()
    with pytest.raises(ValueError, match=_REFUSALS[copy.refused]):
        copy.apply(book)
    assert target.document.root.to_xml() == before
    assert not book.is_modified


@needs_workbook
@pytest.mark.parametrize("copy", [copy for copy in COPIES if copy.sheet in NOT_YET], ids=lambda copy: copy.sheet)
def test_what_is_not_supported_yet_is_refused(copy: _Copy) -> None:
    book = Workbook.from_bytes(WORKBOOK.read_bytes())
    before = book[copy.onto].document.root.to_xml()
    with pytest.raises(ValueError, match=NOT_YET[copy.sheet]):
        copy.apply(book)
    assert book[copy.onto].document.root.to_xml() == before


@needs_workbook
def test_the_measured_copies_cover_what_matters() -> None:
    """Refusals, copies to another sheet, repeated copies, and whole rows
    and columns."""
    assert len(COPIES) == 44
    assert {copy.sheet for copy in COPIES if copy.refused} == {"ArrayCut", "MergedPart"}
    assert any(copy.onto != copy.sheet for copy in COPIES)
    assert any(":" in copy.target for copy in COPIES)
    assert {copy.source for copy in COPIES} >= {"1:3", "A:A"}


# ----------------------------------------------------------------------
# The method's own contract
# ----------------------------------------------------------------------


@pytest.fixture
def sheet(live_empty_xlsx: Path) -> Worksheet:
    book = Workbook.from_bytes(live_empty_xlsx.read_bytes())
    data = book.sheets[0]
    for row, (name, amount) in enumerate([("n", "amount"), ("a", 1), ("b", 2), ("c", 3)], start=1):
        data[f"A{row}"] = name
        data[f"B{row}"] = amount
    return data


def test_a_formula_moves_as_a_copy_does(sheet: Worksheet) -> None:
    sheet["C2"].formula = "=B2*2+$B$2"
    assert sheet.copy_range("C2", "C3:C4").a1 == "C3:C4"
    assert [sheet[f"C{row}"].formula for row in range(2, 5)] == ["B2*2+$B$2", "B3*2+$B$2", "B4*2+$B$2"]


def test_a_thread_comes_with_its_replies(sheet: Worksheet) -> None:
    """Measured in Excel: a thread is copied whole, its replies with it.
    No fixture holds one, since a workbook Excel wrote one in carries who
    wrote it."""
    when = dt.datetime(2026, 9, 26, 12, 0)
    sheet.add_threaded_comment("A2", "first", author="Ann", when=when)
    sheet.add_threaded_reply("A2", "second", author="Bo", when=when)
    sheet.copy_range("A2", "D5")
    copied = sheet.threaded_comment("D5")
    assert copied is not None
    assert (copied.text, copied.author, [reply.text for reply in copied.replies]) == ("first", "Ann", ["second"])
    assert sheet.threaded_comment("A2") is not None


def test_to_another_sheet(sheet: Worksheet) -> None:
    other = sheet.workbook.add_sheet("Other")
    sheet["C2"].formula = "=B2*2"
    sheet.copy_range("B2:C2", "E1", to=other)
    assert (other["E1"].value, other["F1"].formula) == (1, "E1*2")


def test_not_to_another_workbook(sheet: Worksheet, live_empty_xlsx: Path) -> None:
    elsewhere = Workbook.from_bytes(live_empty_xlsx.read_bytes()).sheets[0]
    with pytest.raises(ValueError, match="another workbook"):
        sheet.copy_range("A1:B2", "A1", to=elsewhere)


@pytest.mark.parametrize(
    ("setup", "cells", "destination", "message"),
    [
        ("table", "A1:B4", "D1", "all of the table"),
        ("table", "A2", "B1", "header"),
        ("table", "A2:B2", "A5", "take it in"),
        ("table", "A2:A4", "C1", "take it in"),
        ("table", "A2:B3", "D2", "out of it"),
        ("shape", "A1:B4", "D1", "shape"),
        ("", "1:2", "B5", "column A"),
        ("", "A:A", "C2", "row 1"),
        ("", "A1:B4", "XFD1", "does not fit"),
    ],
)
def test_a_copy_refused_here_changes_nothing(
    sheet: Worksheet, setup: str, cells: str, destination: str, message: str
) -> None:
    if setup == "table":
        sheet.add_table("Amounts", "A1:B4")
    if setup == "shape":
        sheet.add_shape("Box", left=1, top=20, width=20, height=10)
    before = sheet.document.root.to_xml()
    with pytest.raises(ValueError, match=message):
        sheet.copy_range(cells, destination)
    assert sheet.document.root.to_xml() == before
