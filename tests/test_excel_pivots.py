"""Pivot tables, and what an edit to their sheets does to them.

``pivots.xlsx`` was built by Excel with two pivot tables reading
``Data!A1:C11``, each through a cache of its own: Summary on a sheet of its
own, and Beside on the data sheet at F3:G6, where the edits reach it.
``pivots_answers.json`` records, for the file as built and after each of
nineteen edits Excel made, where each pivot table was, what its cache read
and how many caches the workbook kept, all read from the file Excel saved;
or that Excel refused the edit. Each edit here is held to Excel's.
"""

from __future__ import annotations

import json
import shutil
from collections.abc import Callable
from pathlib import Path

import pytest

from pyofficeeditor.excel import CellRef, PivotTable, RangeRef, Workbook
from pyofficeeditor.excel._pivots import cache_parts

Answers = dict[str, dict[str, object]]

#: Each edit Excel made, made the same way here.
EDITS: dict[str, Callable[[Workbook], object]] = {
    "insert row 1": lambda book: book["Data"].insert_rows(1, 1),
    "insert row 3": lambda book: book["Data"].insert_rows(3, 1),
    "insert row 7": lambda book: book["Data"].insert_rows(7, 1),
    "insert rows 5:6": lambda book: book["Data"].insert_rows(5, 2),
    "delete row 3": lambda book: book["Data"].delete_rows(3, 1),
    "delete row 4": lambda book: book["Data"].delete_rows(4, 1),
    "delete rows 3:6": lambda book: book["Data"].delete_rows(3, 4),
    "delete rows 2:7": lambda book: book["Data"].delete_rows(2, 6),
    "delete rows 1:11": lambda book: book["Data"].delete_rows(1, 11),
    "delete rows 1:2": lambda book: book["Data"].delete_rows(1, 2),
    "delete row 7": lambda book: book["Data"].delete_rows(7, 1),
    "insert column A": lambda book: book["Data"].insert_columns(1, 1),
    "insert column F": lambda book: book["Data"].insert_columns(6, 1),
    "insert column H": lambda book: book["Data"].insert_columns(8, 1),
    "delete column B": lambda book: book["Data"].delete_columns(2, 1),
    "delete column F": lambda book: book["Data"].delete_columns(6, 1),
    "delete columns E:H": lambda book: book["Data"].delete_columns(5, 4),
    "insert row 1 on Pivot": lambda book: book["Pivot"].insert_rows(1, 1),
    "rename Data": lambda book: book.rename_sheet("Data", "Q1 Data"),
}


@pytest.fixture(scope="module")
def answers(live_pivots_answers: Path) -> Answers:
    return json.loads(live_pivots_answers.read_text(encoding="utf-8"))


@pytest.fixture
def book(tmp_path: Path, live_pivots_xlsx: Path) -> Workbook:
    target = tmp_path / "pivots.xlsx"
    shutil.copy(live_pivots_xlsx, target)
    return Workbook.open(target)


def state(book: Workbook) -> dict[str, object]:
    """Where each pivot table is and what its cache reads, and how many
    caches there are, as the answers record them."""
    tables: dict[str, dict[str, str]] = {}
    for sheet in book.sheets:
        for pivot in sheet.pivot_tables:
            source = "" if pivot.source_range is None else pivot.source_range.a1
            tables[pivot.name] = {"location": pivot.location.a1, "source": f"{pivot.source_sheet}!{source}"}
    return {"refused": False, "tables": tables, "caches": len(cache_parts(book.package, book.workbook_part))}


class TestReadingExcelsOwn:
    def test_every_pivot_table(self, book: Workbook, answers: Answers) -> None:
        assert state(book) == answers["base"]

    def test_what_one_is(self, book: Workbook) -> None:
        (beside,) = book["Data"].pivot_tables
        assert isinstance(beside, PivotTable)
        assert (beside.name, beside.location, beside.source_sheet, beside.source_range, beside.source_name) == (
            "Beside",
            RangeRef.parse("F3:G6"),
            "Data",
            RangeRef.parse("A1:C11"),
            None,
        )


class TestEditsAsExcelMakesThem:
    @pytest.mark.parametrize("edit", list(EDITS))
    def test_each_edit(self, book: Workbook, answers: Answers, edit: str) -> None:
        said = answers[edit]
        if said["refused"]:
            before = state(book)
            with pytest.raises(ValueError, match="pivot table 'Beside'"):
                EDITS[edit](book)
            assert state(book) == before, "nothing was changed on the way to refusing"
            return
        EDITS[edit](book)
        book.save()
        assert book.path is not None
        assert state(Workbook.open(book.path)) == said


class TestAPivotTableDeletedWithItsRows:
    def test_its_parts_go_and_its_cache_with_them(self, book: Workbook) -> None:
        (beside,) = book["Data"].pivot_tables
        package = book.package
        book["Data"].delete_rows(3, 4)
        assert beside.cache_part is not None
        for part in (beside.part_name, beside.cache_part):
            assert not package.has_part(part), part
        workbook = package.read(book.workbook_part).decode("utf-8")
        assert workbook.count("<pivotCache ") == 1

    def test_the_rest_of_the_workbook_is_as_it_was(self, book: Workbook) -> None:
        book["Data"].delete_rows(3, 4)
        book.save()
        assert book.path is not None
        again = Workbook.open(book.path)
        (summary,) = again["Pivot"].pivot_tables
        assert summary.name == "Summary"
        assert again["Pivot"].get_value(CellRef.parse("A3")) == "Row Labels"
