"""A note's box when the grid under it changes, held to Excel.

``note_boxes.xlsx`` is authored by ``scripts/build_excel_fixtures.py``:
each sheet holds a note or two, of each placement. Excel then made every
change ``note_boxes_answers.json`` records, in order, a column's width or
a row's height set, columns and rows hidden and shown, a filter, and
folded groups of rows and of columns, and saved the result as
``note_boxes_changed.xlsx``. The library makes the same changes to
``note_boxes.xlsx``, and each sheet has to come out as Excel's did: its
columns and rows, what hangs on its cells, and every note's box, its
anchor and where its style draws it.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path

import pytest
import sheet_state

from pyofficeeditor.excel import FilterColumn, Workbook, Worksheet, criteria
from pyofficeeditor.excel._rowcol import RT_VML, related_parts

FIXTURES = Path(__file__).parent / "fixtures" / "excel"
WORKBOOK = FIXTURES / "note_boxes.xlsx"
CHANGED = FIXTURES / "note_boxes_changed.xlsx"
ANSWERS = FIXTURES / "note_boxes_answers.json"

needs_workbook = pytest.mark.skipif(
    not (WORKBOOK.exists() and CHANGED.exists() and ANSWERS.exists()),
    reason="run scripts/build_excel_fixtures.py to author note_boxes.xlsx with real Excel",
)


@dataclass(frozen=True)
class _Change:
    """One change Excel made, as the answers record it."""

    sheet: str
    change: str
    first: int
    last: int
    #: The width in characters or the height in points Excel was given.
    size: float

    def apply(self, book: Workbook, changed: Workbook) -> None:
        sheet = book[self.sheet]
        spans = range(self.first, self.last + 1)
        # Excel stores a width in units of the standard font's digits, which
        # the file does not measure, and a height fitted to whole pixels, so
        # what it stored is set.
        if self.change == "width":
            sheet.set_column_width(self.first, changed[self.sheet].column_width(self.first))
        elif self.change == "height":
            sheet.set_row_height(self.first, changed[self.sheet].row_height(self.first))
        elif self.change in ("hide columns", "show columns"):
            for column in spans:
                sheet.set_column_hidden(column, self.change == "hide columns")
        elif self.change == "hide rows":
            for row in spans:
                sheet.set_row_hidden(row, True)
        elif self.change == "filter":
            sheet.set_auto_filter(f"A{self.first}:A{self.last}", [FilterColumn(0, criteria(">2"))])
        elif self.change == "group rows":
            sheet.group_rows(self.first, self.last, collapsed=True)
        elif self.change == "group columns":
            sheet.group_columns(self.first, self.last, collapsed=True)
        else:
            raise AssertionError(f"no change is called {self.change!r}")


def _changes() -> list[_Change]:
    if not ANSWERS.exists():
        return []
    answers = json.loads(ANSWERS.read_text(encoding="utf-8"))
    return [_Change(**entry) for _, entry in sorted(answers.items())]


CHANGES = _changes()
SHEETS = sorted({change.sheet for change in CHANGES})


def boxes(sheet: Worksheet) -> list[tuple[str, ...]]:
    """Each note's box: the cell it belongs to, its anchor, and where its
    style draws it and how big."""
    package = sheet.workbook.package
    found: list[tuple[str, ...]] = []
    for part in related_parts(sheet, RT_VML):
        for shape in re.findall(r"<v:shape\b.*?</v:shape>", package.read(part).decode("utf-8"), re.DOTALL):
            cell = re.search(r"<x:Row>(\d+)</x:Row>\s*<x:Column>(\d+)</x:Column>", shape)
            anchor = re.search(r"<x:Anchor>\s*(.*?)\s*</x:Anchor>", shape, re.DOTALL)
            style = re.search(r"style='([^']*)'", shape, re.DOTALL)
            measures = re.findall(r"(?<![\w-])(?:margin-left|margin-top|width|height):[^;']*", style.group(1) if style else "")
            found.append((cell.group(0) if cell else "", anchor.group(1) if anchor else "", *measures))
    return found


@pytest.fixture(scope="module")
def by_library() -> Workbook:
    """note_boxes.xlsx with every change Excel made made here too, in the
    same order."""
    book = Workbook.from_bytes(WORKBOOK.read_bytes())
    changed = Workbook.from_bytes(CHANGED.read_bytes())
    for change in CHANGES:
        change.apply(book, changed)
    return book


@pytest.fixture(scope="module")
def by_excel() -> Workbook:
    return Workbook.from_bytes(CHANGED.read_bytes())


@needs_workbook
@pytest.mark.parametrize("name", SHEETS)
def test_a_note_stays_where_excel_keeps_it(name: str, by_library: Workbook, by_excel: Workbook) -> None:
    """Each box, and the grid it is drawn over: the columns and rows as a
    box counts them, a hidden one as none whatever width it stores."""
    mine, excel = by_library[name], by_excel[name]
    assert boxes(mine) == boxes(excel)
    for axis in ("_column_axis", "_row_axis"):
        assert getattr(mine, axis)() == getattr(excel, axis)()
    assert sheet_state.rows(mine) == sheet_state.rows(excel)
    assert mine.comments == excel.comments


@needs_workbook
def test_the_measured_changes_cover_what_matters() -> None:
    """Widths, heights, hiding and showing, a filter, folded groups of rows
    and of columns, and each placement."""
    assert len(CHANGES) == 31
    assert {change.change for change in CHANGES} == {
        "width", "height", "hide columns", "show columns", "hide rows", "filter", "group rows", "group columns",
    }  # fmt: skip
    assert {"PlaceMove", "PlaceSize", "PlaceFree"} <= set(SHEETS)


# ----------------------------------------------------------------------
# The method's own contract
# ----------------------------------------------------------------------


@pytest.fixture
def sheet(live_empty_xlsx: Path) -> Worksheet:
    book = Workbook.from_bytes(live_empty_xlsx.read_bytes())
    data = book.sheets[0]
    data.set_comment("E5", "a note", author="Ada")
    return data


def test_nothing_moves_when_a_change_is_refused(sheet: Worksheet) -> None:
    before = boxes(sheet)
    with pytest.raises(ValueError, match="negative"):
        sheet.set_column_width(2, -1)
    assert boxes(sheet) == before


def test_a_note_right_of_a_wider_column_keeps_its_place(sheet: Worksheet) -> None:
    """Measured in Excel, on the sheet ``ColRight`` of the fixture: column
    B at 30 characters, 215 pixels, takes the box's left edge from 15
    pixels into F to 56 into C."""
    assert boxes(sheet)[0][1].split(", ")[:2] == ["5", "15"]
    sheet.set_column_width(2, 30.7109375)
    assert boxes(sheet)[0][1].split(", ")[:2] == ["2", "56"]
