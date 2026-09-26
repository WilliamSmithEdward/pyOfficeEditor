"""Paste Special, held to Excel's.

``paste_specials.xlsx`` is authored by ``scripts/build_excel_fixtures.py``:
each sheet holds a source block and a destination with things of its own.
Excel then made every paste ``paste_specials_answers.json`` records, in
order, through ``Range.PasteSpecial`` or, for Paste Link,
``Worksheet.Paste Link:=True``, and saved the result as
``paste_specials_pasted.xlsx``; the answers record each paste's source,
destination, choice, operation, whether it skipped blanks and transposed,
and the error Excel refused it with. The library makes the same pastes of
``paste_specials.xlsx``, and each sheet has to come out as Excel's did:
every cell's formula, value and format once calculated, the rows and
columns, the notes and their boxes, links, validation, conditional formats,
merged cells and shapes.
"""

from __future__ import annotations

import datetime as dt
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Literal, cast

import pytest
import sheet_state

from pyofficeeditor.excel import PasteKind, PasteOperation, RangeRef, Workbook, Worksheet
from pyofficeeditor.excel._formulas import number_in_formula, transposed_formula
from pyofficeeditor.excel._reference import CellRef

FIXTURES = Path(__file__).parent / "fixtures" / "excel"
WORKBOOK = FIXTURES / "paste_specials.xlsx"
PASTED = FIXTURES / "paste_specials_pasted.xlsx"
ANSWERS = FIXTURES / "paste_specials_answers.json"
#: The day paste_specials.xlsx was authored on.
AUTHORED = dt.date(2026, 9, 26)

needs_workbook = pytest.mark.skipif(
    not (WORKBOOK.exists() and PASTED.exists() and ANSWERS.exists()),
    reason="run scripts/build_excel_fixtures.py to author paste_specials.xlsx with real Excel",
)


@dataclass(frozen=True)
class _Paste:
    """One paste Excel made, as the answers record it."""

    sheet: str
    source: str
    onto: str
    target: str
    paste: PasteKind
    #: The operation, or empty for none.
    operation: PasteOperation | Literal[""]
    skip_blanks: bool
    transpose: bool
    #: The error Excel refused it with, or empty.
    refused: str

    def apply(self, book: Workbook) -> RangeRef:
        return book[self.sheet].copy_range(
            self.source,
            self.target,
            to=book[self.onto],
            paste=self.paste,
            operation=self.operation or None,
            skip_blanks=self.skip_blanks,
            transpose=self.transpose,
        )


def _pastes() -> list[_Paste]:
    if not ANSWERS.exists():
        return []
    answers = json.loads(ANSWERS.read_text(encoding="utf-8"))
    return [_Paste(**entry) for _, entry in sorted(answers.items())]


PASTES = _pastes()
#: Sheets whose pastes the library refuses for now, though Excel makes
#: them, and what it says: a spill transposed, which spills as far as
#: calculating it says; Paste Link from rows a filter hides; and merged
#: cells a paste covers in part, which Paste Special pastes into in ways
#: of its own.
NOT_YET = {
    "TransposeSpill": "spills",
    "LinkFiltered": "filter hides",
    "LinkMerged": "ways of its own",
    "FormatsMergeAnchor": "ways of its own",
    "CommentsMerged": "ways of its own",
}
SHEETS = sorted(({paste.sheet for paste in PASTES} | {paste.onto for paste in PASTES}) - set(NOT_YET))


@pytest.fixture(scope="module")
def by_library() -> Workbook:
    """paste_specials.xlsx with every paste Excel made made here too, in the
    same order, and calculated."""
    book = Workbook.from_bytes(WORKBOOK.read_bytes())
    for paste in PASTES:
        if not paste.refused and paste.sheet not in NOT_YET:
            paste.apply(book)
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
    assert [shape.name for shape in mine.shapes] == [shape.name for shape in excel.shapes]


#: What the library says when it refuses, by what Excel said.
_REFUSALS = {
    "To do this, all the merged cells need to be the same size.": "merged cells",
    "You can't change part of an array.": "array formula",
    "This selection isn't valid. Make sure the copy and paste areas don't overlap unless they are the same size "
    "and shape.": "overlaps",
}


@needs_workbook
@pytest.mark.parametrize("paste", [paste for paste in PASTES if paste.refused], ids=lambda paste: paste.sheet)
def test_what_excel_refuses_is_refused_and_left_alone(paste: _Paste) -> None:
    book = Workbook.from_bytes(WORKBOOK.read_bytes())
    target = book[paste.onto]
    before = target.document.root.to_xml()
    with pytest.raises(ValueError, match=_REFUSALS[paste.refused]):
        paste.apply(book)
    assert target.document.root.to_xml() == before
    assert not book.is_modified


@needs_workbook
@pytest.mark.parametrize("paste", [paste for paste in PASTES if paste.sheet in NOT_YET], ids=lambda paste: paste.sheet)
def test_what_is_not_supported_yet_is_refused(paste: _Paste) -> None:
    book = Workbook.from_bytes(WORKBOOK.read_bytes())
    before = book[paste.onto].document.root.to_xml()
    with pytest.raises(ValueError, match=NOT_YET[paste.sheet]):
        paste.apply(book)
    assert book[paste.onto].document.root.to_xml() == before


@needs_workbook
def test_the_measured_pastes_cover_what_matters() -> None:
    """Every choice and operation, skipping blanks and transposing, whole
    rows and columns, another sheet, and the refusals."""
    assert len(PASTES) == 109
    assert {paste.paste for paste in PASTES} == {
        "all", "formulas", "values", "formats", "comments", "validation", "all_except_borders", "column_widths",
        "formulas_and_number_formats", "values_and_number_formats", "all_merging_conditional_formats", "link",
    }  # fmt: skip
    assert {paste.operation for paste in PASTES} == {"", "add", "subtract", "multiply", "divide"}
    assert any(paste.skip_blanks for paste in PASTES)
    assert any(paste.transpose for paste in PASTES)
    assert any(paste.onto != paste.sheet for paste in PASTES)
    assert {paste.source for paste in PASTES} >= {"1:2", "A:B"}
    assert {paste.sheet for paste in PASTES if paste.refused} == {
        "ValuesMergeRefused", "ArrayPart", "TransposeOverlap", "TransposeSquare", "OpOntoArray",
    }  # fmt: skip


# ----------------------------------------------------------------------
# The formulas a transposed paste and an operation write
# ----------------------------------------------------------------------


@pytest.mark.parametrize(
    ("formula", "block", "corner", "transposed"),
    [
        ("B2", "A1:B10", "D12", "E13"),
        ("$B2", "A1:B10", "D12", "E$13"),
        ("$F20", "A1:B10", "D12", "$F20"),
        ("SUM($F$20:G21)", "A1:B10", "D12", "SUM($W$17:X18)"),
        ("SUM(5:5)", "A1:B10", "D12", "SUM(H$12:H$1048576)"),
        ("SUM(2:2)", "A1:B3", "G1", "SUM(H:H)"),
        ("SUM($A:$A)", "B3:C12", "E15", "SUM($A:$A)"),
        ("SUM($A$1:$C$6)", "B5:C10", "A1", "SUM(#REF!)"),
        ("Other!D6", "A1:B3", "G1", "Other!L4"),
        ('"B5"&B5', "B5:C10", "A1", '"B5"&A1'),
    ],
)
def test_a_reference_transposed(formula: str, block: str, corner: str, transposed: str) -> None:
    """Measured in Excel: see ``paste_specials.xlsx`` for these and more."""
    assert transposed_formula(formula, RangeRef.parse(block), CellRef.parse(corner), "Sheet1") == transposed


@pytest.mark.parametrize(
    ("value", "written"),
    [
        (1e20, "100000000000000000000"),
        (1e21, "1E+21"),
        (-0.5, "-0.5"),
        (1 / 3, "0.333333333333333"),
        (0.1 + 0.2, "0.3"),
        (1e-19, "0.0000000000000000001"),
        (1.2345e-17, "1.2345E-17"),
        (1.23456789012345e-7, "1.23456789012345E-07"),
        (1000000000000015.0, "1000000000000010"),
        (9007199254740992.0, "9007199254740990"),
    ],
)
def test_a_number_written_into_a_formula(value: float, written: str) -> None:
    """Measured in Excel: an operation writes what is there, or what comes,
    into the formula it makes like this."""
    assert number_in_formula(value) == written


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


def test_values_come_without_their_formulas(sheet: Worksheet) -> None:
    sheet["C2"].formula = "=B2*2"
    sheet["C2"].value = 2
    sheet.copy_range("C2", "D2", paste="values")
    assert (sheet["D2"].formula, sheet["D2"].value) == (None, 2)


def test_an_operation_adds_into_what_is_there(sheet: Worksheet) -> None:
    sheet.copy_range("B2:B4", "B2:B4", paste="values", operation="add")
    assert [sheet[f"B{row}"].value for row in range(2, 5)] == [2, 4, 6]


@pytest.mark.parametrize(
    ("paste", "operation", "transpose", "message"),
    [
        ("everything", None, False, "not a kind of paste"),
        ("all", "power", False, "not an operation"),
        ("formats", "add", False, "combines values"),
        ("link", None, True, "links each cell"),
        ("column_widths", None, True, "not supported yet"),
    ],
)
def test_options_that_do_not_go_together_are_refused(
    sheet: Worksheet, paste: str, operation: str | None, transpose: bool, message: str
) -> None:
    before = sheet.document.root.to_xml()
    with pytest.raises(ValueError, match=message):
        sheet.copy_range(
            "A1:B2",
            "D1",
            paste=cast("PasteKind", paste),
            operation=cast("PasteOperation | None", operation),
            transpose=transpose,
        )
    assert sheet.document.root.to_xml() == before


def test_whole_rows_transposed_are_refused(sheet: Worksheet) -> None:
    with pytest.raises(ValueError, match="not supported yet"):
        sheet.copy_range("1:2", "5:6", transpose=True)
