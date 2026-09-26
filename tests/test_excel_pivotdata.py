"""GETPIVOTDATA, held to Excel.

``pivotdata.xlsx`` is authored by ``scripts/build_excel_fixtures.py``:
pivot tables in each layout Excel offers, and on sheet Q formulas reading
them, each with the answer Excel cached beside it. The engine has to give
every one of them.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyofficeeditor.excel import CellError, Workbook
from pyofficeeditor.excel._calc.engine import Engine
from pyofficeeditor.excel._calc.values import Array
from pyofficeeditor.excel._values import read_value
from pyofficeeditor.exceptions import UnsupportedFormulaError

FIXTURE = Path(__file__).parent / "fixtures" / "excel" / "pivotdata.xlsx"

pytestmark = pytest.mark.skipif(
    not FIXTURE.exists(), reason="run scripts/build_excel_fixtures.py to author pivotdata.xlsx with real Excel"
)


def _same(got: object, want: object) -> bool:
    if isinstance(want, CellError):
        return isinstance(got, CellError) and got.code == want.code
    if isinstance(want, (int, float)) and not isinstance(want, bool):
        return isinstance(got, float) and got == float(want)
    return got == want


def test_getpivotdata_gives_what_excel_cached() -> None:
    with Workbook.open(FIXTURE) as book:
        sheet = book["Q"]
        engine = Engine(book)
        wrong: list[str] = []
        count = 0
        for reference, element in sheet.cell_elements():
            formula = sheet.formula_of(element, reference)
            if formula is None:
                continue
            count += 1
            want = read_value(element, shared_strings=book.shared_strings, styles=None)
            got = engine.evaluate(formula, "Q", reference.row, reference.column)
            if isinstance(got, Array):
                got = got.rows[0][0]
            if not _same(got, want):
                wrong.append(f"{reference.a1} {formula}: Excel {want!r}, engine {got!r}")
    assert count == 147
    assert not wrong, "\n".join(wrong)


def test_the_report_is_read_as_excel_lays_it_out() -> None:
    with Workbook.open(FIXTURE) as book:
        engine = Engine(book)
        reports = {report.name: report for report in engine.pivot_reports("PT1")}
    nested = reports["Nested"]
    # Region then Product down the side, each region's line first, holding
    # its subtotal, then its products, and the grand total last.
    assert [nested.fields[field].name for field in nested.rows] == ["Region", "Product"]
    assert [line.kind for line in nested.row_lines][:5] == ["data"] * 5
    assert nested.row_lines[0].items == (0,)
    assert nested.row_lines[1].items == (0, 0)
    assert nested.row_lines[-1].kind == "grand"
    filtered = reports["Filtered"]
    # The filter sits above the table and counts as part of it.
    assert filtered.extent.top == filtered.location.top - 2
    year = next(field for field, _ in filtered.pages)
    assert filtered.fields[year].name == "Year"


def test_a_report_it_cannot_read_is_left_to_the_cached_value() -> None:
    with Workbook.open(FIXTURE) as book:
        table = next(table for table in book["PT1"].pivot_tables if table.name == "Nested")
        root = book.package.xml(table.part_name).root
        lines = root.child("rowItems")
        assert lines is not None
        root.remove(lines)
        engine = Engine(book)
        with pytest.raises(UnsupportedFormulaError):
            engine.evaluate('GETPIVOTDATA("Sales",PT1!$E$3,"Region","East")', "Q", 1, 1)
