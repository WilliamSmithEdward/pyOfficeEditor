"""What-if data tables, held to Excel.

``datatables.xlsx`` is authored by ``scripts/build_excel_fixtures.py``: a
small loan model with data tables of each kind Range.Table makes, built
twice, on Model and on Moved, the second with other years and principal.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyofficeeditor.excel import CellValue, Workbook
from pyofficeeditor.excel._reference import CellRef, RangeRef

FIXTURE = Path(__file__).parent / "fixtures" / "excel" / "datatables.xlsx"

pytestmark = pytest.mark.skipif(
    not FIXTURE.exists(), reason="run scripts/build_excel_fixtures.py to author datatables.xlsx with real Excel"
)


def _tables(book: Workbook, sheet: str) -> dict[str, CellValue]:
    """Every cell of a sheet's data tables, with its value."""
    found: dict[str, CellValue] = {}
    worksheet = book[sheet]
    for _, element in worksheet.cell_elements():
        formula = element.child("f")
        if formula is None or formula.get("t") != "dataTable":
            continue
        block = RangeRef.parse(formula.get("ref") or "")
        for row in range(block.top, block.bottom + 1):
            for column in range(block.left, block.right + 1):
                address = CellRef(row, column).a1
                found[address] = worksheet[address].value
    return found


def test_data_tables_recalculate_to_what_excel_cached() -> None:
    with Workbook.open(FIXTURE) as book:
        cached = {sheet: _tables(book, sheet) for sheet in ("Model", "Moved")}
        report = book.calculate()
        assert report.complete
        for sheet, values in cached.items():
            assert len(values) == 64
            assert _tables(book, sheet) == values


def test_data_tables_follow_their_inputs() -> None:
    with Workbook.open(FIXTURE) as book:
        wanted = _tables(book, "Moved")
        before = _tables(book, "Model")
        model = book["Model"]
        model["B2"] = 25
        model["B3"] = 300000
        book.calculate()
        got = _tables(book, "Model")
    assert got == wanted
    # All but the error and the text tried, which fail however the model
    # stands.
    assert sum(got[address] != before[address] for address in got) == 62


def test_a_data_table_whose_input_cell_is_gone_keeps_its_values() -> None:
    with Workbook.open(FIXTURE) as book:
        model = book["Model"]
        master = model.cell_element(CellRef.parse("E3"))
        formula = master.child("f")
        assert formula is not None
        formula.set("del1", "1")
        before = model["F8"].value
        model["B3"] = 300000
        report = book.calculate()
        assert "Model!E3" in report.unsupported
        assert model["F8"].value == before
