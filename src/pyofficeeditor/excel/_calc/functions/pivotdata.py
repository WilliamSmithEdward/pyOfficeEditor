"""GETPIVOTDATA: a value a pivot table's report shows, picked by the value
it summarizes and the items of the fields around it.

Measured against Excel 365 on pivot tables in each layout Excel offers:

- The report is the pivot table whose cells, or the filters above them,
  the reference names. Anything else is ``#REF!``, text spelling a
  reference among it.
- A value answers to its own name or to the field it summarizes, a field
  to the name the report gives it, and an item to the name it was given
  in the report, to the text it is shown as, or, given a number or a
  date, to its value. Case never matters.
- A field left out means its total, which is there only where the report
  shows it, and a filter can be named only as the item it is set to.
- A cell the report leaves empty is 0; one it does not show is ``#REF!``.
"""

from __future__ import annotations

from pyofficeeditor.excel._calc.evaluator import Context
from pyofficeeditor.excel._calc.registry import R, V, function
from pyofficeeditor.excel._calc.values import REF, Empty, Reference, Scalar, Value
from pyofficeeditor.excel._pivots import PivotReport
from pyofficeeditor.excel._values import CellError
from pyofficeeditor.exceptions import UnsupportedFormulaError


def _report(context: Context, reference: Reference) -> tuple[str, PivotReport] | None:
    """The report a reference names: the one holding its first cell, or
    else the first it reaches into."""
    area = reference.areas[0]
    reports = context.book.pivot_reports(area.sheet)
    for report in reports:
        extent = report.extent
        if extent.top <= area.top <= extent.bottom and extent.left <= area.left <= extent.right:
            return area.sheet, report
    for report in reports:
        extent = report.extent
        if extent.top <= area.bottom and area.top <= extent.bottom and extent.left <= area.right and area.left <= extent.right:
            return area.sheet, report
    return None


@function("GETPIVOTDATA", V, R, V, V, minimum=2, maximum=254, repeat=2)
def GETPIVOTDATA(context: Context, data_field: Scalar, pivot_table: Value, *pairs: Scalar) -> Value:
    for given in (data_field, *pairs):
        if isinstance(given, CellError):
            return given
    if not isinstance(pivot_table, Reference):
        return REF
    found = _report(context, pivot_table)
    if found is None or len(pairs) % 2:
        return REF
    sheet, report = found
    if not report.readable:
        raise UnsupportedFormulaError("GETPIVOTDATA on a pivot table from a cube, or one laid out another way")
    value = report.value_index(context.text(data_field))
    if value is None:
        return REF
    chosen: dict[int, int] = {}
    for position in range(0, len(pairs), 2):
        field = report.field_index(context.text(pairs[position]))
        wanted = pairs[position + 1]
        if field is None or isinstance(wanted, (Empty, CellError)):
            return REF
        item = report.item_index(field, wanted)
        if item is None:
            return REF
        chosen[field] = item
    cell = report.cell(value, chosen)
    if cell is None:
        return REF
    shown = context.book.cell(sheet, cell.row, cell.column)
    return 0.0 if isinstance(shown, Empty) else shown


__all__: list[str] = []
