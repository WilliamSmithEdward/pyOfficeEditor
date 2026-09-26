"""A table giving up the rows at its bottom, held to what Excel's Resize
did.

``table_resizes.json`` is written by ``scripts/measure_table_resize.py``:
tables over ``A1:C8``, and over ``A1:C9`` with a totals row, resized in
Excel to ``A1:C5``, with the references beside them, and conditional
formats and validations of every span of a column and every kind of rule,
each read from the files saved before and after. Remove Duplicates on a
table resizes it this way once it has moved the rows it keeps up, so each
case here is a rule its references follow.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyofficeeditor._xml import Element
from pyofficeeditor.excel import _rowcol
from pyofficeeditor.excel._conditional import ConditionalFormatting, ConditionalRule
from pyofficeeditor.excel._formulas import TableShrink, delete_in_formula
from pyofficeeditor.excel._reference import RangeRef

CORPUS = Path(__file__).parent / "fixtures" / "excel" / "table_resizes.json"

pytestmark = pytest.mark.skipif(
    not CORPUS.exists(), reason="run scripts/measure_table_resize.py to measure table_resizes.json with real Excel"
)


def _shrink(*, totals: bool) -> TableShrink:
    """The resize measured: data rows 2 to 8 down to 2 to 5, under a header
    row, the totals row on row 9 when there is one."""
    return TableShrink(left=1, right=3, starts=(1, 2), last=8, kept=5, totals=9 if totals else None)


def _areas(sqref: str) -> tuple[RangeRef, ...]:
    return tuple(RangeRef.parse(piece) for piece in sqref.split())


def _element(name: str, bounds: list[str], formulas: list[str], kind: str = "") -> Element:
    """A ``<cfRule>`` or a validation, with its formulas as the file held
    them."""
    element = Element.create(name, {"type": kind} if kind else None)
    for bound, text in zip(bounds, formulas, strict=False):
        node = Element.create(bound)
        node.set_text(text)
        element.append(node)
    return element


def test_the_measured_resize_is_the_one_modelled() -> None:
    corpus = json.loads(CORPUS.read_text(encoding="utf-8"))
    assert corpus["table"] == {"columns": "A:C", "first": 2, "last": 8, "kept": 5, "totals": 9}


def test_each_reference_goes_where_excel_moved_it() -> None:
    corpus = json.loads(CORPUS.read_text(encoding="utf-8"))
    wrong = [
        (case["totals"], case["before"], case["after"], moved)
        for case in corpus["formulas"]
        if (
            moved := delete_in_formula(
                case["before"], _shrink(totals=case["totals"]), formula_sheet="Data", target_sheet="Data"
            )
        )
        != case["after"]
    ]
    assert len(corpus["formulas"]) == 148
    assert wrong == []


def test_each_conditional_format_lands_where_excel_put_it() -> None:
    corpus = json.loads(CORPUS.read_text(encoding="utf-8"))
    reads_absolutely = getattr(_rowcol, "_reads_absolutely")
    wrong: list[tuple[object, ...]] = []
    for case in corpus["formats"]:
        rule = ConditionalRule.read(_element("cfRule", ["formula"] * 2, case["formulas"], case["type"]))
        block = ConditionalFormatting(ranges=_areas(case["before"]), rules=(rule,))
        moved = _shrink(totals=True).moved_areas(block.ranges, joined=True, stretched=reads_absolutely(block))
        if " ".join(area.a1 for area in moved) != case["after"]:
            wrong.append((case["type"], case["formulas"], case["before"], case["after"], moved))
    assert len(corpus["formats"]) == 90
    assert wrong == []


def test_each_validation_lands_where_excel_put_it() -> None:
    corpus = json.loads(CORPUS.read_text(encoding="utf-8"))
    bounds_read_absolutely = getattr(_rowcol, "_bounds_read_absolutely")
    wrong: list[tuple[object, ...]] = []
    for case in corpus["validations"]:
        element = _element("dataValidation", ["formula1", "formula2"], case["formulas"])
        moved = _shrink(totals=True).moved_areas(_areas(case["before"]), stretched=bounds_read_absolutely(element))
        if " ".join(area.a1 for area in moved) != case["after"]:
            wrong.append((case["formulas"], case["before"], case["after"], moved))
    assert len(corpus["validations"]) == 75
    assert wrong == []
