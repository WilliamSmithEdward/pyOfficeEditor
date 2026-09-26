"""Excel's error checking, held to Excel.

``errorchecks.xlsx`` is authored by ``scripts/build_excel_fixtures.py``,
with ``errorchecks_answers.json`` beside it: every cell of it and the rules
``Range.Errors`` says catch it, an error the file records as ignored
included. ``text_checks.json`` is written by ``scripts/measure_text_checks.py``:
some 6,700 strings and what the two rules that read text say of each.
"""

from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

import pytest

from pyofficeeditor.excel import DEFAULT_ERROR_RULES, ERROR_RULES, CellRef, IgnoredError, RangeRef, Workbook
from pyofficeeditor.excel._calc.values import plain_number
from pyofficeeditor.excel._errorchecks import format_kind, r1c1_key, two_digit_year

FIXTURES = Path(__file__).parent / "fixtures" / "excel"
WORKBOOK = FIXTURES / "errorchecks.xlsx"
ANSWERS = FIXTURES / "errorchecks_answers.json"
TEXTS = FIXTURES / "text_checks.json"
#: The day errorchecks.xlsx was authored on.
AUTHORED = dt.date(2026, 9, 25)

needs_workbook = pytest.mark.skipif(
    not WORKBOOK.exists(), reason="run scripts/build_excel_fixtures.py to author errorchecks.xlsx with real Excel"
)


@needs_workbook
def test_the_rules_catch_what_excel_catches() -> None:
    answers = json.loads(ANSWERS.read_text(encoding="utf-8"))
    wanted = {cell: entry["rules"] for cell, entry in answers.items() if entry["rules"]}
    found: dict[str, list[str]] = {}
    with Workbook.open(WORKBOOK) as book:
        for sheet in book.sheets:
            for check in sheet.error_checks(ERROR_RULES, include_ignored=True, today=AUTHORED):
                found.setdefault(f"{sheet.name}!{check.reference.a1}", []).append(check.rule)
    assert found == wanted


@needs_workbook
def test_an_ignored_error_is_left_out_unless_asked_for() -> None:
    with Workbook.open(WORKBOOK) as book:
        sheet = book["Ignored"]
        shown = [(check.reference.a1, check.rule) for check in sheet.error_checks(today=AUTHORED)]
        every = [(check.reference.a1, check.ignored) for check in sheet.error_checks(include_ignored=True, today=AUTHORED)]
        entries = sheet.ignored_errors
    assert shown == [("B7", "twoDigitTextYear")]
    assert every == [("B2", True), ("B3", True), ("B5", True), ("B7", False)]
    assert entries == [
        IgnoredError((RangeRef.parse("B2:B3").normalized,), frozenset({"numberStoredAsText"})),
        IgnoredError((RangeRef(CellRef(5, 2), CellRef(5, 2)).normalized,), frozenset({"evalError"})),
    ]


@needs_workbook
def test_references_to_empty_cells_are_checked_only_when_asked() -> None:
    with Workbook.open(WORKBOOK) as book:
        sheet = book["Empty"]
        by_default = sheet.error_checks(today=AUTHORED)
        asked = sheet.error_checks({"emptyCellReference"}, today=AUTHORED)
    assert "emptyCellReference" not in DEFAULT_ERROR_RULES
    assert by_default == []
    assert [check.reference.a1 for check in asked] == ["A1", "A3", "A5", "E21", "G40"]


def test_the_rules_that_read_text_judge_it_as_excel_does() -> None:
    record = json.loads(TEXTS.read_text(encoding="utf-8"))
    measured = dt.date.fromisoformat(record["measured"])
    wrong = [
        (text, as_number, as_date)
        for text, as_number, as_date in record["cases"]
        if (plain_number(text) is not None) != as_number or two_digit_year(text, measured) != as_date
    ]
    assert len(record["cases"]) > 6000
    assert wrong == []


def test_a_formula_is_judged_by_the_value_cached_for_it(live_empty_xlsx: Path) -> None:
    with Workbook.open(live_empty_xlsx) as book:
        sheet = book.sheets[0]
        sheet["A1"].formula = "1/0"
        # No value is cached for a formula the library wrote until the
        # workbook is calculated.
        before = sheet.error_checks({"evalError"})
        book.calculate()
        after = sheet.error_checks({"evalError"})
    assert before == []
    assert [(check.reference.a1, check.rule) for check in after] == [("A1", "evalError")]


def test_an_unknown_rule_is_refused(live_empty_xlsx: Path) -> None:
    with Workbook.open(live_empty_xlsx) as book, pytest.raises(ValueError, match="not error-checking rules"):
        book.sheets[0].error_checks(["evalError", "spelling"])  # pyright: ignore[reportArgumentType]


def test_formulas_copied_down_compare_alike_and_a_space_makes_them_differ() -> None:
    assert r1c1_key("A1*2", 1, 2) == r1c1_key("A2*2", 2, 2)
    assert r1c1_key("$A$1*2", 1, 2) != r1c1_key("$A$1*2", 1, 2).replace("R1", "R2")
    assert r1c1_key("A1*2", 1, 2) != r1c1_key("A2 *2", 2, 2)
    assert r1c1_key("SUM(A:A)", 1, 2) == r1c1_key("sum(A:A)", 5, 2)


@pytest.mark.parametrize(
    ("code", "kind"),
    [
        ("General", "number"),
        ("0.00", "number"),
        ('0 "kg"', "number"),
        ("m/d/yyyy", "date"),
        ("[h]:mm", "date"),
        ("m/d/yyyy;@", "date"),
        ("@", "neutral"),
        ('"x"', "neutral"),
        (";;;", "neutral"),
    ],
)
def test_a_format_reads_as_a_date_a_number_or_neither(code: str, kind: str) -> None:
    assert format_kind(code) == kind
