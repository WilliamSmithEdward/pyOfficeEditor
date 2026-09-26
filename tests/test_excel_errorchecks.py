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
import random
from collections.abc import Iterable
from pathlib import Path

import pytest

from pyofficeeditor._xml import XmlDocument
from pyofficeeditor.excel import (
    DEFAULT_ERROR_RULES,
    ERROR_RULES,
    CellRef,
    IgnoredError,
    RangeRef,
    Workbook,
    _ignorederrors,
)
from pyofficeeditor.excel._calc.values import plain_number
from pyofficeeditor.excel._errorchecks import format_kind, r1c1_key, two_digit_year
from pyofficeeditor.excel._ignorederrors import ignored_rules
from pyofficeeditor.excel._schema import WORKSHEET_CHILD_ORDER, insert_in_schema_order

FIXTURES = Path(__file__).parent / "fixtures" / "excel"
WORKBOOK = FIXTURES / "errorchecks.xlsx"
ANSWERS = FIXTURES / "errorchecks_answers.json"
TEXTS = FIXTURES / "text_checks.json"
IGNORED = FIXTURES / "ignored_errors.json"
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


# ----------------------------------------------------------------------
# Ignored errors, against ignored_errors.json
# ----------------------------------------------------------------------


def _entries(markup: str | None) -> list[IgnoredError]:
    """The entries of an ``<ignoredErrors>`` as markup."""
    if markup is None:
        return []
    found: list[IgnoredError] = []
    for element in XmlDocument.parse(markup.encode("utf-8")).root.children_named("ignoredError"):
        ranges = tuple(RangeRef.parse(piece).normalized for piece in (element.get("sqref") or "").split())
        found.append(IgnoredError(ranges, frozenset(rule for rule in ERROR_RULES if element.get(rule) == "1")))
    return found


def _held(entries: list[IgnoredError], cells: Iterable[str]) -> dict[str, list[str]]:
    """Each cell and the rules it ignores in Excel's order, a cell ignoring
    nothing left out, as the measurement records them."""
    found: dict[str, list[str]] = {}
    for cell in cells:
        at = CellRef.parse(cell)
        rules = ignored_rules(entries, at.row, at.column)
        if rules:
            found[cell] = [rule for rule in ERROR_RULES if rule in rules]
    return found


def test_ignoring_errors_records_them_as_excel_does(live_empty_xlsx: Path) -> None:
    # Each scenario is a sequence of steps Excel took through
    # Range.Errors(i).Ignore. The cells must ignore what Excel said they
    # did, and the markup must be what Excel saved wherever it saved what
    # it held: it loses marks when a cell alone in its entry takes others.
    record = json.loads(IGNORED.read_text(encoding="utf-8"))
    wrong: list[str] = []
    compared = 0
    for case in record["writes"]:
        cells = list(dict.fromkeys(cell for cell, _, _ in case["steps"]))
        with Workbook.open(live_empty_xlsx) as book:
            sheet = book.sheets[0]
            for cell, rule, ignore in case["steps"]:
                sheet.ignore_errors(cell, rule, ignore=ignore)
            written = sheet.document.root.child("ignoredErrors")
            markup = None if written is None else written.to_xml()
            held = _held(sheet.ignored_errors, cells)
        if held != case["ignored"]:
            wrong.append(f"{case['name']}: ignores {held}, Excel {case['ignored']}")
        if _held(_entries(case["saved"]), cells) == case["ignored"]:
            compared += 1
            if markup != case["saved"]:
                wrong.append(f"{case['name']}: wrote {markup}, Excel {case['saved']}")
    assert wrong == []
    assert compared >= 40


def test_ignored_errors_are_read_as_excel_reads_them(live_empty_xlsx: Path) -> None:
    # Each file held an <ignoredErrors> Excel opened: a later entry
    # replaces what an earlier one said of a cell it covers.
    record = json.loads(IGNORED.read_text(encoding="utf-8"))
    cells = [f"{column}{row}" for row in (2, 3, 4) for column in "BCD"]
    wrong: list[str] = []
    for case in record["reads"]:
        with Workbook.open(live_empty_xlsx) as book:
            sheet = book.sheets[0]
            markup = XmlDocument.parse(case["file"].encode("utf-8")).root
            insert_in_schema_order(sheet.document.root, markup, WORKSHEET_CHILD_ORDER)
            held = _held(sheet.ignored_errors, cells)
        if held != case["ignored"]:
            wrong.append(f"{case['name']}: ignores {held}, Excel {case['ignored']}")
    assert wrong == []


def test_a_range_changes_as_its_cells_would_one_at_a_time() -> None:
    # Large ranges are changed in strides; they must end as a cell at a
    # time, row by row, would leave them.
    ignores = getattr(_ignorederrors, "_Ignores")
    rng = random.Random(20260926)
    for _ in range(400):
        quick, slow = ignores([]), ignores([], shortcuts=False)
        for _ in range(12):
            top, left = rng.randint(1, 6), rng.randint(1, 6)
            area = (top, left, top + rng.randint(0, 3), left + rng.randint(0, 3))
            rule = rng.choice(["evalError", "numberStoredAsText", "formula"])
            ignore = rng.random() < 0.7
            quick.change_block(area, rule, ignore)
            slow.change_block(area, rule, ignore)
            assert [(entry.blocks, entry.rules) for entry in quick.entries] == [
                (entry.blocks, entry.rules) for entry in slow.entries
            ]


def test_a_whole_column_is_ignored_as_one_block(live_empty_xlsx: Path) -> None:
    with Workbook.open(live_empty_xlsx) as book:
        sheet = book.sheets[0]
        sheet.ignore_errors("B5", "numberStoredAsText")
        sheet.ignore_errors("B1:B1048576", "numberStoredAsText")
        entries = sheet.ignored_errors
    assert entries == [IgnoredError((RangeRef.parse("B1:B1048576"),), frozenset({"numberStoredAsText"}))]


@needs_workbook
def test_an_entry_left_alone_keeps_its_markup() -> None:
    with Workbook.open(WORKBOOK) as book:
        sheet = book["Ignored"]
        sheet.ignore_errors("B7", "twoDigitTextYear")
        written = sheet.document.root.child("ignoredErrors")
        shown = sheet.error_checks(today=AUTHORED)
    assert written is not None
    assert written.to_xml() == (
        '<ignoredErrors><ignoredError sqref="B2:B3" numberStoredAsText="1"/><ignoredError sqref="B5" evalError="1"/>'
        '<ignoredError sqref="B7" twoDigitTextYear="1"/></ignoredErrors>'
    )
    assert shown == []


def test_ignored_errors_are_saved_and_can_be_reset(live_empty_xlsx: Path, tmp_path: Path) -> None:
    target = tmp_path / "ignored.xlsx"
    with Workbook.open(live_empty_xlsx) as book:
        sheet = book.sheets[0]
        sheet["B2"] = "5"
        sheet["B3"] = "6"
        sheet.ignore_errors("B2", "numberStoredAsText")
        book.save(target)
    with Workbook.open(target) as book:
        sheet = book.sheets[0]
        shown = [(check.reference.a1, check.rule) for check in sheet.error_checks()]
        entries = sheet.ignored_errors
        sheet.reset_ignored_errors()
        after = sheet.ignored_errors
        container = sheet.document.root.child("ignoredErrors")
    assert shown == [("B3", "numberStoredAsText")]
    assert entries == [IgnoredError((RangeRef.parse("B2"),), frozenset({"numberStoredAsText"}))]
    assert after == []
    assert container is None


def test_ignoring_errors_needs_known_rules_and_some_cells(live_empty_xlsx: Path) -> None:
    with Workbook.open(live_empty_xlsx) as book:
        sheet = book.sheets[0]
        with pytest.raises(ValueError, match="not error-checking rules"):
            sheet.ignore_errors("B2", "spelling")  # pyright: ignore[reportArgumentType]
        with pytest.raises(ValueError, match="at least one range"):
            sheet.ignore_errors("", "evalError")
