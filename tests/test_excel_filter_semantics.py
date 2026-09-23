"""What an autofilter keeps, held to Excel one criterion at a time.

``filter_semantics.json`` is measured by ``scripts/measure_filters.py``:
columns of every kind of cell, and per case one criterion and the rows
Excel hid for it. A case reached Excel one of three ways: through the object
model, whose markup is what Excel stored; from a file, markup written
straight into a package Excel then re-applied; or as a collation pivot, a
``>`` criterion per word against a column of all of them.

Every case is re-applied here from its markup, with the reader and the
evaluator the worksheet uses, over the cells Excel recorded, and has to hide
exactly the rows Excel hid. Relative date periods are worked out from the
day the corpus was measured, as Excel worked them out then.
"""

from __future__ import annotations

import datetime as dt
import json
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import cast

import pytest

from pyofficeeditor._xml import XmlDocument
from pyofficeeditor.excel._filters import (
    AutoFilter,
    Criterion,
    DynamicFilter,
    FilterCell,
    OpaqueCriterion,
    Top10Filter,
    Verdict,
    decide,
    keeps,
    read_auto_filter,
    resolve,
)
from pyofficeeditor.excel._numfmt import format_value, parse
from pyofficeeditor.excel._reference import RangeRef
from pyofficeeditor.excel._values import CellError

FIXTURE = Path(__file__).parent / "fixtures" / "excel" / "filter_semantics.json"


@dataclass(frozen=True)
class Case:
    name: str
    #: "object model", "file" or "special".
    kind: str
    #: The measured column the criterion was applied to.
    column: str
    #: The ``<autoFilter>`` element, as Excel stored it or as written.
    markup: str
    hidden: frozenset[int]
    #: The error Excel's object model raised applying it, or 0.
    error: int
    refused_to_open: bool
    filter_mode: bool

    @classmethod
    def read(cls, name: str, raw: dict[str, object]) -> Case:
        return cls(
            name=name,
            kind=cast("str", raw["kind"]),
            column=cast("str", raw["column"]),
            markup=cast("str", raw["markup"]),
            hidden=frozenset(cast("list[int]", raw.get("hidden", []))),
            error=cast("int", raw.get("error", 0)),
            refused_to_open=bool(raw.get("refused_to_open", False)),
            filter_mode=bool(raw.get("filter_mode", True)),
        )

    @property
    def applied(self) -> bool:
        return not self.refused_to_open and not self.error and self.kind != "special"


@dataclass(frozen=True)
class Corpus:
    built: dt.date
    columns: dict[str, dict[int, FilterCell]]
    cases: dict[str, Case]


def cell_value(raw: object) -> object:
    if isinstance(raw, dict):
        entry = cast("dict[str, str]", raw)
        return entry["text"] if "text" in entry else CellError(entry["error"])
    return raw


def read_column(rows: list[list[object]]) -> dict[int, FilterCell]:
    """A measured column as the filter sees it, by row. The text is this
    library's rendering, fill left out, which the renderer's own corpus
    holds to Excel."""
    cells: dict[int, FilterCell] = {}
    for row, raw, _, code in rows:
        value = cell_value(raw)
        is_date = isinstance(value, float) and parse(cast("str", code)).is_date
        cells[cast("int", row)] = FilterCell(
            value,  # type: ignore[arg-type]
            format_value(value, cast("str", code)),
            is_date,
        )
    return cells


def load() -> Corpus | None:
    if not FIXTURE.is_file():
        return None
    raw = cast("dict[str, object]", json.loads(FIXTURE.read_text(encoding="utf-8")))
    columns = cast("dict[str, list[list[object]]]", raw["columns"])
    cases = cast("dict[str, dict[str, object]]", raw["cases"])
    return Corpus(
        built=dt.date.fromisoformat(cast("str", raw["built"])),
        columns={name: read_column(rows) for name, rows in columns.items()},
        cases={name: Case.read(name, case) for name, case in cases.items()},
    )


CORPUS = load()
needs_corpus = pytest.mark.skipif(CORPUS is None, reason="run scripts/measure_filters.py with real Excel")


def named(wanted: Callable[[Case], bool]) -> list[str]:
    return [] if CORPUS is None else sorted(name for name, case in CORPUS.cases.items() if wanted(case))


def corpus() -> Corpus:
    assert CORPUS is not None
    return CORPUS


def read(case: Case) -> AutoFilter:
    return read_auto_filter(XmlDocument.parse(case.markup.encode("utf-8")).root)


def hidden_by(filters: AutoFilter, columns: list[str]) -> set[int]:
    """The rows re-applying a filter hides, one measured column per offset."""
    block = RangeRef.parse(filters.ref)
    rows = list(range(block.top + 1, block.bottom + 1))
    verdicts: list[list[Verdict]] = []
    for entry in filters.columns:
        cells = corpus().columns[columns[entry.column]]
        verdicts.append(
            keeps(entry.criterion, [cells.get(row, FilterCell(None)) for row in rows], today=corpus().built)
        )
    if not verdicts:
        return set()
    return {row for row, verdict in zip(rows, decide(verdicts), strict=True) if verdict is False}


@needs_corpus
@pytest.mark.parametrize("name", named(lambda case: case.applied))
def test_hides_what_excel_hid(name: str) -> None:
    case = corpus().cases[name]
    assert hidden_by(read(case), [case.column]) == case.hidden


@needs_corpus
@pytest.mark.parametrize("name", named(lambda case: case.refused_to_open))
def test_markup_excel_refuses_is_not_modelled(name: str) -> None:
    """Excel will not open these, so none reads as a criterion this
    library would write or evaluate."""
    [entry] = read(corpus().cases[name]).columns
    assert isinstance(entry.criterion, OpaqueCriterion)
    assert not entry.criterion.ignored


@needs_corpus
@pytest.mark.parametrize("name", named(lambda case: case.kind == "file" and not case.filter_mode))
def test_markup_excel_ignores_keeps_every_row(name: str) -> None:
    filters = read(corpus().cases[name])
    assert not filters.filtering
    assert hidden_by(filters, ["mixed"]) == set()


@needs_corpus
@pytest.mark.parametrize("name", named(lambda case: case.error == 1004))
def test_top_and_average_over_an_error_are_refused(name: str) -> None:
    """Excel's object model refuses these over a column holding an error,
    and so does resolving one here, before anything is written."""
    case = corpus().cases[name]
    cells = list(corpus().columns[case.column].values())
    criterion: Criterion
    if "avg" in name:
        criterion = DynamicFilter("aboveAverage" if "above" in name else "belowAverage")
    else:
        criterion = Top10Filter(3)
    with pytest.raises(ValueError, match="error"):
        resolve(criterion, cells, today=corpus().built)


@needs_corpus
def test_the_filter_range_stops_at_the_data() -> None:
    """Asked for A1:A40 over data ending at row 30, Excel stored A1:A30."""
    assert read(corpus().cases["beyond_data"]).ref == "A1:A30"


@needs_corpus
def test_a_filter_with_no_criteria_hides_nothing() -> None:
    """Row 7 was hidden by hand before the arrows went on, and stays hidden:
    turning a filter on without criteria touches no row."""
    case = corpus().cases["no_criteria"]
    assert hidden_by(read(case), ["mixed"]) == set()
    assert case.hidden == {7}
    assert not case.filter_mode


@needs_corpus
def test_a_row_hidden_by_hand_that_the_criterion_keeps_is_shown() -> None:
    case = corpus().cases["hand_hidden_kept"]
    assert 7 not in case.hidden
    assert hidden_by(read(case), ["mixed"]) == case.hidden


@needs_corpus
def test_two_columns_both_have_to_keep_a_row() -> None:
    case = corpus().cases["two_columns"]
    assert hidden_by(read(case), ["mixed", "parity"]) == case.hidden


@needs_corpus
@pytest.mark.parametrize(
    "name",
    named(lambda case: case.kind == "object model" and case.applied and ("top" in case.markup or "dynamicFilter" in case.markup)),
)
def test_resolving_gives_what_excel_stored(name: str) -> None:
    """A top ten, an average or a date period resolved the way the object
    model resolves it has the threshold, average and bounds Excel stored."""
    case = corpus().cases[name]
    [entry] = read(case).columns
    stored = entry.criterion
    asked: Criterion
    if isinstance(stored, Top10Filter):
        asked = Top10Filter(stored.count, stored.percent, stored.top)
    else:
        assert isinstance(stored, DynamicFilter)
        asked = DynamicFilter(stored.kind)
    cells = list(corpus().columns[case.column].values())
    assert resolve(asked, cells, today=corpus().built) == stored
