"""The formula engine, held to Excel one formula at a time.

``formulas.xlsx`` is measured by ``scripts/measure_formulas.py``: inputs
this library wrote as exact doubles, and thousands of formulas Excel typed
in, calculated and saved, each on a sheet named for what it probes. The
value Excel cached beside each formula is what the engine has to give, to
the last bit, or within the few units in the last place listed for the
functions the corpus narrows no further.

A text that names a day without a year, ``"1/15"``, falls in the year the
corpus was built, which the sheet ``About`` records, and CELL("address")
names the workbook by the name Excel calculated it under, ``inputs.xlsx``.
"""

from __future__ import annotations

import datetime as dt
import math
from pathlib import Path

import pytest

from pyofficeeditor.excel import Workbook
from pyofficeeditor.excel._calc import parse
from pyofficeeditor.excel._calc.engine import Engine
from pyofficeeditor.excel._calc.nodes import Call, walk
from pyofficeeditor.excel._calc.values import Array, Reference, Scalar
from pyofficeeditor.excel._values import CellError, read_value

FIXTURE = Path(__file__).parent / "fixtures" / "excel" / "formulas.xlsx"

#: Sheets that hold inputs, not formulas to check.
INPUTS = frozenset(
    {
        "About",
        "Data",
        "Numbers",
        "Pairs",
        "FileLiterals",
        "Powers",
        "Rounding",
        "Divisions",
        "DatePairs",
        "Reals",
        "Wholes",
        "Times",
        "Ties",
        "Crit",
        "Stats",
        "Db",
        "Samples",
    }
)

#: Formulas whose result the engine does not match, and why. Each must
#: still differ: one that starts matching is taken off this list.
KNOWN: dict[str, str] = {}

#: Functions the corpus holds to within a number of units in the last
#: place of Excel's result rather than to the bit: the most any formula
#: calling one is off by. Every other function is exact.
NEAR: dict[str, int] = {
    # Built from ATAN and SQRT, as Excel builds it, but Excel's own is not
    # odd below 0.35, which no odd formula follows.
    "ASIN": 1,
    # Computed exactly here; Excel, more precisely than doubles but not
    # exactly.
    "IPMT": 1, "PPMT": 1, "CUMIPMT": 2, "CUMPRINC": 4,
    # ERF, ERFC and the normal distribution take Excel's own rounding of the
    # argument, which is most of its error in the tail; the rest of the
    # distributions are the nearest double to the exact value. Excel's own
    # approximations for the error and gamma functions and for the gamma
    # and beta integrals are not yet reproduced, nor GAMMALN from 0.7 to 3.
    "NORM.DIST": 1, "NORM.S.DIST": 4, "NORM.S.INV": 3, "T.DIST": 6, "T.DIST.2T": 13, "T.DIST.RT": 13,
    "TDIST": 13, "T.INV": 1, "T.INV.2T": 35, "TINV": 35, "CHISQ.DIST": 52, "CHISQ.DIST.RT": 2, "CHIDIST": 2,
    "CHISQ.INV": 16, "F.DIST": 4, "F.INV": 1, "GAMMA": 15, "GAMMALN": 1, "GAMMA.DIST": 11, "GAMMADIST": 1,
    "GAMMA.INV": 1, "GAMMAINV": 1, "BETA.DIST": 2, "BETADIST": 2, "BETA.INV": 1, "BETAINV": 1,
    "LOGNORM.DIST": 3, "LOGNORMDIST": 1, "LOGNORM.INV": 2, "LOGINV": 2, "HYPGEOM.DIST": 3, "HYPGEOMDIST": 2,
    "NEGBINOM.DIST": 1, "NEGBINOMDIST": 1, "BINOM.DIST": 38, "POISSON.DIST": 9, "CONFIDENCE": 3,
    "CONFIDENCE.NORM": 3, "CONFIDENCE.T": 1, "T.TEST": 23, "TTEST": 20, "F.TEST": 34, "Z.TEST": 6,
    # A Householder QR of a column of ones and the centred data, exact for
    # two and three points; on more the rounding is not yet pinned down.
    "LINEST": 4, "TREND": 1,
    # 2 to the mean base-2 logarithm in an x87 register: Excel's last bit on
    # 239 of 276 samples, a unit from it on the rest.
    "GEOMEAN": 1,
}  # fmt: skip

def _groups() -> list[str]:
    if not FIXTURE.exists():
        return []
    with Workbook.open(FIXTURE) as book:
        return [name for name in book.sheet_names if name not in INPUTS]


@pytest.fixture(scope="module")
def calculated() -> tuple[Workbook, Engine]:
    if not FIXTURE.exists():
        pytest.skip("run scripts/measure_formulas.py to measure formulas.xlsx with real Excel")
    book = Workbook.open(FIXTURE)
    # Excel took the day the build wrote as text for a date.
    built = book["About"]["A2"].value
    assert isinstance(built, dt.date)
    # The build calculated the book as inputs.xlsx, before saving it under
    # the fixture's name.
    engine = Engine(book, today=built, name="inputs.xlsx")
    engine.calculate()
    return book, engine


def _cached(book: Workbook, sheet: str) -> dict[str, tuple[str, Scalar]]:
    """Every formula of a sheet and the value Excel cached for it."""
    found: dict[str, tuple[str, Scalar]] = {}
    worksheet = book[sheet]
    for reference, element in worksheet.cell_elements():
        formula = worksheet.formula_of(element, reference)
        if formula is None:
            continue
        value = read_value(element, shared_strings=book.shared_strings, styles=None)
        found[reference.a1] = (formula, _scalar(value))
    return found


def _scalar(value: object) -> Scalar:
    """A cached value as the engine computes with it: a number is a float,
    and ``t="str"`` with no text is the empty text."""
    if value is None:
        return ""
    if isinstance(value, (bool, str, CellError)):
        return value
    assert isinstance(value, (int, float))
    return float(value)


def _same(got: object, want: object, units: int = 0) -> bool:
    if isinstance(want, float) and isinstance(got, float):
        return got == want or abs(got - want) <= units * math.ulp(want)
    if isinstance(want, CellError) and isinstance(got, CellError):
        return got.code == want.code
    return type(got) is type(want) and got == want


def _tolerance(formula: str) -> int:
    """The units in the last place a formula may be off by: the most any
    function it calls may be."""
    called = {node.function for node in walk(parse(formula)) if isinstance(node, Call)}
    return max((NEAR.get(name, 0) for name in called), default=0)


@pytest.mark.parametrize("sheet", _groups())
def test_every_formula_gives_what_excel_cached(calculated: tuple[Workbook, Engine], sheet: str) -> None:
    book, engine = calculated
    report = engine.report()
    # A formula the engine could not calculate keeps Excel's own value,
    # which would match: it has to be calculated to count.
    kept = set(report.unsupported) | set(report.dependent) | set(report.circular)
    wrong: list[str] = []
    still_known: list[str] = []
    for address, (formula, want) in _cached(book, sheet).items():
        key = (sheet, int("".join(char for char in address if char.isdigit())), 1)
        got = engine.result(key)
        assert not isinstance(got, (Array, Reference))
        name = f"{sheet}!{address}"
        if name in kept:
            wrong.append(f"{name} ={formula}: not calculated, {report.unsupported.get(name, 'reads one that was not')}")
            continue
        if name in KNOWN:
            if _same(got, want):
                still_known.append(name)
            continue
        if not _same(got, want, _tolerance(formula)):
            wrong.append(f"{name} ={formula}: Excel {want!r}, engine {got!r}")
    assert not wrong, "\n".join(wrong)
    assert not still_known, f"these now match Excel, so take them off KNOWN: {still_known}"
