"""Reading and writing a cell's value.

A ``<c>`` element does not say what it holds in any direct way. Its ``t``
attribute names one of six encodings, one of which (a number) is also how
Excel stores every date, and another of which (a shared string) is an index
into a table in a different part. Getting any of it wrong produces a
plausible wrong answer rather than an error, which is why this module is
separate and tested against bytes Excel wrote.

The encodings, as they appear in a real worksheet::

    <c r="B2"><v>120</v></c>                        number  (t absent)
    <c r="A2" t="s"><v>6</v></c>                    shared string, by index
    <c r="E2" t="b"><v>1</v></c>                    boolean
    <c r="B8" t="e"><v>#DIV/0!</v></c>              error
    <c r="F2" s="2"><v>46037</v></c>                a date, and only s says so
    <c r="A1" t="str"><v>text</v></c>               a formula's string result
    <c r="A1" t="inlineStr"><is><t>x</t></is></c>   text stored in the cell
    <c r="A1"/>                                     nothing

An error is returned as :class:`CellError` rather than as its text, because
a cell holding the literal string ``#DIV/0!`` and a cell holding the error
are different cells and a caller has to be able to tell them apart.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass

from pyofficeeditor._xml import Element
from pyofficeeditor.excel._sharedstrings import SharedStrings, entry_text, needs_space_preserved
from pyofficeeditor.excel._styles import (
    ISO_DATE_FORMAT,
    ISO_DATETIME_FORMAT,
    ISO_TIME_FORMAT,
    Styles,
)
from pyofficeeditor.excel._xstring import decode, encode_text

#: Excel's own error strings.  A cell of type ``e`` holds one of these.
ERROR_CODES = frozenset(
    {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!"}
)

#: The 1900 date system's anchor.  Serial 1 is 1900-01-01, so the epoch sits
#: one day earlier.
_EPOCH_1900 = dt.datetime(1899, 12, 31)
#: The same anchor shifted back a day, used from serial 61 onward.  Excel
#: believes 1900 was a leap year and reserves serial 60 for a 29 February
#: that never happened, so every later serial is one greater than the true
#: day count.  Lotus 1-2-3 had the bug and Excel kept it for compatibility;
#: it is in the file format, not a rounding error.
_EPOCH_1900_AFTER_BUG = dt.datetime(1899, 12, 30)
#: The 1904 system, which old Mac workbooks use.  Serial 0 is 1904-01-01.
_EPOCH_1904 = dt.datetime(1904, 1, 1)

#: The serial Excel gives to its nonexistent 29 February 1900.
PHANTOM_LEAP_DAY_SERIAL = 60

_SECONDS_PER_DAY = 86400


@dataclass(frozen=True)
class CellError:
    """An Excel error value, such as ``#DIV/0!``.

    A distinct type, so a formula that failed is never mistaken for a cell
    whose text happens to look like an error.
    """

    code: str

    def __str__(self) -> str:
        return self.code

    @property
    def is_known(self) -> bool:
        """Whether this is one of the errors Excel documents."""
        return self.code in ERROR_CODES


#: Everything a cell can hold.
CellValue = str | int | float | bool | dt.datetime | dt.date | dt.time | CellError | None


class DateOutOfRangeError(ValueError):
    """A date Excel's serial numbering cannot represent."""


def serial_to_datetime(serial: float, *, epoch_1904: bool = False) -> dt.datetime:
    """Turn a stored number into the moment it denotes.

    In the 1900 system serials 1 to 59 are 1900-01-01 to 1900-02-28, serial
    60 is Excel's imaginary 29 February 1900, and 61 onward run from
    1900-03-01. Serial 60 has no real moment, so it is refused rather than
    quietly reported as 1 March.
    """
    if epoch_1904:
        return _EPOCH_1904 + dt.timedelta(days=serial)
    if serial < 0:
        raise DateOutOfRangeError(f"serial {serial} is negative; the 1900 date system starts at 1")
    if int(serial) == PHANTOM_LEAP_DAY_SERIAL:
        raise DateOutOfRangeError(
            "serial 60 is Excel's 29 February 1900, a day that never existed. "
            "Excel keeps it for Lotus 1-2-3 compatibility; it denotes no real date."
        )
    epoch = _EPOCH_1900 if serial < PHANTOM_LEAP_DAY_SERIAL else _EPOCH_1900_AFTER_BUG
    return epoch + dt.timedelta(days=serial)


def datetime_to_serial(value: dt.datetime | dt.date | dt.time, *, epoch_1904: bool = False) -> float:
    """Turn a moment into the number Excel stores for it.

    A bare :class:`datetime.time` becomes a fraction of a day with no date
    part, which is how Excel stores a time on its own.

    The branches go most specific first, which is not a style choice:
    ``datetime`` is a subclass of ``date``, so testing for ``date`` first
    would swallow every ``datetime`` and drop its time of day.
    """
    moment: dt.datetime
    if isinstance(value, dt.datetime):
        moment = value
    elif isinstance(value, dt.date):
        moment = dt.datetime(value.year, value.month, value.day)
    else:
        return (
            value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000
        ) / _SECONDS_PER_DAY

    if moment.tzinfo is not None:
        raise ValueError(
            "a timezone-aware datetime has no Excel serial; Excel stores wall-clock "
            "times with no zone. Convert it first and decide which zone you meant."
        )

    if epoch_1904:
        if moment < _EPOCH_1904:
            raise DateOutOfRangeError(f"{moment} is before 1904-01-01, the 1904 system's epoch")
        delta = moment - _EPOCH_1904
        return delta.days + delta.seconds / _SECONDS_PER_DAY + delta.microseconds / 1e6 / _SECONDS_PER_DAY

    if moment < _EPOCH_1900 + dt.timedelta(days=1):
        raise DateOutOfRangeError(
            f"{moment} is before 1900-01-01; Excel's 1900 date system cannot store it"
        )
    epoch = _EPOCH_1900 if moment < dt.datetime(1900, 3, 1) else _EPOCH_1900_AFTER_BUG
    delta = moment - epoch
    return delta.days + delta.seconds / _SECONDS_PER_DAY + delta.microseconds / 1e6 / _SECONDS_PER_DAY


def format_number(value: int | float) -> str:
    """A number as Excel writes it into ``<v>``.

    ``repr`` is deliberate for a float: it is the shortest string that reads
    back as the same double, so a value survives the round trip exactly.
    """
    if isinstance(value, bool):  # pragma: no cover - callers check first
        raise TypeError("a boolean is not a number in a worksheet; it has its own cell type")
    if isinstance(value, int):
        return str(value)
    if value != value or value in (float("inf"), float("-inf")):
        raise ValueError(f"{value!r} has no worksheet representation; Excel stores an error instead")
    if value.is_integer() and abs(value) < 1e15:
        return str(int(value))
    return repr(value)


def parse_number(text: str) -> int | float:
    """A ``<v>`` number, kept as an int when it was written as one."""
    try:
        return int(text)
    except ValueError:
        return float(text)


def read_value(
    cell: Element,
    *,
    shared_strings: SharedStrings | None,
    styles: Styles | None,
    epoch_1904: bool = False,
) -> CellValue:
    """The Python value a ``<c>`` element holds.

    ``shared_strings`` and ``styles`` may be ``None`` when the workbook has
    no such part. A text cell then cannot be resolved and raises, because
    returning an index as if it were the text would be worse.
    """
    cell_type = cell.get("t") or "n"

    if cell_type == "inlineStr":
        inline = cell.child("is")
        return "" if inline is None else entry_text(inline)

    value = cell.child("v")
    if value is None:
        return None
    raw = value.text

    if cell_type == "s":
        if shared_strings is None:
            raise ValueError(
                f"cell {cell.get('r')} is a shared string but the workbook has no "
                f"sharedStrings part to resolve index {raw!r} against."
            )
        return shared_strings[int(raw)]

    if cell_type == "str":
        return decode(raw)

    if cell_type == "b":
        return raw not in ("0", "", "false", "FALSE")

    if cell_type == "e":
        return CellError(raw)

    if cell_type == "d":
        # ECMA-376 2nd edition allows an ISO 8601 date in the cell itself.
        # Excel does not write it; other producers do.
        return dt.datetime.fromisoformat(raw)

    if raw == "":
        return None
    number = parse_number(raw)

    style_index = _style_index(cell)
    if styles is not None and styles.is_date(style_index):
        return _as_temporal(number, styles, style_index, epoch_1904)
    return number


def _as_temporal(
    number: int | float, styles: Styles, style_index: int | None, epoch_1904: bool
) -> dt.datetime | dt.date | dt.time | int | float:
    """A number under a date format, as the narrowest type that fits.

    A calendar format with no fractional part gives a :class:`date`; add a
    time and it is a :class:`datetime`; a clock-only format below one day is
    a :class:`time`. A serial Excel's numbering cannot express is returned as
    the number it is, because refusing to read the cell would be worse than
    handing back what the file says.
    """
    calendar = styles.shows_calendar(style_index)
    clock = styles.shows_clock(style_index)
    try:
        moment = serial_to_datetime(number, epoch_1904=epoch_1904)
    except DateOutOfRangeError:
        return number

    if clock and not calendar:
        return moment.time() if number < 1 else moment
    if calendar and not clock and moment.time() == dt.time(0, 0):
        return moment.date()
    return moment


def write_value(
    cell: Element,
    value: CellValue,
    *,
    shared_strings: SharedStrings | None,
    styles: Styles | None,
    epoch_1904: bool = False,
) -> None:
    """Put a Python value into a ``<c>`` element.

    Any previous value, inline string or formula is removed first, because a
    cell that keeps a stale ``<f>`` beside a new ``<v>`` shows the formula's
    old result until Excel recalculates.
    """
    for name in ("v", "is", "f"):
        existing = cell.child(name)
        while existing is not None:
            cell.remove(existing)
            existing = cell.child(name)

    if value is None:
        cell.unset("t")
        return

    # bool before int: in Python True is an int, in a worksheet it is not.
    if isinstance(value, bool):
        cell.set("t", "b")
        _set_v(cell, "1" if value else "0")
        return

    if isinstance(value, CellError):
        cell.set("t", "e")
        _set_v(cell, value.code)
        return

    if isinstance(value, str):
        if shared_strings is None:
            cell.set("t", "inlineStr")
            inline = Element.create("is")
            text = Element.create("t")
            if needs_space_preserved(value):
                text.set("xml:space", "preserve")
            text.set_text(encode_text(value))
            inline.append(text)
            cell.append(inline)
            return
        cell.set("t", "s")
        _set_v(cell, str(shared_strings.index_for(value)))
        return

    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        if styles is None:
            raise ValueError(
                "writing a date needs the styles part, because a date is a number and "
                "only its number format makes it a date."
            )
        serial = datetime_to_serial(value, epoch_1904=epoch_1904)
        cell.unset("t")
        _set_v(cell, format_number(serial))
        cell.set("s", str(styles.ensure_number_format(_format_for(value))))
        return

    # The annotation says this is all that is left, and the check is kept
    # anyway: a type annotation is a promise a caller can break, and writing
    # str(a list) into a cell would be a silently corrupt workbook rather
    # than an error. Hence the suppression, not a wider signature.
    if isinstance(value, (int, float)):  # pyright: ignore[reportUnnecessaryIsInstance]
        cell.unset("t")
        _set_v(cell, format_number(value))
        return

    raise TypeError(
        f"{type(value).__name__} is not something a worksheet cell can hold. "
        f"Use str, int, float, bool, datetime, date, time, CellError or None."
    )


def _format_for(value: dt.datetime | dt.date | dt.time) -> str:
    if isinstance(value, dt.time) and not isinstance(value, dt.datetime):
        return ISO_TIME_FORMAT
    if isinstance(value, dt.datetime) and value.time() != dt.time(0, 0):
        return ISO_DATETIME_FORMAT
    return ISO_DATE_FORMAT


def write_shared_index(cell: Element, index: int) -> None:
    """Point a cell at a shared string by index, as :func:`write_value`
    does for text, for an entry written some other way, such as runs."""
    for name in ("v", "is", "f"):
        existing = cell.child(name)
        while existing is not None:
            cell.remove(existing)
            existing = cell.child(name)
    cell.set("t", "s")
    _set_v(cell, str(index))


def _set_v(cell: Element, text: str) -> None:
    value = Element.create("v")
    value.set_text(text)
    cell.append(value)


def _style_index(cell: Element) -> int | None:
    raw = cell.get("s")
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


__all__ = [
    "ERROR_CODES",
    "PHANTOM_LEAP_DAY_SERIAL",
    "CellError",
    "CellValue",
    "DateOutOfRangeError",
    "datetime_to_serial",
    "format_number",
    "parse_number",
    "read_value",
    "serial_to_datetime",
    "write_shared_index",
    "write_value",
]
