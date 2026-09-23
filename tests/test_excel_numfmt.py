"""The text Excel shows for a value under a number format.

``number_formats.json`` is Excel's own answer, measured by
``scripts/measure_number_formats.py``: some five hundred format codes, each
applied to every value in its suite, and ``Range.Text`` read back from a
column too wide to cut anything short. The renderer is held to every cell
of it. The named tests after the corpus pin the rules those measurements
forced, so the rules can be read here rather than dug out of the corpus.
"""

from __future__ import annotations

import datetime as dt
import json
import re
from pathlib import Path
from typing import TypedDict, cast

import pytest

from pyofficeeditor.excel import Workbook
from pyofficeeditor.excel._numfmt import (
    BUILTIN_DISPLAY_CODES,
    FILL_MARK,
    OVERFLOW,
    format_value,
    format_value_marked,
    general,
    parse,
)
from pyofficeeditor.excel._styles import is_date_format
from pyofficeeditor.excel._values import CellError


class Suite(TypedDict):
    name: str
    date1904: bool
    refused: list[str]
    values: list[tuple[str, object]]
    texts: dict[str, list[str]]


class Builtin(TypedDict):
    code: str
    texts: list[str]


class Corpus(TypedDict):
    excel: str
    builtin_samples: list[float]
    builtins: dict[str, Builtin]
    #: Per committed workbook, ``Sheet!A1`` to the text Excel showed.
    cells: dict[str, dict[str, str]]
    suites: list[Suite]


#: Wider than the 255-character column the corpus was measured in, where
#: Excel paints ``#`` because the text does not fit, not because it has none.
COLUMN_WIDTH = 255


@pytest.fixture(scope="module")
def corpus(number_formats_json: Path) -> Corpus:
    return json.loads(number_formats_json.read_text(encoding="utf-8"))


def decode_text(stored: str) -> str:
    """Undo the corpus's run-length squeeze of fills and ``#`` columns."""
    return re.sub("\x01(.)(\\d+)\x02", lambda m: m.group(1) * int(m.group(2)), stored)


def cell_value(raw: object) -> object:
    """A corpus value as a cell holds it."""
    if isinstance(raw, dict):
        entry = cast("dict[str, str]", raw)
        return entry["text"] if "text" in entry else CellError(entry["error"])
    return raw


def shows_as(excel: str, marked: str) -> bool:
    """Whether Excel's text is what the renderer produced, a fill standing
    for any number of its character."""
    if excel and set(excel) == {"#"}:
        return marked == OVERFLOW or len(marked.replace(FILL_MARK, "")) > COLUMN_WIDTH
    pattern = ""
    index = 0
    while index < len(marked):
        if marked[index] == FILL_MARK:
            pattern += re.escape(marked[index + 1]) + "*"
            index += 2
            continue
        pattern += re.escape(marked[index])
        index += 1
    return re.fullmatch(pattern, excel) is not None


@pytest.mark.parametrize("name", ["catalog", "calendar_1904", "details", "conditions"])
def test_every_measured_render(corpus: Corpus, name: str) -> None:
    suite = next(suite for suite in corpus["suites"] if suite["name"] == name)
    misses: list[str] = []
    for code, texts in suite["texts"].items():
        for (label, raw), stored in zip(suite["values"], texts, strict=True):
            got = format_value_marked(cell_value(raw), code, epoch_1904=suite["date1904"])
            if not shows_as(decode_text(stored), got):
                misses.append(f"{code!r} {label}: Excel {decode_text(stored)[:40]!r}, here {got[:40]!r}")
    assert not misses, f"{len(misses)} renders differ:\n" + "\n".join(misses[:25])


def test_every_builtin_id(corpus: Corpus) -> None:
    samples = corpus["builtin_samples"]
    misses: list[str] = []
    for key, entry in corpus["builtins"].items():
        code = BUILTIN_DISPLAY_CODES.get(int(key), "General")
        if code != entry["code"]:
            misses.append(f"id {key}: table says {code!r}, Excel says {entry['code']!r}")
        for sample, stored in zip(samples, entry["texts"], strict=True):
            if not shows_as(decode_text(stored), format_value_marked(sample, code)):
                misses.append(f"id {key} {sample}: Excel {decode_text(stored)!r}")
    assert not misses, "\n".join(misses)


def test_cell_text_for_every_cell_of_the_committed_workbooks(
    corpus: Corpus, number_formats_json: Path
) -> None:
    """From the file's bytes: the style a cell points at, the code its
    format id means, and the text that code gives its value."""
    misses: list[str] = []
    for name, cells in corpus["cells"].items():
        book = Workbook.open(number_formats_json.parent / name)
        for key, stored in cells.items():
            sheet, _, address = key.rpartition("!")
            shown = book[sheet][address].text
            if shown != decode_text(stored):
                misses.append(f"{name} {key}: Excel {decode_text(stored)!r}, here {shown!r}")
    assert sum(len(cells) for cells in corpus["cells"].values()) > 700
    assert not misses, "\n".join(misses)


def test_corpus_covers_what_excel_refused(corpus: Corpus) -> None:
    """Formats Excel would not accept are recorded, not silently dropped:
    ``@`` anywhere but the last section, and a condition on a third."""
    refused = {code for suite in corpus["suites"] for code in suite["refused"]}
    assert {"@;@", "@;0", '"a"0;"b"0;[<-5]"c"0'} <= refused


# ---------------------------------------------------------------------------
# The rules, one at a time
# ---------------------------------------------------------------------------


def test_general_is_eleven_characters_whatever_the_width() -> None:
    assert general(1234.5678) == "1234.5678"
    assert general(1 / 3) == "0.333333333"
    assert general(-1 / 3) == "-0.333333333"
    assert general(123456789012) == "1.23457E+11"
    assert general(12345678901) == "12345678901"
    assert general(0.00000123456789) == "1.23457E-06"
    assert general(0.1 + 0.2) == "0.3"


def test_rounding_is_half_away_from_zero_on_fifteen_digits() -> None:
    # 1.005 is a hair below as a double; Excel rounds what it shows.
    assert format_value(1.005, "0.00") == "1.01"
    assert format_value(2.5, "0") == "3"
    assert format_value(-2.5, "0") == "-3"


def test_a_negative_that_rounds_to_zero_loses_its_sign() -> None:
    assert format_value(-0.0001, "0.00") == "0.00"
    assert format_value(-0.0001, "0%") == "0%"
    # A mixed fraction keeps it.
    assert format_value(-0.0001, "# ?/?") == "-0    "


def test_sections_split_at_zero_without_conditions() -> None:
    assert format_value(-5, "0.00;(0.00)") == "(5.00)"
    assert format_value(0, '0;-0;"zero"') == "zero"
    assert format_value(-5, '"x"') == "-x"
    assert format_value(-1, '"$"General') == "-$1"


def test_a_negative_section_shows_the_magnitude() -> None:
    # A condition no number of zero or more meets.
    assert format_value(-5, '[<0]"neg"0') == "neg5"
    assert format_value(-5, "[=-5]0;0") == "5"
    # One that zero meets shows the sign.
    assert format_value(-5, "[<=0]0;0") == "-5"
    # The catch-all after one inequality whose limit is zero or below.
    assert format_value(-5, "[<-10]0;0") == "5"
    assert format_value(-5, "[>0.5]0;0") == "-5"
    assert format_value(-5, "[=-10]0;0") == "-5"


def test_a_vanishing_negative_is_sectioned_again_as_its_magnitude() -> None:
    assert format_value(-0.49, '[<=0]"a"0;"b"0') == "b0"
    assert format_value(-0.5, '[<=0]"a"0;"b"0') == "-a1"
    # The precision of the section it first landed in decides.
    assert format_value(-0.04, '[<=0]"a"0.0;"b"0.0') == "b0.0"
    assert format_value(-0.05, '[<=0]"a"0.0;"b"0.0') == "-a0.1"
    assert format_value(-0.25, '"a"0.0;[<-5]"b"0;"c"0') == "a0.3"


def test_implied_conditions() -> None:
    # One conditioned section: Excel adds a General catch-all.
    assert format_value(150, "[<100]0.0") == "150"
    # Three sections, the first conditioned: the second takes negatives.
    assert format_value(-3, '[>10]"a"0;"b"0;"c"0') == "b3"
    assert format_value(0, '[>10]"a"0;"b"0;"c"0') == "c0"
    # No section takes the number.
    assert format_value(0, '[<0]"a"0;[>0]"b"0') == OVERFLOW


def test_a_fraction_is_a_convergent_of_the_double() -> None:
    # Not the best approximation: 4/9 is closer to 0.456 than 1/2.
    assert format_value(123.456, "# ?/?") == "123 1/2"
    assert format_value(1234.5678, "# ??/??") == "1234 46/81"
    # 999.99 is a hair above as a double, so it rounds up to 1000.
    assert format_value(999.99, "# ##/##") == "1000"
    assert format_value(0.5, "# ?/8") == " 4/8"


def test_hash_and_question_mark_pad_a_fraction_differently() -> None:
    assert format_value(0.25, "# ?/?") == " 1/4"
    assert format_value(0.25, "# #/#") == "1/4"
    assert format_value(5, "# ?/?") == "5    "
    assert format_value(5, "# #/#") == "5"
    assert format_value(0.25, "# - ?/?") == "   1/4"


def test_improper_fraction_limits() -> None:
    assert format_value(3276.7, "?/10") == "32767/10"
    assert format_value(3276.8, "?/10") == OVERFLOW
    assert format_value(2147483647, "?/?") == OVERFLOW
    assert format_value(1073741824.5, "?/?") == "1073741825/1"
    assert format_value(-0.0001, "?/?") == "0/1"


def test_exponents_step_by_the_integer_placeholders() -> None:
    assert format_value(12345.6789, "0.00E+00") == "1.23E+04"
    assert format_value(12345.6789, "##0.0E+0") == "12.3E+3"
    assert format_value(1, "00.0E+00") == "01.0E+00"
    assert format_value(0, "###.0E+0") == "000.0E+0"


def test_time_rounds_to_what_is_shown_before_the_parts() -> None:
    assert format_value(46027.999999, "m/d/yyyy") == "1/6/2026"
    assert format_value(0.999995, "h:mm:ss") == "0:00:00"
    assert format_value(0.999994, "h:mm:ss") == "23:59:59"


def test_the_1900_calendar_keeps_its_fictional_leap_day() -> None:
    assert format_value(60, "m/d/yyyy") == "2/29/1900"
    assert format_value(61, "m/d/yyyy") == "3/1/1900"
    assert format_value(0, "m/d/yyyy") == "1/0/1900"
    assert format_value(2958465, "m/d/yyyy") == "12/31/9999"
    assert format_value(2958466, "m/d/yyyy") == OVERFLOW


def test_negative_dates_depend_on_the_date_system() -> None:
    assert format_value(-1, "m/d/yyyy") == OVERFLOW
    assert format_value(-0.5, "h:mm;-h:mm") == OVERFLOW
    assert format_value(-1, "m/d/yyyy", epoch_1904=True) == "-1/2/1904"
    assert format_value(-0.25, "h:mm", epoch_1904=True) == "-6:00"
    assert format_value(2957004, "m/d/yyyy", epoch_1904=True) == OVERFLOW


def test_an_elapsed_count_alone_is_a_plain_number() -> None:
    assert format_value(-1, "[h]") == "-24"
    assert format_value(1e15, "[h]") == "24000000000000000"
    assert format_value(-1, "[h]:mm") == OVERFLOW


def test_date_tokens() -> None:
    assert format_value(46027.75, "dddd, mmmm d, yyyy h:mm AM/PM") == "Monday, January 5, 2026 6:00 PM"
    assert format_value(46027, "e") == "2026"
    assert format_value(46027, "mmmmm") == "J"
    assert format_value(0.25, "h am/pm") == "6 AM"
    assert format_value(0.75, "h a/P") == "6 P"
    assert format_value(46027.5, "[$-F800]yyyy") == "Monday, January 5, 2026"
    assert format_value(46027.5, "[$-F400]h:mm") == "12:00:00 PM"


def test_fill_is_dropped_and_padding_is_a_space() -> None:
    assert format_value(5, "_(* #,##0_)") == " 5 "
    assert format_value_marked(5, "* #,##0") == FILL_MARK + " 5"


def test_text_goes_to_the_text_section() -> None:
    assert format_value("abc", '0;0;0;"<"@">"') == "<abc>"
    assert format_value("abc", "0.00") == "abc"
    assert format_value(True, '"<"@">"') == "<TRUE>"
    # A number under a text-only format shows as General.
    assert format_value(5, '"<"@">"') == "5"
    # A last section with @ is the text section wherever it is.
    assert format_value(-5, "0;@") == "-5"


def test_errors_and_empty_cells() -> None:
    assert format_value(CellError("#N/A"), "0.00") == "#N/A"
    assert format_value(None, "0.00") == ""


def test_dates_and_times_as_python_values() -> None:
    assert format_value(dt.datetime(2026, 1, 5, 18, 0), "m/d/yyyy h:mm") == "1/5/2026 18:00"
    assert format_value(dt.date(2026, 1, 5), "m/d/yyyy", epoch_1904=True) == "1/5/2026"
    assert format_value(dt.time(6, 0), "h:mm AM/PM") == "6:00 AM"
    assert format_value(dt.date(1900, 2, 28), "m/d/yyyy") == "2/28/1900"


def test_a_parsed_format_knows_whether_it_is_a_date() -> None:
    assert parse("m/d/yyyy").is_date
    assert parse("[h]:mm").is_date
    assert not parse('0.00" days"').is_date
    assert not parse("General").is_date


def test_the_cell_layer_agrees_which_codes_are_dates(corpus: Corpus) -> None:
    """``Cell.value`` reads a date from ``_styles``, and ``Cell.text`` and
    the filters from the renderer; across every code Excel was measured
    with, the two agree. ``e``, the era year, was where they did not:
    Excel shows 1900 for 0 under it."""
    codes = {code for suite in corpus["suites"] for code in suite["texts"]}
    codes |= {entry["code"] for entry in corpus["builtins"].values()}
    assert {"e", "ee", "yyyy e"} <= codes
    assert {code for code in codes if is_date_format(code) != parse(code).is_date} == set()
    assert not is_date_format("0.00E+00")
    assert not is_date_format("General")
