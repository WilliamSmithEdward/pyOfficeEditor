"""Number formats: the text Excel shows for a value.

Held to what Excel produced rather than to the format-code documentation:
some five hundred format codes applied to values chosen to sit on
every boundary, and ``Range.Text`` read back for each cell, in both date
systems. The rules below are the ones those measurements forced, several of
which the documentation does not mention.

**General is eleven characters, whatever the column width.** A narrow
column shows ``1234.568``, but the text a filter matches, and the text this
renders, is ``1234.5678``. The rule: up to eleven characters, not counting
a minus sign, choosing scientific notation only when it keeps more
significant digits than the decimal form would.

**Rounding is half away from zero on the value as Excel holds it**, fifteen
significant digits, so ``1.005`` shows as ``1.01`` although the binary
double is a hair below it.

**Which section a number uses** is decided by conditions, stated or
implied. Without any, one section serves everything, two split at zero with
the second showing the magnitude, and a third takes zero. Excel accepts
conditions on the first two sections only; an unconditioned section among
conditioned ones gets an implied condition, and a number no section accepts
paints ``####``.

**A section shows the minus sign unless it is a negative section.** A
stated condition that no non-negative number meets, ``[<0]`` or ``[=-5]``,
makes one, and so does the implied ``<0`` of a three-section layout. The
catch-all after a single conditioned section is a negative section when
that condition is an inequality with a limit of zero or below.

**A negative that rounds to zero is sectioned again as its magnitude.**
``-0.0001`` under ``0.00`` shows ``0.00``, and under ``[<=0]"a"0;"b"0``
shows ``b0``, because ``a0`` would have lost the sign. A mixed fraction
keeps its sign at zero, so it never goes round again.

**A fraction is a continued-fraction convergent**, the last one whose
denominator fits the placeholders. Not the best approximation: ``0.456``
with one digit is ``1/2``, although ``4/9`` is closer. ``?`` pads with
spaces where ``#`` pads with nothing, which also decides whether the text
between the whole number and the fraction survives when either is absent.
An improper fraction's numerator is a 32-bit count, a 16-bit one over a
stated denominator.

**Time rounds to the precision displayed**, the nearest second unless
fractional seconds are shown, before the parts are taken. So
``46027.999999`` is the next day under ``m/d/yyyy``.

**The 1900 calendar has its fictional leap day**, and cannot show a
negative date or time: ``####``. The 1904 calendar shows one as a minus sign
and the date of its magnitude. Either way a lone elapsed count, ``[h]`` or
``[s]``, is just a number and has no limits.

**Fill is dropped and padding is a space.** ``*x`` repeats x to the column
width, which is not a property of the value, so it renders as nothing;
``_x`` renders as one space. That is the width-independent text Excel
matches a filter against, once the outer spaces are trimmed.

The locale is Excel's en-US: builtin id 14 is ``m/d/yyyy``, day and month
names are English, the decimal separator is a point, and the system long
date and time, ``[$-F800]`` and ``[$-F400]``, are ``dddd, mmmm d, yyyy``
and ``h:mm:ss AM/PM``.
"""

from __future__ import annotations

import datetime as dt
import functools
import itertools
import re
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Context, Decimal
from typing import Literal

#: What each builtin number format id displays as in en-US Excel,
#: measured by writing one cell per id and reading back the code Excel
#: reports. The spec leaves 5-8, 14, 22, 27-44 and 50-81 to the locale;
#: this is the locale Excel was measured in. Ids past 81 show as General.
BUILTIN_DISPLAY_CODES: dict[int, str] = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    5: "$#,##0_);($#,##0)",
    6: "$#,##0_);[Red]($#,##0)",
    7: "$#,##0.00_);($#,##0.00)",
    8: "$#,##0.00_);[Red]($#,##0.00)",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "m/d/yyyy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yyyy h:mm",
    **dict.fromkeys(range(23, 27), "General"),
    **dict.fromkeys(range(27, 32), "m/d/yyyy"),
    **dict.fromkeys(range(32, 36), "h:mm:ss"),
    36: "m/d/yyyy",
    37: "#,##0_);(#,##0)",
    38: "#,##0_);[Red](#,##0)",
    39: "#,##0.00_);(#,##0.00)",
    40: "#,##0.00_);[Red](#,##0.00)",
    41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
    42: '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
    44: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mm:ss.0",
    48: "##0.0E+0",
    49: "@",
    **dict.fromkeys(range(50, 59), "m/d/yyyy"),
    59: "0",
    60: "0.00",
    61: "#,##0",
    62: "#,##0.00",
    63: "$#,##0_);($#,##0)",
    64: "$#,##0_);[Red]($#,##0)",
    65: "$#,##0.00_);($#,##0.00)",
    66: "$#,##0.00_);[Red]($#,##0.00)",
    67: "0%",
    68: "0.00%",
    69: "# ?/?",
    70: "# ??/??",
    71: "m/d/yyyy",
    72: "m/d/yyyy",
    73: "d-mmm-yy",
    74: "d-mmm",
    75: "mmm-yy",
    76: "h:mm",
    77: "h:mm:ss",
    78: "m/d/yyyy h:mm",
    79: "mm:ss",
    80: "[h]:mm:ss",
    81: "mm:ss.0",
}

MONTHS = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)
DAYS = ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")

#: The last serial a date can have, 31 December 9999, in each date system.
MAX_DATE_SERIAL = 2958465
MAX_DATE_SERIAL_1904 = 2957003

#: The en-US system long date and time, which ``[$-F800]`` and ``[$-F400]``
#: (or ``[$-x-sysdate]`` and ``[$-x-systime]``) stand for whatever follows.
SYSTEM_LONG_DATE = "dddd, mmmm d, yyyy"
SYSTEM_TIME = "h:mm:ss AM/PM"

#: What a cell shows when its value has no text under its format: a date
#: out of range, a fraction too large to count, a number no section takes.
#: Excel paints the cell's width in ``#``; this is the width-free form.
OVERFLOW = "########"

#: Where a fill would go, in :func:`format_value_marked`'s output: this character
#: followed by the fill character.
FILL_MARK = "\x00"
_FILL = re.compile("\x00.", re.S)

#: General's budget, not counting a minus sign.
GENERAL_WIDTH = 11

_INT32 = 2**31 - 1
_INT16 = 2**15 - 1

#: Enough precision for any double's digits and then some decimals.
_EXACT = Context(prec=400, rounding=ROUND_HALF_UP)

_COLORS = frozenset(
    {"black", "blue", "cyan", "green", "magenta", "red", "white", "yellow"}
)
_CONDITION = re.compile(r"^(<=|>=|<>|<|>|=)\s*(-?[\d.]+(?:[eE][+-]?\d+)?)$")
_ELAPSED = re.compile(r"^(h+|m+|s+)$", re.IGNORECASE)
_LONG_DATE_LOCALES = frozenset({"F800", "X-SYSDATE"})
_TIME_LOCALES = frozenset({"F400", "X-SYSTIME"})
_PADDING = {"0": "0", "?": " ", "#": ""}


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------

TokenKind = Literal[
    "literal", "digit", "point", "comma", "percent", "exponent", "slash",
    "text", "general", "date", "ampm", "elapsed", "subsecond",
]

#: A condition: an operator and the number it compares against.
Condition = tuple[str, float]


@dataclass
class Token:
    kind: TokenKind
    value: str = ""


@dataclass
class Section:
    tokens: list[Token] = field(default_factory=lambda: [])
    condition: Condition | None = None
    #: Whether this section is a date or time section.
    is_date: bool = False
    #: Whether it holds the ``@`` text placeholder.
    has_text: bool = False
    #: Whether it holds ``General``.
    has_general: bool = False
    #: ``"long_date"`` or ``"time"`` for a system format, else empty.
    system: str = ""


@dataclass(frozen=True)
class Rule:
    """A section a number may use, and when."""

    section: Section
    #: What the number must satisfy; None takes whatever reaches it.
    condition: Condition | None
    #: Whether the section shows the magnitude, its literals supplying any
    #: sign, rather than a minus sign of its own.
    magnitude: bool


@dataclass
class Format:
    """A parsed format code: its sections as written, the rules a number
    goes through, and the section text uses, if any."""

    sections: list[Section]
    rules: list[Rule]
    text: Section | None

    @property
    def is_date(self) -> bool:
        """Whether a number under this format reads as a date or time: its
        first numeric section is one."""
        return self.rules[0].section.is_date


@functools.lru_cache(maxsize=1024)
def parse(code: str) -> Format:
    """A format code, split into sections and tokens and given its rules."""
    sections = [_parse_section(text) for text in _split_sections(code or "General")]
    text: Section | None = None
    numeric = sections
    if sections[-1].has_text:
        # Excel accepts ``@`` in the last section only, and that section is
        # then the text section whichever position it has.
        text, numeric = sections[-1], sections[:-1]
    elif len(sections) >= 4:
        text, numeric = sections[3], sections[:3]
    return Format(sections, _rules(numeric[:3]), text)


def _split_sections(code: str) -> list[str]:
    sections: list[str] = []
    current = ""
    index = 0
    while index < len(code):
        char = code[index]
        if char == '"':
            close = code.find('"', index + 1)
            end = len(code) if close < 0 else close + 1
            current += code[index:end]
            index = end
            continue
        if char in "\\_*" and index + 1 < len(code):
            current += code[index : index + 2]
            index += 2
            continue
        if char == "[":
            close = code.find("]", index + 1)
            end = len(code) if close < 0 else close + 1
            current += code[index:end]
            index = end
            continue
        if char == ";":
            sections.append(current)
            current = ""
            index += 1
            continue
        current += char
        index += 1
    sections.append(current)
    return sections


def _parse_section(text: str) -> Section:
    section = Section()
    tokens = section.tokens
    index = 0
    lower = text.lower()
    while index < len(text):
        char = text[index]
        low = char.lower()
        if char == '"':
            close = text.find('"', index + 1)
            end = len(text) if close < 0 else close
            tokens.append(Token("literal", text[index + 1 : end]))
            index = end + 1
            continue
        if char == "\\":
            if index + 1 < len(text):
                tokens.append(Token("literal", text[index + 1]))
            index += 2
            continue
        if char == "_":
            # Padding the width of the next character: one space.
            tokens.append(Token("literal", " "))
            index += 2
            continue
        if char == "*":
            # Fill to the column width, which the value alone cannot fix. It
            # rides along as a marked literal so every emitter places it, and
            # format_value() drops it at the end.
            fill = text[index + 1] if index + 1 < len(text) else " "
            tokens.append(Token("literal", FILL_MARK + fill))
            index += 2
            continue
        if char == "[":
            close = text.find("]", index + 1)
            if close < 0:
                index += 1
                continue
            _bracket(section, text[index + 1 : close])
            index = close + 1
            continue
        if lower.startswith("general", index):
            tokens.append(Token("general"))
            section.has_general = True
            index += len("general")
            continue
        if char == "@":
            tokens.append(Token("text"))
            section.has_text = True
            index += 1
            continue
        if lower.startswith("am/pm", index):
            tokens.append(Token("ampm", text[index : index + 5]))
            section.is_date = True
            index += 5
            continue
        if lower.startswith("a/p", index):
            tokens.append(Token("ampm", text[index : index + 3]))
            section.is_date = True
            index += 3
            continue
        if low == "e" and index + 1 < len(text) and text[index + 1] in "+-":
            tokens.append(Token("exponent", text[index + 1]))
            index += 2
            continue
        if low in "ymdhse":
            # ``e`` is the era year, which in this locale is the year.
            run = index
            while run < len(text) and text[run].lower() == low:
                run += 1
            tokens.append(Token("date", low * (run - index) if low != "e" else "yyyy"))
            section.is_date = True
            index = run
            continue
        if char in "0#?":
            tokens.append(Token("digit", char))
            index += 1
            continue
        if char == ".":
            tokens.append(Token("point", "."))
            index += 1
            continue
        if char == ",":
            tokens.append(Token("comma", ","))
            index += 1
            continue
        if char == "%":
            tokens.append(Token("percent", "%"))
            index += 1
            continue
        if char == "/":
            tokens.append(Token("slash", "/"))
            index += 1
            continue
        tokens.append(Token("literal", char))
        index += 1

    if section.system:
        # A system format shows the system's date or time whatever the
        # section itself spells out.
        stand_in = _parse_section(SYSTEM_LONG_DATE if section.system == "long_date" else SYSTEM_TIME)
        section.tokens = stand_in.tokens
        section.is_date = True
    elif section.is_date:
        _resolve_date_tokens(section)
    return section


def _bracket(section: Section, body: str) -> None:
    lowered = body.lower()
    if lowered in _COLORS or re.fullmatch(r"color\s*\d+", lowered):
        return
    if _ELAPSED.match(body):
        section.tokens.append(Token("elapsed", body.lower()))
        section.is_date = True
        return
    condition = _CONDITION.match(body.strip())
    if condition:
        section.condition = (condition.group(1), float(condition.group(2)))
        return
    if body.startswith("$"):
        symbol, _, locale = body[1:].partition("-")
        if locale.upper() in _LONG_DATE_LOCALES:
            section.system = "long_date"
        elif locale.upper() in _TIME_LOCALES:
            section.system = "time"
        if symbol:
            section.tokens.append(Token("literal", symbol))
        return
    # DBNum and other directives change nothing in this locale.


def _resolve_date_tokens(section: Section) -> None:
    """Tell minutes from months, and pick up fractional seconds.

    ``m`` and ``mm`` are minutes straight after an hour token or straight
    before a seconds token, skipping literals; months otherwise. A point
    followed by zeros after a seconds token is fractional seconds. A token
    straight after an elapsed one of the same unit, ``[h]h``, is elapsed too.
    """
    tokens = section.tokens
    for previous, token in itertools.pairwise(tokens):
        if previous.kind == "elapsed" and token.kind == "date" and token.value[:1] == previous.value[:1]:
            token.kind = "elapsed"

    dated = [i for i, token in enumerate(tokens) if token.kind in ("date", "elapsed")]
    for position, index in enumerate(dated):
        token = tokens[index]
        if token.kind != "date" or token.value not in ("m", "mm"):
            continue
        before = tokens[dated[position - 1]] if position > 0 else None
        after = tokens[dated[position + 1]] if position + 1 < len(dated) else None
        after_hour = before is not None and before.value[:1] == "h"
        before_second = after is not None and after.value[:1] == "s"
        if after_hour or before_second:
            # "n" marks minutes, so the renderer need not re-decide.
            token.value = "n" * len(token.value)

    rebuilt: list[Token] = []
    index = 0
    while index < len(tokens):
        token = tokens[index]
        if token.kind == "point" and rebuilt and _last_is_seconds(rebuilt):
            digits = 0
            run = index + 1
            while run < len(tokens) and tokens[run].kind == "digit" and tokens[run].value == "0":
                digits += 1
                run += 1
            if digits:
                rebuilt.append(Token("subsecond", str(min(digits, 3))))
                index = run
                continue
        if token.kind == "digit":
            # A stray digit placeholder in a date section is a literal zero.
            rebuilt.append(Token("literal", token.value if token.value != "#" else ""))
            index += 1
            continue
        rebuilt.append(token)
        index += 1
    section.tokens = rebuilt


def _last_is_seconds(tokens: list[Token]) -> bool:
    for token in reversed(tokens):
        if token.kind in ("date", "elapsed"):
            return token.value[:1] == "s"
        if token.kind != "literal":
            return False
    return False


_GENERAL_SECTION = _parse_section("General")


def _rules(numeric: list[Section]) -> list[Rule]:
    """The order sections are tried in, and what each accepts.

    The layouts are the ones Excel allows: conditions on the first two
    sections at most, and an implied condition for an unconditioned
    section among conditioned ones.
    """
    if not numeric:
        # Only a text section: numbers show as General.
        return [Rule(_GENERAL_SECTION, None, False)]
    first = numeric[0].condition
    second = numeric[1].condition if len(numeric) > 1 else None
    third = [Rule(section, None, False) for section in numeric[2:]]
    if second is not None:
        # An unconditioned first section takes the positives.
        head = (
            Rule(numeric[0], first, _negative_only(first))
            if first is not None
            else Rule(numeric[0], (">", 0.0), False)
        )
        return [head, Rule(numeric[1], second, _negative_only(second)), *third]
    if first is not None:
        head = Rule(numeric[0], first, _negative_only(first))
        if len(numeric) == 3:
            return [head, Rule(numeric[1], ("<", 0.0), True), *third]
        # One conditioned section alone gets a General catch-all from Excel.
        catch_all = numeric[1] if len(numeric) == 2 else _GENERAL_SECTION
        return [head, Rule(catch_all, None, _catch_all_is_negative(first))]
    if len(numeric) == 1:
        return [Rule(numeric[0], None, False)]
    if len(numeric) == 2:
        return [Rule(numeric[0], (">=", 0.0), False), Rule(numeric[1], None, True)]
    return [Rule(numeric[0], (">", 0.0), False), Rule(numeric[1], ("<", 0.0), True), *third]


def _negative_only(condition: Condition) -> bool:
    """Whether no number of zero or more meets the condition."""
    operator, limit = condition
    return (operator == "<" and limit <= 0) or (operator in ("<=", "=") and limit < 0)


def _catch_all_is_negative(condition: Condition) -> bool:
    """Whether the section after one conditioned section shows magnitudes.

    Measured rather than derived: it does when the condition is an
    inequality whose limit is zero or below, whatever its direction.
    """
    operator, limit = condition
    return operator != "=" and limit <= 0


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------


def format_value(value: object, code: str, *, epoch_1904: bool = False) -> str:
    """The text Excel shows for a value under a number format code.

    ``value`` is what a cell holds: ``None`` for empty, a ``str``, a
    ``bool``, a number, a date or time, or an error object whose ``str``
    is its code. Errors show as themselves under every format.

    Fill (``*x``) renders as nothing, because how much of it there is
    depends on the column's width rather than on the value.
    """
    return _FILL.sub("", format_value_marked(value, code, epoch_1904=epoch_1904))


def format_value_marked(value: object, code: str, *, epoch_1904: bool = False) -> str:
    """As :func:`format_value`, with each fill left in as ``FILL_MARK`` plus its
    character, for comparing against text Excel rendered in a real column."""
    if value is None:
        return ""
    fmt = parse(code or "General")
    if isinstance(value, bool):
        return _render_text("TRUE" if value else "FALSE", fmt)
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return _render_number(_to_serial(value, epoch_1904), fmt, epoch_1904)
    if isinstance(value, (int, float)):
        return _render_number(float(value), fmt, epoch_1904)
    if isinstance(value, str):
        return _render_text(value, fmt)
    return str(value)


def _to_serial(value: dt.datetime | dt.date | dt.time, epoch_1904: bool) -> float:
    if isinstance(value, dt.time):
        return (value.hour * 3600 + value.minute * 60 + value.second
                + value.microsecond / 1e6) / 86400
    moment = value if isinstance(value, dt.datetime) else dt.datetime(value.year, value.month, value.day)
    if epoch_1904:
        delta = moment - dt.datetime(1904, 1, 1)
    else:
        delta = moment - dt.datetime(1899, 12, 30)
        if moment < dt.datetime(1900, 3, 1):
            delta += dt.timedelta(days=-1)
    return delta.days + delta.seconds / 86400 + delta.microseconds / 86400e6


def _render_text(text: str, fmt: Format) -> str:
    if fmt.text is None:
        return text
    out: list[str] = []
    for token in fmt.text.tokens:
        if token.kind == "text":
            out.append(text)
        elif token.kind == "literal":
            out.append(token.value)
    return "".join(out)


def _render_number(number: float, fmt: Format, epoch_1904: bool) -> str:
    if number != number or number in (float("inf"), float("-inf")):
        return "#NUM!"
    chosen = _choose(fmt, number)
    if chosen is None:
        return OVERFLOW
    section, value = chosen
    if section.is_date:
        if number < 0 and not epoch_1904 and not _elapsed_only(section):
            return OVERFLOW
        return _render_date(section, value, epoch_1904)
    return _render_numeric(section, value)


def _choose(fmt: Format, number: float) -> tuple[Section, float] | None:
    """The section a number uses and the value it renders there: the
    number itself, or its magnitude in a negative section."""
    rule = _first_rule(fmt, number)
    if rule is None:
        return None
    if rule.magnitude:
        return rule.section, abs(number)
    if number < 0 and _rounds_to_zero(rule.section, number):
        # The section would drop the sign: the number goes round again as
        # the positive it now reads as.
        again = _first_rule(fmt, -number)
        if again is None:
            return None
        return again.section, -number
    return rule.section, number


def _first_rule(fmt: Format, number: float) -> Rule | None:
    for rule in fmt.rules:
        if rule.condition is None or _holds(rule.condition, number):
            return rule
    return None


def _holds(condition: Condition, number: float) -> bool:
    operator, limit = condition
    if operator == "<":
        return number < limit
    if operator == "<=":
        return number <= limit
    if operator == ">":
        return number > limit
    if operator == ">=":
        return number >= limit
    if operator == "=":
        return number == limit
    return number != limit


def _is_general(section: Section) -> bool:
    """Whether a section renders a number the General way: it says
    General, or it has no digit placeholders at all."""
    if section.has_general or section.has_text:
        return True
    return not any(token.kind == "digit" for token in section.tokens)


def _rounds_to_zero(section: Section, number: float) -> bool:
    """Whether the section would show a negative number without its sign,
    because what it displays rounds to zero."""
    if section.is_date or _is_general(section):
        return False
    tokens = section.tokens
    if any(token.kind == "exponent" for token in tokens):
        return False
    if _slash_between_digits(tokens):
        layout = _fraction_layout(tokens)
        if layout.has_integer:
            return False
        numbers = _fraction_numbers(layout, number)
        return numbers is not None and numbers[1] == 0
    whole, fraction = _digit_text(_digit_layout(tokens), number)
    return not whole.strip("0") and not fraction.strip("0")


def _render_numeric(section: Section, value: float) -> str:
    tokens = section.tokens
    if _is_general(section):
        return _render_general_section(section, value)
    if any(token.kind == "exponent" for token in tokens):
        return _render_scientific(section, value)
    if _slash_between_digits(tokens):
        return _render_fraction(section, value)
    return _render_digits(section, value)


def _render_general_section(section: Section, value: float) -> str:
    """A section that is General, possibly with literals around it, or one
    with no placeholders at all, such as ``"pos"``. A minus sign goes in
    front of the lot."""
    out: list[str] = ["-"] if value < 0 else []
    for token in section.tokens:
        if token.kind in ("general", "text"):
            out.append(general(abs(value)))
        elif token.kind == "literal":
            out.append(token.value)
    return "".join(out)


# ---------------------------------------------------------------------------
# General
# ---------------------------------------------------------------------------


def general(number: float) -> str:
    """A number in the General format, the eleven-character way."""
    if number == 0:
        return "0"
    sign = "-" if number < 0 else ""
    size = abs(number)
    decimal = _general_decimal(size)
    scientific = _general_scientific(size)
    if decimal is None:
        return sign + scientific
    if _significant(scientific) > _significant(decimal):
        return sign + scientific
    return sign + decimal


def _general_decimal(size: float) -> str | None:
    """The decimal form in eleven characters, or None if it cannot fit."""
    integer_digits = max(_exponent(size) + 1, 1)
    if integer_digits > GENERAL_WIDTH:
        return None
    places = max(GENERAL_WIDTH - integer_digits - 1, 0)
    text = _fixed(_held(size), places)
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    if len(text) > GENERAL_WIDTH or text in ("0", ""):
        return None
    return text


def _general_scientific(size: float) -> str:
    exponent = _exponent(size)
    for _ in range(2):
        tail = f"E{'-' if exponent < 0 else '+'}{abs(exponent):02d}"
        places = max(GENERAL_WIDTH - len(tail) - 2, 0)
        mantissa = _round(_held(size).scaleb(-exponent), places)
        if mantissa >= 10:
            exponent += 1
            continue
        text = f"{mantissa:.{places}f}"
        if "." in text:
            text = text.rstrip("0").rstrip(".")
        return text + tail
    return f"1E{'-' if exponent < 0 else '+'}{abs(exponent):02d}"


def _significant(text: str) -> int:
    mantissa = text.split("E")[0].replace(".", "")
    return len(mantissa.strip("0"))


# ---------------------------------------------------------------------------
# Rounding
# ---------------------------------------------------------------------------


def _held(number: float) -> Decimal:
    """A number as Excel holds it: fifteen significant digits."""
    return Decimal(format(number, ".15g"))


def _exponent(size: float) -> int:
    """The power of ten of a positive number's leading digit, as held."""
    return _held(size).adjusted()


def _round(value: Decimal, places: int) -> Decimal:
    return value.quantize(Decimal(1).scaleb(-places), context=_EXACT)


def _fixed(value: Decimal, places: int) -> str:
    """A non-negative number rounded half away from zero to fixed places."""
    return f"{_round(value, places):.{places}f}"


# ---------------------------------------------------------------------------
# Digit placeholders
# ---------------------------------------------------------------------------


@dataclass
class _DigitLayout:
    point: int | None
    integer_slots: list[str]
    decimal_slots: list[str]
    percents: int
    scale_commas: set[int]
    thousands: bool


def _digit_layout(tokens: list[Token]) -> _DigitLayout:
    point = next((i for i, t in enumerate(tokens) if t.kind == "point"), None)
    scale_commas = _scaling_commas(tokens)
    return _DigitLayout(
        point=point,
        integer_slots=[
            t.value for i, t in enumerate(tokens)
            if t.kind == "digit" and (point is None or i < point)
        ],
        decimal_slots=[
            t.value for i, t in enumerate(tokens)
            if t.kind == "digit" and point is not None and i > point
        ],
        percents=sum(1 for t in tokens if t.kind == "percent"),
        scale_commas=scale_commas,
        thousands=_thousands(tokens, scale_commas),
    )


def _digit_text(layout: _DigitLayout, value: float) -> tuple[str, str]:
    """The integer and decimal digits a value shows: scaled by each percent
    and each scaling comma, then rounded to the decimal placeholders."""
    scaled = _held(abs(value)) * (100 ** layout.percents)
    scaled = _EXACT.divide(scaled, Decimal(1000) ** len(layout.scale_commas))
    whole, _, fraction = _fixed(scaled, len(layout.decimal_slots)).partition(".")
    return whole, fraction


def _render_digits(section: Section, value: float) -> str:
    layout = _digit_layout(section.tokens)
    whole, fraction = _digit_text(layout, value)
    shows_nonzero = bool(whole.strip("0") or fraction.strip("0"))
    sign = "-" if value < 0 and shows_nonzero else ""
    if whole == "0":
        whole = ""
    integer = _place_integer(whole, layout.integer_slots, layout.thousands)
    decimals = _place_decimals(fraction, layout.decimal_slots)
    return sign + _assemble(section.tokens, integer, decimals, layout.point, layout.scale_commas)


def _scaling_commas(tokens: list[Token]) -> set[int]:
    """Commas that divide by a thousand.

    A comma whose next non-comma token is a digit placeholder separates
    thousands. Any other comma scales, provided a digit placeholder comes
    before it or the point comes straight after it: ``0,`` and ``0.0,,``
    and ``,.00`` all scale.
    """
    found: set[int] = set()
    digits = [i for i, t in enumerate(tokens) if t.kind == "digit"]
    if not digits:
        return found
    for i, token in enumerate(tokens):
        if token.kind != "comma":
            continue
        following = next((t for t in tokens[i + 1 :] if t.kind != "comma"), None)
        if following is not None and following.kind == "digit":
            continue
        if digits[0] < i or (following is not None and following.kind == "point"):
            found.add(i)
    return found


def _thousands(tokens: list[Token], scale_commas: set[int]) -> bool:
    point = next((i for i, t in enumerate(tokens) if t.kind == "point"), len(tokens))
    return any(
        token.kind == "comma" and i not in scale_commas and i < point
        and any(t.kind == "digit" for t in tokens[:i])
        for i, token in enumerate(tokens)
    )


def _place_integer(whole: str, slots: list[str], thousands: bool) -> list[str]:
    """One string per integer placeholder, left to right.

    Digits fill from the right. Surplus digits all go to the first
    placeholder. A placeholder without a digit shows 0 for ``0``, a space
    for ``?`` and nothing for ``#``.
    """
    digits = list(whole)
    placed = [""] * len(slots)
    for slot in range(len(slots) - 1, -1, -1):
        placed[slot] = digits.pop() if digits else _PADDING[slots[slot]]
    if digits and slots:
        placed[0] = "".join(digits) + placed[0]
    if thousands:
        joined = "".join(placed)
        stripped = joined.lstrip(" ")
        grouped = _group(stripped)
        return [joined[: len(joined) - len(stripped)] + grouped] + [""] * (len(slots) - 1)
    return placed


def _group(digits: str) -> str:
    if not digits:
        return digits
    head = len(digits) % 3 or 3
    parts = [digits[:head]] + [digits[i : i + 3] for i in range(head, len(digits), 3)]
    return ",".join(parts)


def _place_decimals(fraction: str, slots: list[str]) -> list[str]:
    """One string per decimal placeholder, left to right. Trailing zeros
    are dropped for ``#`` and become spaces for ``?``."""
    placed = list(fraction.ljust(len(slots), "0"))[: len(slots)]
    for slot in range(len(slots) - 1, -1, -1):
        if placed[slot] != "0" or slots[slot] == "0":
            break
        placed[slot] = _PADDING[slots[slot]]
    return placed


def _assemble(
    tokens: list[Token], integer: list[str], decimals: list[str],
    point: int | None, scale_commas: set[int],
) -> str:
    out: list[str] = []
    integer_index = 0
    decimal_index = 0
    for i, token in enumerate(tokens):
        kind = token.kind
        if kind == "digit":
            if point is None or i < point:
                out.append(integer[integer_index])
                integer_index += 1
            else:
                out.append(decimals[decimal_index])
                decimal_index += 1
        elif kind in ("point", "percent", "slash", "literal"):
            out.append(token.value)
        # Commas are grouping, applied already, or scaling, applied already.
    return "".join(out)


# ---------------------------------------------------------------------------
# Scientific
# ---------------------------------------------------------------------------


def _render_scientific(section: Section, value: float) -> str:
    """``0.00E+00`` and relatives. The exponent is a multiple of the number
    of integer placeholders, so ``##0.0E+0`` is engineering notation and
    ``00.0E+00`` steps by two."""
    tokens = section.tokens
    marker = next(i for i, t in enumerate(tokens) if t.kind == "exponent")
    mantissa_tokens = tokens[:marker]
    exponent_tokens = tokens[marker + 1 :]
    point = next((i for i, t in enumerate(mantissa_tokens) if t.kind == "point"), None)
    integer_slots = [t.value for i, t in enumerate(mantissa_tokens)
                     if t.kind == "digit" and (point is None or i < point)]
    decimal_slots = [t.value for i, t in enumerate(mantissa_tokens)
                     if t.kind == "digit" and point is not None and i > point]
    places = len(decimal_slots)
    step = max(len(integer_slots), 1)

    size = abs(value)
    if size == 0:
        exponent = 0
        # Every integer placeholder shows a zero, # included.
        mantissa_text = "0" * (step - 1) + _fixed(Decimal(0), places)
    else:
        held = _held(size)
        exponent = held.adjusted()
        exponent -= exponent % step
        rounded = _round(held.scaleb(-exponent), places)
        if rounded >= 10**step:
            exponent += step
            rounded = _round(held.scaleb(-exponent), places)
        mantissa_text = f"{rounded:.{places}f}"

    whole, _, fraction = mantissa_text.partition(".")
    integer = _place_integer(whole, integer_slots, False)
    decimals = _place_decimals(fraction, decimal_slots)
    mantissa_out = _assemble(mantissa_tokens, integer, decimals, point, set())

    exponent_sign = "-" if exponent < 0 else ("+" if tokens[marker].value == "+" else "")
    exponent_digits = sum(1 for t in exponent_tokens if t.kind == "digit")
    exponent_out = ""
    placed = False
    for token in exponent_tokens:
        if token.kind == "digit":
            if not placed:
                exponent_out += str(abs(exponent)).rjust(exponent_digits, "0")
                placed = True
        elif token.kind == "literal":
            exponent_out += token.value
    sign = "-" if value < 0 else ""
    return f"{sign}{mantissa_out}E{exponent_sign}{exponent_out}"


# ---------------------------------------------------------------------------
# Fractions
# ---------------------------------------------------------------------------


def _slash_between_digits(tokens: list[Token]) -> bool:
    for i, token in enumerate(tokens):
        if token.kind != "slash":
            continue
        before = any(t.kind == "digit" for t in tokens[:i])
        after = any(t.kind == "digit" or (t.kind == "literal" and t.value.isdigit())
                    for t in tokens[i + 1 :])
        if before and after:
            return True
    return False


@dataclass
class _FractionLayout:
    #: Literals before the whole number and its placeholders.
    head: list[Token]
    #: Literals between the whole number and the numerator.
    separator: str
    integer_slots: list[str]
    numerator_slots: list[str]
    denominator_slots: list[str]
    #: A stated denominator, such as the 8 of ``# ?/8``, else empty.
    fixed: str
    trailing: str

    @property
    def has_integer(self) -> bool:
        return bool(self.integer_slots)


def _fraction_layout(tokens: list[Token]) -> _FractionLayout:
    slash = next(i for i, t in enumerate(tokens) if t.kind == "slash")
    before = tokens[:slash]
    after = tokens[slash + 1 :]

    # The numerator is the last run of digit placeholders before the slash;
    # any placeholders before that are the whole number.
    start = slash
    while start > 0 and before[start - 1].kind == "digit":
        start -= 1
    integer_tokens = before[:start]
    last_digit = max((i for i, t in enumerate(integer_tokens) if t.kind == "digit"), default=-1)
    head = integer_tokens[: last_digit + 1]
    separator = "".join(t.value for t in integer_tokens[last_digit + 1 :] if t.kind == "literal")
    if last_digit < 0:
        head, separator = integer_tokens, ""

    denominator: list[Token] = []
    trailing: list[Token] = []
    for position, token in enumerate(after):
        if token.kind == "digit" or (token.kind == "literal" and token.value.isdigit()):
            denominator.append(token)
        else:
            trailing = after[position:]
            break
    # A denominator with any literal digit in it is stated, and the whole
    # run is the number: "100" tokenizes as a literal 1 and two 0 placeholders.
    fixed = ""
    if any(t.kind == "literal" for t in denominator):
        fixed = "".join(t.value for t in denominator)
    return _FractionLayout(
        head=head,
        separator=separator,
        integer_slots=[t.value for t in integer_tokens if t.kind == "digit"],
        numerator_slots=[t.value for t in before[start:]],
        denominator_slots=[] if fixed else [t.value for t in denominator],
        fixed=fixed,
        trailing="".join(t.value for t in trailing if t.kind == "literal"),
    )


def _fraction_numbers(layout: _FractionLayout, value: float) -> tuple[int, int, int] | None:
    """The whole number, numerator and denominator a value shows as, or
    None when an improper fraction's count overflows.

    Unlike everywhere else, the fraction comes from the binary double, not
    the fifteen digits Excel shows: ``999.99`` is a hair above, so under
    ``# ##/##`` it rounds up to 1000 where 0.99 exactly would be 98/99.
    """
    size = abs(value)
    whole = int(size)
    remainder = size - whole
    if layout.fixed:
        denominator = int(layout.fixed)
        numerator = int(_round(_held(remainder * denominator), 0))
    else:
        limit = 10 ** max(len(layout.denominator_slots), 1) - 1
        numerator, denominator = _convergent(remainder, limit)
    if numerator == denominator:
        whole += 1
        numerator = 0
    if layout.has_integer:
        # Past 2**53 a double is a whole number, shown to fifteen digits.
        return (int(_held(size)) if size >= 1e15 else whole), numerator, denominator
    if layout.fixed:
        numerator += whole * denominator
        return None if numerator > _INT16 else (0, numerator, denominator)
    if whole >= _INT32:
        return None
    numerator += whole * denominator
    if numerator > _INT32:
        # Too fine to count: the nearest whole number over one.
        return 0, int(_round(_held(size), 0)), 1
    return 0, numerator, denominator


def _render_fraction(section: Section, value: float) -> str:
    """``# ?/?``, ``# ??/??``, ``?/?`` and stated denominators like ``# ?/8``."""
    layout = _fraction_layout(section.tokens)
    numbers = _fraction_numbers(layout, value)
    if numbers is None:
        return OVERFLOW
    whole, numerator, denominator = numbers
    numerator_text = _numerator_text(str(numerator), layout.numerator_slots)
    denominator_text = layout.fixed or _denominator_text(str(denominator), layout.denominator_slots)

    if not layout.has_integer:
        sign = "-" if value < 0 and numerator else ""
        head = "".join(t.value for t in layout.head if t.kind == "literal")
        return f"{sign}{head}{numerator_text}/{denominator_text}{layout.trailing}"

    # A mixed fraction keeps its sign even when it shows zero, "-0    ".
    sign = "-" if value < 0 else ""
    zero_fraction = numerator == 0
    placed = _place_integer(str(whole) if whole or zero_fraction else "", layout.integer_slots, False)
    head = _assemble(layout.head, placed, [], None, set())
    if zero_fraction:
        # The fraction and what leads to it blank out, as spaces the width
        # of 0/1 if any placeholder pads, or as nothing if all are #.
        if all(slot == "#" for slot in layout.numerator_slots + layout.denominator_slots):
            body = ""
        else:
            blank = _numerator_text("0", layout.numerator_slots) + "/" + (
                layout.fixed or _denominator_text("1", layout.denominator_slots))
            body = " " * (len(layout.separator) + len(blank))
    else:
        separator = layout.separator
        if not "".join(placed):
            # No whole number: the separator pads like the numerator does.
            pads = layout.numerator_slots[:1] != ["#"]
            separator = " " * len(separator) if pads else ""
        body = f"{separator}{numerator_text}/{denominator_text}"
    return f"{sign}{head}{body}{layout.trailing}"


def _numerator_text(digits: str, slots: list[str]) -> str:
    """Right-aligned in its placeholders; surplus digits spill left."""
    remaining = list(digits)
    placed: list[str] = []
    for slot in reversed(slots):
        placed.append(remaining.pop() if remaining else _PADDING[slot])
    return "".join(remaining) + "".join(reversed(placed))


def _denominator_text(digits: str, slots: list[str]) -> str:
    """Left-aligned in its placeholders, padded after with spaces for ``?``."""
    return digits + "".join(" " if slot != "#" else "" for slot in slots[len(digits) :])


def _convergent(value: float, limit: int) -> tuple[int, int]:
    """The last continued-fraction convergent with a denominator in limit."""
    if value <= 0:
        return 0, 1
    # h(-2), h(-1) = 0, 1 and k(-2), k(-1) = 1, 0: the standard seeds.
    h_prev, h = 0, 1
    k_prev, k = 1, 0
    remainder = value
    best = (0, 1)
    for _ in range(64):
        a = int(remainder)
        h_prev, h = h, a * h + h_prev
        k_prev, k = k, a * k + k_prev
        if k > limit:
            break
        best = (h, k)
        fraction = remainder - a
        if fraction < 1e-12:
            break
        remainder = 1 / fraction
    return best


# ---------------------------------------------------------------------------
# Dates and times
# ---------------------------------------------------------------------------


def _elapsed_only(section: Section) -> bool:
    """Whether the section counts elapsed time and shows nothing else, such
    as ``[h]`` or ``[s] "sec"``: a plain number with no date limits."""
    kinds = [t.kind for t in section.tokens if t.kind in ("date", "elapsed", "ampm", "subsecond")]
    return bool(kinds) and all(kind == "elapsed" for kind in kinds)


def _render_date(section: Section, value: float, epoch_1904: bool) -> str:
    if value < 0:
        # Only the 1904 system and a lone elapsed count get here: a minus
        # sign and the magnitude's text.
        magnitude = _render_date(section, -value, epoch_1904)
        return magnitude if magnitude == OVERFLOW else "-" + magnitude

    elapsed_only = _elapsed_only(section)
    last = MAX_DATE_SERIAL_1904 if epoch_1904 else MAX_DATE_SERIAL
    if not elapsed_only and value >= last + 1:
        return OVERFLOW
    places = max((int(t.value) for t in section.tokens if t.kind == "subsecond"), default=0)
    total_seconds = _round(_EXACT.multiply(_held(value), Decimal(86400)), places)
    if elapsed_only:
        return "".join(
            _elapsed(t.value, total_seconds) if t.kind == "elapsed" else t.value
            for t in section.tokens
            if t.kind in ("elapsed", "literal")
        )
    # Rounding to the displayed precision can carry into the next day.
    days = int(total_seconds) // 86400
    if days > last:
        return OVERFLOW
    seconds_of_day = total_seconds - days * 86400
    hour = int(seconds_of_day // 3600)
    minute = int((seconds_of_day % 3600) // 60)
    second = int(seconds_of_day % 60)
    fraction = seconds_of_day - int(seconds_of_day)

    year, month, day = _calendar(days, epoch_1904)
    weekday = (days + (5 if epoch_1904 else 6)) % 7
    twelve = any(t.kind == "ampm" for t in section.tokens)

    out: list[str] = []
    for token in section.tokens:
        kind = token.kind
        text = token.value
        if kind == "literal":
            out.append(text)
        elif kind == "ampm":
            out.append(_meridiem(text, hour < 12))
        elif kind == "elapsed":
            out.append(_elapsed(text, total_seconds))
        elif kind == "subsecond":
            out.append("." + f"{fraction:.{int(text)}f}".split(".")[1])
        elif kind == "date":
            out.append(_date_part(text, year, month, day, hour, minute, second, weekday, twelve))
        elif kind in ("point", "comma", "percent", "slash"):
            out.append(text)
    return "".join(out)


def _meridiem(spelling: str, morning: bool) -> str:
    """AM/PM in any case shows as ``AM`` or ``PM``; A/P shows each letter
    in the case it was written."""
    if len(spelling) == 5:
        return "AM" if morning else "PM"
    return spelling[0] if morning else spelling[2]


def _elapsed(token: str, total_seconds: Decimal) -> str:
    """An elapsed count, ``[h]``, ``[mm]`` or ``[ss]``, of whole units, to
    the fifteen significant digits Excel shows of any number."""
    amount = int(total_seconds) // {"h": 3600, "m": 60, "s": 1}[token[0]]
    if amount >= 10**15:
        amount = int(_held(float(amount)))
    return str(amount).rjust(len(token), "0")


def _date_part(
    token: str, year: int, month: int, day: int, hour: int, minute: int,
    second: int, weekday: int, twelve: bool,
) -> str:
    letter = token[0]
    size = len(token)
    if letter == "y":
        return f"{year % 100:02d}" if size <= 2 else f"{year:04d}"
    if letter == "m":
        if size == 1:
            return str(month)
        if size == 2:
            return f"{month:02d}"
        if size == 3:
            return MONTHS[month - 1][:3]
        if size == 5:
            return MONTHS[month - 1][0]
        return MONTHS[month - 1]
    if letter == "d":
        if size == 1:
            return str(day)
        if size == 2:
            return f"{day:02d}"
        if size == 3:
            return DAYS[weekday][:3]
        return DAYS[weekday]
    if letter == "h":
        shown = (hour % 12 or 12) if twelve else hour
        return str(shown) if size == 1 else f"{shown:02d}"
    if letter == "n":
        return str(minute) if size == 1 else f"{minute:02d}"
    if letter == "s":
        return str(second) if size == 1 else f"{second:02d}"
    return ""


def date_parts(serial: float, *, epoch_1904: bool = False) -> tuple[int, int, int, int, int, int] | None:
    """The year, month, day, hour, minute and second a serial shows as.

    Rounded to the second first, as a date format rounds it, and with the
    1900 system's quirks: serial 0 is 0 January 1900 and serial 60 its 29
    February. None for a serial no date format can show.
    """
    if serial < 0:
        return None
    total = int(_round(_EXACT.multiply(_held(serial), Decimal(86400)), 0))
    days, seconds = divmod(total, 86400)
    if days > (MAX_DATE_SERIAL_1904 if epoch_1904 else MAX_DATE_SERIAL):
        return None
    year, month, day = _calendar(days, epoch_1904)
    return year, month, day, seconds // 3600, seconds % 3600 // 60, seconds % 60


def _calendar(days: int, epoch_1904: bool) -> tuple[int, int, int]:
    """A serial day as year, month, day, with the 1900 system's quirks.

    Day 0 is 0 January 1900 and day 60 is 29 February 1900, a day that
    never happened and that Excel keeps for Lotus compatibility.
    """
    if epoch_1904:
        moment = dt.date(1904, 1, 1) + dt.timedelta(days=days)
        return moment.year, moment.month, moment.day
    if days == 0:
        return 1900, 1, 0
    if days == 60:
        return 1900, 2, 29
    if days < 60:
        moment = dt.date(1899, 12, 31) + dt.timedelta(days=days)
    else:
        moment = dt.date(1899, 12, 30) + dt.timedelta(days=days)
    return moment.year, moment.month, moment.day


__all__ = [
    "BUILTIN_DISPLAY_CODES",
    "DAYS",
    "GENERAL_WIDTH",
    "MAX_DATE_SERIAL",
    "MAX_DATE_SERIAL_1904",
    "MONTHS",
    "OVERFLOW",
    "SYSTEM_LONG_DATE",
    "SYSTEM_TIME",
    "Format",
    "Rule",
    "Section",
    "Token",
    "date_parts",
    "format_value",
    "general",
    "parse",
]
