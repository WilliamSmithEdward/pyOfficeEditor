"""Number formats, because a date in a worksheet is just a number.

Excel stores ``2026-01-15`` as ``46037``. Nothing in the cell says it is a
date; the only clue is the style index, and following it takes three hops:

    <c r="F2" s="2"><v>46037</v></c>
        s="2"          -> styles.xml cellXfs[2]
        numFmtId="164" -> numFmts entry formatCode="yyyy\\-mm\\-dd"
        the code has date tokens, so the number is a date

A reader that skips those hops reports 46037, which is not a date read
badly but a wrong answer. This module is that lookup, plus enough writing to
make a date land as a date.

Scope: number formats only. Fonts, fills, borders and alignment are a later
increment; this does not pretend to model them, and it never disturbs them.
"""

from __future__ import annotations

import re

from pyofficeeditor._xml import Element, XmlDocument

#: Where Excel starts numbering custom formats.  0 to 163 are reserved for
#: the builtins, whether or not a given build defines them all.
FIRST_CUSTOM_NUMBER_FORMAT = 164

#: The builtin format codes ECMA-376 states outright.  The gaps (5 to 8,
#: 23 to 36, 41 to 44, 50 to 58) are locale-dependent and the spec does not
#: fix their codes, so they are absent here and handled by
#: :data:`_BUILTIN_DATE_IDS` instead.
BUILTIN_NUMBER_FORMATS: dict[int, str] = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
}

#: Builtin ids that mean a date or a time.  14 to 22 and 45 to 47 are the
#: ones with codes above.  27 to 36 and 50 to 58 are the locale-specific
#: date formats the spec reserves for East Asian builds: their codes are not
#: fixed, but every one of them is a date format, so a workbook using them
#: must still read as dates rather than as serial numbers.
_BUILTIN_DATE_IDS = frozenset(
    {*range(14, 23), *range(27, 37), *range(45, 48), *range(50, 59)}
)

#: A number format code this library writes for a plain date.  ISO order,
#: because it is unambiguous and locale-independent.
ISO_DATE_FORMAT = "yyyy-mm-dd"
#: Likewise for a date and time.
ISO_DATETIME_FORMAT = "yyyy-mm-dd hh:mm:ss"
#: Likewise for a time alone.
ISO_TIME_FORMAT = "hh:mm:ss"

_DATE_TOKENS = frozenset("ymdhs")
#: Bracketed sections that are elapsed-time tokens rather than colors or
#: locale ids: ``[h]:mm:ss`` measures duration.
_ELAPSED_TOKENS = frozenset({"h", "hh", "m", "mm", "s", "ss"})
_ELAPSED_BODY = re.compile(r"^(h+|m+|s+)$", re.IGNORECASE)


def format_tokens(code: str) -> set[str]:
    """The date and time letters a format code actually uses as tokens.

    The letters that matter are ``y m d h s``, but only outside the parts
    of a code that are literal text. A code carries literals three ways,
    and all three have to be skipped or a currency symbol becomes a month:

    - ``"..."`` a quoted run, so ``#,##0 "days"`` is not a date
    - ``\\x`` an escaped character, so ``yyyy\\-mm`` has literal hyphens
    - ``[...]`` a color, locale or condition, so ``[Red]0.00`` is not a date,
      while ``[h]:mm`` is, because an elapsed-time token uses brackets too

    Only the first section of a ``;``-separated code is examined, which is
    the positive-number form and the one that fixes the type.
    """
    section = _first_section(code)
    found: set[str] = set()
    index = 0
    while index < len(section):
        character = section[index]
        if character == '"':
            close = section.find('"', index + 1)
            index = len(section) if close < 0 else close + 1
            continue
        if character == "\\":
            index += 2
            continue
        if character == "[":
            close = section.find("]", index + 1)
            if close < 0:
                return found
            body = section[index + 1 : close]
            if _ELAPSED_BODY.match(body) and body.lower() in _ELAPSED_TOKENS:
                found.add(body[0].lower())
            index = close + 1
            continue
        lowered = character.lower()
        if lowered in _DATE_TOKENS:
            found.add(lowered)
        index += 1
    return found


def is_date_format(code: str) -> bool:
    """Whether a number format code renders a date, a time, or both."""
    return bool(format_tokens(code))


def has_calendar_tokens(code: str) -> bool:
    """Whether the code shows a calendar date, as opposed to a clock time.

    ``y`` and ``d`` are unambiguous. ``m`` is the awkward one: Excel reads it
    as minutes next to an hour or a second token and as a month otherwise,
    and reproducing that positional rule exactly is not worth it here. A
    code with ``m`` but no ``h`` or ``s`` is treated as a month, which is
    what ``mm`` and ``mmm`` mean in practice; with ``h`` or ``s`` present the
    clock reading wins.
    """
    tokens = format_tokens(code)
    if tokens & {"y", "d"}:
        return True
    return "m" in tokens and not (tokens & {"h", "s"})


def has_clock_tokens(code: str) -> bool:
    """Whether the code shows a time of day or an elapsed duration."""
    tokens = format_tokens(code)
    if tokens & {"h", "s"}:
        return True
    return "m" in tokens and not (tokens & {"y", "d"}) and not has_calendar_tokens(code)


def _first_section(code: str) -> str:
    """The positive-number section of a format code.

    Splitting on ``;`` is not a plain ``str.split``: a semicolon inside a
    quoted run or an escape is literal text.
    """
    out: list[str] = []
    index = 0
    while index < len(code):
        character = code[index]
        if character == '"':
            close = code.find('"', index + 1)
            end = len(code) if close < 0 else close + 1
            out.append(code[index:end])
            index = end
            continue
        if character == "\\":
            out.append(code[index : index + 2])
            index += 2
            continue
        if character == ";":
            break
        out.append(character)
        index += 1
    return "".join(out)


class Styles:
    """``xl/styles.xml``, as far as number formats go.

    Built over the part's tree, so every edit is scoped: adding a number
    format rewrites the ``numFmts`` and ``cellXfs`` elements and leaves the
    fonts, fills and borders byte-identical.
    """

    def __init__(self, document: XmlDocument) -> None:
        self._document = document
        self._root = document.root

    # ------------------------------------------------------------------
    # Reading
    # ------------------------------------------------------------------

    @property
    def custom_formats(self) -> dict[int, str]:
        """The workbook's own format codes, by id."""
        container = self._root.child("numFmts")
        if container is None:
            return {}
        found: dict[int, str] = {}
        for entry in container.children_named("numFmt"):
            raw_id = entry.get("numFmtId")
            code = entry.get("formatCode")
            if raw_id is None or code is None:
                continue
            try:
                found[int(raw_id)] = code
            except ValueError:
                continue
        return found

    def _cell_formats(self) -> list[Element]:
        container = self._root.child("cellXfs")
        return [] if container is None else list(container.children_named("xf"))

    @property
    def cell_format_count(self) -> int:
        return len(self._cell_formats())

    def number_format_id(self, style_index: int | None) -> int:
        """The ``numFmtId`` a cell's ``s`` attribute resolves to.

        A cell with no ``s`` uses format 0, General.  A cell whose ``s``
        points past the table is treated the same way rather than raising,
        because Excel itself renders such a cell with the default format and
        refusing to read the file would be worse than agreeing with Excel.
        """
        if style_index is None:
            return 0
        formats = self._cell_formats()
        if not 0 <= style_index < len(formats):
            return 0
        raw = formats[style_index].get("numFmtId")
        if raw is None:
            return 0
        try:
            return int(raw)
        except ValueError:
            return 0

    def number_format(self, style_index: int | None) -> str:
        """The format code a cell's ``s`` attribute resolves to.

        A locale-dependent builtin whose code the spec does not fix comes
        back as an empty string; :meth:`is_date` still answers correctly for
        those, which is what callers actually need.
        """
        format_id = self.number_format_id(style_index)
        custom = self.custom_formats.get(format_id)
        if custom is not None:
            return custom
        return BUILTIN_NUMBER_FORMATS.get(format_id, "")

    def is_date(self, style_index: int | None) -> bool:
        """Whether a cell with this style holds a date or a time.

        This is the question the cell layer asks of every number.
        """
        format_id = self.number_format_id(style_index)
        if format_id in _BUILTIN_DATE_IDS:
            return True
        return is_date_format(self.number_format(style_index))

    def shows_calendar(self, style_index: int | None) -> bool:
        """Whether the style shows a calendar date.

        A locale-dependent builtin has no code here, so it is answered from
        the id: every one of those reserved ids is a calendar format.
        """
        format_id = self.number_format_id(style_index)
        if format_id in _BUILTIN_DATE_IDS and not BUILTIN_NUMBER_FORMATS.get(format_id):
            return True
        return has_calendar_tokens(self.number_format(style_index))

    def shows_clock(self, style_index: int | None) -> bool:
        """Whether the style shows a time of day or a duration."""
        return has_clock_tokens(self.number_format(style_index))

    # ------------------------------------------------------------------
    # Writing
    # ------------------------------------------------------------------

    def ensure_number_format(self, code: str) -> int:
        """The ``s`` index for a cell formatted with ``code``, adding it if
        the workbook does not have it yet.

        Returns an index into ``cellXfs``. An existing cell format with the
        right ``numFmtId`` is reused, so writing a thousand dates adds one
        entry rather than a thousand. A new cell format copies the default
        one's font, fill and border so the cell does not also change
        appearance.
        """
        format_id = self._ensure_format_id(code)
        formats = self._cell_formats()
        for index, entry in enumerate(formats):
            if self.number_format_id(index) == format_id and _is_plain(entry):
                return index
        return self._append_cell_format(format_id)

    def _ensure_format_id(self, code: str) -> int:
        """The ``numFmtId`` for a format code, adding a ``numFmt`` if needed."""
        for builtin_id, builtin_code in BUILTIN_NUMBER_FORMATS.items():
            if builtin_code == code:
                return builtin_id
        existing = self.custom_formats
        for format_id, existing_code in existing.items():
            if existing_code == code:
                return format_id
        new_id = max([FIRST_CUSTOM_NUMBER_FORMAT - 1, *existing], default=0) + 1
        new_id = max(new_id, FIRST_CUSTOM_NUMBER_FORMAT)

        container = self._root.child("numFmts")
        if container is None:
            container = Element.create("numFmts", {"count": "0"})
            # numFmts is the first child of styleSheet when it is present.
            first = next(iter(self._root.elements()), None)
            if first is None:
                self._root.append(container)
            else:
                self._root.insert_before(first, container)
        container.append(
            Element.create("numFmt", {"numFmtId": str(new_id), "formatCode": code})
        )
        container.set("count", str(sum(1 for _ in container.children_named("numFmt"))))
        return new_id

    def _append_cell_format(self, format_id: int) -> int:
        container = self._root.child("cellXfs")
        if container is None:
            raise ValueError(
                "styles.xml has no cellXfs element, so there is no cell format table to add to."
            )
        template = next(container.children_named("xf"), None)
        attributes = {
            "numFmtId": str(format_id),
            "fontId": template.get("fontId", "0") or "0" if template else "0",
            "fillId": template.get("fillId", "0") or "0" if template else "0",
            "borderId": template.get("borderId", "0") or "0" if template else "0",
            "xfId": "0",
            "applyNumberFormat": "1",
        }
        container.append(Element.create("xf", attributes))
        count = sum(1 for _ in container.children_named("xf"))
        container.set("count", str(count))
        return count - 1


def _is_plain(entry: Element) -> bool:
    """Whether a cell format carries nothing but a number format.

    Reusing an entry that also applies a bold font would silently restyle
    the cell, so only the plain ones are candidates for reuse.
    """
    return not any(
        entry.get(name) not in (None, "0")
        for name in ("fontId", "fillId", "borderId", "applyFont", "applyFill", "applyBorder", "applyAlignment")
    )


__all__ = [
    "BUILTIN_NUMBER_FORMATS",
    "FIRST_CUSTOM_NUMBER_FORMAT",
    "ISO_DATETIME_FORMAT",
    "ISO_DATE_FORMAT",
    "ISO_TIME_FORMAT",
    "Styles",
    "format_tokens",
    "has_calendar_tokens",
    "has_clock_tokens",
    "is_date_format",
]
