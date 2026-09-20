"""``xl/styles.xml``: the tables everything about a cell's look lives in.

A cell carries an index, never its formatting. Resolving what it looks like
means following that index through five tables::

    <c r="F2" s="2"><v>46037</v></c>
        s="2"          -> cellXfs[2]
        numFmtId="164" -> numFmts entry formatCode="yyyy\\-mm\\-dd"
        fontId="0"     -> fonts[0]
        fillId="0"     -> fills[0]
        borderId="0"   -> borders[0]

The number-format hop is the one that decides a cell's *type*, not just its
appearance: nothing else distinguishes ``46037`` from ``2026-01-15``, and a
reader that skips it reports the number. So that part of this module is
load-bearing for :mod:`pyofficeeditor.excel._values`, and the rest describes
appearance.

Writing works by reuse, never by mutation. An entry is shared, so changing
one repaints every cell pointing at it, which is almost never what a caller
asking to embolden one cell meant. ``ensure_*`` therefore finds a matching
entry or appends a new one, and existing entries are left exactly as they
were. That also keeps a part's bytes intact when nothing new was needed.

The component value objects live in :mod:`pyofficeeditor.excel._formats`.
"""

from __future__ import annotations

import re

from pyofficeeditor._xml import Element, XmlDocument
from pyofficeeditor.excel._dxf import Dxf
from pyofficeeditor.excel._formats import (
    PATTERN_GRAY125,
    PATTERN_NONE,
    Alignment,
    Border,
    CellFormat,
    Fill,
    Font,
    Protection,
)
from pyofficeeditor.excel._schema import STYLESHEET_CHILD_ORDER, insert_in_schema_order

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
    """The workbook's style tables, read and appended to.

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
        """How many entries ``cellXfs`` has, which is the range a cell's
        ``s`` may index."""
        return len(self._cell_formats())

    @property
    def font_count(self) -> int:
        return len(self._table("fonts"))

    @property
    def fill_count(self) -> int:
        return len(self._table("fills"))

    @property
    def border_count(self) -> int:
        return len(self._table("borders"))

    def number_format_id(self, style_index: int | None) -> int:
        """The ``numFmtId`` a cell's ``s`` attribute resolves to.

        A cell with no ``s`` is not unformatted: it uses ``cellXfs`` entry 0,
        so that is what ``None`` resolves to rather than a bare default. A
        cell whose ``s`` points past the table is treated the same way rather
        than raising, because Excel itself renders such a cell with the
        default format and refusing the file would be worse than agreeing
        with Excel.
        """
        formats = self._cell_formats()
        index = 0 if style_index is None else style_index
        if not 0 <= index < len(formats):
            return 0
        raw = formats[index].get("numFmtId")
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

    # ------------------------------------------------------------------
    # Whole cell formats
    # ------------------------------------------------------------------

    def _table(self, name: str) -> list[Element]:
        container = self._root.child(name)
        if container is None:
            return []
        return list(container.children_named(_ENTRY_NAMES[name]))

    def _ensure_table(self, name: str) -> Element:
        container = self._root.child(name)
        if container is None:
            container = Element.create(name, {"count": "0"})
            insert_in_schema_order(self._root, container, STYLESHEET_CHILD_ORDER)
        return container

    @staticmethod
    def _refresh_count(container: Element, entry: str) -> int:
        count = sum(1 for _ in container.children_named(entry))
        container.set("count", str(count))
        return count

    def font(self, index: int) -> Font:
        """The font at an index in the workbook's font table."""
        entries = self._table("fonts")
        if not 0 <= index < len(entries):
            return Font()
        return Font.read(entries[index])

    def fill(self, index: int) -> Fill:
        """The fill at an index in the workbook's fill table."""
        entries = self._table("fills")
        if not 0 <= index < len(entries):
            return Fill()
        return Fill.read(entries[index])

    def border(self, index: int) -> Border:
        """The border at an index in the workbook's border table."""
        entries = self._table("borders")
        if not 0 <= index < len(entries):
            return Border()
        return Border.read(entries[index])

    def cell_format(self, style_index: int | None) -> CellFormat:
        """Everything a cell's ``s`` attribute resolves to.

        A cell with no ``s`` is not unformatted: it uses ``cellXfs`` entry 0,
        which normally names the workbook's default font rather than nothing
        at all. Resolving ``None`` to an empty format instead would make
        "add bold to this cell" quietly change its typeface, since the new
        font would carry no name or size.
        """
        formats = self._cell_formats()
        index = 0 if style_index is None else style_index
        if not 0 <= index < len(formats):
            return CellFormat(number_format=self.number_format(None))
        entry = formats[index]
        return CellFormat(
            number_format=self.number_format(index),
            font=self.font(_index_of(entry, "fontId")),
            fill=self.fill(_index_of(entry, "fillId")),
            border=self.border(_index_of(entry, "borderId")),
            alignment=Alignment.read(entry.child("alignment")),
            protection=Protection.read(entry.child("protection")),
            style_id=_index_of(entry, "xfId"),
        )

    def ensure_font(self, font: Font) -> int:
        """The index of a font, appending it if the workbook lacks it."""
        container = self._ensure_table("fonts")
        entries = list(container.children_named("font"))
        for index, entry in enumerate(entries):
            if Font.read(entry) == font:
                return index
        container.append(font.write())
        self._refresh_count(container, "font")
        return len(entries)

    def ensure_fill(self, fill: Fill) -> int:
        """The index of a fill, appending it if the workbook lacks it.

        Indices 0 and 1 are reserved for ``none`` and ``gray125``. They are
        created only when the table is empty: shifting existing entries to
        make room would renumber every ``fillId`` in the workbook and repaint
        every cell, so a table that already has entries is left as its
        producer wrote it.
        """
        container = self._ensure_table("fills")
        entries = list(container.children_named("fill"))
        if not entries:
            for pattern in (PATTERN_NONE, PATTERN_GRAY125):
                container.append(Fill(pattern=pattern).write())
            entries = list(container.children_named("fill"))
            self._refresh_count(container, "fill")
        for index, entry in enumerate(entries):
            if Fill.read(entry) == fill:
                return index
        container.append(fill.write())
        self._refresh_count(container, "fill")
        return len(entries)

    def ensure_border(self, border: Border) -> int:
        """The index of a border, appending it if the workbook lacks it."""
        container = self._ensure_table("borders")
        entries = list(container.children_named("border"))
        if not entries:
            container.append(Border().write())
            entries = list(container.children_named("border"))
            self._refresh_count(container, "border")
        for index, entry in enumerate(entries):
            if Border.read(entry) == border:
                return index
        container.append(border.write())
        self._refresh_count(container, "border")
        return len(entries)

    def dxf_count(self) -> int:
        """How many differential formats the workbook has."""
        return len(self._table("dxfs"))

    def dxf(self, index: int) -> Dxf:
        """The differential format a ``dxfId`` points at.

        An index the table does not have returns an empty :class:`Dxf`
        rather than raising: a rule whose ``dxfId`` dangles formats nothing,
        which is what Excel shows, and refusing to read the workbook over it
        would be worse than reporting what it does.
        """
        entries = self._table("dxfs")
        if not 0 <= index < len(entries):
            return Dxf()
        return Dxf.read(entries[index])

    def ensure_dxf(self, dxf: Dxf) -> int:
        """The ``dxfId`` for a differential format, appending it if new.

        Unlike the fill and border tables this one reserves nothing: index 0
        is an ordinary entry, because no cell refers to a dxf by default.
        """
        container = self._ensure_table("dxfs")
        entries = list(container.children_named("dxf"))
        for index, entry in enumerate(entries):
            if Dxf.read(entry) == dxf:
                return index
        container.append(dxf.write())
        self._refresh_count(container, "dxf")
        return len(entries)

    def ensure_cell_format(self, wanted: CellFormat) -> int:
        """The ``s`` index for a whole format, building what is missing.

        The four component tables are filled first, then an existing
        ``cellXfs`` entry with the same components is reused. Reuse is the
        point: formatting is shared, so a hundred cells given the same format
        add one entry, and an existing entry is never modified because other
        cells may point at it.
        """
        number_format_id = self._ensure_format_id(wanted.number_format)
        font_id = self.ensure_font(wanted.font)
        fill_id = self.ensure_fill(wanted.fill)
        border_id = self.ensure_border(wanted.border)

        for index in range(len(self._cell_formats())):
            if self.cell_format(index) == wanted:
                return index

        container = self._ensure_table("cellXfs")
        entry = Element.create(
            "xf",
            {
                "numFmtId": str(number_format_id),
                "fontId": str(font_id),
                "fillId": str(fill_id),
                "borderId": str(border_id),
                "xfId": str(wanted.style_id),
            },
        )
        # The apply flags say this entry overrides its named style for that
        # aspect. Excel writes them, and some readers honour them, so a
        # format that sets a component says so.
        if number_format_id != 0:
            entry.set("applyNumberFormat", "1")
        if font_id != 0:
            entry.set("applyFont", "1")
        if fill_id != 0:
            entry.set("applyFill", "1")
        if border_id != 0:
            entry.set("applyBorder", "1")
        if not wanted.alignment.is_empty:
            entry.set("applyAlignment", "1")
        if not wanted.protection.is_default:
            entry.set("applyProtection", "1")
        # CT_Xf is a sequence: alignment, protection, extLst.
        if not wanted.alignment.is_empty:
            entry.append(wanted.alignment.write())
        if not wanted.protection.is_default:
            entry.append(wanted.protection.write())

        container.append(entry)
        return self._refresh_count(container, "xf") - 1

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


#: The child element name inside each of the stylesheet's tables.
_ENTRY_NAMES = {
    "numFmts": "numFmt",
    "fonts": "font",
    "fills": "fill",
    "borders": "border",
    "cellStyleXfs": "xf",
    "cellXfs": "xf",
    "cellStyles": "cellStyle",
    "dxfs": "dxf",
}


def _index_of(entry: Element, attribute: str) -> int:
    raw = entry.get(attribute)
    if raw is None:
        return 0
    try:
        return int(raw)
    except ValueError:
        return 0


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
