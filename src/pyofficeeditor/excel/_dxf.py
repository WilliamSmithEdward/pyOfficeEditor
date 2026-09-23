"""A differential format: only what a conditional-formatting rule changes.

A ``dxf`` is not a :class:`CellFormat` with a different name. It is a patch
applied over whatever the cell already has, and four things about it are the
opposite of what the cell-level tables do. Each was measured against a
workbook Excel wrote, and each produces a file that opens and looks wrong
rather than one that fails.

**The fill colour is ``bgColor``.** A solid cell fill puts its colour in
``fgColor``, and this library says so in :class:`Fill`. A dxf fill does not:
Excel writes ``<fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill>
</fill>``, with no ``patternType`` at all. Writing ``fgColor`` here gives a
rule that matches and highlights nothing.

**Flags are three-valued.** In a cell font, ``<b/>`` means bold and no
``<b>`` means not bold. In a dxf, no ``<b>`` means *inherit*, and turning
bold off for matching cells needs ``<b val="0"/>`` written out. Excel emits
both together, ``<b val="0"/><i/>`` for a rule that italicises without
touching weight, so the two-valued reading loses the distinction between
"leave it alone" and "make it not bold".

**The number format is inline.** ``<numFmt numFmtId="14" formatCode="0.00%"/>``
carries its own ``formatCode``. The id is not a reference into the workbook's
``numFmts`` table and need not agree with it: 14 is the built-in
``m/d/yyyy``, sitting next to a percentage. Resolving the id gives a date
where the file says percent, so only ``formatCode`` is read.

**The border is partial.** ``CT_Border`` at cell level always writes all of
``<left><right><top><bottom><diagonal>``, empty or not, and :class:`Border`
does that. A dxf writes only the sides the rule sets, so a rule that colours
one edge writes one child.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from pyofficeeditor._xml import Element
from pyofficeeditor.excel._formats import (
    Alignment,
    Border,
    Color,
    Protection,
    Side,
    Underline,
    as_underline,
)
from pyofficeeditor.excel._xstring import decode, encode_attribute

#: The order ``CT_Dxf`` requires its children to appear in.
DXF_CHILD_ORDER: tuple[str, ...] = (
    "font",
    "numFmt",
    "fill",
    "alignment",
    "border",
    "protection",
    "extLst",
)

_DXF_BORDER_SIDES = ("left", "right", "top", "bottom", "diagonal", "vertical", "horizontal")

#: What to put in a new dxf's required ``numFmtId``. It is the first id in
#: Excel's custom range, which is what an inline format code is, and nothing
#: reads it. Spelled out here rather than imported from ``_styles``: that
#: module imports :class:`Dxf` from this one, and the pair would be a cycle.
_CUSTOM_NUMBER_FORMAT_ID = 164


def _tri_state(parent: Element, name: str) -> bool | None:
    """A dxf's boolean font child: present, present-and-off, or absent.

    ``None`` means the element is not there at all, which in a differential
    format means *inherit* rather than *false*.
    """
    child = parent.child(name)
    if child is None:
        return None
    raw = child.get("val")
    return True if raw is None else raw in ("1", "true")


def _write_tri_state(name: str, value: bool | None) -> Element | None:
    if value is None:
        return None
    element = Element.create(name)
    if not value:
        element.set("val", "0")
    return element


@dataclass(frozen=True)
class DxfFont:
    """The typeface changes a rule makes, and only those.

    Every flag is three-valued: ``None`` leaves the cell's own setting alone.
    """

    bold: bool | None = None
    italic: bool | None = None
    strike: bool | None = None
    underline: Underline | None = None
    color: Color | None = None

    @property
    def is_empty(self) -> bool:
        return (
            self.bold is None
            and self.italic is None
            and self.strike is None
            and self.underline is None
            and (self.color is None or self.color.is_empty)
        )

    @classmethod
    def read(cls, element: Element | None) -> DxfFont | None:
        if element is None:
            return None
        underline: Underline | None = None
        u = element.child("u")
        if u is not None:
            raw = u.get("val")
            underline = as_underline(raw) if raw is not None else "single"
        return cls(
            bold=_tri_state(element, "b"),
            italic=_tri_state(element, "i"),
            strike=_tri_state(element, "strike"),
            underline=underline,
            color=Color.read(element.child("color")),
        )

    def write(self) -> Element:
        element = Element.create("font")
        for name, value in (("b", self.bold), ("i", self.italic), ("strike", self.strike)):
            node = _write_tri_state(name, value)
            if node is not None:
                element.append(node)
        if self.underline is not None:
            node = Element.create("u")
            if self.underline != "single":
                node.set("val", self.underline)
            element.append(node)
        if self.color is not None and not self.color.is_empty:
            element.append(self.color.write())
        return element


@dataclass(frozen=True)
class DxfFill:
    """The background a rule paints.

    ``background`` is the one that shows. See the module docstring: a dxf
    puts its colour in ``bgColor`` where a cell fill uses ``fgColor``.
    """

    background: Color | None = None
    foreground: Color | None = None
    pattern: str | None = None

    @property
    def is_empty(self) -> bool:
        return (
            (self.background is None or self.background.is_empty)
            and (self.foreground is None or self.foreground.is_empty)
            and self.pattern is None
        )

    @classmethod
    def read(cls, element: Element | None) -> DxfFill | None:
        if element is None:
            return None
        pattern_fill = element.child("patternFill")
        if pattern_fill is None:
            return cls()
        return cls(
            background=Color.read(pattern_fill.child("bgColor")),
            foreground=Color.read(pattern_fill.child("fgColor")),
            pattern=pattern_fill.get("patternType"),
        )

    def write(self) -> Element:
        element = Element.create("fill")
        pattern_fill = Element.create("patternFill")
        if self.pattern is not None:
            pattern_fill.set("patternType", self.pattern)
        # fgColor first, which is the order CT_PatternFill declares even
        # though a dxf usually carries only the second.
        if self.foreground is not None and not self.foreground.is_empty:
            pattern_fill.append(self.foreground.write("fgColor"))
        if self.background is not None and not self.background.is_empty:
            pattern_fill.append(self.background.write("bgColor"))
        element.append(pattern_fill)
        return element


@dataclass(frozen=True)
class Dxf:
    """What a rule changes about a cell that matches it.

    Build one with :meth:`of` rather than by hand; the field types exist so
    a dxf read out of a file comes back exactly as it went in.
    """

    font: DxfFont | None = None
    number_format: str | None = None
    fill: DxfFill | None = None
    alignment: Alignment | None = None
    border: Border | None = None
    protection: Protection | None = None
    #: The id written beside the format code. ``CT_NumFmt`` requires the
    #: attribute, so one is always written, but it carries no meaning here
    #: and is kept only so a dxf read from a file goes back unchanged. It is
    #: out of the comparison deliberately: two rules that format the same way
    #: are the same rule, and letting an arbitrary id split them would append
    #: a duplicate entry for every rule added.
    number_format_id: int | None = field(default=None, compare=False)

    @classmethod
    def of(
        cls,
        *,
        fill: Color | str | None = None,
        color: Color | str | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        strike: bool | None = None,
        underline: Underline | None = None,
        number_format: str | None = None,
        border: Border | None = None,
    ) -> Dxf:
        """The shorthand a caller actually wants.

        ``fill`` is the background the rule paints and ``color`` the text
        colour, both accepting ``"FFC7CE"`` as well as a :class:`Color`.
        """
        font = DxfFont(
            bold=bold,
            italic=italic,
            strike=strike,
            underline=underline,
            color=_as_color(color),
        )
        background = _as_color(fill)
        return cls(
            font=None if font.is_empty else font,
            number_format=number_format,
            fill=None if background is None else DxfFill(background=background),
            border=border,
        )

    @property
    def is_empty(self) -> bool:
        return (
            (self.font is None or self.font.is_empty)
            and self.number_format is None
            and (self.fill is None or self.fill.is_empty)
            and self.alignment is None
            and (self.border is None or self.border.is_empty)
            and self.protection is None
        )

    @classmethod
    def read(cls, element: Element) -> Dxf:
        number_format = element.child("numFmt")
        alignment = element.child("alignment")
        protection = element.child("protection")
        border = element.child("border")
        return cls(
            font=DxfFont.read(element.child("font")),
            # Only formatCode decides the format. The id beside it is not a
            # reference into the workbook's numFmts table; see the module
            # docstring. It is carried so the entry writes back unchanged.
            number_format=None if number_format is None else _decoded(number_format.get("formatCode")),
            number_format_id=None if number_format is None else _read_id(number_format),
            fill=DxfFill.read(element.child("fill")),
            alignment=None if alignment is None else Alignment.read(alignment),
            border=None if border is None else Border.read(border),
            protection=None if protection is None else Protection.read(protection),
        )

    def write(self) -> Element:
        element = Element.create("dxf")
        if self.font is not None and not self.font.is_empty:
            element.append(self.font.write())
        if self.number_format is not None:
            # numFmtId is required by CT_NumFmt even though nothing reads it,
            # so one is always written: the id this dxf came in with, or the
            # start of the custom range, which is what an inline format code
            # is. Leaving it out produces a schema-invalid dxf.
            identifier = (
                _CUSTOM_NUMBER_FORMAT_ID
                if self.number_format_id is None
                else self.number_format_id
            )
            element.append(
                Element.create(
                    "numFmt",
                    {"numFmtId": str(identifier), "formatCode": encode_attribute(self.number_format)},
                )
            )
        if self.fill is not None and not self.fill.is_empty:
            element.append(self.fill.write())
        if self.alignment is not None:
            element.append(self.alignment.write())
        if self.border is not None and not self.border.is_empty:
            element.append(write_partial_border(self.border))
        if self.protection is not None:
            element.append(self.protection.write())
        return element


def write_partial_border(border: Border) -> Element:
    """A border with only the sides that were set.

    :meth:`Border.write` always emits all five of a cell border's sides,
    which is what ``CT_Border`` at cell level requires. Inside a dxf that
    would say the rule clears every edge it did not mention.
    """
    element = Element.create("border")
    if border.diagonal_up:
        element.set("diagonalUp", "1")
    if border.diagonal_down:
        element.set("diagonalDown", "1")
    if border.outline:
        element.set("outline", "1")
    for name in _DXF_BORDER_SIDES:
        side: Side = getattr(border, name)
        if not side.is_empty:
            element.append(side.write(name))
    return element


def _decoded(raw: str | None) -> str | None:
    return None if raw is None else decode(raw)


def _read_id(element: Element) -> int | None:
    raw = element.get("numFmtId")
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


def _as_color(value: Color | str | None) -> Color | None:
    if value is None:
        return None
    return Color.from_rgb(value) if isinstance(value, str) else value


__all__ = [
    "DXF_CHILD_ORDER",
    "Dxf",
    "DxfFill",
    "DxfFont",
    "write_partial_border",
]
