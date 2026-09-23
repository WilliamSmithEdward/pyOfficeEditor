"""Named cell styles: Normal, Good, Heading 1 and the rest of Excel's
gallery, and the styles a workbook defines for itself.

A style is two entries that have to agree. ``<cellStyle name="Good"
xfId="11" builtinId="26"/>`` in ``cellStyles`` names an entry in
``cellStyleXfs``, which holds the style's number format, font, fill and
border the way any cell format does, and says with ``apply...="0"`` what
the style leaves out: Good sets a font and a fill, so it writes
``applyNumberFormat="0"`` and the like for the rest. A cell in a style
points at it with ``xfId`` and carries a copy of what the style sets.

Measured, with every style Excel lists applied by Excel:

- A workbook defines a built-in style only once something uses it, from
  a definition Excel keeps. The definitions below are those, as Excel
  wrote them. They do not follow the Normal style: with Normal in Arial
  10, Good is still in the theme's body font at 11 points and Title in
  its heading font at 18. A font names the theme's typeface and says
  which of the two it is, so it follows a change of theme.
- Applying a style replaces what it sets and keeps the rest. Good over a
  bold, centred cell showing two decimals gives Good's font and fill and
  keeps the two decimals and the centring.
- A cell's own ``apply...="1"`` marks what differs from its style.
- ``cellStyles`` is kept in name order; a new style's format goes at the
  end of ``cellStyleXfs``.
- Comma and Currency name the built-in formats 43, 41, 44 and 42, and
  write them out with the codes Excel shows in en-US.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

from pyofficeeditor.excel._formats import Border, CellFormat, Color, Fill, Font, Side

#: What a style can set, in the order ``CT_Xf`` names them.
Aspect = Literal["number_format", "font", "fill", "border", "alignment", "protection"]
ASPECTS: tuple[Aspect, ...] = ("number_format", "font", "fill", "border", "alignment", "protection")

#: The attribute of a style's format that says whether it sets an aspect.
APPLY_ATTRIBUTES: dict[Aspect, str] = {
    "number_format": "applyNumberFormat",
    "font": "applyFont",
    "fill": "applyFill",
    "border": "applyBorder",
    "alignment": "applyAlignment",
    "protection": "applyProtection",
}

#: The typefaces of Excel's own theme, for a workbook without one.
DEFAULT_HEADING_FONT = "Aptos Display"
DEFAULT_BODY_FONT = "Aptos Narrow"


@dataclass(frozen=True)
class CellStyle:
    """A named cell style, as the Cell Styles gallery lists it."""

    name: str
    format: CellFormat = field(default_factory=CellFormat)
    #: What the style sets. A cell given the style keeps its own for the rest.
    aspects: tuple[Aspect, ...] = ASPECTS
    #: Excel's number for one of its own styles, and None for a workbook's.
    builtin_id: int | None = None
    #: Whether the gallery leaves it out.
    hidden: bool = False


# ---------------------------------------------------------------------------
# Excel's own definitions
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class _Typeface:
    """A built-in style's font, before the theme names its typeface."""

    color: Color
    size: float = 11
    bold: bool = False
    italic: bool = False
    heading: bool = False

    def font(self, *, heading_font: str, body_font: str) -> Font:
        return Font(
            name=heading_font if self.heading else body_font,
            size=self.size,
            bold=self.bold,
            italic=self.italic,
            color=self.color,
            family=2,
            scheme="major" if self.heading else "minor",
        )


@dataclass(frozen=True)
class _Definition:
    builtin_id: int
    aspects: tuple[Aspect, ...]
    typeface: _Typeface
    fill: Fill = field(default_factory=Fill.none)
    border: Border = field(default_factory=Border)
    #: The built-in format id and the code it is written out with, if the
    #: style sets one.
    number_format: tuple[int, str] | None = None


def _rgb(value: str) -> Color:
    return Color(rgb=value)


def _theme(index: int, tint: float | None = None) -> Color:
    return Color(theme=index, tint=tint)


def _solid(color: Color, *, tinted: bool = False) -> Fill:
    # Excel writes the tinted accents with a system background colour and
    # the rest without one; both are kept as it writes them.
    return Fill(pattern="solid", foreground=color, background=Color(indexed=65) if tinted else None)


def _box(style: Literal["thin", "double"], color: Color) -> Border:
    side = Side(style=style, color=color)
    return Border(left=side, right=side, top=side, bottom=side)


_TEXT = _Typeface(color=_theme(1))
_FONT_AND_FILL: tuple[Aspect, ...] = ("font", "fill")
_FONT_FILL_BORDER: tuple[Aspect, ...] = ("font", "fill", "border")
_FONT_AND_BORDER: tuple[Aspect, ...] = ("font", "border")
_FONT: tuple[Aspect, ...] = ("font",)
_NUMBER: tuple[Aspect, ...] = ("number_format",)

#: The tints Excel gives the 20%, 40% and 60% accent styles, as it writes them.
_ACCENT_TINTS = {20: 0.79998168889431442, 40: 0.59999389629810485, 60: 0.39997558519241921}

_DEFINITIONS: dict[str, _Definition] = {
    "Comma": _Definition(3, _NUMBER, _TEXT, number_format=(43, '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)')),
    "Currency": _Definition(
        4, _NUMBER, _TEXT, number_format=(44, '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)')
    ),
    "Percent": _Definition(5, _NUMBER, _TEXT, number_format=(9, "0%")),
    "Comma [0]": _Definition(6, _NUMBER, _TEXT, number_format=(41, '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)')),
    "Currency [0]": _Definition(
        7, _NUMBER, _TEXT, number_format=(42, '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)')
    ),
    "Note": _Definition(10, ("fill", "border"), _TEXT, _solid(_rgb("FFFFFFCC")), _box("thin", _rgb("FFB2B2B2"))),
    "Warning Text": _Definition(11, _FONT, _Typeface(color=_rgb("FFFF0000"))),
    "Title": _Definition(15, _FONT, _Typeface(color=_theme(3), size=18, heading=True)),
    "Heading 1": _Definition(
        16, _FONT_AND_BORDER, _Typeface(color=_theme(3), size=15, bold=True),
        border=Border(bottom=Side(style="thick", color=_theme(4))),
    ),
    "Heading 2": _Definition(
        17, _FONT_AND_BORDER, _Typeface(color=_theme(3), size=13, bold=True),
        border=Border(bottom=Side(style="thick", color=_theme(4, 0.499984740745262))),
    ),
    "Heading 3": _Definition(
        18, _FONT_AND_BORDER, _Typeface(color=_theme(3), bold=True),
        border=Border(bottom=Side(style="medium", color=_theme(4, 0.39997558519241921))),
    ),
    "Heading 4": _Definition(19, _FONT, _Typeface(color=_theme(3), bold=True)),
    "Input": _Definition(
        20, _FONT_FILL_BORDER, _Typeface(color=_rgb("FF3F3F76")), _solid(_rgb("FFFFCC99")),
        _box("thin", _rgb("FF7F7F7F")),
    ),
    "Output": _Definition(
        21, _FONT_FILL_BORDER, _Typeface(color=_rgb("FF3F3F3F"), bold=True), _solid(_rgb("FFF2F2F2")),
        _box("thin", _rgb("FF3F3F3F")),
    ),
    "Calculation": _Definition(
        22, _FONT_FILL_BORDER, _Typeface(color=_rgb("FFFA7D00"), bold=True), _solid(_rgb("FFF2F2F2")),
        _box("thin", _rgb("FF7F7F7F")),
    ),
    "Check Cell": _Definition(
        23, _FONT_FILL_BORDER, _Typeface(color=_theme(0), bold=True), _solid(_rgb("FFA5A5A5")),
        _box("double", _rgb("FF3F3F3F")),
    ),
    "Linked Cell": _Definition(
        24, _FONT_AND_BORDER, _Typeface(color=_rgb("FFFA7D00")),
        border=Border(bottom=Side(style="double", color=_rgb("FFFF8001"))),
    ),
    "Total": _Definition(
        25, _FONT_AND_BORDER, _Typeface(color=_theme(1), bold=True),
        border=Border(top=Side(style="thin", color=_theme(4)), bottom=Side(style="double", color=_theme(4))),
    ),
    "Good": _Definition(26, _FONT_AND_FILL, _Typeface(color=_rgb("FF006100")), _solid(_rgb("FFC6EFCE"))),
    "Bad": _Definition(27, _FONT_AND_FILL, _Typeface(color=_rgb("FF9C0006")), _solid(_rgb("FFFFC7CE"))),
    "Neutral": _Definition(28, _FONT_AND_FILL, _Typeface(color=_rgb("FF9C5700")), _solid(_rgb("FFFFEB9C"))),
    "Explanatory Text": _Definition(53, _FONT, _Typeface(color=_rgb("FF7F7F7F"), italic=True)),
}

for _accent in range(1, 7):
    _theme_index = 3 + _accent
    _first = 29 + 4 * (_accent - 1)
    _DEFINITIONS[f"Accent{_accent}"] = _Definition(
        _first, _FONT_AND_FILL, _Typeface(color=_theme(0)), _solid(_theme(_theme_index))
    )
    for _offset, (_percent, _tint) in enumerate(_ACCENT_TINTS.items(), start=1):
        _DEFINITIONS[f"{_percent}% - Accent{_accent}"] = _Definition(
            _first + _offset, _FONT_AND_FILL, _TEXT, _solid(_theme(_theme_index, _tint), tinted=True)
        )

#: The name of each of Excel's own styles, by its number: Normal is 0.
BUILTIN_STYLE_NAMES: dict[int, str] = {0: "Normal", **{d.builtin_id: name for name, d in _DEFINITIONS.items()}}


def builtin_style(name: str, *, heading_font: str, body_font: str) -> tuple[CellStyle, int | None] | None:
    """One of Excel's own styles as Excel defines it in a workbook, and the
    built-in number format it names, or None for a name Excel does not
    have. Normal is not among them: every workbook has its own."""
    definition = _DEFINITIONS.get(_canonical(name))
    if definition is None:
        return None
    number = definition.number_format
    style = CellStyle(
        name=_canonical(name),
        format=CellFormat(
            number_format="General" if number is None else number[1],
            font=definition.typeface.font(heading_font=heading_font, body_font=body_font),
            fill=definition.fill,
            border=definition.border,
        ),
        aspects=definition.aspects,
        builtin_id=definition.builtin_id,
    )
    return style, None if number is None else number[0]


def is_builtin_name(name: str) -> bool:
    """Whether Excel has a style of its own by this name, Normal included."""
    return name.casefold() == "normal" or _canonical(name) in _DEFINITIONS


def _canonical(name: str) -> str:
    """A style's name as Excel spells it, which is how it is matched: Excel
    finds ``good`` as Good."""
    folded = name.casefold()
    for known in _DEFINITIONS:
        if known.casefold() == folded:
            return known
    return name


__all__ = [
    "APPLY_ATTRIBUTES",
    "ASPECTS",
    "BUILTIN_STYLE_NAMES",
    "DEFAULT_BODY_FONT",
    "DEFAULT_HEADING_FONT",
    "Aspect",
    "CellStyle",
    "builtin_style",
    "is_builtin_name",
]
