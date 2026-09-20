"""Fonts, fills, borders and alignment, as values.

A cell does not carry its formatting. It carries an index::

    <c r="A1" s="1" t="s"><v>0</v></c>

and ``1`` indexes ``cellXfs`` in ``xl/styles.xml``, whose entry in turn
indexes a font, a fill, a border and a number format in four more tables::

    <cellXfs><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"
                 applyFont="1"/></cellXfs>
    <fonts><font><sz val="11"/>...</font>
           <font><b/><sz val="11"/>...</font></fonts>

Two consequences shape this module.

**Formatting is shared, so it cannot be edited in place.** Two hundred cells
may point at one ``xf``. Changing that entry restyles all of them, which is
almost never what a caller asking to embolden one cell meant. So formatting
is modelled as immutable values: read a cell's format, derive a new one, and
let :class:`~pyofficeeditor.excel.Styles` find or append the entry that
matches. Nothing existing is ever mutated.

**Setting one aspect must preserve the others.** A bold header that gains a
yellow fill has to stay bold. A caller who sets only the fill gets a format
carrying the cell's existing font, border and number format, which is why
these are whole values with functional updates rather than a bag of setters.

Everything ``CT_Font``, ``CT_Border``, ``CT_CellAlignment`` and
``CT_CellProtection`` can hold is modelled, so nothing is silently dropped.
A gradient fill is the one exception: it is kept as its own source text and
written back unchanged, because a caller has no way to ask for one here and
discarding it would be worse.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import Literal

from pyofficeeditor._xml import Element

#: Fill index 0 must be ``none`` and index 1 must be ``gray125``. Excel
#: writes both into every workbook whether or not anything uses them, and
#: repairs a file that is missing them, so a new fill is appended from index
#: 2 onward (measured: every Excel-authored fixture has exactly these two).
RESERVED_FILL_COUNT = 2

PATTERN_NONE = "none"
PATTERN_GRAY125 = "gray125"
PATTERN_SOLID = "solid"

#: ``ST_UnderlineValues``.
Underline = Literal["single", "double", "singleAccounting", "doubleAccounting", "none"]
#: ``ST_VerticalAlignRun``.
ScriptPosition = Literal["baseline", "superscript", "subscript"]
#: ``ST_HorizontalAlignment``.
HorizontalAlignment = Literal[
    "general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"
]
#: ``ST_VerticalAlignment``.
VerticalAlignment = Literal["top", "center", "bottom", "justify", "distributed"]
#: ``ST_BorderStyle``.
BorderStyle = Literal[
    "none",
    "thin",
    "medium",
    "dashed",
    "dotted",
    "thick",
    "double",
    "hair",
    "mediumDashed",
    "dashDot",
    "mediumDashDot",
    "dashDotDot",
    "mediumDashDotDot",
    "slantDashDot",
]

#: The order Excel writes a font's children in. ``CT_Font`` is a choice, so
#: any order validates, but matching Excel keeps a diff between a font this
#: library wrote and one Excel wrote down to the values.
_FONT_CHILD_ORDER = (
    "b",
    "i",
    "strike",
    "condense",
    "extend",
    "outline",
    "shadow",
    "u",
    "vertAlign",
    "sz",
    "color",
    "name",
    "family",
    "charset",
    "scheme",
)

#: ``CT_Border`` is a sequence, so these really are ordered. Excel writes the
#: legacy ``left``/``right`` rather than ``start``/``end``; both are read.
_BORDER_SIDE_ORDER = ("left", "right", "top", "bottom", "diagonal", "vertical", "horizontal")


@dataclass(frozen=True)
class Color:
    """A color, in whichever of the four ways the file spells it.

    Excel picks between them: a theme color for anything from the palette, an
    explicit ``rgb`` for a custom one, ``indexed`` in older files, and
    ``auto`` for "whatever the system says". They are kept apart rather than
    converted, because resolving a theme index to an RGB value needs the
    theme part and would make a round trip lossy.
    """

    rgb: str | None = None
    theme: int | None = None
    indexed: int | None = None
    auto: bool = False
    tint: float | None = None

    @classmethod
    def from_rgb(cls, rgb: str) -> Color:
        """A color from ``RRGGBB`` or ``AARRGGBB``.

        Excel stores eight hex digits with alpha first, and a six-digit value
        is taken as fully opaque.
        """
        cleaned = rgb.lstrip("#").upper()
        if len(cleaned) == 6:
            cleaned = "FF" + cleaned
        if len(cleaned) != 8 or any(c not in "0123456789ABCDEF" for c in cleaned):
            raise ValueError(f"{rgb!r} is not a hex color; use RRGGBB or AARRGGBB.")
        return cls(rgb=cleaned)

    @property
    def is_empty(self) -> bool:
        return self.rgb is None and self.theme is None and self.indexed is None and not self.auto

    @classmethod
    def read(cls, element: Element | None) -> Color | None:
        if element is None:
            return None
        theme = _read_int(element, "theme")
        indexed = _read_int(element, "indexed")
        tint = _read_float(element, "tint")
        return cls(
            rgb=element.get("rgb"),
            theme=theme,
            indexed=indexed,
            auto=element.get("auto") in ("1", "true"),
            tint=tint,
        )

    def write(self, name: str = "color") -> Element:
        element = Element.create(name)
        if self.auto:
            element.set("auto", "1")
        if self.indexed is not None:
            element.set("indexed", str(self.indexed))
        if self.rgb is not None:
            element.set("rgb", self.rgb)
        if self.theme is not None:
            element.set("theme", str(self.theme))
        if self.tint is not None:
            element.set("tint", _format_float(self.tint))
        return element


@dataclass(frozen=True)
class Font:
    """A cell's typeface."""

    name: str | None = None
    size: float | None = None
    bold: bool = False
    italic: bool = False
    strike: bool = False
    underline: Underline | None = None
    script: ScriptPosition | None = None
    color: Color | None = None
    family: int | None = None
    charset: int | None = None
    scheme: str | None = None
    condense: bool = False
    extend: bool = False
    outline: bool = False
    shadow: bool = False

    @classmethod
    def read(cls, element: Element) -> Font:
        underline: Underline | None = None
        u = element.child("u")
        if u is not None:
            raw = u.get("val")
            # <u/> with no val means single, which is the common spelling.
            underline = _as_underline(raw) if raw is not None else "single"
        script_element = element.child("vertAlign")
        return cls(
            name=_child_value(element, "name"),
            size=_child_float(element, "sz"),
            bold=_flag(element, "b"),
            italic=_flag(element, "i"),
            strike=_flag(element, "strike"),
            underline=underline,
            script=_as_script(script_element.get("val")) if script_element is not None else None,
            color=Color.read(element.child("color")),
            family=_child_int(element, "family"),
            charset=_child_int(element, "charset"),
            scheme=_child_value(element, "scheme"),
            condense=_flag(element, "condense"),
            extend=_flag(element, "extend"),
            outline=_flag(element, "outline"),
            shadow=_flag(element, "shadow"),
        )

    def write(self) -> Element:
        element = Element.create("font")
        children: dict[str, Element] = {}
        for name, flag in (
            ("b", self.bold),
            ("i", self.italic),
            ("strike", self.strike),
            ("condense", self.condense),
            ("extend", self.extend),
            ("outline", self.outline),
            ("shadow", self.shadow),
        ):
            if flag:
                children[name] = Element.create(name)
        if self.underline is not None:
            node = Element.create("u")
            if self.underline != "single":
                node.set("val", self.underline)
            children["u"] = node
        if self.script is not None:
            children["vertAlign"] = Element.create("vertAlign", {"val": self.script})
        if self.size is not None:
            children["sz"] = Element.create("sz", {"val": _format_float(self.size)})
        if self.color is not None and not self.color.is_empty:
            children["color"] = self.color.write()
        if self.name is not None:
            children["name"] = Element.create("name", {"val": self.name})
        if self.family is not None:
            children["family"] = Element.create("family", {"val": str(self.family)})
        if self.charset is not None:
            children["charset"] = Element.create("charset", {"val": str(self.charset)})
        if self.scheme is not None:
            children["scheme"] = Element.create("scheme", {"val": self.scheme})
        for name in _FONT_CHILD_ORDER:
            node = children.get(name)
            if node is not None:
                element.append(node)
        return element


@dataclass(frozen=True)
class Fill:
    """A cell's background.

    A solid fill is the ordinary case and the one a caller asks for. Excel
    stores its color in ``fgColor``, not ``bgColor``, which is the opposite
    of what the names suggest and a reliable way to produce a cell that looks
    unfilled.
    """

    pattern: str | None = None
    foreground: Color | None = None
    background: Color | None = None
    #: A gradient fill, kept as its own source text. Nothing here can build
    #: one, so it is written back exactly as it was read.
    gradient: str | None = None

    @classmethod
    def solid(cls, color: Color | str) -> Fill:
        """A plain block of one color."""
        resolved = Color.from_rgb(color) if isinstance(color, str) else color
        return cls(pattern=PATTERN_SOLID, foreground=resolved)

    @classmethod
    def none(cls) -> Fill:
        return cls(pattern=PATTERN_NONE)

    @property
    def is_gradient(self) -> bool:
        return self.gradient is not None

    @classmethod
    def read(cls, element: Element) -> Fill:
        gradient = element.child("gradientFill")
        if gradient is not None:
            return cls(gradient=gradient.to_xml())
        pattern = element.child("patternFill")
        if pattern is None:
            return cls()
        return cls(
            pattern=pattern.get("patternType"),
            foreground=Color.read(pattern.child("fgColor")),
            background=Color.read(pattern.child("bgColor")),
        )

    def write(self) -> Element:
        element = Element.create("fill")
        if self.gradient is not None:
            from pyofficeeditor._xml import XmlDocument

            element.append(XmlDocument.parse(self.gradient.encode()).root)
            return element
        pattern = Element.create("patternFill")
        if self.pattern is not None:
            pattern.set("patternType", self.pattern)
        if self.foreground is not None and not self.foreground.is_empty:
            pattern.append(self.foreground.write("fgColor"))
        if self.background is not None and not self.background.is_empty:
            pattern.append(self.background.write("bgColor"))
        element.append(pattern)
        return element


@dataclass(frozen=True)
class Side:
    """One edge of a border."""

    style: BorderStyle | None = None
    color: Color | None = None

    @property
    def is_empty(self) -> bool:
        return self.style is None or self.style == "none"

    @classmethod
    def read(cls, element: Element | None) -> Side:
        if element is None:
            return cls()
        raw = element.get("style")
        return cls(
            style=_as_border_style(raw) if raw is not None else None,
            color=Color.read(element.child("color")),
        )

    def write(self, name: str) -> Element:
        element = Element.create(name)
        if self.style is not None:
            element.set("style", self.style)
        if self.color is not None and not self.color.is_empty:
            element.append(self.color.write())
        return element


@dataclass(frozen=True)
class Border:
    """A cell's edges.

    Every side is always written, even an empty one: Excel's own borders
    carry ``<left/><right/><top/><bottom/><diagonal/>`` whether or not any
    has a style, and ``CT_Border`` is a sequence, so they go in order.
    """

    left: Side = field(default_factory=Side)
    right: Side = field(default_factory=Side)
    top: Side = field(default_factory=Side)
    bottom: Side = field(default_factory=Side)
    diagonal: Side = field(default_factory=Side)
    vertical: Side = field(default_factory=Side)
    horizontal: Side = field(default_factory=Side)
    diagonal_up: bool = False
    diagonal_down: bool = False
    outline: bool = False

    @classmethod
    def all_sides(cls, style: BorderStyle = "thin", color: Color | str | None = None) -> Border:
        """The same edge on all four sides."""
        resolved = Color.from_rgb(color) if isinstance(color, str) else color
        side = Side(style=style, color=resolved)
        return cls(left=side, right=side, top=side, bottom=side)

    @property
    def is_empty(self) -> bool:
        return all(
            getattr(self, name).is_empty for name in _BORDER_SIDE_ORDER
        ) and not (self.diagonal_up or self.diagonal_down or self.outline)

    @classmethod
    def read(cls, element: Element) -> Border:
        return cls(
            # start/end are the modern spellings of left/right.
            left=Side.read(element.child("left") or element.child("start")),
            right=Side.read(element.child("right") or element.child("end")),
            top=Side.read(element.child("top")),
            bottom=Side.read(element.child("bottom")),
            diagonal=Side.read(element.child("diagonal")),
            vertical=Side.read(element.child("vertical")),
            horizontal=Side.read(element.child("horizontal")),
            diagonal_up=element.get("diagonalUp") in ("1", "true"),
            diagonal_down=element.get("diagonalDown") in ("1", "true"),
            outline=element.get("outline") in ("1", "true"),
        )

    def write(self) -> Element:
        element = Element.create("border")
        if self.diagonal_up:
            element.set("diagonalUp", "1")
        if self.diagonal_down:
            element.set("diagonalDown", "1")
        if self.outline:
            element.set("outline", "1")
        for name in ("left", "right", "top", "bottom", "diagonal"):
            side: Side = getattr(self, name)
            element.append(side.write(name))
        for name in ("vertical", "horizontal"):
            side = getattr(self, name)
            if not side.is_empty:
                element.append(side.write(name))
        return element


@dataclass(frozen=True)
class Alignment:
    """How the value sits inside the cell."""

    horizontal: HorizontalAlignment | None = None
    vertical: VerticalAlignment | None = None
    wrap_text: bool = False
    shrink_to_fit: bool = False
    indent: int | None = None
    relative_indent: int | None = None
    text_rotation: int | None = None
    justify_last_line: bool = False
    reading_order: int | None = None

    @property
    def is_empty(self) -> bool:
        return self == Alignment()

    @classmethod
    def read(cls, element: Element | None) -> Alignment:
        if element is None:
            return cls()
        return cls(
            horizontal=_as_horizontal(element.get("horizontal")),
            vertical=_as_vertical(element.get("vertical")),
            wrap_text=element.get("wrapText") in ("1", "true"),
            shrink_to_fit=element.get("shrinkToFit") in ("1", "true"),
            indent=_read_int(element, "indent"),
            relative_indent=_read_int(element, "relativeIndent"),
            text_rotation=_read_int(element, "textRotation"),
            justify_last_line=element.get("justifyLastLine") in ("1", "true"),
            reading_order=_read_int(element, "readingOrder"),
        )

    def write(self) -> Element:
        element = Element.create("alignment")
        if self.horizontal is not None:
            element.set("horizontal", self.horizontal)
        if self.vertical is not None:
            element.set("vertical", self.vertical)
        if self.text_rotation is not None:
            element.set("textRotation", str(self.text_rotation))
        if self.wrap_text:
            element.set("wrapText", "1")
        if self.indent is not None:
            element.set("indent", str(self.indent))
        if self.relative_indent is not None:
            element.set("relativeIndent", str(self.relative_indent))
        if self.justify_last_line:
            element.set("justifyLastLine", "1")
        if self.shrink_to_fit:
            element.set("shrinkToFit", "1")
        if self.reading_order is not None:
            element.set("readingOrder", str(self.reading_order))
        return element


@dataclass(frozen=True)
class Protection:
    """Whether the cell resists editing once the sheet is protected."""

    locked: bool = True
    hidden: bool = False

    @property
    def is_default(self) -> bool:
        return self == Protection()

    @classmethod
    def read(cls, element: Element | None) -> Protection:
        if element is None:
            return cls()
        locked = element.get("locked")
        hidden = element.get("hidden")
        return cls(
            locked=True if locked is None else locked in ("1", "true"),
            hidden=False if hidden is None else hidden in ("1", "true"),
        )

    def write(self) -> Element:
        element = Element.create("protection")
        element.set("locked", "1" if self.locked else "0")
        element.set("hidden", "1" if self.hidden else "0")
        return element


@dataclass(frozen=True)
class CellFormat:
    """Everything a cell's ``s`` index resolves to.

    Derive a new one from an existing one rather than building from scratch,
    or the aspects you did not mention revert to the default::

        cell.format = cell.format.with_font(bold=True)
    """

    number_format: str = "General"
    font: Font = field(default_factory=Font)
    fill: Fill = field(default_factory=Fill)
    border: Border = field(default_factory=Border)
    alignment: Alignment = field(default_factory=Alignment)
    protection: Protection = field(default_factory=Protection)
    #: The named style this format builds on, an index into ``cellStyleXfs``.
    style_id: int = 0

    def with_font(self, **changes: object) -> CellFormat:
        """A copy with some of the font's aspects changed."""
        return replace(self, font=replace(self.font, **changes))  # type: ignore[arg-type]

    def with_fill(self, fill: Fill | Color | str) -> CellFormat:
        """A copy with a different fill.  A color means a solid fill."""
        resolved = fill if isinstance(fill, Fill) else Fill.solid(fill)
        return replace(self, fill=resolved)

    def with_border(self, border: Border) -> CellFormat:
        return replace(self, border=border)

    def with_alignment(self, **changes: object) -> CellFormat:
        """A copy with some of the alignment's aspects changed."""
        return replace(self, alignment=replace(self.alignment, **changes))  # type: ignore[arg-type]

    def with_number_format(self, code: str) -> CellFormat:
        return replace(self, number_format=code)


def _flag(parent: Element, name: str) -> bool:
    """A boolean font child, which is true when present unless it says
    otherwise: ``<b/>`` and ``<b val="1"/>`` both mean bold."""
    child = parent.child(name)
    if child is None:
        return False
    raw = child.get("val")
    return True if raw is None else raw in ("1", "true")


def _child_value(parent: Element, name: str) -> str | None:
    child = parent.child(name)
    return None if child is None else child.get("val")


def _child_int(parent: Element, name: str) -> int | None:
    raw = _child_value(parent, name)
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


def _child_float(parent: Element, name: str) -> float | None:
    raw = _child_value(parent, name)
    if raw is None:
        return None
    try:
        return float(raw)
    except ValueError:
        return None


def _read_int(element: Element, name: str) -> int | None:
    raw = element.get(name)
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


def _read_float(element: Element, name: str) -> float | None:
    raw = element.get(name)
    if raw is None:
        return None
    try:
        return float(raw)
    except ValueError:
        return None


def _format_float(value: float) -> str:
    """A number as Excel writes a style value: no trailing ``.0``."""
    if value == int(value):
        return str(int(value))
    return repr(value)


def _as_underline(raw: str) -> Underline | None:
    return raw if raw in ("single", "double", "singleAccounting", "doubleAccounting", "none") else None  # type: ignore[return-value]


def _as_script(raw: str | None) -> ScriptPosition | None:
    return raw if raw in ("baseline", "superscript", "subscript") else None  # type: ignore[return-value]


def _as_horizontal(raw: str | None) -> HorizontalAlignment | None:
    allowed = (
        "general",
        "left",
        "center",
        "right",
        "fill",
        "justify",
        "centerContinuous",
        "distributed",
    )
    return raw if raw in allowed else None  # type: ignore[return-value]


def _as_vertical(raw: str | None) -> VerticalAlignment | None:
    return raw if raw in ("top", "center", "bottom", "justify", "distributed") else None  # type: ignore[return-value]


def _as_border_style(raw: str) -> BorderStyle | None:
    allowed = (
        "none",
        "thin",
        "medium",
        "dashed",
        "dotted",
        "thick",
        "double",
        "hair",
        "mediumDashed",
        "dashDot",
        "mediumDashDot",
        "dashDotDot",
        "mediumDashDotDot",
        "slantDashDot",
    )
    return raw if raw in allowed else None  # type: ignore[return-value]


__all__ = [
    "PATTERN_GRAY125",
    "PATTERN_NONE",
    "PATTERN_SOLID",
    "RESERVED_FILL_COUNT",
    "Alignment",
    "Border",
    "BorderStyle",
    "CellFormat",
    "Color",
    "Fill",
    "Font",
    "HorizontalAlignment",
    "Protection",
    "ScriptPosition",
    "Side",
    "Underline",
    "VerticalAlignment",
]
