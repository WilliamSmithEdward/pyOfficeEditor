"""How a sheet prints: margins, orientation, headers, and what to repeat.

Five things here are not where a reader would first look for them, and each
was measured from a workbook Excel wrote rather than taken from the schema.

**Margins are inches.** ``<pageMargins left="0.25"/>`` against the points the
object model takes, so ``Application.InchesToPoints(0.25)`` through VBA
lands as ``0.25`` in the file. Writing points gives a quarter-inch margin
eighteen inches wide.

**Fitting to a page is a sheet property, not a page-setup one.**
``fitToWidth`` and ``fitToHeight`` sit on ``<pageSetup>`` but do nothing
until ``<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>`` says to use them.
Setting the two numbers alone leaves the sheet printing at its scale.

**A header is one string with the sections coded into it.** Excel writes
``&Lleft&Cmiddle&Rright`` in a single ``<oddHeader>``, so the three boxes of
the dialog are not three elements. ``&P``, ``&N``, ``&D`` and the rest are
the same language and are carried through as text: reading them is a
different job from placing them.

**Defaults are absent.** ``fitToWidth`` is missing when it is 1 and
``orientation`` when it is ``default``, so a reader expecting them present
reports the wrong setup for most sheets.

**The print area is a defined name.** ``_xlnm.Print_Area`` scoped to the
sheet, not an attribute, which is why it already moves when rows are
inserted. :class:`~pyofficeeditor.excel.Workbook` owns those, so the
worksheet's print-area property goes through them.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Literal

from pyofficeeditor._xml import Element

#: ``ST_Orientation``. ``default`` means Excel decides, and is unwritten.
Orientation = Literal["default", "portrait", "landscape"]

#: ``ST_PageOrder``.
PageOrder = Literal["downThenOver", "overThenDown"]

#: ``ST_CellComments``: how comments print, if at all.
CommentPrinting = Literal["none", "asDisplayed", "atEnd"]

#: ``ST_PrintError``: what a printed error cell shows.
ErrorPrinting = Literal["displayed", "blank", "dash", "NA"]

#: The paper sizes worth naming. ``paperSize`` is a number, and these are
#: the ones a caller is likely to want; any other number passes through.
PAPER_SIZES: dict[str, int] = {
    "letter": 1,
    "legal": 5,
    "executive": 7,
    "A3": 8,
    "A4": 9,
    "A5": 11,
    "B4": 12,
    "B5": 13,
    "tabloid": 3,
}

#: Excel's own defaults, which it writes into every new sheet.
DEFAULT_MARGINS = (0.7, 0.7, 0.75, 0.75, 0.3, 0.3)

_SECTION = re.compile(r"&([LCR])")


@dataclass(frozen=True)
class PageMargins:
    """The white space around the page, in inches.

    Inches because that is what the file holds. The object model takes
    points, and ``Application.InchesToPoints`` is the whole of the
    difference.
    """

    left: float = 0.7
    right: float = 0.7
    top: float = 0.75
    bottom: float = 0.75
    header: float = 0.3
    footer: float = 0.3

    @classmethod
    def points(
        cls,
        left: float = 50.4,
        right: float = 50.4,
        top: float = 54.0,
        bottom: float = 54.0,
        header: float = 21.6,
        footer: float = 21.6,
    ) -> PageMargins:
        """The same margins given in points, as the object model takes them."""
        return cls(*(value / 72 for value in (left, right, top, bottom, header, footer)))

    @classmethod
    def read(cls, element: Element) -> PageMargins:
        return cls(
            left=_as_float(element.get("left"), 0.7),
            right=_as_float(element.get("right"), 0.7),
            top=_as_float(element.get("top"), 0.75),
            bottom=_as_float(element.get("bottom"), 0.75),
            header=_as_float(element.get("header"), 0.3),
            footer=_as_float(element.get("footer"), 0.3),
        )

    def write(self) -> Element:
        # CT_PageMargins requires all six, so none is omitted as a default.
        return Element.create(
            "pageMargins",
            {
                "left": _number(self.left),
                "right": _number(self.right),
                "top": _number(self.top),
                "bottom": _number(self.bottom),
                "header": _number(self.header),
                "footer": _number(self.footer),
            },
        )


@dataclass(frozen=True)
class PageSetup:
    """Orientation, paper, scaling and page numbering.

    ``fit_to_width`` and ``fit_to_height`` do nothing on their own: see
    :attr:`Worksheet.fit_to_page`, which is the ``sheetPr`` flag that turns
    them on.
    """

    orientation: Orientation = "default"
    paper_size: int | None = None
    scale: int | None = None
    fit_to_width: int | None = None
    fit_to_height: int | None = None
    first_page_number: int | None = None
    use_first_page_number: bool = False
    page_order: PageOrder | None = None
    black_and_white: bool = False
    draft: bool = False
    cell_comments: CommentPrinting | None = None
    errors: ErrorPrinting | None = None
    horizontal_dpi: int | None = None
    vertical_dpi: int | None = None
    #: The relationship naming the saved printer settings. Carried so a
    #: sheet Excel set up keeps them; this library never makes one.
    printer_settings_id: str | None = None

    @property
    def paper(self) -> str | None:
        """The paper size's name, when it is one this knows."""
        for name, number in PAPER_SIZES.items():
            if number == self.paper_size:
                return name
        return None

    @classmethod
    def on(cls, paper: str, **rest: object) -> PageSetup:
        """A setup on a named paper size, such as ``A4`` or ``letter``."""
        number = PAPER_SIZES.get(paper)
        if number is None:
            raise ValueError(
                f"{paper!r} is not a paper size this names; expected one of "
                f"{', '.join(sorted(PAPER_SIZES))}, or pass paper_size with Excel's number."
            )
        return cls(paper_size=number, **rest)  # type: ignore[arg-type]

    @classmethod
    def read(cls, element: Element) -> PageSetup:
        return cls(
            orientation=_as_orientation(element.get("orientation")),
            paper_size=_as_int(element.get("paperSize")),
            scale=_as_int(element.get("scale")),
            fit_to_width=_as_int(element.get("fitToWidth")),
            fit_to_height=_as_int(element.get("fitToHeight")),
            first_page_number=_as_int(element.get("firstPageNumber")),
            use_first_page_number=element.get("useFirstPageNumber") in ("1", "true"),
            page_order=_literal(element.get("pageOrder")),
            black_and_white=element.get("blackAndWhite") in ("1", "true"),
            draft=element.get("draft") in ("1", "true"),
            cell_comments=_literal(element.get("cellComments")),
            errors=_literal(element.get("errors")),
            horizontal_dpi=_as_int(element.get("horizontalDpi")),
            vertical_dpi=_as_int(element.get("verticalDpi")),
            printer_settings_id=element.get("r:id"),
        )

    def write(self) -> Element:
        # The order CT_PageSetup declares, which Excel follows.
        element = Element.create("pageSetup")
        for name, value in (
            ("paperSize", self.paper_size),
            ("scale", self.scale),
            ("firstPageNumber", self.first_page_number),
            ("fitToWidth", self.fit_to_width),
            ("fitToHeight", self.fit_to_height),
        ):
            # 1 is the default for fitToWidth and fitToHeight, and Excel
            # leaves it out.
            if value is None:
                continue
            if name in ("fitToWidth", "fitToHeight") and value == 1:
                continue
            element.set(name, str(value))
        if self.page_order is not None:
            element.set("pageOrder", self.page_order)
        if self.orientation != "default":
            element.set("orientation", self.orientation)
        if self.use_first_page_number:
            element.set("useFirstPageNumber", "1")
        if self.black_and_white:
            element.set("blackAndWhite", "1")
        if self.draft:
            element.set("draft", "1")
        if self.cell_comments is not None:
            element.set("cellComments", self.cell_comments)
        if self.errors is not None:
            element.set("errors", self.errors)
        for name, value in (
            ("horizontalDpi", self.horizontal_dpi),
            ("verticalDpi", self.vertical_dpi),
        ):
            if value is not None:
                element.set(name, str(value))
        if self.printer_settings_id is not None:
            element.set("r:id", self.printer_settings_id)
        return element


@dataclass(frozen=True)
class PrintOptions:
    """What else appears on the page, and where it sits."""

    horizontal_centered: bool = False
    vertical_centered: bool = False
    headings: bool = False
    gridlines: bool = False

    @property
    def is_empty(self) -> bool:
        return not (
            self.horizontal_centered
            or self.vertical_centered
            or self.headings
            or self.gridlines
        )

    @classmethod
    def read(cls, element: Element) -> PrintOptions:
        return cls(
            horizontal_centered=element.get("horizontalCentered") in ("1", "true"),
            vertical_centered=element.get("verticalCentered") in ("1", "true"),
            headings=element.get("headings") in ("1", "true"),
            gridlines=element.get("gridLines") in ("1", "true"),
        )

    def write(self) -> Element:
        element = Element.create("printOptions")
        for name, flag in (
            ("horizontalCentered", self.horizontal_centered),
            ("verticalCentered", self.vertical_centered),
            ("headings", self.headings),
            ("gridLines", self.gridlines),
        ):
            if flag:
                element.set(name, "1")
        return element


@dataclass(frozen=True)
class HeaderFooterText:
    """The three boxes of one header or footer.

    They are one string in the file, split by ``&L``, ``&C`` and ``&R``.
    The other codes, ``&P`` for the page number and the rest, are carried
    through as written.
    """

    left: str = ""
    center: str = ""
    right: str = ""

    @property
    def is_empty(self) -> bool:
        return not (self.left or self.center or self.right)

    @classmethod
    def parse(cls, raw: str) -> HeaderFooterText:
        """Split Excel's single string into its three sections.

        Text before any marker belongs to the centre, which is where Excel
        puts an unmarked header.
        """
        if not raw:
            return cls()
        sections = {"L": "", "C": "", "R": ""}
        pieces = _SECTION.split(raw)
        leading = pieces[0]
        if leading:
            sections["C"] = leading
        for index in range(1, len(pieces) - 1, 2):
            marker = pieces[index]
            sections[marker] = sections[marker] + pieces[index + 1]
        return cls(left=sections["L"], center=sections["C"], right=sections["R"])

    def render(self) -> str:
        """The single string Excel stores, sections in L, C, R order."""
        out = ""
        for marker, text in (("L", self.left), ("C", self.center), ("R", self.right)):
            if text:
                out += f"&{marker}{text}"
        return out


@dataclass(frozen=True)
class HeaderFooter:
    """The headers and footers, including the alternates.

    Excel keeps a separate first page and separate even pages only when the
    matching flag is set, so an ``evenHeader`` with ``differentOddEven``
    unset is written and never shown.
    """

    odd_header: HeaderFooterText = HeaderFooterText()
    odd_footer: HeaderFooterText = HeaderFooterText()
    even_header: HeaderFooterText = HeaderFooterText()
    even_footer: HeaderFooterText = HeaderFooterText()
    first_header: HeaderFooterText = HeaderFooterText()
    first_footer: HeaderFooterText = HeaderFooterText()
    different_odd_even: bool = False
    different_first: bool = False
    scale_with_document: bool = True
    align_with_margins: bool = True

    @property
    def is_empty(self) -> bool:
        return all(
            text.is_empty
            for text in (
                self.odd_header,
                self.odd_footer,
                self.even_header,
                self.even_footer,
                self.first_header,
                self.first_footer,
            )
        ) and not (self.different_odd_even or self.different_first)

    @classmethod
    def read(cls, element: Element) -> HeaderFooter:
        def section(name: str) -> HeaderFooterText:
            child = element.child(name)
            return HeaderFooterText.parse("" if child is None else (child.text or ""))

        scale = element.get("scaleWithDoc")
        align = element.get("alignWithMargins")
        return cls(
            odd_header=section("oddHeader"),
            odd_footer=section("oddFooter"),
            even_header=section("evenHeader"),
            even_footer=section("evenFooter"),
            first_header=section("firstHeader"),
            first_footer=section("firstFooter"),
            different_odd_even=element.get("differentOddEven") in ("1", "true"),
            different_first=element.get("differentFirst") in ("1", "true"),
            scale_with_document=True if scale is None else scale in ("1", "true"),
            align_with_margins=True if align is None else align in ("1", "true"),
        )

    def write(self) -> Element:
        element = Element.create("headerFooter")
        if self.different_odd_even:
            element.set("differentOddEven", "1")
        if self.different_first:
            element.set("differentFirst", "1")
        if not self.scale_with_document:
            element.set("scaleWithDoc", "0")
        if not self.align_with_margins:
            element.set("alignWithMargins", "0")
        # CT_HeaderFooter declares them in this order.
        for name, text in (
            ("oddHeader", self.odd_header),
            ("oddFooter", self.odd_footer),
            ("evenHeader", self.even_header),
            ("evenFooter", self.even_footer),
            ("firstHeader", self.first_header),
            ("firstFooter", self.first_footer),
        ):
            if text.is_empty:
                continue
            node = Element.create(name)
            node.set_text(text.render())
            element.append(node)
        return element


def _as_float(raw: str | None, fallback: float) -> float:
    if raw is None:
        return fallback
    try:
        return float(raw)
    except ValueError:
        return fallback


def _as_int(raw: str | None) -> int | None:
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


def _as_orientation(raw: str | None) -> Orientation:
    return raw if raw else "default"  # type: ignore[return-value]


def _literal(raw: str | None) -> Any:
    """An enumerated attribute, carried as written.

    Validating it here would refuse a value a future Excel adds, and the
    value is written back exactly as it came in either way.
    """
    return raw


def _number(value: float) -> str:
    """A margin as Excel writes it: no trailing zeros."""
    if float(value) == int(value):
        return str(int(value))
    return repr(round(float(value), 10)).rstrip("0").rstrip(".")


__all__ = [
    "DEFAULT_MARGINS",
    "PAPER_SIZES",
    "CommentPrinting",
    "ErrorPrinting",
    "HeaderFooter",
    "HeaderFooterText",
    "Orientation",
    "PageMargins",
    "PageOrder",
    "PageSetup",
    "PrintOptions",
]
