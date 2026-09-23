"""Charts: what each one plots, read from its part, and chart sheets.

A chart on a worksheet is a graphic frame in the sheet's drawing, pointing
at a chart part; a chart sheet is a tab that holds nothing but a drawing
with one such frame. Everything a chart reads from the workbook is a
sheet-qualified reference in a ``<c:f>``: a series' name, its categories
and its values, and a title linked to a cell. Excel shows a series as one
``=SERIES(...)`` formula of those, and :attr:`ChartSeries.formula` gives
the same.

Measured, those references move as a cell's do when rows or columns are
inserted or deleted, become ``Data!#REF!`` when what they read is deleted
outright, and follow a sheet's rename; the moving is done in
:mod:`~pyofficeeditor.excel._rowcol`. Excel's object model goes on
reporting a deleted reference's old address until the file is reopened, so
the file is what the moving is held to.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from typing import TYPE_CHECKING

from pyofficeeditor._xml import Element, XmlDocument, local_name
from pyofficeeditor.exceptions import PackageError

if TYPE_CHECKING:
    from pyofficeeditor.excel.workbook import Workbook
    from pyofficeeditor.opc import OpcPackage

RT_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
RT_CHARTSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
RT_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"


@dataclass(frozen=True)
class ChartSeries:
    """One series: where its name, its categories and its values come from."""

    #: Its place in the chart's plotting order, counted from 1 as Excel
    #: counts it.
    order: int
    #: The cell its name comes from, as a sheet-qualified reference.
    name_reference: str | None = None
    #: Its name: the text last read from that cell, or one typed in.
    name: str | None = None
    #: The categories along the axis, or a scatter chart's x values.
    categories: str | None = None
    #: The values plotted, or a scatter chart's y values.
    values: str | None = None
    #: A bubble chart's bubble sizes.
    bubble_sizes: str | None = None

    @property
    def formula(self) -> str:
        """The series as Excel's formula bar shows it, ``=SERIES(name,
        categories,values,order)``."""
        if self.name_reference is not None:
            name = self.name_reference
        elif self.name is not None:
            name = '"' + self.name.replace('"', '""') + '"'
        else:
            name = ""
        parts = [name, self.categories or "", self.values or "", str(self.order)]
        if self.bubble_sizes is not None:
            parts.append(self.bubble_sizes)
        return "=SERIES(" + ",".join(parts) + ")"


@dataclass(frozen=True)
class Chart:
    """A chart: what kind it is, what it plots and what its title says."""

    #: The chart's name on its sheet, such as "Chart 1", or a chart
    #: sheet's name.
    name: str
    part_name: str
    #: Each kind of plot it holds, in order, as its element names them less
    #: "Chart": ``column`` and ``bar`` for the two directions of a bar
    #: chart, ``line``, ``pie``, ``scatter`` and the rest. A combination
    #: chart holds more than one.
    kinds: tuple[str, ...]
    series: tuple[ChartSeries, ...]
    #: The title's text, if it is text typed in.
    title: str | None = None
    #: The cell the title is linked to, if it is.
    title_reference: str | None = None
    #: Every reference the chart makes, in the order its part holds them:
    #: the title's, then each series' name, categories and values, and any
    #: an extension adds, such as a range data labels are taken from.
    references: tuple[str, ...] = ()


def read_chart(root: Element, *, name: str, part_name: str) -> Chart:
    """A chart from its part's root, ``<c:chartSpace>``."""
    chart = root.child("chart")
    references = tuple(element.text for element in root.descendants("f"))
    if chart is None:
        return Chart(name=name, part_name=part_name, kinds=(), series=(), references=references)
    title, title_reference = _title(chart.child("title"))
    kinds: list[str] = []
    series: list[ChartSeries] = []
    plot_area = chart.child("plotArea")
    for plot in [] if plot_area is None else list(plot_area.elements()):
        kind = local_name(plot.name)
        if not kind.endswith("Chart"):
            continue
        kinds.append(_kind(plot))
        series.extend(_series(entry) for entry in plot.children_named("ser"))
    series.sort(key=lambda one: one.order)
    return Chart(
        name=name,
        part_name=part_name,
        kinds=tuple(kinds),
        series=tuple(series),
        title=title,
        title_reference=title_reference,
        references=references,
    )


def charts_in_drawing(package: OpcPackage, drawing_part: str) -> list[Chart]:
    """The charts a drawing holds, in its order, each named as its frame is."""
    relationships = package.relationships(drawing_part)
    found: list[Chart] = []
    for frame in package.xml(drawing_part).root.descendants("graphicFrame"):
        reference = next(frame.descendants("chart"), None)
        identifier = None if reference is None else reference.get("r:id")
        if identifier is None:
            continue
        try:
            relationship = relationships.by_id(identifier)
        except PackageError:
            continue
        part = relationship.target_part
        if relationship.is_external or not package.has_part(part):
            continue
        naming = next(frame.descendants("cNvPr"), None)
        name = "" if naming is None else naming.get("name") or ""
        found.append(read_chart(package.xml(part).root, name=name, part_name=part))
    return found


class ChartSheet:
    """A tab that is one chart and holds no cells.

    It keeps its place among the sheets, and its name in formulas that
    count sheets by position, but it is not a
    :class:`~pyofficeeditor.excel.Worksheet`: it has no cells to read or
    rows to insert.
    """

    def __init__(self, workbook: Workbook, name: str, part_name: str, document: XmlDocument) -> None:
        self._workbook = workbook
        self._name = name
        self._part_name = part_name
        self._document = document

    @property
    def name(self) -> str:
        return self._name

    @property
    def part_name(self) -> str:
        return self._part_name

    @property
    def document(self) -> XmlDocument:
        return self._document

    @property
    def workbook(self) -> Workbook:
        return self._workbook

    @property
    def chart(self) -> Chart | None:
        """The chart the sheet shows, named as the sheet is."""
        package = self._workbook.package
        for relationship in package.relationships(self._part_name).by_type(RT_DRAWING):
            if relationship.is_external or not package.has_part(relationship.target_part):
                continue
            for chart in charts_in_drawing(package, relationship.target_part):
                return replace(chart, name=self._name)
        return None

    def record_rename(self, name: str) -> None:
        """Take the name the workbook has just given this sheet."""
        self._name = name

    def __repr__(self) -> str:
        return f"<ChartSheet {self._name!r}>"


def _kind(plot: Element) -> str:
    kind = local_name(plot.name)[: -len("Chart")]
    if kind in ("bar", "bar3D"):
        direction = plot.child("barDir")
        if direction is not None and direction.get("val") == "col":
            return "column" + kind[len("bar") :]
    return kind


def _series(entry: Element) -> ChartSeries:
    order_element = entry.child("order")
    if order_element is None:
        order_element = entry.child("idx")
    raw = None if order_element is None else order_element.get("val")
    order = int(raw) + 1 if raw is not None and raw.isdigit() else 1
    name_reference, name = _text_source(entry.child("tx"))
    return ChartSeries(
        order=order,
        name_reference=name_reference,
        name=name,
        categories=_reference(entry.child("cat")) or _reference(entry.child("xVal")),
        values=_reference(entry.child("val")) or _reference(entry.child("yVal")),
        bubble_sizes=_reference(entry.child("bubbleSize")),
    )


def _reference(container: Element | None) -> str | None:
    """The formula a series' data comes from, whichever kind of reference
    holds it."""
    if container is None:
        return None
    formula = next(container.descendants("f"), None)
    return None if formula is None else formula.text


def _text_source(container: Element | None) -> tuple[str | None, str | None]:
    """A name or title's reference, and its text: the text last read from
    the cell, or the text typed in."""
    if container is None:
        return None, None
    reference = container.child("strRef")
    if reference is not None:
        formula = reference.child("f")
        cached = next(reference.descendants("v"), None)
        return (None if formula is None else formula.text), (None if cached is None else cached.text)
    literal = container.child("v")
    if literal is not None:
        return None, literal.text
    rich = container.child("rich")
    if rich is not None:
        paragraphs = [
            "".join(run.text for run in paragraph.descendants("t")) for paragraph in rich.children_named("p")
        ]
        return None, "\n".join(paragraphs)
    return None, None


def _title(title: Element | None) -> tuple[str | None, str | None]:
    """A title's typed-in text and the cell it is linked to."""
    if title is None:
        return None, None
    reference, text = _text_source(title.child("tx"))
    return (None, reference) if reference is not None else (text, None)


__all__ = [
    "RT_CHART",
    "RT_CHARTSHEET",
    "Chart",
    "ChartSeries",
    "ChartSheet",
    "charts_in_drawing",
    "read_chart",
]
