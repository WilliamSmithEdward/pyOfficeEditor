"""A worksheet, and the cells in it.

The part looks like this, and the two ordering rules are load-bearing::

    <sheetData>
      <row r="1" spans="1:6"><c r="A1" t="s"><v>0</v></c>...</row>
      <row r="2" spans="1:6">...</row>
    </sheetData>

Rows go in ascending ``r`` order and a row's cells in ascending column
order. Excel does not repair a worksheet that breaks either rule; it
refuses to open it. So writing a cell means finding its place, not
appending.

A cell that holds nothing is simply absent. ``sheet["Z99"]`` therefore
always returns a :class:`Cell`, and that cell creates its element only when
something is written to it. Reading a hundred empty cells adds nothing to
the file.

:class:`Worksheet` owns the part and does the work, addressed by
:class:`~pyofficeeditor.excel.CellRef`. :class:`Cell` is a view onto one
address that reads nicely; everything it does, the worksheet exposes too.
"""

from __future__ import annotations

import datetime as dt
import math
from collections.abc import Iterable, Iterator, Sequence
from dataclasses import dataclass, replace
from pathlib import Path
from typing import TYPE_CHECKING, Literal

from pyofficeeditor._xml import Element, XmlDocument, local_name
from pyofficeeditor.excel._calc.engine import Engine
from pyofficeeditor.excel._calc.values import Empty
from pyofficeeditor.excel._chartbuild import CHART_KINDS, CT_CHART, ChartKind, SeriesData, chart_frame, chart_part
from pyofficeeditor.excel._charts import RT_CHART, Chart, charts_in_drawing
from pyofficeeditor.excel._comments import (
    CT_COMMENTS,
    CT_THREADED_COMMENTS,
    EMPTY_COMMENTS,
    EMPTY_THREADED_COMMENTS,
    RT_COMMENTS,
    RT_THREADED_COMMENTS,
    Comment,
    ThreadedComment,
    delete_comment,
    delete_thread,
    is_empty,
    note_anchor,
    note_shapes,
    note_vml,
    person_id,
    placeholder_text,
    read_comments,
    read_threads,
    set_resolved,
    shape_count,
    shown_as,
    shows,
    thread_ids,
    with_note_shape_type,
    write_comment,
    write_thread,
)
from pyofficeeditor.excel._conditional import ConditionalFormatting, ConditionalRule
from pyofficeeditor.excel._dimensions import (
    Freeze,
    column_entry,
    format_width,
    isolate_column,
    says_nothing,
    sheet_standard_width,
)
from pyofficeeditor.excel._dxf import Dxf
from pyofficeeditor.excel._errorchecks import ErrorCheck, ErrorRule, IgnoredError, check_errors, read_ignored_errors
from pyofficeeditor.excel._filters import (
    AutoFilter,
    FilterCell,
    FilterColumn,
    FilterOutcome,
    decide,
    filter_column_element,
    keeps,
    read_auto_filter,
    read_filter_column,
    resolve,
)
from pyofficeeditor.excel._formats import (
    Alignment,
    Border,
    CellFormat,
    Color,
    Fill,
    Font,
)
from pyofficeeditor.excel._formulas import Deletion, quote_sheet_name, shared_formula_for
from pyofficeeditor.excel._hyperlinks import RT_HYPERLINK, Hyperlink
from pyofficeeditor.excel._names import FILTER_DATABASE, PRINT_AREA, PRINT_TITLES, DefinedName
from pyofficeeditor.excel._numfmt import format_value
from pyofficeeditor.excel._numfmt import parse as parse_format
from pyofficeeditor.excel._pagesetup import (
    HeaderFooter,
    PageMargins,
    PageSetup,
    PrintOptions,
)
from pyofficeeditor.excel._pictures import RT_IMAGE, image_info, picture_anchor
from pyofficeeditor.excel._pivots import PivotTable, read_pivot_tables
from pyofficeeditor.excel._placement import swallowed_nodes
from pyofficeeditor.excel._protection import SheetProtection
from pyofficeeditor.excel._reference import MAX_COLUMN, MAX_ROW, CellRef, RangeRef, column_letter
from pyofficeeditor.excel._richtext import TextRun, completed, read_runs, rich_entry, shown
from pyofficeeditor.excel._rowcol import (
    RT_DRAWING,
    RT_VML,
    delete_columns,
    delete_rows,
    insert_columns,
    insert_rows,
    related_parts,
)
from pyofficeeditor.excel._schema import (
    AUTO_FILTER_CHILD_ORDER,
    SHEET_PR_CHILD_ORDER,
    WORKSHEET_CHILD_ORDER,
    insert_in_schema_order,
)
from pyofficeeditor.excel._shapes import (
    CAPTIONED_CONTROLS,
    CT_CONTROL_PROPERTIES,
    CT_DRAWING,
    CT_VML,
    DEFAULT_ROW_POINTS,
    EMPTY_DRAWING,
    EMPTY_VML,
    FIRST_CONTROL_ID,
    NS_MARKUP_COMPATIBILITY,
    NS_SPREADSHEET_DRAWING,
    NS_X14,
    RT_CONTROL_PROPERTIES,
    FormControl,
    Shape,
    ShapeKind,
    SheetGrid,
    VmlControl,
    anchor_box,
    anchor_cells,
    anchors_in,
    check_control_kind,
    control_drawing,
    control_entry,
    control_properties,
    control_vml,
    emu,
    find_shape_element,
    has_vml_shape,
    is_two_cell,
    move_anchor,
    new_anchor,
    qualified_macro,
    read_control,
    read_drawing,
    replace_text,
    set_vml_macro,
    shape_copies,
    shape_names,
    update_vml_control,
    vml_blocks,
    vml_controls,
    vml_has_shapes,
    vml_shape_ids,
    with_vml_block,
    with_vml_shape,
    without_vml_shape,
)
from pyofficeeditor.excel._tables import (
    CT_TABLE,
    RT_TABLE,
    Table,
    TableStyle,
    build_table_part,
    insert_table_child,
    unique_column_names,
)
from pyofficeeditor.excel._validation import DataValidation
from pyofficeeditor.excel._values import (
    CellValue,
    datetime_to_serial,
    read_value,
    write_shared_index,
    write_value,
)
from pyofficeeditor.excel._xstring import decode, encode_text
from pyofficeeditor.exceptions import PackageError

#: What ``<sheet state=...>`` can say. A ``veryHidden`` sheet is not in
#: Excel's unhide list, so only code can bring it back.
SheetVisibility = Literal["visible", "hidden", "veryHidden"]

#: How deep an outline goes: eight levels on Excel's outline bar, the
#: first of them ungrouped.
MAX_OUTLINE_LEVEL = 7

if TYPE_CHECKING:
    from pyofficeeditor.excel.workbook import Workbook
    from pyofficeeditor.opc import Relationship


@dataclass(frozen=True)
class _SheetRecord:
    """What the sheet's own ``<control>`` or ``<oleObject>`` says about one
    object drawn in VML, whose drawing shape is only a hidden twin."""

    #: ``formControl``, ``activeX`` or ``oleObject``.
    kind: ShapeKind = "formControl"
    name: str = ""
    #: The procedure a click runs, as ``[0]!Clicked``.
    macro: str = ""
    alt_text: str = ""
    #: What a Forms control is wired to, read from its own part.
    control: FormControl | None = None
    #: The ``<controlPr><anchor>``, which places an ActiveX control that
    #: has no drawing twin.
    anchor: Element | None = None


def _members(shape: Shape) -> Iterator[Shape]:
    """A shape and, for a group, every shape inside it, at any depth."""
    yield shape
    for child in shape.children:
        yield from _members(child)


def _finished(shape: Shape, records: dict[int, _SheetRecord], facts: dict[int, VmlControl]) -> Shape:
    """A shape read from the drawing, with what the sheet and the VML know
    about each control or OLE object in it filled in, at any depth in a
    group."""
    children = tuple(_finished(child, records, facts) for child in shape.children)
    record = records.get(shape.shape_id)
    if record is None:
        return replace(shape, children=children) if shape.children else shape
    fact = facts.get(shape.shape_id, VmlControl())
    return replace(
        shape,
        kind=record.kind,
        macro=record.macro,
        control=record.control,
        alt_text=record.alt_text or shape.alt_text,
        hidden=fact.hidden,
        # The caption Excel draws is the VML's. A control this library made
        # has no text in its drawing twin, only there.
        text=shape.text or (fact.caption or ""),
        children=children,
    )


def _set_or_drop(element: Element, attribute: str, value: str) -> None:
    """Set an attribute, or drop it for an empty value, which is how Excel
    writes one that says nothing."""
    if value:
        element.set(attribute, value)
    else:
        element.unset(attribute)


def _inside_fallback(element: Element) -> bool:
    """Whether an element sits in an ``mc:Fallback``, the copy of a record
    kept for Excel versions that cannot read the ``mc:Choice`` beside it."""
    node = element.parent
    while node is not None:
        if local_name(node.name) == "Fallback":
            return True
        node = node.parent
    return False


def _shows_text(shape: Shape) -> bool:
    """Whether a shape has text to change: an AutoShape, a text box, and a
    Forms control that shows a caption."""
    if shape.kind in ("shape", "textBox"):
        return True
    return shape.kind == "formControl" and shape.control is not None and shape.control.kind in CAPTIONED_CONTROLS


def _relationship_ids(node: Element, ids: set[str]) -> set[str]:
    """Which of a part's relationship ids a node and its descendants use.

    Found by value rather than by attribute name, because a picture names
    its image in ``r:embed``, a chart in ``r:id``, a hyperlink in ``r:id``
    on ``a:hlinkClick`` and a SmartArt diagram in four attributes of its
    own, and an attribute list would miss the next one.
    """
    found = {value for value in node.attributes.values() if value in ids}
    for element in node.descendants():
        found.update(value for value in element.attributes.values() if value in ids)
    return found


def _set_transform(body: Element, box: tuple[float, float, float, float]) -> None:
    """Move a shape's own transform with its anchor.

    A chart's frame and a form control's twin carry an all-zero transform,
    which is Excel saying the anchor alone places them, and are left so.
    """
    if local_name(body.name) == "graphicFrame":
        return
    properties = body.child("spPr") or body.child("grpSpPr")
    transform = None if properties is None else properties.child("xfrm")
    offset = None if transform is None else transform.child("off")
    extent = None if transform is None else transform.child("ext")
    if offset is None or extent is None:
        return
    values = (offset.get("x"), offset.get("y"), extent.get("cx"), extent.get("cy"))
    if all(value in (None, "0") for value in values):
        return
    offset.set("x", str(emu(box[0])))
    offset.set("y", str(emu(box[1])))
    extent.set("cx", str(emu(box[2])))
    extent.set("cy", str(emu(box[3])))


def _format_dimension(value: float) -> str:
    """A width or height as Excel writes it: no trailing ``.0``."""
    if value == int(value):
        return str(int(value))
    return repr(value)


def _as_ranges(
    reference: str | RangeRef | Sequence[str | RangeRef],
) -> tuple[RangeRef, ...]:
    """One range, several, or a space-separated ``sqref`` string.

    A plain string may itself name several areas, because that is how an
    ``sqref`` spells a rule applied to a multi-area selection.
    """
    if isinstance(reference, RangeRef):
        return (reference.normalized,)
    if isinstance(reference, str):
        pieces = reference.split()
        if not pieces:
            return ()
        return tuple(RangeRef.parse(piece).normalized for piece in pieces)
    found: list[RangeRef] = []
    for item in reference:
        found.extend(_as_ranges(item))
    return tuple(found)


class Worksheet:
    """One worksheet of a workbook."""

    def __init__(
        self,
        workbook: Workbook,
        name: str,
        part_name: str,
        document: XmlDocument,
    ) -> None:
        self._workbook = workbook
        self._name = name
        self._part_name = part_name
        self._document = document
        self._root = document.root

        data = self._root.child("sheetData")
        if data is None:
            data = Element.create("sheetData")
            insert_in_schema_order(self._root, data, WORKSHEET_CHILD_ORDER)
        self._data: Element = data

        #: Every ``<row>``, by number, and the highest number among them,
        #: which is where a sheet written top to bottom adds the next.
        self._rows: dict[int, Element] = {}
        self._highest_row = 0
        self.reindex_rows()
        self._masters: dict[str, tuple[CellRef, str]] | None = None

    # ------------------------------------------------------------------
    # Identity
    # ------------------------------------------------------------------

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

    def rename(self, name: str) -> None:
        """Rename the sheet, as :meth:`Workbook.rename_sheet` does: the name
        is checked, the workbook's entry changes, and every formula and
        defined name that referred to the old name follows it."""
        if name != self._name:
            self._workbook.rename_sheet(self._name, name)

    def record_rename(self, name: str) -> None:
        """Take the name the workbook has just given this sheet.

        A sheet's name lives in the workbook part, not in its own, so this
        only updates what the object reports. :meth:`Workbook.rename_sheet`
        calls it once the entry has moved; to rename a sheet, call
        :meth:`rename`.
        """
        self._name = name

    # ------------------------------------------------------------------
    # Addressing
    # ------------------------------------------------------------------

    def cell(self, row: int, column: int) -> Cell:
        """The cell at a 1-based row and column."""
        return Cell(self, CellRef(row, column))

    def __getitem__(self, key: str) -> Cell:
        """``sheet["B2"]``.  Use :meth:`range` for a block."""
        return Cell(self, CellRef.parse(key))

    def __setitem__(self, key: str, value: CellValue) -> None:
        """``sheet["B2"] = 7``."""
        self.set_value(CellRef.parse(key), value)

    def range(self, reference: str) -> Range:
        """``sheet.range("A1:C3")``."""
        return Range(self, RangeRef.parse(reference))

    # ------------------------------------------------------------------
    # Values and formulas, by reference
    # ------------------------------------------------------------------

    def has_cell(self, reference: CellRef) -> bool:
        """Whether the part actually carries this cell."""
        return self._find_cell(reference) is not None

    def get_value(self, reference: CellRef) -> CellValue:
        """What a cell holds, as a Python value.

        A number under a date format comes back as a date, a shared string as
        its text, an error as :class:`~pyofficeeditor.excel.CellError`, and an
        absent or empty cell as ``None``.

        For a formula cell this is the result Excel last computed, because
        this library does not evaluate formulas. Change one of a formula's
        inputs and the cached result is stale until Excel reopens the file,
        which is why a modified workbook is saved with ``fullCalcOnLoad``.
        """
        element = self._find_cell(reference)
        if element is None:
            return None
        return read_value(
            element,
            shared_strings=self._workbook.shared_strings,
            styles=self._workbook.styles,
            epoch_1904=self._workbook.epoch_1904,
        )

    def set_value(self, reference: CellRef, value: CellValue) -> None:
        """Put a Python value in a cell, creating it if it is absent."""
        element = self._ensure_cell(reference)
        write_value(
            element,
            value,
            shared_strings=self._workbook.ensure_shared_strings(),
            styles=self._workbook.styles,
            epoch_1904=self._workbook.epoch_1904,
        )
        self._workbook.mark_values_changed()
        self._invalidate()

    def get_rich_text(self, reference: CellRef) -> tuple[TextRun, ...] | None:
        """A text cell's text as runs, each with the font Excel shows it
        in, or ``None`` for a cell that holds no text.

        Plain text is one run in the cell's font, and so is a run with no
        font of its own at the start, which is how Excel writes the first.
        A run with no font after one that has a font shows in the
        workbook's default font, not the cell's, and a run's font takes
        what it leaves unsaid from that default font too. Excel shows them
        that way; see :mod:`~pyofficeeditor.excel._richtext`.
        """
        element = self._find_cell(reference)
        if element is None:
            return None
        kind = element.get("t")
        container: Element | None = None
        if kind == "inlineStr":
            container = element.child("is")
        elif kind == "s":
            shared = self._workbook.shared_strings
            raw = element.child("v")
            if shared is None or raw is None or not raw.text.strip().isdigit():
                return None
            index = int(raw.text)
            if not 0 <= index < len(shared):
                return None
            container = shared.entry(index)
        if container is None:
            return None
        styles = self._workbook.styles
        default = styles.font(0) if styles is not None else Font()
        runs = read_runs(container)
        fonts = shown(runs, self.get_format(reference).font, default)
        return tuple(TextRun(text, font) for (text, _), font in zip(runs, fonts, strict=True))

    def set_rich_text(self, reference: CellRef, runs: Sequence[TextRun | str]) -> None:
        """Put text in several fonts in a cell, as Excel writes it.

        Each run is a :class:`TextRun`, or a plain string in the font the
        cell has before the call. A run's font takes the typeface, size and
        colour it leaves unsaid from that font too, so ``Font(bold=True)``
        is them in bold; bold, italic and the other on and off settings are
        the run's own. As Excel does it, the cell takes the first run's
        font, and the rest carry their own written out in full. One run is
        plain text in that font.
        """
        pieces = [TextRun(run) if isinstance(run, str) else run for run in runs]
        if not pieces or not "".join(piece.text for piece in pieces):
            raise ValueError("rich text needs some text; clear the cell to leave it empty.")
        current = self.get_format(reference)
        fonts = [completed(piece.font, current.font) for piece in pieces]
        if fonts[0] != current.font:
            self.set_format(reference, replace(current, font=fonts[0]))
        if len(pieces) == 1:
            self.set_value(reference, pieces[0].text)
            return
        entry = rich_entry(
            [(pieces[0].text, None)] + [(piece.text, font) for piece, font in zip(pieces[1:], fonts[1:], strict=True)]
        )
        index = self._workbook.ensure_shared_strings().index_for_entry(entry)
        write_shared_index(self._ensure_cell(reference), index)
        self._workbook.mark_values_changed()
        self._invalidate()

    def get_formula(self, reference: CellRef) -> str | None:
        """A cell's formula, without the leading ``=``.

        A cell in a shared-formula group carries no text of its own, so its
        formula is derived from the group's master by shifting the relative
        references. That derivation is why this exists rather than callers
        reading ``<f>`` themselves. The text is read as Excel spells it,
        ``_x005F_`` and all, and given back as the characters it spells.
        """
        element = self._find_cell(reference)
        if element is None:
            return None
        return self.formula_of(element, reference)

    def formula_of(self, element: Element, reference: CellRef) -> str | None:
        """The formula a ``<c>`` element of this sheet holds, as
        :meth:`get_formula` gives it, for a caller walking the elements."""
        formula = element.child("f")
        if formula is None:
            return None
        text = formula.text
        if text:
            return decode(text)
        if (formula.get("t") or "") != "shared":
            return None
        index = formula.get("si")
        if index is None:
            return None
        master = self._shared_master(index)
        if master is None:
            return None
        master_cell, master_text = master
        return decode(shared_formula_for(master_text, master_cell, reference))

    def cell_elements(self) -> Iterator[tuple[CellRef, Element]]:
        """Every ``<c>`` element the sheet has, in reading order, with its
        address. A cell whose ``r`` Excel could not have written is left
        out."""
        for number in sorted(self._rows):
            for element in self._rows[number].children_named("c"):
                raw = element.get("r")
                if raw is None:
                    continue
                try:
                    yield CellRef.parse(raw), element
                except ValueError:
                    continue

    def cell_element(self, reference: CellRef) -> Element:
        """The ``<c>`` element of a cell, created if the sheet has none."""
        return self._ensure_cell(reference)

    def set_formula(self, reference: CellRef, formula: str | None) -> None:
        """Put a formula in a cell, or remove the one it has.

        The cell's cached result goes too: a stale ``<v>`` beside a new
        ``<f>`` is the old answer, and Excel shows it until it recalculates.
        """
        element = self._ensure_cell(reference)
        self._drop_children(element, "f")
        self._workbook.mark_values_changed()

        if formula is None:
            self._invalidate()
            return

        self._drop_children(element, "v")
        element.unset("t")
        node = Element.create("f")
        node.set_text(encode_text(formula[1:] if formula.startswith("=") else formula))
        element.insert(0, node)
        self._invalidate()

    def evaluate(
        self,
        formula: str,
        at: str | CellRef = "A1",
        *,
        today: dt.date | None = None,
        now: dt.datetime | None = None,
    ) -> CellValue:
        """What ``formula`` gives in a cell of this sheet, without putting it
        there: ``sheet.evaluate("SUM(B2:B9)")``.

        The cell matters where a range stands for one value, as in
        ``=A1:A10*2``, which takes the row of ``at``, and to ROW() and the
        like. Formulas the workbook's cells hold are calculated as needed,
        not taken from their cached values. Raises ``FormulaSyntaxError``
        for text Excel would refuse, and ``UnsupportedFormulaError`` when
        the formula needs something the engine does not have.
        """
        cell = CellRef.parse(at) if isinstance(at, str) else at
        value = Engine(self._workbook, today=today, now=now).evaluate(formula, self._name, cell.row, cell.column)
        if isinstance(value, Empty):
            return 0
        if isinstance(value, float) and value.is_integer() and abs(value) < 2**53:
            # As the cell would read back once saved.
            return int(value)
        return value

    def error_checks(
        self,
        rules: Iterable[ErrorRule] | None = None,
        *,
        include_ignored: bool = False,
        today: dt.date | None = None,
    ) -> list[ErrorCheck]:
        """The cells Excel's error checking marks with a green triangle, and
        the rule that catches each: ``evalError``, ``numberStoredAsText``
        and the rest, named as the file names them.

        ``rules`` are the rules to check, by default those Excel checks
        unless told otherwise, which is all but ``emptyCellReference``. An
        error the file records as ignored is left out, as Excel hides its
        triangle, unless ``include_ignored`` is set. A formula's value is
        the one cached for it, so calculate a changed workbook first.
        ``today`` settles whether February 29 is a day this year, for text
        such as ``2/29``.
        """
        return check_errors(self, rules, include_ignored=include_ignored, today=today)

    @property
    def ignored_errors(self) -> list[IgnoredError]:
        """The cells whose errors the file records as ignored, and under
        which rules."""
        return read_ignored_errors(self)

    def get_text(self, reference: CellRef) -> str:
        """The text Excel shows for a cell: its value under its number format.

        This is ``Range.Text`` in a column wide enough to show all of it, so
        it depends on the cell and not on the column: General keeps its
        eleven characters where a narrow column would cut it to fewer, and a
        fill (``*x``), whose length is the column's, is left out. It is also
        the text a value filter matches, once its outer spaces are trimmed.

        A value the format cannot show, such as a negative date in the 1900
        date system, is ``########``. A formula cell shows its cached result,
        and an empty cell shows nothing.
        """
        element = self._find_cell(reference)
        if element is None:
            return ""
        # No styles, so a number stays the number Excel formats, not a date.
        value = read_value(
            element,
            shared_strings=self._workbook.shared_strings,
            styles=None,
            epoch_1904=self._workbook.epoch_1904,
        )
        styles = self._workbook.styles
        code = "General" if styles is None else styles.display_format(self.style_index(reference))
        return format_value(value, code, epoch_1904=self._workbook.epoch_1904)

    def style_index(self, reference: CellRef) -> int | None:
        """A cell's ``s`` attribute, an index into the workbook's cell
        formats."""
        element = self._find_cell(reference)
        if element is None:
            return None
        raw = element.get("s")
        if raw is None:
            return None
        try:
            return int(raw)
        except ValueError:
            return None

    def number_format(self, reference: CellRef) -> str:
        """The format code a cell is displayed with."""
        styles = self._workbook.styles
        return "" if styles is None else styles.number_format(self.style_index(reference))

    def get_format(self, reference: CellRef) -> CellFormat:
        """Everything about how a cell looks: number format, font, fill,
        border, alignment and protection.

        A cell with no format of its own reports the default rather than
        ``None``, so a caller can derive from it without a special case.
        """
        styles = self._workbook.styles
        if styles is None:
            return CellFormat()
        return styles.cell_format(self.style_index(reference))

    def set_format(self, reference: CellRef, wanted: CellFormat) -> None:
        """Give a cell a format, reusing an existing one where possible.

        Formatting is shared: many cells point at one entry, so nothing
        existing is modified. The workbook's tables gain whatever the format
        needs, and the cell's ``s`` is pointed at the entry that matches.
        """
        styles = self._workbook.styles
        if styles is None:
            raise ValueError(
                "this workbook has no styles part, so there is nowhere to record a format."
            )
        index = styles.ensure_cell_format(wanted)
        element = self._ensure_cell(reference)
        element.set("s", str(index))
        self._invalidate()

    def get_cell_style(self, reference: CellRef) -> str | None:
        """The name of a cell's style, Normal for a cell with none of its
        own, or None if the style it points at has no name."""
        styles = self._workbook.styles
        if styles is None:
            return None
        return styles.style_name(self.get_format(reference).style_id)

    def set_cell_style(self, reference: CellRef, name: str) -> None:
        """Give a cell a named style, as Excel does.

        Measured: the style's font, fill and the rest replace the cell's for
        what the style sets, and the cell keeps its own for the rest, so
        Good over a bold, centred cell showing two decimals gives Good's
        font and fill and keeps the decimals and the centring. One of
        Excel's own styles is defined the first time something uses it,
        from Excel's definition.
        """
        styles = self._workbook.styles
        if styles is None:
            raise ValueError("this workbook has no styles part, so it has no cell styles.")
        style_id = self._workbook.ensure_cell_style(name)
        style = styles.style_format(style_id)
        sets = styles.style_aspects(style_id)
        current = self.get_format(reference)
        self.set_format(
            reference,
            CellFormat(
                number_format=style.number_format if "number_format" in sets else current.number_format,
                font=style.font if "font" in sets else current.font,
                fill=style.fill if "fill" in sets else current.fill,
                border=style.border if "border" in sets else current.border,
                alignment=style.alignment if "alignment" in sets else current.alignment,
                protection=style.protection if "protection" in sets else current.protection,
                style_id=style_id,
            ),
        )

    def clear_cell(self, reference: CellRef) -> None:
        """Remove a cell, leaving the sheet as if it were never set."""
        self._remove_cell(reference)
        self._workbook.mark_values_changed()
        self._invalidate()

    # ------------------------------------------------------------------
    # Inserting rows and columns
    # ------------------------------------------------------------------

    def insert_rows(self, at: int, count: int = 1, *, copy_format: bool = True) -> None:
        """Insert empty rows, pushing everything at or below ``at`` down.

        Everything that records a cell address moves with the data: the
        cells themselves, every formula in the workbook that reads from this
        sheet, shared-formula groups, merged ranges, hyperlinks, tables, the
        sheet's filter and dimension, its page breaks, and the workbook's
        defined names.

        The new rows are formatted like the row above them, as Excel's
        Insert does: its height, style and outline level, each cell's style,
        and the conditional formats, data validations and sparklines it
        carries. Never its values, and never hidden. ``copy_format=False``
        inserts plain rows instead.

        Refused, as Excel refuses it, when it would cut through a pivot
        table.
        """
        insert_rows(self, at, count, copy_format=copy_format)
        self._workbook.mark_values_changed()

    def insert_columns(self, at: int, count: int = 1, *, copy_format: bool = True) -> None:
        """Insert empty columns, pushing everything at or right of ``at``
        over, formatted like the column to their left unless
        ``copy_format`` is false. The same shifting, copying and refusal as
        :meth:`insert_rows`."""
        insert_columns(self, at, count, copy_format=copy_format)
        self._workbook.mark_values_changed()

    def delete_rows(self, at: int, count: int = 1) -> None:
        """Delete rows, closing the gap behind them.

        Everything that referred to the deleted rows is repointed: a formula
        reading one of them becomes #REF!, while a range that only partly
        overlapped shrinks instead. Merges, tables, hyperlinks and defined
        names shrink or go the same way.

        A shape, chart, group or Forms control that moves and sizes with
        its cells goes with them when all of its rows do, as Excel deletes
        it, parts and all. One that moves without sizing, a picture as Excel
        inserts one, lands on the boundary at its own size, and one that
        does not move stays where it is.

        Refused, like an insertion, when the sheet carries something that
        addresses cells and this library cannot move. Also refused when it
        would remove a table's header row, since a table's column names come
        from there, and when it would take an ActiveX control or an embedded
        object whole, whose binary part this does not take apart.
        """
        swallowed = self._swallowed_by(Deletion.rows(at, count))
        delete_rows(self, at, count)
        self._remove_swallowed(swallowed)
        self._workbook.mark_values_changed()

    def delete_columns(self, at: int, count: int = 1) -> None:
        """Delete columns, closing the gap.  The same repointing and the same
        refusals as :meth:`delete_rows`, and the same shapes go.

        A filter criterion on a deleted column goes with it, and the filter
        is applied again, as Excel does, so the rows only it hid show. That
        holds for the sheet's filter and for each table's.
        """
        before = self.auto_filter
        tables_before = {table.name: table.auto_filter for table in self.tables}
        swallowed = self._swallowed_by(Deletion.columns(at, count))
        delete_columns(self, at, count)
        self._remove_swallowed(swallowed)
        self._workbook.mark_values_changed()
        after = self.auto_filter
        if before is not None and after is not None and before.filtering and len(after.columns) < len(before.columns):
            if after.filtering:
                self.apply_auto_filter()
            else:
                self._set_filter_mode(False)
                self._show_filtered_rows(RangeRef.parse(after.ref).normalized)
        for table in self.tables:
            earlier, now = tables_before.get(table.name), table.auto_filter
            if earlier is None or now is None or not earlier.filtering or len(now.columns) >= len(earlier.columns):
                continue
            if now.filtering:
                self.apply_table_filter(table.name)
            else:
                self._show_filtered_rows(self._table_filter_block(table))

    # ------------------------------------------------------------------
    # Column and row dimensions
    # ------------------------------------------------------------------

    def column_width(self, column: int) -> float | None:
        """A column's stored width, or ``None`` if it uses the default.

        This is the file's own unit, not the number Excel's Column Width
        dialog shows: a VBA ``ColumnWidth = 18`` stores ``18.6328125``. The
        unit counts ``0`` glyphs in the default font plus padding, so
        converting needs that font's maximum digit width in pixels, which is
        not in the file. Copying a width from one column to another, or
        reading one back, is exact; translating a number a person typed is
        not something this library can do honestly.
        """
        container = self._root.child("cols")
        if container is None:
            return None
        entry = column_entry(container, column)
        if entry is None:
            return None
        raw = entry.get("width")
        if raw is None:
            return None
        try:
            return float(raw)
        except ValueError:
            return None

    def set_column_width(self, column: int, width: float | None) -> None:
        """Set a column's width, or ``None`` to restore the standard width.

        ``customWidth`` goes with a width, as Excel writes it for one set by
        hand. Restoring the standard width drops the column's entry when it
        says nothing else, and otherwise writes the standard width into it:
        an entry with no width at all is a column of width 0.
        """
        container = self._ensure_cols()
        standard = self._standard_width()
        entry = isolate_column(container, column, width=standard)
        if width is None:
            entry.set("width", format_width(standard))
            entry.unset("customWidth")
            entry.unset("bestFit")
        else:
            if width < 0:
                raise ValueError(f"a column width cannot be negative; got {width}.")
            entry.set("width", _format_dimension(width))
            entry.set("customWidth", "1")
        self._tidy_cols(container)
        self._invalidate()

    def column_hidden(self, column: int) -> bool:
        container = self._root.child("cols")
        if container is None:
            return False
        entry = column_entry(container, column)
        return entry is not None and entry.get("hidden") in ("1", "true")

    def set_column_hidden(self, column: int, hidden: bool) -> None:
        container = self._ensure_cols()
        entry = isolate_column(container, column, width=self._standard_width())
        if hidden:
            entry.set("hidden", "1")
        else:
            entry.unset("hidden")
        self._tidy_cols(container)
        self._invalidate()

    def row_height(self, row: int) -> float | None:
        """A row's height in points, or ``None`` if it uses the default.

        Unlike a column width, this is exact: a height of 24 stores as 24.
        """
        element = self._rows.get(row)
        if element is None:
            return None
        raw = element.get("ht")
        if raw is None:
            return None
        try:
            return float(raw)
        except ValueError:
            return None

    def set_row_height(self, row: int, height: float | None) -> None:
        """Set a row's height in points, or ``None`` to restore the default.

        ``customHeight`` goes with it: without the flag Excel ignores the
        value, measured, where a width it honours either way.
        """
        element = self._ensure_row(row)
        if height is None:
            element.unset("ht")
            element.unset("customHeight")
        else:
            if height < 0:
                raise ValueError(f"a row height cannot be negative; got {height}.")
            element.set("ht", _format_dimension(height))
            element.set("customHeight", "1")
        self._invalidate()

    def row_hidden(self, row: int) -> bool:
        element = self._rows.get(row)
        return element is not None and element.get("hidden") in ("1", "true")

    def set_row_hidden(self, row: int, hidden: bool) -> None:
        element = self._ensure_row(row)
        if hidden:
            element.set("hidden", "1")
        else:
            element.unset("hidden")
        self._invalidate()

    @property
    def default_row_height(self) -> float | None:
        """The height rows use when they set none of their own."""
        element = self._root.child("sheetFormatPr")
        if element is None:
            return None
        raw = element.get("defaultRowHeight")
        if raw is None:
            return None
        try:
            return float(raw)
        except ValueError:
            return None

    def _ensure_cols(self) -> Element:
        container = self._root.child("cols")
        if container is None:
            container = Element.create("cols")
            insert_in_schema_order(self._root, container, WORKSHEET_CHILD_ORDER)
        return container

    def _tidy_cols(self, container: Element) -> None:
        """Drop entries that no longer say anything, and the block itself
        when nothing is left.  Excel omits an empty ``<cols/>``."""
        standard = self._standard_width()
        for entry in list(container.children_named("col")):
            if says_nothing(entry, standard):
                container.remove(entry)
        if next(container.children_named("col"), None) is None:
            self._root.remove(container)

    def _standard_width(self) -> float:
        """The width a column at this sheet's standard width stores, which
        an entry made for one has to carry."""
        return sheet_standard_width(self)

    # ------------------------------------------------------------------
    # Frozen panes
    # ------------------------------------------------------------------

    @property
    def freeze(self) -> Freeze:
        """Which rows and columns are pinned while the rest scrolls."""
        view = self._root.child("sheetViews")
        if view is None:
            return Freeze()
        first = view.child("sheetView")
        if first is None:
            return Freeze()
        return Freeze.read(first.child("pane"))

    def freeze_panes(self, reference: str | CellRef | None) -> Freeze:
        """Pin everything above and left of a cell.

        ``freeze_panes("B2")`` pins row 1 and column A, which is how Excel's
        own command is described. ``"A2"`` pins the first row only, ``"B1"``
        the first column only, and ``None`` or ``"A1"`` unfreezes.
        """
        if reference is None:
            wanted = Freeze()
        else:
            cell = CellRef.parse(reference) if isinstance(reference, str) else reference
            wanted = Freeze.at(cell)

        view = self._ensure_sheet_view()
        existing = view.child("pane")
        if existing is not None:
            view.remove(existing)
        if wanted.is_frozen:
            # CT_SheetView is a sequence and pane is its first child.
            view.insert(0, wanted.write())
        self._invalidate()
        return wanted

    def _ensure_sheet_view(self) -> Element:
        container = self._root.child("sheetViews")
        if container is None:
            container = Element.create("sheetViews")
            insert_in_schema_order(self._root, container, WORKSHEET_CHILD_ORDER)
        view = container.child("sheetView")
        if view is None:
            view = Element.create("sheetView", {"workbookViewId": "0"})
            container.append(view)
        return view

    # ------------------------------------------------------------------
    # Tables
    # ------------------------------------------------------------------

    @property
    def tables(self) -> list[Table]:
        """Every table on this sheet, in the order ``<tableParts>`` lists them."""
        container = self._root.child("tableParts")
        if container is None:
            return []
        package = self._workbook.package
        relationships = package.relationships(self._part_name)
        found: list[Table] = []
        for entry in container.children_named("tablePart"):
            relationship_id = entry.get("r:id") or entry.get("id")
            if relationship_id is None:
                continue
            try:
                part = relationships.by_id(relationship_id).target_part
            except PackageError:
                continue
            if not package.has_part(part):
                continue
            found.append(Table(self, part, package.xml(part)))
        return found

    def table(self, name: str) -> Table:
        """One table, by name."""
        for table in self.tables:
            if table.name.casefold() == name.casefold():
                return table
        available = ", ".join(t.name for t in self.tables) or "none"
        raise KeyError(f"{self._name!r} has no table {name!r}. It has: {available}")

    def add_table(
        self,
        name: str,
        reference: str | RangeRef,
        *,
        totals_row: bool = False,
        style: TableStyle | None = None,
    ) -> Table:
        """Turn a block into a table.

        The block's first row must already hold the column names, because a
        table's column names have to equal the text in its header cells and
        Excel reconciles the two by rewriting the part. A blank header is
        filled with ``Column1``-style names and a duplicate gets a digit,
        which is what Excel does rather than refusing the table.

        A headerless table is not supported: Excel makes one by inserting a
        row above the block, which shifts every row below it and every
        formula that referred to them. That is a different operation and
        belongs with row insertion.
        """
        block = (
            RangeRef.parse(reference) if isinstance(reference, str) else reference
        ).normalized
        self._workbook.check_new_table_name(name)

        for existing in self.tables:
            if existing.ref.intersects(block):
                raise ValueError(
                    f"{block.a1} overlaps the table {existing.name!r} at {existing.ref.a1}. "
                    f"Excel does not allow two tables to share a cell."
                )
        sheet_filter = self.auto_filter
        if sheet_filter is not None and RangeRef.parse(sheet_filter.ref).normalized.intersects(block):
            raise ValueError(
                f"{block.a1} overlaps the sheet's autofilter at {sheet_filter.ref}; Excel refuses a "
                "workbook where the two overlap. Take the filter off with clear_auto_filter() first."
            )

        header_row = block.top
        headers = [
            str(self.get_value(CellRef(header_row, column)) or "")
            for column in range(block.left, block.right + 1)
        ]
        names = unique_column_names(headers)
        for offset, column_name in enumerate(names):
            reference_cell = CellRef(header_row, block.left + offset)
            if self.get_value(reference_cell) != column_name:
                self.set_value(reference_cell, column_name)

        if totals_row and block.height < 3:
            raise ValueError(
                f"{block.a1} is {block.height} rows, which leaves no data between the "
                f"header and a totals row."
            )

        package = self._workbook.package
        identifier = self._workbook.next_table_id()
        part_name = self._workbook.free_table_part_name()
        document = build_table_part(
            identifier=identifier,
            name=name,
            ref=block,
            column_names=names,
            has_totals_row=totals_row,
            style=style or TableStyle(),
        )
        package.write(part_name, document.to_bytes(), content_type=CT_TABLE)
        relationship = package.relationships(self._part_name).add_part(RT_TABLE, part_name)

        container = self._root.child("tableParts")
        if container is None:
            container = Element.create("tableParts")
            insert_in_schema_order(self._root, container, WORKSHEET_CHILD_ORDER)
        container.append(Element.create("tablePart", {"r:id": relationship.id}))
        container.set("count", str(sum(1 for _ in container.children_named("tablePart"))))

        self._invalidate()
        return Table(self, part_name, package.xml(part_name))

    def remove_table(self, name: str) -> None:
        """Delete a table, leaving the cells and their values in place.

        Which is what Excel's "Convert to Range" does: the table stops being
        a table and the data stays.
        """
        table = self.table(name)
        package = self._workbook.package
        relationships = package.relationships(self._part_name)

        container = self._root.child("tableParts")
        if container is not None:
            for entry in list(container.children_named("tablePart")):
                relationship_id = entry.get("r:id") or entry.get("id")
                if relationship_id is None:
                    continue
                try:
                    if relationships.by_id(relationship_id).target_part != table.part_name:
                        continue
                except PackageError:
                    continue
                container.remove(entry)
                relationships.remove(relationship_id)
                break
            remaining = sum(1 for _ in container.children_named("tablePart"))
            if remaining:
                container.set("count", str(remaining))
            else:
                self._root.remove(container)

        package.remove_part(table.part_name)
        self._invalidate()

    # ------------------------------------------------------------------
    # Merged ranges
    # ------------------------------------------------------------------

    @property
    def merged_ranges(self) -> list[RangeRef]:
        """Every merged block on the sheet, in the order the part lists them."""
        container = self._root.child("mergeCells")
        if container is None:
            return []
        found: list[RangeRef] = []
        for entry in container.children_named("mergeCell"):
            reference = entry.get("ref")
            if reference is None:
                continue
            try:
                found.append(RangeRef.parse(reference).normalized)
            except ValueError:
                continue
        return found


    # ------------------------------------------------------------------
    # Protection, appearance and visibility
    # ------------------------------------------------------------------

    @property
    def protection(self) -> SheetProtection | None:
        """How the sheet is protected, or ``None`` when it is not."""
        element = self._root.child("sheetProtection")
        return None if element is None else SheetProtection.read(element)

    def protect(
        self,
        protection: SheetProtection | None = None,
        *,
        password: str | None = None,
    ) -> SheetProtection:
        """Lock the sheet.

        With no arguments this locks what Excel's own Protect locks. Pass a
        :class:`SheetProtection` to say what stays allowed.

        A password is a deterrent rather than a secret: the file carries a
        hash, and anything that can read the file can remove it.
        """
        resolved = protection if protection is not None else SheetProtection()
        if password is not None:
            resolved = resolved.with_password(password)
        existing = self._root.child("sheetProtection")
        element = resolved.write()
        if existing is not None:
            self._root.insert_before(existing, element)
            self._root.remove(existing)
        else:
            insert_in_schema_order(self._root, element, WORKSHEET_CHILD_ORDER)
        self._invalidate()
        return resolved

    def unprotect(self) -> bool:
        """Remove the protection, reporting whether there was any."""
        element = self._root.child("sheetProtection")
        if element is None:
            return False
        self._root.remove(element)
        self._invalidate()
        return True

    @property
    def tab_color(self) -> Color | None:
        """The colour of the sheet's tab, or ``None`` for the default."""
        properties = self._root.child("sheetPr")
        if properties is None:
            return None
        return Color.read(properties.child("tabColor"))

    @tab_color.setter
    def tab_color(self, value: Color | str | None) -> None:
        properties = self._sheet_properties()
        existing = properties.child("tabColor")
        if value is None:
            if existing is not None:
                properties.remove(existing)
                self._tidy_sheet_properties(properties)
                self._invalidate()
            return
        color = Color.from_rgb(value) if isinstance(value, str) else value
        element = color.write("tabColor")
        if existing is not None:
            properties.insert_before(existing, element)
            properties.remove(existing)
        else:
            # tabColor is the first child of CT_SheetPr.
            insert_in_schema_order(properties, element, SHEET_PR_CHILD_ORDER)
        self._invalidate()

    @property
    def show_gridlines(self) -> bool:
        """Whether the grid is drawn. Printing has its own setting."""
        return self._view_flag("showGridLines", default=True)

    @show_gridlines.setter
    def show_gridlines(self, value: bool) -> None:
        self._set_view_flag("showGridLines", value, default=True)

    @property
    def show_headings(self) -> bool:
        """Whether the row numbers and column letters are shown."""
        return self._view_flag("showRowColHeaders", default=True)

    @show_headings.setter
    def show_headings(self, value: bool) -> None:
        self._set_view_flag("showRowColHeaders", value, default=True)

    @property
    def zoom(self) -> int:
        """The view's zoom, as a percentage. 100 when unset."""
        view = self._first_view()
        if view is None:
            return 100
        raw = view.get("zoomScale")
        if raw is None:
            return 100
        try:
            return int(raw)
        except ValueError:
            return 100

    @zoom.setter
    def zoom(self, value: int) -> None:
        if not 10 <= value <= 400:
            raise ValueError(f"zoom {value} is outside 10 to 400, which is what Excel allows.")
        view = self._ensure_view()
        if value == 100:
            view.unset("zoomScale")
            view.unset("zoomScaleNormal")
        else:
            view.set("zoomScale", str(value))
            # Excel writes both, and the normal-view one is what it
            # restores when switching back from page-break preview.
            view.set("zoomScaleNormal", str(value))
        self._invalidate()

    @property
    def visible(self) -> SheetVisibility:
        """``visible``, ``hidden`` or ``veryHidden``.

        A ``veryHidden`` sheet is not in Excel's unhide list; only code can
        bring it back.
        """
        state = self._workbook.sheet_entry(self._name).get("state")
        return state if state in ("hidden", "veryHidden") else "visible"  # type: ignore[return-value]

    @visible.setter
    def visible(self, value: SheetVisibility) -> None:
        if value not in ("visible", "hidden", "veryHidden"):
            raise ValueError(
                f"{value!r} is not a visibility Excel has; expected visible, hidden "
                f"or veryHidden."
            )
        if value != "visible" and not self._workbook.has_another_visible_sheet(self._name):
            raise ValueError(
                f"{self._name!r} is the only visible sheet, and Excel refuses a workbook "
                f"in which every sheet is hidden. Show another sheet first."
            )
        entry = self._workbook.sheet_entry(self._name)
        if value == "visible":
            entry.unset("state")
        else:
            entry.set("state", value)
        self._workbook.mark_changed()

    def _sheet_properties(self) -> Element:
        properties = self._root.child("sheetPr")
        if properties is None:
            properties = Element.create("sheetPr")
            insert_in_schema_order(self._root, properties, WORKSHEET_CHILD_ORDER)
        return properties

    def _tidy_sheet_properties(self, properties: Element) -> None:
        if not properties.attributes and not any(
            isinstance(child, Element) for child in properties.children
        ):
            self._root.remove(properties)

    def _first_view(self) -> Element | None:
        container = self._root.child("sheetViews")
        return None if container is None else container.child("sheetView")

    def _ensure_view(self) -> Element:
        container = self._root.child("sheetViews")
        if container is None:
            container = Element.create("sheetViews")
            insert_in_schema_order(self._root, container, WORKSHEET_CHILD_ORDER)
        view = container.child("sheetView")
        if view is None:
            view = Element.create("sheetView", {"workbookViewId": "0"})
            container.append(view)
        return view

    def _view_flag(self, name: str, *, default: bool) -> bool:
        view = self._first_view()
        if view is None:
            return default
        raw = view.get(name)
        return default if raw is None else raw in ("1", "true")

    def _set_view_flag(self, name: str, value: bool, *, default: bool) -> None:
        view = self._ensure_view()
        if value == default:
            view.unset(name)
        else:
            view.set(name, "1" if value else "0")
        self._invalidate()




    # ------------------------------------------------------------------
    # Autofilter
    # ------------------------------------------------------------------

    @property
    def auto_filter(self) -> AutoFilter | None:
        """The sheet's autofilter, or ``None`` if it has none.

        A column filtered by colour, by icon or by markup this library does
        not model reads with an :class:`OpaqueCriterion`, carried exactly as
        found and written back unchanged.
        """
        element = self._root.child("autoFilter")
        return None if element is None else read_auto_filter(element)

    def set_auto_filter(
        self,
        reference: str | RangeRef,
        columns: Sequence[FilterColumn] = (),
        *,
        apply: bool = True,
        today: dt.date | None = None,
    ) -> FilterOutcome:
        """Filter a range, and hide the rows its criteria exclude.

        What Excel's ``Range.AutoFilter`` does, down to what it stores: the
        range stops at the last row holding data; a top ten or an average is
        worked out from its column and a relative date period from
        ``today``, the local date unless given; a column whose criterion did
        not change keeps its markup exactly. Filtering a different range
        takes the old filter off first and shows its rows, as Excel does.

        The rows are hidden here because Excel does not re-apply a filter
        when it opens a workbook: it trusts the hidden flags. ``apply=False``
        records the criteria and touches no row. A filter with no criteria,
        only the dropdown arrows, touches no row either, as in Excel.

        Refused, with nothing changed: a column outside the range or given
        twice, a range overlapping a table, and a top ten or an average over
        a column holding an error, all of which Excel refuses too.
        """
        block = self._filter_block(reference)
        entries = list(columns)
        self._check_filter_columns(entries, block)
        entries.sort(key=lambda entry: entry.column)
        for table in self.tables:
            if table.ref.intersects(block):
                raise ValueError(
                    f"{block.a1} overlaps the table {table.name!r} at {table.ref.a1}; Excel refuses a "
                    "workbook where a sheet's filter and a table overlap. A table filters itself."
                )
        when = today or dt.date.today()
        resolved = self._resolve_columns(block, entries, when, strict=True)
        filters = AutoFilter(ref=block.a1, columns=tuple(resolved))

        # Everything that can refuse has: from here on the sheet changes.
        element = self._root.child("autoFilter")
        previous = None if element is None else read_auto_filter(element)
        if element is not None and previous is not None and previous.ref != block.a1:
            self._take_filter_off(element, previous)
            element, previous = None, None
        if element is None:
            element = Element.create("autoFilter", {"ref": block.a1})
            insert_in_schema_order(self._root, element, WORKSHEET_CHILD_ORDER)
        self._write_filter_columns(element, resolved)
        self._set_filter_mode(filters.filtering)
        self._put_builtin(FILTER_DATABASE, f"{quote_sheet_name(self._name)}!{block.absolute.a1}", hidden=True)

        outcome = FilterOutcome()
        if apply and filters.filtering:
            outcome = self._apply_filter(filters, when)
        elif apply and previous is not None and previous.filtering:
            outcome = FilterOutcome(shown=self._show_filtered_rows(block))
        self._invalidate()
        return outcome

    def apply_auto_filter(self, *, today: dt.date | None = None) -> FilterOutcome:
        """Apply the sheet's filter again, as Excel's ``ApplyFilter`` does.

        Every row of the range is decided afresh. A top ten, an average and
        a relative date period are worked out again from the column and
        ``today`` and stored; over a column holding an error Excel cannot
        rank or average, so a top ten keeps every row and an average uses
        what was stored.
        """
        element = self._root.child("autoFilter")
        if element is None:
            raise ValueError(f"{self._name!r} has no autofilter to apply.")
        current = read_auto_filter(element)
        block = RangeRef.parse(current.ref).normalized
        when = today or dt.date.today()
        refreshed = self._resolve_columns(block, current.columns, when, strict=False)
        self._write_filter_columns(element, refreshed)
        filters = AutoFilter(ref=current.ref, columns=tuple(refreshed))
        self._set_filter_mode(filters.filtering)
        outcome = self._apply_filter(filters, when) if filters.filtering else FilterOutcome()
        self._invalidate()
        return outcome

    def clear_auto_filter(self, *, show_rows: bool = True) -> None:
        """Take the filter off, and show the rows it was hiding.

        As Excel does: every row of the range shows, rows a collapsed
        outline was hiding included, and ``_xlnm._FilterDatabase`` stays
        defined. ``show_rows=False`` leaves the rows as they are.
        """
        element = self._root.child("autoFilter")
        if element is None:
            return
        previous = read_auto_filter(element)
        if show_rows:
            self._take_filter_off(element, previous)
        else:
            self._root.remove(element)
        self._set_filter_mode(False)
        self._invalidate()

    def set_table_filter(
        self,
        name: str,
        columns: Sequence[FilterColumn] = (),
        *,
        apply: bool = True,
        today: dt.date | None = None,
    ) -> FilterOutcome:
        """Filter a table, and hide the rows its criteria exclude.

        A table has a filter of its own, over its header and data rows with
        the totals row left out, and a column counts from the table's first
        column. Otherwise this is :meth:`set_auto_filter`: the same
        criteria, worked out and stored the same way, the same refusals
        before anything changes, and the same rows hidden. As in Excel it
        turns the table's dropdowns on if they were off, sets no
        ``filterMode`` and defines no ``_xlnm._FilterDatabase``. With no
        columns it is Excel's Clear: the dropdowns stay and every row
        shows.

        A table with no header row is refused: the dropdowns sit on it, and
        Excel turns the header row back on to filter, which this library
        does not do.
        """
        return self._filter_table(self.table(name), columns, apply=apply, today=today, strict=True)

    def apply_table_filter(self, name: str, *, today: dt.date | None = None) -> FilterOutcome:
        """Apply a table's filter again, as :meth:`apply_auto_filter` does
        for the sheet's."""
        table = self.table(name)
        current = table.auto_filter
        if current is None:
            raise ValueError(f"the table {table.name!r} has its dropdowns off, so no filter to apply.")
        return self._filter_table(table, current.columns, apply=True, today=today, strict=False)

    def clear_table_filter(self, name: str, *, show_rows: bool = True) -> None:
        """Turn a table's dropdowns off, and show the rows its filter hid.

        What unticking Filter Button does in Excel. To keep the dropdowns
        and drop only the criteria, call :meth:`set_table_filter` with no
        columns. ``show_rows=False`` leaves the rows as they are.
        """
        table = self.table(name)
        root = table.document.root
        element = root.child("autoFilter")
        if element is None:
            return
        previous = read_auto_filter(element)
        root.remove(element)
        if show_rows and previous.filtering:
            self._show_filtered_rows(self._table_filter_block(table))
        self._invalidate()

    def _filter_table(
        self,
        table: Table,
        columns: Sequence[FilterColumn],
        *,
        apply: bool,
        today: dt.date | None,
        strict: bool,
    ) -> FilterOutcome:
        if table.header_row_count == 0:
            raise ValueError(
                f"the table {table.name!r} has no header row, and a table's dropdowns sit on it; Excel turns "
                "the header row back on to filter, which this library does not do."
            )
        block = self._table_filter_block(table)
        entries = list(columns)
        self._check_filter_columns(entries, block)
        entries.sort(key=lambda entry: entry.column)
        when = today or dt.date.today()
        resolved = self._resolve_columns(block, entries, when, strict=strict)

        # Everything that can refuse has: from here on the table changes.
        root = table.document.root
        element = root.child("autoFilter")
        previous = None if element is None else read_auto_filter(element)
        if element is None:
            element = Element.create("autoFilter", {"ref": block.a1})
            insert_table_child(root, element)
        elif element.get("ref") != block.a1:
            element.set("ref", block.a1)
        self._write_filter_columns(element, resolved)
        filters = AutoFilter(ref=block.a1, columns=tuple(resolved))
        outcome = FilterOutcome()
        if apply and filters.filtering:
            outcome = self._apply_filter(filters, when)
        elif apply and previous is not None and previous.filtering:
            outcome = FilterOutcome(shown=self._show_filtered_rows(block))
        self._invalidate()
        return outcome

    @staticmethod
    def _table_filter_block(table: Table) -> RangeRef:
        """What a table's filter covers: its header and data rows."""
        block = table.ref
        bottom = max(block.top, block.bottom - table.totals_row_count)
        return RangeRef(CellRef(block.top, block.left), CellRef(bottom, block.right))

    def _resolve_columns(
        self, block: RangeRef, entries: Sequence[FilterColumn], today: dt.date, *, strict: bool
    ) -> list[FilterColumn]:
        """Each column with what Excel works out and stores filled in: a top
        ten's threshold, an average, a relative period's bounds."""
        epoch = self._workbook.epoch_1904
        return [
            replace(
                entry,
                criterion=resolve(
                    entry.criterion, self._filter_cells(block, entry.column),
                    today=today, epoch_1904=epoch, strict=strict,
                ),
            )
            for entry in entries
        ]

    def _filter_block(self, reference: str | RangeRef) -> RangeRef:
        """The range a filter covers: as given, down to the last row that
        holds a value in it, header row included."""
        block = (RangeRef.parse(reference) if isinstance(reference, str) else reference).normalized
        last = block.top
        for number in sorted((n for n in self._rows if block.top < n <= block.bottom), reverse=True):
            if any(
                self._value_in(cell) for cell in self._rows[number].children_named("c")
                if block.left <= self._column_of(cell) <= block.right
            ):
                last = number
                break
        return RangeRef(CellRef(block.top, block.left), CellRef(last, block.right))

    @staticmethod
    def _value_in(cell: Element) -> bool:
        return cell.child("v") is not None or cell.child("is") is not None or cell.child("f") is not None

    @staticmethod
    def _column_of(cell: Element) -> int:
        try:
            return CellRef.parse(cell.get("r") or "").column
        except ValueError:
            return 0

    @staticmethod
    def _check_filter_columns(entries: Sequence[FilterColumn], block: RangeRef) -> None:
        seen: set[int] = set()
        for entry in entries:
            if not isinstance(entry, FilterColumn):  # pyright: ignore[reportUnnecessaryIsInstance]
                raise TypeError(f"{entry!r} is not a FilterColumn.")
            if entry.column >= block.width:
                raise ValueError(
                    f"column {entry.column} is outside {block.a1}, which is {block.width} wide; "
                    "a filter column counts from the range's left edge, starting at 0."
                )
            if entry.column in seen:
                raise ValueError(f"column {entry.column} is given twice; a column has one criterion.")
            seen.add(entry.column)

    def _write_filter_columns(self, element: Element, entries: Sequence[FilterColumn]) -> None:
        """Replace the filter's columns, keeping the markup of any that did
        not change, and the sort state and extensions after them in place."""
        existing: dict[int, Element] = {}
        for column in list(element.children_named("filterColumn")):
            existing.setdefault(read_filter_column(column).column, column)
            element.remove(column)
        epoch = self._workbook.epoch_1904
        for entry in entries:
            kept = existing.get(entry.column)
            if kept is None or read_filter_column(kept) != entry:
                kept = filter_column_element(entry, epoch_1904=epoch)
            insert_in_schema_order(element, kept, AUTO_FILTER_CHILD_ORDER)

    def _set_filter_mode(self, on: bool) -> None:
        """``filterMode`` on ``sheetPr``, which Excel sets while a filter
        hides rows by criteria."""
        properties = self._root.child("sheetPr")
        if on:
            self._sheet_properties().set("filterMode", "1")
        elif properties is not None and properties.unset("filterMode"):
            self._tidy_sheet_properties(properties)

    def _take_filter_off(self, element: Element, previous: AutoFilter) -> None:
        """Remove a filter, showing its rows if it was filtering."""
        self._root.remove(element)
        if previous.filtering:
            try:
                block = RangeRef.parse(previous.ref).normalized
            except ValueError:
                return
            self._show_filtered_rows(block)

    def _apply_filter(self, filters: AutoFilter, today: dt.date) -> FilterOutcome:
        """Hide the rows a criterion excludes and show the rows all keep."""
        block = RangeRef.parse(filters.ref).normalized
        rows = list(range(block.top + 1, block.bottom + 1))
        epoch = self._workbook.epoch_1904
        columns: list[list[bool | None]] = []
        unevaluated: list[int] = []
        for entry in filters.columns:
            verdicts = keeps(entry.criterion, self._filter_cells(block, entry.column), today=today, epoch_1904=epoch)
            if any(verdict is None for verdict in verdicts):
                unevaluated.append(entry.column)
            columns.append(verdicts)
        decided = decide(columns)
        hide = [row for row, verdict in zip(rows, decided, strict=True) if verdict is False]
        show = [row for row, verdict in zip(rows, decided, strict=True) if verdict is True]
        undecided = [row for row, verdict in zip(rows, decided, strict=True) if verdict is None]
        self._hide_rows(hide)
        opened = False
        for number in show:
            existing = self._rows.get(number)
            if existing is not None and existing.unset("hidden"):
                opened = True
        if opened:
            self._tidy_collapsed()
        return FilterOutcome(tuple(hide), tuple(show), tuple(undecided), tuple(unevaluated))

    def _show_filtered_rows(self, block: RangeRef) -> tuple[int, ...]:
        """Show every row under a filter's header, as Excel's ShowAllData
        does, and drop the collapsed mark of any outline group that leaves
        fully open."""
        shown: list[int] = []
        for number in sorted(n for n in self._rows if block.top < n <= block.bottom):
            element = self._rows[number]
            if element.get("hidden") in ("1", "true"):
                element.unset("hidden")
                shown.append(number)
        if shown:
            self._tidy_collapsed()
        return tuple(shown)

    def _tidy_collapsed(self) -> None:
        """Clear ``collapsed`` on a summary row whose detail rows all show."""
        for number, element in self._rows.items():
            if element.get("collapsed") not in ("1", "true"):
                continue
            detail = self._detail_rows(number)
            if detail and not any(self.row_hidden(row) for row in detail):
                element.unset("collapsed")

    def _hide_rows(self, numbers: Sequence[int]) -> None:
        """Hide rows, creating in one pass the elements of any that have
        none, since a row with no element has nowhere to say it is hidden."""
        missing = [number for number in numbers if number not in self._rows]
        if missing:
            self._ensure_rows(missing)
        for number in numbers:
            self._rows[number].set("hidden", "1")

    def _ensure_rows(self, numbers: Sequence[int]) -> None:
        """Create row elements for several rows at once, in order.

        One at a time costs a scan of every row per row; a filter hiding the
        empty rows of a long range would take minutes.
        """
        created = {number: Element.create("row", {"r": str(number)}) for number in numbers if number not in self._rows}
        if not created:
            return
        merged = sorted({**self._rows, **created}.items())
        self._data.clear()
        for _, element in merged:
            self._data.append(element)
        self._rows = dict(merged)
        self._highest_row = max(self._highest_row, *created)

    def _filter_cells(self, block: RangeRef, offset: int) -> list[FilterCell]:
        """One column of a filter's range, under the header, as the filter
        sees it: the value, the text its format shows, and whether a
        formula's cached result can be trusted."""
        column = block.left + offset
        styles = self._workbook.styles
        shared = self._workbook.shared_strings
        epoch = self._workbook.epoch_1904
        stale = self._workbook.values_changed
        codes: dict[str | None, str] = {}
        cells: list[FilterCell] = []
        for number in range(block.top + 1, block.bottom + 1):
            element = self._find_cell(CellRef(number, column))
            if element is None:
                cells.append(FilterCell(None))
                continue
            value = read_value(element, shared_strings=shared, styles=None, epoch_1904=epoch)
            if isinstance(value, (dt.datetime, dt.date, dt.time)):
                try:
                    value = datetime_to_serial(value, epoch_1904=epoch)
                except ValueError:
                    value = str(value)
            style = element.get("s")
            code = codes.get(style)
            if code is None:
                index = None if style is None else int(style) if style.isdigit() else None
                code = "General" if styles is None else styles.display_format(index)
                codes[style] = code
            number_value = isinstance(value, (int, float)) and not isinstance(value, bool)
            formula = element.child("f") is not None
            cells.append(
                FilterCell(
                    value,  # type: ignore[arg-type]
                    format_value(value, code, epoch_1904=epoch),
                    is_date=number_value and parse_format(code).is_date,
                    unknown=formula and (stale or not self._value_in_cache(element)),
                )
            )
        return cells

    @staticmethod
    def _value_in_cache(cell: Element) -> bool:
        return cell.child("v") is not None or cell.child("is") is not None

    # ------------------------------------------------------------------
    # Comments
    # ------------------------------------------------------------------

    @property
    def comments(self) -> list[Comment]:
        """Every note on the sheet, in the order the file holds them.

        Excel's Review tab calls these notes, and keeps "comments" for the
        threaded kind; the file calls them comments. The text and author
        come from the comments part, and whether each shows from the box
        its VML draws. A threaded comment is in :attr:`threaded_comments`,
        not here, although Excel keeps a placeholder note beside it.
        """
        part = self._comments_part()
        if part is None:
            return []
        boxes = self._note_boxes()
        threaded = self._thread_ids()
        found: list[Comment] = []
        for comment in read_comments(self._workbook.package.xml(part).root):
            if comment.ref in threaded:
                continue
            try:
                box = boxes.get(CellRef.parse(comment.ref))
            except ValueError:
                box = None
            found.append(replace(comment, visible=box is not None and shows(box)))
        return found

    def comment(self, reference: str | CellRef) -> Comment | None:
        """A cell's note, or ``None`` if it has none."""
        wanted = _as_cell(reference).a1
        return next((found for found in self.comments if found.ref == wanted), None)

    def set_comment(
        self, reference: str | CellRef, text: str, *, author: str = "", visible: bool = False
    ) -> Comment:
        """Put a note on a cell, replacing the one it has.

        Three things change together, as Excel changes them: the text in
        the comments part, a box in the VML part, and the sheet's
        ``<legacyDrawing>``, each made if the sheet has none. A new note
        gets the box Excel gives one, 15 pixels right of the cell and 10
        above it; a note already there keeps its box and only has its
        visibility set. The text is written in Excel's own note font.
        """
        cell = _as_cell(reference)
        if cell.a1 in self._thread_ids():
            raise ValueError(
                f"{cell.a1} has a threaded comment, and Excel keeps a cell's note and its thread "
                "apart. Remove the thread with remove_comment() first."
            )
        write_comment(self._comments_root(), cell.a1, text, author)
        self._note_box(cell, visible)
        self._invalidate()
        return Comment(ref=cell.a1, text=text, author=author, visible=visible)

    def _comments_root(self) -> Element:
        """The root of the sheet's comments part, made if it has none."""
        package = self._workbook.package
        part = self._comments_part()
        if part is None:
            part = self._workbook.free_part_name("xl/comments{n}.xml")
            package.write(part, EMPTY_COMMENTS.encode("utf-8"), content_type=CT_COMMENTS)
            package.relationships(self._part_name).add_part(RT_COMMENTS, part)
        return package.xml(part).root

    def _note_box(self, cell: CellRef, visible: bool) -> None:
        """Give a cell's note its box in the VML part, placed as Excel
        places a new one, or set the visibility of the box it has."""
        package = self._workbook.package
        vml_part = self._vml_part()
        vml = package.read(vml_part).decode("utf-8")
        box = note_shapes(vml).get(cell)
        if box is None:
            anchor, left, top = note_anchor(cell, *self._pixel_grid(cell))
            markup = note_vml(
                self._next_control_id(), cell, anchor, left, top,
                visible=visible, z_index=shape_count(vml) + 1,
            )
            vml = with_vml_shape(with_note_shape_type(vml), markup)
        else:
            vml = vml.replace(box, shown_as(box, visible), 1)
        package.write(vml_part, vml.encode("utf-8"))

    def remove_comment(self, reference: str | CellRef) -> bool:
        """Take a cell's note or thread off, box and all; whether it had
        one.

        A thread goes with its replies and the placeholder note Excel keeps
        beside it. A part left empty goes, and so does a VML part left
        drawing nothing, with the ``<legacyDrawing>`` that named it, so a
        sheet whose last comment is removed is as if it never had one.
        """
        cell = _as_cell(reference)
        package = self._workbook.package
        relationships = package.relationships(self._part_name)
        threads = self._threads_part()
        if threads is not None:
            thread_root = package.xml(threads).root
            if delete_thread(thread_root, cell.a1) and next(
                thread_root.children_named("threadedComment"), None
            ) is None:
                for relationship in relationships.by_type(RT_THREADED_COMMENTS):
                    if relationship.target_part == threads:
                        relationships.remove(relationship.id)
                package.remove_part(threads)

        part = self._comments_part()
        if part is None:
            return False
        root = package.xml(part).root
        if not delete_comment(root, cell.a1):
            return False
        if is_empty(root):
            for relationship in relationships.by_type(RT_COMMENTS):
                if relationship.target_part == part:
                    relationships.remove(relationship.id)
            package.remove_part(part)

        for vml_part in self._legacy_vml_parts():
            vml = package.read(vml_part).decode("utf-8")
            box = note_shapes(vml).get(cell)
            if box is None:
                break
            vml = vml.replace(box, "", 1)
            if vml_has_shapes(vml):
                package.write(vml_part, vml.encode("utf-8"))
            else:
                self._remove_vml_part(vml_part)
            break
        self._invalidate()
        return True

    def _comments_part(self) -> str | None:
        for part in related_parts(self, RT_COMMENTS):
            return part
        return None

    @property
    def threaded_comments(self) -> list[ThreadedComment]:
        """Every threaded comment on the sheet, each with its replies, in
        the order the file holds them.

        Threaded comments are what Excel's Review tab calls comments since
        2019: a conversation on a cell, each entry with an author and a
        time, which a thread can be marked resolved. Notes, the older kind,
        are in :attr:`comments`.
        """
        part = self._threads_part()
        if part is None:
            return []
        return read_threads(self._workbook.package.xml(part).root, self._workbook.persons())

    def threaded_comment(self, reference: str | CellRef) -> ThreadedComment | None:
        """A cell's thread, or ``None`` if it has none."""
        wanted = _as_cell(reference).a1
        return next((found for found in self.threaded_comments if found.ref == wanted), None)

    def add_threaded_comment(
        self, reference: str | CellRef, text: str, *, author: str, when: dt.datetime | None = None
    ) -> ThreadedComment:
        """Start a thread on a cell.

        Written as Excel writes one: the comment in the sheet's threads
        part, its author in the workbook's person list, and a placeholder
        note beside it, in a box, for a version of Excel that cannot show
        threads. ``when`` is stored in UTC, now unless given; a moment with
        no time zone is taken as UTC already. The author is a person no
        account stands behind, as Excel writes someone not signed in.

        A cell with a note or a thread already is refused: Excel keeps one
        of either.
        """
        cell = _as_cell(reference)
        if cell.a1 in self._thread_ids() or self.comment(cell) is not None:
            raise ValueError(
                f"{cell.a1} already has a comment; Excel keeps one note or one thread on a cell. "
                "Add a reply with add_threaded_reply(), or remove it first."
            )
        root = self._threads_root()
        person = person_id(self._workbook.persons_root(), author)
        identifier = write_thread(root, cell.a1, text, person, when or _now())
        return self._placeholder(cell, identifier)

    def add_threaded_reply(
        self, reference: str | CellRef, text: str, *, author: str, when: dt.datetime | None = None
    ) -> ThreadedComment:
        """Reply to a cell's thread, at its end, and bring its placeholder
        note up to date as Excel does."""
        cell = _as_cell(reference)
        parent = self._thread_ids().get(cell.a1)
        if parent is None:
            raise ValueError(f"{cell.a1} has no thread to reply to; start one with add_threaded_comment().")
        person = person_id(self._workbook.persons_root(), author)
        write_thread(self._threads_root(), cell.a1, text, person, when or _now(), parent=parent)
        return self._placeholder(cell, parent)

    def resolve_threaded_comment(self, reference: str | CellRef, resolved: bool = True) -> ThreadedComment:
        """Mark a cell's thread resolved, or open it again."""
        cell = _as_cell(reference)
        part = self._threads_part()
        if part is None or not set_resolved(self._workbook.package.xml(part).root, cell.a1, resolved):
            raise ValueError(f"{cell.a1} has no thread to resolve.")
        self._invalidate()
        found = self.threaded_comment(cell)
        assert found is not None
        return found

    def _placeholder(self, cell: CellRef, identifier: str) -> ThreadedComment:
        """Write the note Excel keeps beside a thread, and its box."""
        thread = self.threaded_comment(cell)
        assert thread is not None
        write_comment(
            self._comments_root(), cell.a1, placeholder_text(thread), f"tc={identifier}",
            run=False, uid=identifier,
        )
        self._note_box(cell, visible=False)
        self._invalidate()
        return thread

    def _threads_part(self) -> str | None:
        for part in related_parts(self, RT_THREADED_COMMENTS):
            return part
        return None

    def _threads_root(self) -> Element:
        """The root of the sheet's threads part, made if it has none."""
        package = self._workbook.package
        part = self._threads_part()
        if part is None:
            part = self._workbook.free_part_name("xl/threadedComments/threadedComment{n}.xml")
            package.write(part, EMPTY_THREADED_COMMENTS.encode("utf-8"), content_type=CT_THREADED_COMMENTS)
            package.relationships(self._part_name).add_part(RT_THREADED_COMMENTS, part)
        return package.xml(part).root

    def _thread_ids(self) -> dict[str, str]:
        """Each thread's id, by the cell it is on."""
        part = self._threads_part()
        return {} if part is None else thread_ids(self._workbook.package.xml(part).root)

    def _note_boxes(self) -> dict[CellRef, str]:
        package = self._workbook.package
        for part in self._legacy_vml_parts():
            return note_shapes(package.read(part).decode("utf-8", errors="replace"))
        return {}

    def _pixel_grid(self, cell: CellRef) -> tuple[list[int], list[int]]:
        """Column widths and row heights in pixels, from the first to far
        enough past a cell to hold a note's box, at 96 pixels an inch."""
        grid = SheetGrid.of(self._root)
        columns = [
            round(grid.column_widths.get(index, grid.default_column) / 0.75)
            for index in range(min(cell.column + 40, MAX_COLUMN))
        ]
        rows = [
            round(grid.row_heights.get(index, grid.default_row) / 0.75)
            for index in range(min(cell.row + 40, MAX_ROW))
        ]
        return columns, rows

    def _legacy_vml_parts(self) -> list[str]:
        """The VML parts that draw on the sheet: its notes and its controls.

        A header or footer picture is VML too, related by the same type,
        and ``<legacyDrawingHF>`` names its part. Measured: Excel lists that
        relationship first and numbers the picture ``_x0000_s1025``, the
        same as the sheet's first note, so a note, a control or an edit to
        one would otherwise land in the header, and Excel refuses the file.
        """
        package = self._workbook.package
        try:
            relationships = package.relationships(self._part_name)
        except PackageError:
            return []
        drawings = {
            element.get("r:id") or element.get("id")
            for element in self._root.children_named("legacyDrawing")
        }
        pictures = {
            element.get("r:id") or element.get("id")
            for element in self._root.children_named("legacyDrawingHF")
        } - drawings
        return [
            relationship.target_part
            for relationship in relationships.by_type(RT_VML)
            if relationship.id not in pictures
            and not relationship.is_external
            and package.has_part(relationship.target_part)
        ]

    def _remove_vml_part(self, part: str) -> None:
        """Take a sheet's VML part away, with its relationship and the
        ``<legacyDrawing>`` that names it."""
        package = self._workbook.package
        relationships = package.relationships(self._part_name)
        for relationship in relationships.by_type(RT_VML):
            if relationship.target_part != part:
                continue
            for element in list(self._root.children_named("legacyDrawing")):
                if (element.get("r:id") or element.get("id")) == relationship.id:
                    self._root.remove(element)
            relationships.remove(relationship.id)
        package.remove_part(part)

    # ------------------------------------------------------------------
    # Shapes
    # ------------------------------------------------------------------

    @property
    def pivot_tables(self) -> list[PivotTable]:
        """Every pivot table on the sheet, with where it is and what its
        cache was read from."""
        return read_pivot_tables(self._workbook.package, self._part_name)

    @property
    def charts(self) -> list[Chart]:
        """Every chart on the sheet, in the order the drawing holds them,
        each with what it plots. A chart may plot another sheet's cells."""
        package = self._workbook.package
        found: list[Chart] = []
        for name in related_parts(self, RT_DRAWING):
            found.extend(charts_in_drawing(package, name))
        return found

    @property
    def shapes(self) -> list[Shape]:
        """Every shape on the sheet, in the order the drawing holds them,
        then each ActiveX control the sheet records with no drawing twin.

        A form control is finished off here rather than in the drawing, at
        any depth in a group: the drawing calls it an ordinary shape, hidden
        and with no macro, and only the sheet's own ``<control>`` records
        and the VML say otherwise. An OLE object is the same, with an
        ``<oleObject>`` record.
        """
        grid = SheetGrid.of(self._root)
        package = self._workbook.package
        found: list[Shape] = []
        for name in related_parts(self, RT_DRAWING):
            document = package.xml(name)
            images = {
                relationship.id: relationship.target_part
                for relationship in package.relationships(name).by_type(RT_IMAGE)
                if not relationship.is_external
            }
            found.extend(read_drawing(document.root, grid, images))

        records = self._sheet_records()
        if not records:
            return found
        facts = self._vml_facts()
        finished = [_finished(shape, records, facts) for shape in found]
        seen = {member.shape_id for shape in finished for member in _members(shape)}
        for shape_id, record in records.items():
            if shape_id in seen or record.kind != "activeX":
                continue
            # An ActiveX control with no drawing twin: its record still names
            # it and places it.
            anchor = record.anchor
            left, top, width, height = (0.0, 0.0, 0.0, 0.0) if anchor is None else anchor_box(anchor, grid)
            finished.append(
                Shape(
                    name=record.name or f"ActiveX {shape_id}",
                    kind="activeX",
                    left=left,
                    top=top,
                    width=width,
                    height=height,
                    cells="" if anchor is None else anchor_cells(anchor, grid),
                    alt_text=record.alt_text,
                    hidden=facts.get(shape_id, VmlControl()).hidden,
                    macro=record.macro,
                    shape_id=shape_id,
                )
            )
        return finished

    def shape(self, name: str) -> Shape:
        """One shape by name."""
        shapes = self.shapes
        for found in shapes:
            if found.name == name:
                return found
        available = ", ".join(s.name for s in shapes) or "none"
        raise KeyError(f"no shape named {name!r} on {self._name!r}. It has: {available}")

    def cell_origin(self, reference: str | CellRef) -> tuple[float, float]:
        """A cell's top-left corner in points: the ``left`` and ``top`` that
        put a shape on that cell, worked out from this sheet's own column
        widths and row heights the way an anchor is."""
        cell = CellRef.parse(reference) if isinstance(reference, str) else reference
        grid = SheetGrid.of(self._root)
        return grid.x(cell.column - 1), grid.y(cell.row - 1)

    def update_shape(
        self,
        name: str,
        *,
        new_name: str | None = None,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
        text: str | None = None,
        alt_text: str | None = None,
        hidden: bool | None = None,
        linked_cell: str | None = None,
        list_range: str | None = None,
    ) -> Shape:
        """Change a shape in place, keeping its style and relationships.

        Whatever is left as ``None`` stays as it is. The box is in points,
        and a side not given keeps where the anchor has it, which is where
        Excel draws the shape. New ``text`` takes the font, size, colour and
        alignment of the text it replaces, and a link to a cell's value is
        dropped, since the shape would otherwise go on showing the cell.
        ``alt_text`` of ``""`` clears it.

        A Forms control's four parts change together: its drawing twin, its
        record on the sheet, its own part and its VML. ``linked_cell`` and
        ``list_range`` belong to a control, and ``""`` clears either. A
        control takes ``text`` only where it shows a caption: a button, a
        tick box, an option button, a group box or a label.

        Raises ``ValueError``, having changed nothing, for an ActiveX control,
        an OLE object or a group, for text on a shape that shows none, for a
        side that is negative or not a number, for a name another shape has,
        and for a move of a shape whose anchor does not name both corners.
        """
        current = self.shape(name)
        control = current.kind == "formControl"
        if current.kind == "activeX":
            raise ValueError(f"{name!r} is an ActiveX control, which update_shape does not change.")
        if current.kind == "oleObject":
            raise ValueError(f"{name!r} is an OLE object, which update_shape does not change.")
        if current.kind == "group":
            raise ValueError(f"{name!r} is a group, which update_shape does not change.")
        if text is not None and not _shows_text(current):
            what = f"{current.control.kind} control" if current.control is not None else current.kind
            raise ValueError(f"{name!r} is a {what} and shows no text to change.")
        if not control and (linked_cell is not None or list_range is not None):
            raise ValueError("linked_cell and list_range belong to a form control.")
        for side, value in (("left", left), ("top", top), ("width", width), ("height", height)):
            if value is not None and not (math.isfinite(value) and value >= 0):
                raise ValueError(f"{side} is in points and cannot be {value!r}.")
        if new_name is not None and new_name != name:
            self._check_new_shape_name(new_name)
        located = self._located(name)
        if located is None:
            raise KeyError(f"{name!r} is on {self._name!r} but in none of its drawing parts.")
        _, node, copies = located
        anchors = anchors_in(node)
        moving = any(value is not None for value in (left, top, width, height))
        if moving and not (anchors and all(is_two_cell(anchor) for anchor in anchors)):
            raise ValueError(f"{name!r} has no two-cell anchor to move in place.")
        grid = SheetGrid.of(self._root)
        was = anchor_box(anchors[0], grid) if anchors else (current.left, current.top, current.width, current.height)
        box = (
            was[0] if left is None else left,
            was[1] if top is None else top,
            was[2] if width is None else width,
            was[3] if height is None else height,
        )
        package = self._workbook.package
        records: list[Element] = []
        vml_updates: list[tuple[str, str]] = []
        if control:
            if current.control is None or not current.control.part_name:
                raise ValueError(f"{name!r} has no control part to update.")
            records = [
                record for record in self._outside_cells("control")
                if record.get("shapeId") == str(current.shape_id)
            ]
            if not records:
                raise ValueError(f"{name!r} has no record on the sheet to update.")
            after = replace(
                current,
                name=name if new_name is None else new_name,
                left=box[0],
                top=box[1],
                width=box[2],
                height=box[3],
                text=current.text if text is None else text,
                hidden=current.hidden if hidden is None else hidden,
            )
            for part in self._legacy_vml_parts():
                original = package.read(part).decode("utf-8")
                if not has_vml_shape(original, current.shape_id):
                    continue
                vml_updates.append((part, update_vml_control(
                    original, after, grid,
                    rename=new_name is not None, move=moving, caption=text is not None,
                    visibility=hidden is not None, linked_cell=linked_cell, list_range=list_range,
                )))
            if not vml_updates:
                raise ValueError(f"{name!r} has no VML shape to update.")

        # Every check has passed. From here the parts change together.
        for body in copies:
            naming = next(body.descendants("cNvPr"), None)
            if naming is not None:
                if new_name is not None:
                    naming.set("name", new_name)
                if alt_text is not None:
                    _set_or_drop(naming, "descr", alt_text)
                if hidden is not None and not control:
                    _set_or_drop(naming, "hidden", "1" if hidden else "")
            if text is not None and (not control or body.child("txBody") is not None):
                replace_text(body, text)
                if body.get("textlink"):
                    body.set("textlink", "")
            if moving:
                _set_transform(body, box)
        if moving:
            for anchor in anchors:
                move_anchor(anchor, grid, box)
        if control:
            assert current.control is not None
            for record in records:
                if new_name is not None:
                    record.set("name", new_name)
                properties = record.child("controlPr")
                if properties is None:
                    continue
                if alt_text is not None:
                    _set_or_drop(properties, "altText", alt_text)
                placed = properties.child("anchor")
                if moving and placed is not None and is_two_cell(placed):
                    move_anchor(placed, grid, box, wrapper="")
            wiring = package.xml(current.control.part_name).root
            for attribute, wanted in (("fmlaLink", linked_cell), ("fmlaRange", list_range)):
                if wanted is not None:
                    _set_or_drop(wiring, attribute, wanted)
            for part, updated in vml_updates:
                package.write(part, updated.encode("utf-8"))
        self._invalidate()
        return self.shape(name if new_name is None else new_name)

    def _located(self, name: str) -> tuple[str, Element, list[Element]] | None:
        """Where a listed shape is: its drawing part, the top-level node that
        holds it, and each copy of it in that node."""
        package = self._workbook.package
        for part in related_parts(self, RT_DRAWING):
            for node in package.xml(part).root.elements():
                copies = shape_copies(node, name)
                if copies:
                    return part, node, copies
        return None

    def _vml_facts(self) -> dict[int, VmlControl]:
        """What the sheet's VML says of each control, read in one pass."""
        package = self._workbook.package
        facts: dict[int, VmlControl] = {}
        for part in self._legacy_vml_parts():
            facts.update(vml_controls(package.read(part).decode("utf-8", errors="replace")))
        return facts

    def add_shape(
        self,
        name: str,
        *,
        left: float,
        top: float,
        width: float,
        height: float,
        kind: ShapeKind = "shape",
        geometry: str = "",
        text: str = "",
        macro: str = "",
    ) -> Shape:
        """Put a drawing shape on the sheet.

        ``kind`` is ``"shape"`` for an AutoShape, ``"textBox"`` or
        ``"line"``. ``geometry`` is a preset name such as ``roundRect`` or
        ``ellipse``; left empty each kind gets its own default, which is
        ``rect`` for a shape and a text box and ``line`` for a line.

        The box is in points, the unit the object model uses,
        and it is stored as a two-cell anchor: which cells the shape spans
        is worked out from this sheet's own column widths and row heights,
        because Excel clamps an offset to the cell holding it and a corner
        left to be clamped lands somewhere other than where it was put.

        That placement is close rather than exact, and the error grows
        with how far across the sheet the shape sits. Excel reports a
        shape's position from its anchor, and turning points back into a
        column needs the standard font's maximum digit width, which the
        file does not carry: see :mod:`pyofficeeditor.excel._shapes` for
        what is used instead. Measured against Excel, a shape put at 300
        points came back at 300 on one sheet and 298.5 on another whose
        columns had been resized. Rows are exact, because a row height is
        already in points, as long as Excel's default row height is the
        one the file records: a row with none of its own follows the
        display scaling, 15 points at 100% where a file written at 150%
        says 14.5, and a shape below it moves by the difference.

        A sheet with no drawing part gets one, with its content type and
        its relationship. Every other part of the package is left alone.

        For a Forms-toolbar control use :meth:`add_form_control`: it is
        four parts that have to agree rather than one, and it takes the
        wiring this does not.
        """
        if kind == "formControl":
            raise ValueError(
                "a form control is made with add_form_control: it needs a control "
                "part, a VML shape and a record on the sheet as well as a drawing."
            )
        self._check_new_shape_name(name)
        part, document = self._drawing_part()
        shape = Shape(
            name=name,
            kind=kind,
            geometry=geometry,
            left=left,
            top=top,
            width=width,
            height=height,
            text=text,
            macro=macro,
            shape_id=self._next_shape_id(),
        )
        markup = new_anchor(shape, SheetGrid.of(self._root))
        document.root.append(XmlDocument.parse(markup.encode("utf-8")).root)
        self._workbook.package.write(part, document.to_bytes(), content_type=CT_DRAWING)
        self._invalidate()
        return self.shape(name)

    def add_form_control(
        self,
        name: str,
        *,
        left: float,
        top: float,
        width: float,
        height: float,
        kind: str = "Button",
        text: str = "",
        macro: str = "",
        linked_cell: str = "",
        list_range: str = "",
        value: int = 0,
        minimum: int | None = None,
        maximum: int | None = None,
    ) -> Shape:
        """Put a Forms-toolbar control on the sheet.

        ``kind`` is Excel's own ``objectType``: Button, CheckBox, Drop,
        List, Radio, Spin, Scroll, GBox or Label.

        ``value`` is the control's stored state, and it loses to
        ``linked_cell``: a control with one takes its state from that cell
        when the workbook opens, so a tick box stored ticked and linked to
        an empty cell opens unticked. Setting both sets the cell.

        ``minimum`` and ``maximum`` bound a spinner or a scroll bar. Left
        alone they get Excel's own defaults, 30000 for a spinner and 100
        for a scroll bar, rather than 0: a control that cannot exceed zero
        sits at zero whatever value it was given, in a file that is
        perfectly valid and silently useless.

        Four parts have to agree for Excel to draw one, and all four are
        written here:

        - the drawing, which holds the anchor, wrapped in an
          ``mc:AlternateContent``
        - the sheet's own ``<control>``, which is where the macro that a
          click runs actually lives
        - a control part of its own, holding the linked cell, the list
          range and the current value
        - the VML, which is what Excel draws the control from

        Leave any of them out and the control is invisible, inert, or the
        file does not open.
        """
        self._check_new_shape_name(name)
        # Before anything is written: this makes four parts, and the last
        # of them is where an unknown kind would otherwise be noticed,
        # leaving the other three behind for a control that never existed.
        check_control_kind(kind)
        package = self._workbook.package
        grid = SheetGrid.of(self._root)
        shape = Shape(
            name=name,
            kind="formControl",
            left=left,
            top=top,
            width=width,
            height=height,
            text=text,
            macro=qualified_macro(macro),
            shape_id=self._next_control_id(),
            control=FormControl(
                kind=kind,
                linked_cell=linked_cell,
                list_range=list_range,
                value=value,
                minimum=minimum,
                maximum=maximum,
            ),
        )

        drawing_part, drawing = self._drawing_part()
        drawing.root.append(
            XmlDocument.parse(control_drawing(shape, grid).encode("utf-8")).root
        )
        package.write(drawing_part, drawing.to_bytes(), content_type=CT_DRAWING)

        control_part = self._free_control_part_name()
        package.write(
            control_part,
            control_properties(shape.control or FormControl()).encode("utf-8"),
            content_type=CT_CONTROL_PROPERTIES,
        )
        relationship = package.relationships(self._part_name).add_part(
            RT_CONTROL_PROPERTIES, control_part
        )

        vml_part = self._vml_part()
        package.write(
            vml_part,
            with_vml_shape(
                package.read(vml_part).decode("utf-8"), control_vml(shape, grid)
            ).encode("utf-8"),
        )

        container = self._controls_container()
        container.append(
            XmlDocument.parse(
                control_entry(shape, relationship.id, grid).encode("utf-8")
            ).root
        )
        self._invalidate()
        return self.shape(name)

    def remove_shape(self, name: str) -> None:
        """Take a shape off the sheet, and each part only it used.

        A drawing shape is one anchor and whatever it points at: a picture's
        image, a chart with the chart's own style and colour parts, a
        hyperlink. Each goes once nothing else uses it, so an image another
        picture shows stays. A group goes with its members, pictures, charts
        and controls included.

        A form control is four things, and leaving any of them behind is
        worse than leaving all of them: an orphaned relationship pointing at
        a part that is gone is the failure that stays invisible until Excel
        next opens the file, and then it is the whole workbook that gets
        repaired rather than the control that goes missing.

        Raises ``ValueError`` for an ActiveX control, an OLE object, or a
        group holding either: its parts include a binary this does not take
        apart, and measured, Excel draws an OLE object again from its record
        and its VML when only its drawing shape is gone.
        """
        shape = self.shape(name)
        members = list(_members(shape))
        if any(member.kind == "activeX" for member in members):
            raise ValueError(
                f"ActiveX control removal takes its binary part apart, which this does not do; "
                f"remove {name!r} in Excel."
            )
        if any(member.kind == "oleObject" for member in members):
            raise ValueError(
                f"OLE object removal takes its embedded part apart, which this does not do; "
                f"remove {name!r} in Excel."
            )
        located = self._located(name)
        if located is not None:
            part, node, _ = located
            self._drop_drawing_node(part, node)
        self._remove_control_records([member for member in members if member.kind == "formControl"])
        self._invalidate()

    def _swallowed_by(self, deletion: Deletion) -> list[tuple[str, Element, list[Shape]]]:
        """What a deletion takes whole: each drawing node whose object moves
        and sizes with its cells and lies inside the deleted rows or columns,
        with the shapes it holds.

        Raises ``ValueError``, having changed nothing, when one of them is
        or holds an ActiveX control or an embedded object: Excel deletes
        those too, and their binary parts are not something this takes
        apart.
        """
        package = self._workbook.package
        nodes = [
            (part, node)
            for part in related_parts(self, RT_DRAWING)
            for node in swallowed_nodes(package.xml(part).root, deletion)
        ]
        if not nodes:
            return []
        listed = {shape.shape_id: shape for shape in self.shapes}
        found: list[tuple[str, Element, list[Shape]]] = []
        for part, node in nodes:
            naming = next(node.descendants("cNvPr"), None)
            number = "" if naming is None else (naming.get("id") or "")
            shape = listed.get(int(number)) if number.isdigit() else None
            members = [] if shape is None else list(_members(shape))
            refused = next((member for member in members if member.kind in ("activeX", "oleObject")), None)
            if refused is not None:
                what = "an ActiveX control" if refused.kind == "activeX" else "an OLE object"
                raise ValueError(
                    f"this deletion takes all of {refused.name!r}, {what}, which Excel deletes with its "
                    "cells and this does not take apart; move or remove it in Excel first."
                )
            found.append((part, node, members))
        return found

    def _remove_swallowed(self, swallowed: list[tuple[str, Element, list[Shape]]]) -> None:
        """Take out what :meth:`_swallowed_by` found, once the deletion that
        swallowed it has gone through: each node with the parts only it used,
        and each Forms control's record, part and VML shape."""
        if not swallowed:
            return
        for part, node, _ in swallowed:
            self._drop_drawing_node(part, node)
        self._remove_control_records(
            [member for _, _, members in swallowed for member in members if member.kind == "formControl"]
        )
        self._invalidate()

    def _drop_drawing_node(self, part: str, node: Element) -> None:
        """Take one top-level node out of a drawing, with each relationship
        only it used and each part only those relationships reached."""
        package = self._workbook.package
        relationships = package.relationships(part)
        ids = {relationship.id for relationship in relationships}
        used = _relationship_ids(node, ids)
        root = package.xml(part).root
        root.remove(node)
        if not used:
            return
        still = _relationship_ids(root, ids)
        orphaned: list[str] = []
        for relationship in list(relationships):
            if relationship.id not in used or relationship.id in still:
                continue
            relationships.remove(relationship.id)
            if not relationship.is_external:
                orphaned.append(relationship.target_part)
        for target in orphaned:
            if not self._workbook.is_referenced(target):
                self._workbook.remove_with_dependents(target)

    def add_picture(
        self,
        name: str,
        image: bytes | str | Path,
        *,
        left: float,
        top: float,
        width: float | None = None,
        height: float | None = None,
        description: str = "",
    ) -> Shape:
        """Put a picture on the sheet, as Excel's ``Shapes.AddPicture`` does.

        ``image`` is a PNG, a JPEG or a GIF, as bytes or as a path. With no
        size the picture takes the one Excel gives it: its pixels at 96 to
        the inch, or a PNG's own count to the inch where it gives one. With
        one of ``width`` and ``height`` the other keeps the image's
        proportions; with both it is stretched to them, and Excel then no
        longer locks its aspect. ``description`` is its alternative text.

        The image is stored once in the workbook however many pictures show
        it, in a media part named as Excel names one.
        """
        self._check_new_shape_name(name)
        data = Path(image).read_bytes() if isinstance(image, (str, Path)) else bytes(image)
        info = image_info(data)
        natural_width, natural_height = info.size
        keeps_aspect = width is None or height is None
        if width is None and height is None:
            width, height = natural_width, natural_height
        elif width is None:
            assert height is not None
            width = height * natural_width / natural_height if natural_height else 0.0
        elif height is None:
            height = width * natural_height / natural_width if natural_width else 0.0
        if width <= 0 or height <= 0:
            raise ValueError(f"a picture needs a size above nothing; this one would be {width:g} by {height:g}.")

        package = self._workbook.package
        media = self._workbook.media_part(data, info)
        drawing_part, drawing = self._drawing_part()
        relationships = package.relationships(drawing_part)
        relationship = next(
            (one for one in relationships.by_type(RT_IMAGE) if not one.is_external and one.target_part == media),
            None,
        ) or relationships.add_part(RT_IMAGE, media)
        markup = picture_anchor(
            shape_id=self._next_shape_id(),
            name=name,
            description=description,
            relationship=relationship.id,
            extension=info.extension,
            left=left,
            top=top,
            width=width,
            height=height,
            grid=SheetGrid.of(self._root),
            keeps_aspect=keeps_aspect,
        )
        drawing.root.append(XmlDocument.parse(markup.encode("utf-8")).root)
        package.write(drawing_part, drawing.to_bytes(), content_type=CT_DRAWING)
        self._invalidate()
        return self.shape(name)

    def add_chart(
        self,
        kind: ChartKind,
        data: str | RangeRef,
        *,
        left: float,
        top: float,
        width: float = 360.0,
        height: float = 216.0,
        title: str | None = None,
        name: str | None = None,
        series_in: Literal["columns", "rows"] | None = None,
    ) -> Chart:
        """Put a chart of a block of cells on the sheet, as Excel's Insert
        Chart puts one.

        ``data`` is a block whose first row names the series and whose
        first column holds the categories, on this sheet or, written
        ``Data!A1:C6``, on another. Its series run down its columns when it
        is taller than it is wide, and along its rows otherwise, a square
        block included, as Excel lays a block out, measured; ``series_in``
        says which instead. A scatter chart takes the first column as its x
        values.

        The chart looks as Excel's own of the kind does, measured markup and
        all, with Excel's automatic title unless ``title`` gives one. Its
        name is ``Chart 1``, ``Chart 2`` and so on unless ``name`` gives one.
        """
        if kind not in CHART_KINDS:
            raise ValueError(f"{kind!r} is not a kind of chart this can add; it adds {', '.join(CHART_KINDS)}.")
        if width <= 0 or height <= 0:
            raise ValueError(f"a chart needs a size above nothing; this one would be {width:g} by {height:g}.")
        source, block = self._chart_source(data)
        if block.height < 2 or block.width < 2:
            raise ValueError(
                f"{block.a1} needs a row of series names and a column of categories beside at least one "
                f"value; it is {block.height} by {block.width}."
            )
        down = (block.height > block.width) if series_in is None else series_in == "columns"
        series = source.chart_series(block, down=down)
        chart_name = name if name is not None else self._next_chart_name()
        self._check_new_shape_name(chart_name)

        package = self._workbook.package
        part = self._workbook.free_part_name("xl/charts/chart{n}.xml")
        package.write(part, chart_part(kind, series, title=title).encode("utf-8"), content_type=CT_CHART)
        drawing_part, drawing = self._drawing_part()
        relationship = package.relationships(drawing_part).add_part(RT_CHART, part)
        markup = chart_frame(
            shape_id=self._next_shape_id(),
            name=chart_name,
            relationship=relationship.id,
            left=left,
            top=top,
            width=width,
            height=height,
            grid=SheetGrid.of(self._root),
        )
        drawing.root.append(XmlDocument.parse(markup.encode("utf-8")).root)
        package.write(drawing_part, drawing.to_bytes(), content_type=CT_DRAWING)
        self._invalidate()
        return next(chart for chart in self.charts if chart.part_name == part)

    def chart_series(self, block: RangeRef, *, down: bool) -> list[SeriesData]:
        """The series a block of this sheet's cells makes, as Excel reads a
        block given to a chart, each with the values its references read
        now: the names from the first row and the categories from the first
        column when the series run down, the other way about when they run
        along."""
        sheet = quote_sheet_name(self._name)

        def span(cells: list[CellRef]) -> str:
            first, last = cells[0], cells[-1]
            corner = f"${first.letter}${first.row}"
            return f"{sheet}!{corner}" if first == last else f"{sheet}!{corner}:${last.letter}${last.row}"

        top, left = block.top, block.left
        if down:
            headers = [CellRef(top, column) for column in range(left + 1, block.right + 1)]
            categories = [CellRef(row, left) for row in range(top + 1, block.bottom + 1)]
            columns = [[CellRef(row, column) for row in range(top + 1, block.bottom + 1)] for column in range(left + 1, block.right + 1)]
        else:
            headers = [CellRef(row, left) for row in range(top + 1, block.bottom + 1)]
            categories = [CellRef(top, column) for column in range(left + 1, block.right + 1)]
            columns = [[CellRef(row, column) for column in range(left + 1, block.right + 1)] for row in range(top + 1, block.bottom + 1)]

        category_values = [self.get_value(cell) for cell in categories]
        numeric = all(
            isinstance(value, (int, float)) and not isinstance(value, bool)
            for value in category_values
            if value is not None
        ) and any(value is not None for value in category_values)
        category_cache: list[str | float | None] = (
            [value if isinstance(value, (int, float)) and not isinstance(value, bool) else None for value in category_values]
            if numeric
            else [self.get_text(cell) or None for cell in categories]
        )
        found: list[SeriesData] = []
        for header, cells in zip(headers, columns, strict=True):
            values = [self.get_value(cell) for cell in cells]
            found.append(
                SeriesData(
                    values=span(cells),
                    value_cache=[
                        float(value) if isinstance(value, (int, float)) and not isinstance(value, bool) else None
                        for value in values
                    ],
                    value_format=self.number_format(cells[0]) or "General",
                    name=span([header]),
                    name_cache=self.get_text(header),
                    categories=span(categories),
                    category_cache=category_cache,
                    category_format=self.number_format(categories[0]) or "General",
                    numeric_categories=numeric,
                )
            )
        return found

    def _chart_source(self, data: str | RangeRef) -> tuple[Worksheet, RangeRef]:
        """The sheet and block a chart's data names."""
        if isinstance(data, RangeRef):
            return self, data.normalized
        sheet_name, _, address = data.rpartition("!")
        if sheet_name.startswith("'") and sheet_name.endswith("'"):
            sheet_name = sheet_name[1:-1].replace("''", "'")
        source = self._workbook.sheet(sheet_name) if sheet_name else self
        return source, RangeRef.parse(address.replace("$", "")).normalized

    def _next_chart_name(self) -> str:
        """``Chart N``, with the smallest N no shape on the sheet has."""
        taken = {shape.name for shape in self.shapes}
        number = 1
        while f"Chart {number}" in taken:
            number += 1
        return f"Chart {number}"

    def picture_data(self, name: str) -> bytes:
        """The image a picture shows, as the bytes its media part holds."""
        shape = self.shape(name)
        if shape.kind != "picture" or not shape.image:
            raise ValueError(f"{name!r} is not a picture with an image in this workbook.")
        return self._workbook.package.read(shape.image)

    def _unwrap_control(self, control: Element) -> None:
        """Take one ``<control>`` off the sheet, wrapper and all.

        Excel wraps each one in an ``mc:AlternateContent`` of its own
        inside ``<controls>``, so removing the element found leaves an
        empty wrapper behind. This walks up to whatever child of
        ``<controls>`` holds it and removes that, then drops ``<controls>``
        and the ``mc:AlternateContent`` around it once the last control has
        gone. Measured: Excel refuses a sheet that keeps an empty one.
        """
        node: Element = control
        while True:
            parent = node.parent
            if parent is None:
                return
            if parent.name.rpartition(":")[2] != "controls":
                node = parent
                continue

            parent.remove(node)
            if any(True for _ in parent.elements()):
                return
            # The last one: take the empty container and its wrapper too,
            # rather than leaving a <controls/> Excel never writes.
            container = parent.parent
            if container is None:
                return
            container.remove(parent)
            outer = container.parent
            if outer is None or any(True for _ in container.elements()):
                return
            outer.remove(container)
            # A wrapper with no choice left offers nothing, and a fallback
            # in it offered the same controls to older versions.
            holder = outer.parent
            if (
                holder is not None
                and local_name(outer.name) == "AlternateContent"
                and not any(local_name(child.name) == "Choice" for child in outer.elements())
            ):
                holder.remove(outer)
            return

    def _remove_control_records(self, shapes: list[Shape]) -> None:
        """The three parts of each form control besides its drawing: its
        record on the sheet, its own part and its VML shape. The sheet is
        walked once and the VML rewritten once, however many controls go."""
        if not shapes:
            return
        package = self._workbook.package
        ids = {str(shape.shape_id) for shape in shapes}
        for control in list(self._outside_cells("control")):
            if control.get("shapeId") in ids:
                self._unwrap_control(control)

        relationships = package.relationships(self._part_name)
        for shape in shapes:
            control = shape.control
            if control is None:
                continue
            if control.relationship:
                try:
                    relationships.remove(control.relationship)
                except PackageError:
                    # Already gone, which is the state this is trying to
                    # reach. Raising here would abandon the removal half
                    # done, leaving the VML shape and the control part behind.
                    pass
            if control.part_name and package.has_part(control.part_name):
                package.remove_part(control.part_name)

        for part in self._legacy_vml_parts():
            text = package.read(part).decode("utf-8")
            stripped = text
            for shape in shapes:
                stripped = without_vml_shape(stripped, shape.shape_id)
            if stripped != text:
                package.write(part, stripped.encode("utf-8"))

    def _check_new_shape_name(self, name: str) -> None:
        if not name.strip():
            raise ValueError("a shape needs a name; Excel names every shape it makes.")
        taken = {shape.name for shape in self.shapes}
        package = self._workbook.package
        for part in related_parts(self, RT_DRAWING):
            taken |= shape_names(package.xml(part).root)
        if name in taken:
            raise ValueError(
                f"{self._name!r} already has a shape named {name!r}, perhaps inside a group. "
                f"Excel allows two shapes to share a name, but then neither can be reached by it."
            )

    def _next_shape_id(self) -> int:
        """One past the highest id in use, and never into control territory."""
        used = [shape.shape_id for shape in self.shapes]
        return max([one for one in used if one < FIRST_CONTROL_ID] + [1]) + 1

    def _next_control_id(self) -> int:
        """The next id for a form control or a note, which share one
        sequence apart from drawing shapes.

        Measured: Excel numbers them from the block of 1024 ids the sheet's
        VML part claims, so 1025 on the first sheet to have one and 2049 on
        the second, and a note and a button on one sheet take consecutive
        ids. Counting the drawing's controls alone gave a note's id to the
        next button.
        """
        used = {shape.shape_id for shape in self.shapes if shape.shape_id >= FIRST_CONTROL_ID}
        block = 1
        package = self._workbook.package
        for part in self._legacy_vml_parts():
            vml = package.read(part).decode("utf-8", errors="replace")
            block = min(vml_blocks(vml), default=1)
            used |= vml_shape_ids(vml)
            break
        base = block * 1024
        return max((number for number in used if base < number < base + 1024), default=base) + 1

    def _drawing_part(self) -> tuple[str, XmlDocument]:
        """The sheet's drawing part, made if it has none."""
        package = self._workbook.package
        for part in related_parts(self, RT_DRAWING):
            return part, package.xml(part)

        part = self._workbook.free_part_name("xl/drawings/drawing{n}.xml")
        package.write(part, EMPTY_DRAWING.encode("utf-8"), content_type=CT_DRAWING)
        relationship = package.relationships(self._part_name).add_part(RT_DRAWING, part)
        element = Element.create("drawing", {"r:id": relationship.id})
        insert_in_schema_order(self._root, element, WORKSHEET_CHILD_ORDER)
        return part, package.xml(part)

    def _vml_part(self) -> str:
        """The sheet's legacy drawing, made if it has none.

        The content type is a Default by extension rather than an Override,
        which is how Excel writes it, and :meth:`OpcPackage.set_default`
        keeps it ahead of the Overrides because Excel refuses a package
        where a Default comes after one.
        """
        package = self._workbook.package
        for part in self._legacy_vml_parts():
            return part

        part = self._workbook.free_part_name("xl/drawings/vmlDrawing{n}.vml")
        package.content_types.set_default("vml", CT_VML)
        package.write(part, with_vml_block(EMPTY_VML, self._workbook.free_vml_block()).encode("utf-8"))
        relationship = package.relationships(self._part_name).add_part(RT_VML, part)
        element = Element.create("legacyDrawing", {"r:id": relationship.id})
        insert_in_schema_order(self._root, element, WORKSHEET_CHILD_ORDER)
        return part

    def _declare_control_namespaces(self) -> None:
        """Declare the prefixes a control's markup uses on the sheet root.

        A worksheet that has never held one declares neither, and both are
        needed the moment it does:

        - ``xdr``, because the anchor inside ``<controlPr>`` names its
          corners with ``<xdr:col>`` and friends
        - ``x14``, because the wrapper is an ``mc:Choice Requires="x14"``
          and ``Requires`` names a prefix that has to be bound

        Either one missing leaves the part malformed rather than merely
        unusual, and Excel refuses the workbook instead of repairing it.
        Excel writes both on any sheet that carries a control.
        """
        for prefix, uri in (
            ("xmlns:xdr", NS_SPREADSHEET_DRAWING),
            ("xmlns:x14", NS_X14),
            ("xmlns:mc", NS_MARKUP_COMPATIBILITY),
        ):
            if not self._root.has(prefix):
                self._root.set(prefix, uri)

    def _controls_container(self) -> Element:
        """The sheet's ``<controls>``, made if it has none.

        Excel wraps it in an ``mc:AlternateContent`` requiring ``x14``, and
        wraps each ``<control>`` inside it in another, so this reaches for
        the inner container by name at any depth rather than by position.
        """
        self._declare_control_namespaces()
        for found in self._outside_cells("controls"):
            return found

        container = Element.create("controls")
        wrapper = XmlDocument.parse(
            b'<mc:AlternateContent'
            b' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
            b'<mc:Choice Requires="x14"/></mc:AlternateContent>'
        ).root
        choice = wrapper.require("Choice")
        choice.append(container)
        insert_in_schema_order(self._root, wrapper, WORKSHEET_CHILD_ORDER)
        return container

    def _free_control_part_name(self) -> str:
        return self._workbook.free_part_name("xl/ctrlProps/ctrlProp{n}.xml")

    def set_shape_macro(self, name: str, macro: str) -> None:
        """Point a shape at a procedure, or clear it with ``""``.

        The two kinds keep it in different places, and a form control keeps
        it in two at once:

        - a drawing shape carries ``macro="Clicked"`` on its own element
        - a form control carries nothing there. The sheet's ``<controlPr
          macro="[0]!Clicked">`` is what a click runs, and the VML holds a
          second copy in ``<x:FmlaMacro>``

        Measured: Excel reads the sheet's copy. A workbook whose VML
        disagrees opens cleanly and runs what the sheet says, so the VML is
        not load-bearing here, unlike a comment's owning cell, where a
        disagreement makes Excel refuse the file. Both are written anyway,
        because Excel writes both and a stale name left behind is a trap
        for whatever reads it next.

        The procedure is not checked for existence. Excel does not check
        either: a button pointing at a Sub nobody wrote is a normal state
        for a workbook being assembled.

        An ActiveX control raises ``ValueError``: a click on one runs its
        own event procedure in the sheet's code, and it has no macro to
        point anywhere.
        """
        shape = self.shape(name)
        if shape.kind == "activeX":
            raise ValueError(
                f"{name!r} is an ActiveX control, whose click runs its own event procedure; "
                "it has no macro to set."
            )
        if shape.control is not None:
            self._set_control_macro(shape, macro)
        else:
            self._set_drawing_macro(shape, macro)
        self._invalidate()

    def _set_drawing_macro(self, shape: Shape, macro: str) -> None:
        """Set ``macro`` on the shape's own element in the drawing part."""
        for part in related_parts(self, RT_DRAWING):
            document = self._workbook.package.xml(part)
            element = find_shape_element(document.root, shape.name)
            if element is None:
                continue
            # Excel writes macro="" rather than dropping the attribute, and
            # a shape that never had one has no attribute at all. Both are
            # left as Excel leaves them.
            if macro or element.has("macro"):
                element.set("macro", macro)
            return
        raise KeyError(
            f"{shape.name!r} is on {self._name!r} but not in any of its drawing parts, "
            "so there is nothing to attach a macro to."
        )

    def _set_control_macro(self, shape: Shape, macro: str) -> None:
        """Set it on the sheet's own record, and on the VML beside it."""
        for control in self._outside_cells("control"):
            if control.get("shapeId") != str(shape.shape_id):
                continue
            properties = next(control.descendants("controlPr"), None)
            if properties is None:
                raise KeyError(
                    f"the <control> for {shape.name!r} on {self._name!r} has no "
                    "<controlPr>, so there is nowhere to record a macro."
                )
            wanted = qualified_macro(macro, properties.get("macro") or "")
            if wanted:
                properties.set("macro", wanted)
            else:
                properties.unset("macro")
            self._set_vml_macro(shape.shape_id, wanted)
            return
        raise KeyError(
            f"{shape.name!r} reads as a form control on {self._name!r} but the sheet "
            f"has no <control> with shapeId {shape.shape_id}."
        )

    def _set_vml_macro(self, shape_id: int, macro: str) -> None:
        """The second copy, in the legacy drawing.

        VML is not XML this library parses -- an HTML-ish dialect with
        unquoted attributes and unclosed tags -- so it is rewritten as text,
        the same way row and column shifting reaches its anchors.
        """
        package = self._workbook.package
        for part in self._legacy_vml_parts():
            raw = package.read(part)
            try:
                text = raw.decode("utf-8")
            except UnicodeDecodeError:  # pragma: no cover - Excel writes UTF-8
                continue
            rewritten = set_vml_macro(text, shape_id, macro)
            if rewritten != text:
                package.write(part, rewritten.encode("utf-8"))
                return

    def _outside_cells(self, name: str) -> Iterator[Element]:
        """Each element of one name in the sheet, at any depth, skipping the
        cells. The records of controls and OLE objects come after them, in
        the sheet's own children or in an ``mc:AlternateContent`` among
        those, and a sheet's cells can run to millions of elements."""
        for child in self._root.elements():
            local = local_name(child.name)
            if local == "sheetData":
                continue
            if local == name:
                yield child
            yield from child.descendants(name)

    def _sheet_records(self) -> dict[int, _SheetRecord]:
        """What the sheet's own records of its controls and OLE objects say,
        by shape id.

        Three things make this awkward, and all three are measured. A
        control carries no macro on its drawing shape: the sheet's
        ``<control>`` holds ``macro="[1]!Clicked"``, where the bracketed
        number names the workbook. Excel wraps ``<controls>`` in an
        ``mc:AlternateContent`` of its own, so looking for it among the
        worksheet's children finds nothing; the search is by name at any
        depth instead. And what the control is wired to is in neither
        place: the ``<control>`` points at a part of its own by
        relationship id, and the linked cell, the list range and the
        current value are in there.

        A record whose relationship points anywhere but at a control part
        is an ActiveX control's. Such a record can come twice, the full one
        in an ``mc:Choice`` and a bare copy in the ``mc:Fallback``, and the
        full one is the one kept.

        An OLE object is recorded the same way, as an ``<oleObject>`` in
        ``<oleObjects>`` with an ``<objectPr>`` where a control has its
        ``<controlPr>``. Measured: Excel marks its drawing twin hidden as
        well, and draws it from the VML.
        """
        package = self._workbook.package
        relationships = {one.id: one for one in package.relationships(self.part_name)}
        found: dict[int, _SheetRecord] = {}
        for tag, settings in (("control", "controlPr"), ("oleObject", "objectPr")):
            for record in self._outside_cells(tag):
                try:
                    shape_id = int(record.get("shapeId") or "")
                except ValueError:
                    continue
                if shape_id in found and _inside_fallback(record):
                    continue
                properties = record.child(settings)
                relationship = relationships.get(record.get("r:id") or "")
                kind: ShapeKind = "oleObject"
                if tag == "control":
                    active_x = relationship is not None and relationship.type != RT_CONTROL_PROPERTIES
                    kind = "activeX" if active_x else "formControl"
                found[shape_id] = _SheetRecord(
                    kind=kind,
                    name=record.get("name") or "",
                    macro="" if properties is None else (properties.get("macro") or ""),
                    alt_text="" if properties is None else (properties.get("altText") or ""),
                    control=(
                        self._control_part(relationship)
                        if kind == "formControl" and relationship is not None
                        else None
                    ),
                    anchor=None if properties is None else properties.child("anchor"),
                )
        return found

    def _control_part(self, relationship: Relationship) -> FormControl | None:
        """The control part a ``<control>`` points at, read.

        A missing part is not an error here. Excel writes one, but a package
        assembled by something else may not, and a shape with no wiring to
        report is better than a read that raises.
        """
        package = self._workbook.package
        name = relationship.target_part
        if relationship.is_external or not package.has_part(name):
            return None
        return read_control(package.xml(name).root, part_name=name, relationship=relationship.id)

    # ------------------------------------------------------------------
    # Hyperlinks
    # ------------------------------------------------------------------

    @property
    def hyperlinks(self) -> list[Hyperlink]:
        """Every link on the sheet, with external targets resolved.

        A link that leaves the workbook keeps its address in a relationship
        rather than in the element, so reading one means following ``r:id``.
        """
        container = self._root.child("hyperlinks")
        if container is None:
            return []
        targets = self._external_targets()
        found: list[Hyperlink] = []
        for element in container.children_named("hyperlink"):
            relationship_id = element.get("r:id")
            found.append(
                Hyperlink.read(
                    element,
                    target=None if relationship_id is None else targets.get(relationship_id),
                )
            )
        return found

    def hyperlink_at(self, reference: str | CellRef) -> Hyperlink | None:
        """The link covering a cell, if any."""
        cell = CellRef.parse(reference) if isinstance(reference, str) else reference
        for link in self.hyperlinks:
            if cell in link.ref:
                return link
        return None

    def add_hyperlink(
        self,
        reference: str | RangeRef,
        target: str | None = None,
        *,
        location: str | None = None,
        display: str | None = None,
        tooltip: str | None = None,
    ) -> Hyperlink:
        """Link a cell or a range.

        ``target`` is an address outside the workbook, such as a URL or a
        ``mailto:``; it is written as an external relationship because that
        is where Excel keeps it. ``location`` points inside the workbook,
        such as ``Sheet1!A1`` or a defined name. Give both for a URL with a
        fragment.

        The cell's own value is untouched: a hyperlink decorates whatever is
        there, and Excel shows ``display`` only when it is set.
        """
        if target is None and location is None:
            raise ValueError(
                "a hyperlink needs somewhere to go: pass target for an address outside "
                "the workbook, or location for a cell inside it."
            )
        block = (
            RangeRef.parse(reference) if isinstance(reference, str) else reference
        ).normalized

        self.remove_hyperlink(block)

        relationship_id: str | None = None
        if target is not None:
            relationship = self._workbook.package.relationships(self._part_name).add(
                RT_HYPERLINK, target, external=True
            )
            relationship_id = relationship.id

        link = Hyperlink(
            ref=block,
            target=target,
            location=location,
            display=display,
            tooltip=tooltip,
            relationship_id=relationship_id,
        )
        container = self._root.child("hyperlinks")
        if container is None:
            container = Element.create("hyperlinks")
            insert_in_schema_order(self._root, container, WORKSHEET_CHILD_ORDER)
        container.append(link.write())
        self._invalidate()
        return link

    def remove_hyperlink(self, reference: str | RangeRef) -> int:
        """Remove the links overlapping a range, and their relationships.

        Reports how many went. A link left pointing at a relationship that
        is gone, or a relationship with nothing pointing at it, is the kind
        of thing Excel repairs rather than opens.
        """
        container = self._root.child("hyperlinks")
        if container is None:
            return 0
        block = (
            RangeRef.parse(reference) if isinstance(reference, str) else reference
        ).normalized
        relationships = self._workbook.package.relationships(self._part_name)
        removed = 0
        for element in list(container.children_named("hyperlink")):
            existing = Hyperlink.read(element)
            if not existing.ref.intersects(block):
                continue
            if existing.relationship_id is not None:
                try:
                    relationships.remove(existing.relationship_id)
                except (KeyError, PackageError):
                    pass
            container.remove(element)
            removed += 1
        if removed:
            if not any(container.children_named("hyperlink")):
                self._root.remove(container)
            self._invalidate()
        return removed

    def _external_targets(self) -> dict[str, str]:
        """Every external relationship this sheet has, by id."""
        try:
            relationships = self._workbook.package.relationships(self._part_name)
        except PackageError:
            return {}
        return {
            relationship.id: relationship.target
            for relationship in relationships.by_type(RT_HYPERLINK)
        }

    # ------------------------------------------------------------------
    # Outline grouping
    # ------------------------------------------------------------------

    def group_rows(self, first: int, last: int, *, collapsed: bool = False) -> None:
        """Group rows into an outline, one level deeper than they were.

        ``collapsed`` hides them and marks the summary row as folded, which
        is what the outline's minus button does; the summary row, below the
        group or above it as :attr:`summary_below` says, stays visible. The
        sheet records how deep its outline goes, as Excel writes it.
        """
        if first < 1 or last < first or last > MAX_ROW:
            raise ValueError(f"rows {first} to {last} are not a range to group.")
        numbers = range(first, last + 1)
        self._check_outline_depth(self._level_of(self._rows[n]) for n in numbers if n in self._rows)
        self._ensure_rows(numbers)
        for number in numbers:
            row = self._rows[number]
            row.set("outlineLevel", str(self._level_of(row) + 1))
            if collapsed:
                row.set("hidden", "1")
        summary = last + 1 if self.summary_below else first - 1
        if collapsed and 1 <= summary <= MAX_ROW:
            self._ensure_rows([summary])
            self._rows[summary].set("collapsed", "1")
        self._record_outline_depth()
        self._invalidate()

    def ungroup_rows(self, first: int, last: int) -> None:
        """Take one level of grouping off, and show the rows it leaves
        ungrouped.

        Excel's own Ungroup leaves a folded group's rows hidden, with
        nothing left to unfold them by; they are shown here instead, and
        the summary row loses its folded mark once it has nothing to fold.
        """
        for number in range(first, last + 1):
            row = self._rows.get(number)
            if row is None:
                continue
            level = self._level_of(row) - 1
            if level > 0:
                row.set("outlineLevel", str(level))
            else:
                row.unset("outlineLevel")
                row.unset("hidden")
        summary = last + 1 if self.summary_below else first - 1
        element = self._rows.get(summary)
        if element is not None and element.get("collapsed") in ("1", "true") and not self._detail_rows(summary):
            element.unset("collapsed")
        self._record_outline_depth()
        self._invalidate()

    def row_outline_level(self, number: int) -> int:
        """How deep a row is in the outline. 0 when it is not grouped."""
        row = self.rows_by_number().get(number)
        return 0 if row is None else self._level_of(row)

    def group_columns(self, first: int, last: int, *, collapsed: bool = False) -> None:
        """Group columns into an outline, one level deeper; ``collapsed``
        as for :meth:`group_rows`, the summary column placed by
        :attr:`summary_right`."""
        if first < 1 or last < first or last > MAX_COLUMN:
            raise ValueError(f"columns {first} to {last} are not a range to group.")
        numbers = range(first, last + 1)
        self._check_outline_depth(self.column_outline_level(number) for number in numbers)
        container = self._ensure_cols()
        standard = self._standard_width()
        for number in numbers:
            entry = isolate_column(container, number, width=standard)
            entry.set("outlineLevel", str(self._level_of(entry) + 1))
            if collapsed:
                entry.set("hidden", "1")
        summary = last + 1 if self.summary_right else first - 1
        if collapsed and 1 <= summary <= MAX_COLUMN:
            isolate_column(container, summary, width=standard).set("collapsed", "1")
        self._record_outline_depth()
        self._invalidate()

    def ungroup_columns(self, first: int, last: int) -> None:
        """Take one level of grouping off; see :meth:`ungroup_rows`."""
        container = self._ensure_cols()
        standard = self._standard_width()
        for number in range(first, last + 1):
            entry = isolate_column(container, number, width=standard)
            level = self._level_of(entry) - 1
            if level > 0:
                entry.set("outlineLevel", str(level))
            else:
                entry.unset("outlineLevel")
                entry.unset("hidden")
        summary = last + 1 if self.summary_right else first - 1
        if 1 <= summary <= MAX_COLUMN:
            entry = column_entry(container, summary)
            if entry is not None and entry.get("collapsed") in ("1", "true") and not self._detail_columns(summary):
                isolate_column(container, summary, width=standard).unset("collapsed")
        self._tidy_cols(container)
        self._record_outline_depth()
        self._invalidate()

    def column_outline_level(self, number: int) -> int:
        """How deep a column is in the outline. 0 when it is not grouped."""
        container = self._root.child("cols")
        if container is None:
            return 0
        for entry in container.children_named("col"):
            first = _as_number(entry.get("min"))
            last = _as_number(entry.get("max"))
            if first is not None and last is not None and first <= number <= last:
                return self._level_of(entry)
        return 0

    @property
    def summary_below(self) -> bool:
        """Whether a group's summary row sits below it, as Excel defaults."""
        return self._outline_flag("summaryBelow")

    @summary_below.setter
    def summary_below(self, value: bool) -> None:
        self._set_outline_flag("summaryBelow", value)

    @property
    def summary_right(self) -> bool:
        """Whether a group's summary column sits to its right."""
        return self._outline_flag("summaryRight")

    @summary_right.setter
    def summary_right(self, value: bool) -> None:
        self._set_outline_flag("summaryRight", value)

    @staticmethod
    def _level_of(element: Element) -> int:
        return _as_number(element.get("outlineLevel")) or 0

    @staticmethod
    def _check_outline_depth(levels: Iterable[int]) -> None:
        """Refuse to group past the seventh level, the deepest an outline
        goes in Excel and in the schema."""
        if any(level >= MAX_OUTLINE_LEVEL for level in levels):
            raise ValueError(
                f"an outline goes {MAX_OUTLINE_LEVEL} levels deep, and part of this range is already there."
            )

    def _record_outline_depth(self) -> None:
        """``outlineLevelRow`` and ``outlineLevelCol`` on ``sheetFormatPr``:
        how deep the outline goes, which Excel writes while there is one
        and drops when there is none."""
        rows = max((self._level_of(row) for row in self._rows.values()), default=0)
        container = self._root.child("cols")
        columns = 0 if container is None else max(
            (self._level_of(entry) for entry in container.children_named("col")), default=0
        )
        head = self._root.child("sheetFormatPr")
        if head is None:
            if not rows and not columns:
                return
            head = Element.create("sheetFormatPr", {"defaultRowHeight": f"{DEFAULT_ROW_POINTS:g}"})
            insert_in_schema_order(self._root, head, WORKSHEET_CHILD_ORDER)
        for name, level in (("outlineLevelRow", rows), ("outlineLevelCol", columns)):
            if level and head.get(name) != str(level):
                head.set(name, str(level))
            elif not level:
                head.unset(name)

    def _detail_rows(self, number: int) -> list[int]:
        """The rows a summary row folds: the run beside it, on the side
        :attr:`summary_below` puts the detail, deeper in the outline."""
        level = self.row_outline_level(number)
        step = -1 if self.summary_below else 1
        detail: list[int] = []
        row = number + step
        while row in self._rows and self._level_of(self._rows[row]) > level:
            detail.append(row)
            row += step
        return detail

    def _detail_columns(self, number: int) -> list[int]:
        """The columns a summary column folds, as :meth:`_detail_rows`."""
        level = self.column_outline_level(number)
        step = -1 if self.summary_right else 1
        detail: list[int] = []
        column = number + step
        while 1 <= column <= MAX_COLUMN and self.column_outline_level(column) > level:
            detail.append(column)
            column += step
        return detail

    def _outline_flag(self, name: str) -> bool:
        properties = self._root.child("sheetPr")
        if properties is None:
            return True
        outline = properties.child("outlinePr")
        if outline is None:
            return True
        raw = outline.get(name)
        return True if raw is None else raw in ("1", "true")

    def _set_outline_flag(self, name: str, value: bool) -> None:
        properties = self._sheet_properties()
        outline = properties.child("outlinePr")
        if value:
            if outline is not None:
                outline.unset(name)
                if not outline.attributes:
                    properties.remove(outline)
                self._tidy_sheet_properties(properties)
                self._invalidate()
            return
        if outline is None:
            outline = Element.create("outlinePr")
            insert_in_schema_order(properties, outline, SHEET_PR_CHILD_ORDER)
        outline.set(name, "0")
        self._invalidate()

    # ------------------------------------------------------------------
    # Printing
    # ------------------------------------------------------------------

    @property
    def page_margins(self) -> PageMargins:
        """The margins, in inches. Excel's own defaults when unset."""
        element = self._root.child("pageMargins")
        return PageMargins() if element is None else PageMargins.read(element)

    @page_margins.setter
    def page_margins(self, value: PageMargins) -> None:
        self._replace_child("pageMargins", value.write())

    @property
    def page_setup(self) -> PageSetup:
        """Orientation, paper, scaling and numbering."""
        element = self._root.child("pageSetup")
        return PageSetup() if element is None else PageSetup.read(element)

    @page_setup.setter
    def page_setup(self, value: PageSetup) -> None:
        self._replace_child("pageSetup", value.write())

    @property
    def print_options(self) -> PrintOptions:
        """Centring, and whether headings and gridlines print."""
        element = self._root.child("printOptions")
        return PrintOptions() if element is None else PrintOptions.read(element)

    @print_options.setter
    def print_options(self, value: PrintOptions) -> None:
        if value.is_empty:
            self._drop_child("printOptions")
            return
        self._replace_child("printOptions", value.write())

    @property
    def header_footer(self) -> HeaderFooter:
        """The headers and footers, split into their three boxes."""
        element = self._root.child("headerFooter")
        return HeaderFooter() if element is None else HeaderFooter.read(element)

    @header_footer.setter
    def header_footer(self, value: HeaderFooter) -> None:
        if value.is_empty:
            self._drop_child("headerFooter")
            return
        self._replace_child("headerFooter", value.write())

    @property
    def fit_to_page(self) -> bool:
        """Whether the fit-to-width and fit-to-height numbers are used.

        They sit on ``<pageSetup>`` and do nothing until this says so, so
        setting them without this leaves the sheet printing at its scale.
        """
        properties = self._root.child("sheetPr")
        if properties is None:
            return False
        setup = properties.child("pageSetUpPr")
        return setup is not None and setup.get("fitToPage") in ("1", "true")

    @fit_to_page.setter
    def fit_to_page(self, value: bool) -> None:
        properties = self._sheet_properties()
        setup = properties.child("pageSetUpPr")
        if not value:
            if setup is not None:
                setup.unset("fitToPage")
                if not setup.attributes:
                    properties.remove(setup)
                self._tidy_sheet_properties(properties)
                self._invalidate()
            return
        if setup is None:
            setup = Element.create("pageSetUpPr")
            insert_in_schema_order(properties, setup, SHEET_PR_CHILD_ORDER)
        setup.set("fitToPage", "1")
        self._invalidate()

    @property
    def print_area(self) -> tuple[RangeRef, ...]:
        """The ranges that print, or empty for the whole used range.

        Stored as the built-in defined name ``_xlnm.Print_Area`` scoped to
        this sheet, which is why it already moves when rows are inserted.
        """
        return self._builtin_ranges(PRINT_AREA)

    @print_area.setter
    def print_area(self, value: str | RangeRef | Sequence[str | RangeRef] | None) -> None:
        if value is None:
            self._drop_builtin(PRINT_AREA)
            return
        ranges = _as_ranges(value)
        if not ranges:
            raise ValueError("a print area needs at least one range; pass None to clear it.")
        self._set_builtin(PRINT_AREA, ranges)

    @property
    def print_titles(self) -> str | None:
        """The rows and columns repeated on every page, as written.

        ``$1:$1`` repeats the first row, ``$A:$A`` the first column, and
        ``$A:$A,$1:$1`` both. Kept as text because the two are whole-axis
        references rather than ranges.
        """
        found = self._builtin(PRINT_TITLES)
        return None if found is None else found.refers_to

    @print_titles.setter
    def print_titles(self, value: str | None) -> None:
        if value is None:
            self._drop_builtin(PRINT_TITLES)
            return
        qualified = ",".join(
            piece if "!" in piece else f"{quote_sheet_name(self._name)}!{piece}"
            for piece in value.split(",")
        )
        self._put_builtin(PRINT_TITLES, qualified)

    def _builtin(self, name: str) -> DefinedName | None:
        """The sheet-scoped built-in name, or ``None`` when unset.

        ``Workbook.defined_name`` raises for a name that is not there, and
        an unset print area is the ordinary case rather than an error.
        """
        try:
            return self._workbook.defined_name(name, scope=self._name)
        except KeyError:
            return None

    def _put_builtin(self, name: str, refers_to: str, *, hidden: bool = False) -> None:
        """Define it, replacing whatever was there."""
        self._drop_builtin(name)
        self._workbook.add_defined_name(name, refers_to, scope=self._name, builtin=True, hidden=hidden)

    def _drop_builtin(self, name: str) -> None:
        if self._builtin(name) is not None:
            self._workbook.remove_defined_name(name, scope=self._name)

    def _builtin_ranges(self, name: str) -> tuple[RangeRef, ...]:
        found = self._builtin(name)
        if found is None:
            return ()
        blocks: list[RangeRef] = []
        for piece in found.refers_to.split(","):
            _, _, reference = piece.rpartition("!")
            try:
                blocks.append(RangeRef.parse(reference))
            except ValueError:
                continue
        return tuple(blocks)

    def _set_builtin(self, name: str, ranges: tuple[RangeRef, ...]) -> None:
        sheet = quote_sheet_name(self._name)
        self._put_builtin(
            name, ",".join(f"{sheet}!{block.absolute.a1}" for block in ranges)
        )

    def _replace_child(self, name: str, element: Element) -> None:
        existing = self._root.child(name)
        if existing is not None:
            self._root.insert_before(existing, element)
            self._root.remove(existing)
        else:
            insert_in_schema_order(self._root, element, WORKSHEET_CHILD_ORDER)
        self._invalidate()

    def _drop_child(self, name: str) -> None:
        existing = self._root.child(name)
        if existing is not None:
            self._root.remove(existing)
            self._invalidate()

    # ------------------------------------------------------------------
    # Data validation
    # ------------------------------------------------------------------

    @property
    def data_validations(self) -> list[DataValidation]:
        """Every validation on the sheet, in the order the part lists them."""
        container = self._root.child("dataValidations")
        if container is None:
            return []
        return [
            DataValidation.read(element)
            for element in container.children_named("dataValidation")
        ]

    def data_validation_at(self, reference: str | CellRef) -> DataValidation | None:
        """The validation covering a cell, if any.

        The first one found, which is what Excel applies: a cell carries one
        validation, and Excel replaces rather than stacks.
        """
        cell = CellRef.parse(reference) if isinstance(reference, str) else reference
        for validation in self.data_validations:
            if any(cell in block for block in validation.ranges):
                return validation
        return None

    def add_data_validation(
        self,
        reference: str | RangeRef | Sequence[str | RangeRef],
        validation: DataValidation,
    ) -> DataValidation:
        """Apply a validation to a range, or to several as one entry.

        Any validation already covering those cells is removed first, which
        is what Excel does: a cell has one rule, and leaving two behind
        makes which one applies depend on document order.
        """
        ranges = _as_ranges(reference)
        if not ranges:
            raise ValueError("a data validation needs at least one range.")
        if validation.kind == "none":
            raise ValueError(
                "this validation has no type, so it would accept anything. Build one "
                "with DataValidation.any_of, whole_number, decimal, date, time, "
                "text_length or custom."
            )

        for block in ranges:
            self._drop_validations_over(block)

        resolved = replace(validation, ranges=ranges)
        container = self._root.child("dataValidations")
        if container is None:
            container = Element.create("dataValidations")
            insert_in_schema_order(self._root, container, WORKSHEET_CHILD_ORDER)
        container.append(resolved.write())
        container.set("count", str(sum(1 for _ in container.children_named("dataValidation"))))
        self._invalidate()
        return resolved

    def clear_data_validations(self, reference: str | RangeRef | None = None) -> int:
        """Remove validations, and report how many entries went.

        With no argument every one goes. With a range, an entry is dropped
        only when all of its ranges fall inside it, so a rule that also
        covers cells outside is left alone rather than silently narrowed.
        """
        container = self._root.child("dataValidations")
        if container is None:
            return 0
        block = (
            None
            if reference is None
            else (RangeRef.parse(reference) if isinstance(reference, str) else reference).normalized
        )
        removed = 0
        for element in list(container.children_named("dataValidation")):
            if block is not None:
                covered = DataValidation.read(element).ranges
                if not covered or not all(block.contains(area) for area in covered):
                    continue
            container.remove(element)
            removed += 1
        if removed:
            self._tidy_validations(container)
            self._invalidate()
        return removed

    def _drop_validations_over(self, block: RangeRef) -> None:
        """Narrow or remove whatever already validates these cells."""
        container = self._root.child("dataValidations")
        if container is None:
            return
        for element in list(container.children_named("dataValidation")):
            existing = DataValidation.read(element)
            kept = tuple(area for area in existing.ranges if not area.intersects(block))
            if len(kept) == len(existing.ranges):
                continue
            if not kept:
                container.remove(element)
                continue
            # Excel does not split a range around a hole, and neither does
            # this: an area that merely overlaps is dropped whole, which is
            # visible in the result rather than silently partial.
            replacement = replace(existing, ranges=kept).write()
            container.insert_before(element, replacement)
            container.remove(element)
        self._tidy_validations(container)

    def _tidy_validations(self, container: Element) -> None:
        remaining = sum(1 for _ in container.children_named("dataValidation"))
        if remaining:
            container.set("count", str(remaining))
        else:
            self._root.remove(container)

    # ------------------------------------------------------------------
    # Conditional formatting
    # ------------------------------------------------------------------

    @property
    def conditional_formats(self) -> list[ConditionalFormatting]:
        """Every ``<conditionalFormatting>`` block, in document order."""
        return [
            ConditionalFormatting.read(element)
            for element in self._root.children_named("conditionalFormatting")
        ]

    def conditional_rules_at(self, reference: str | CellRef) -> list[ConditionalRule]:
        """The rules covering a cell, most important first.

        Excel applies rules in ascending ``priority``, so that is the order
        they come back in, and a lower number wins.
        """
        cell = CellRef.parse(reference) if isinstance(reference, str) else reference
        found: list[ConditionalRule] = []
        for block in self.conditional_formats:
            if any(cell in area for area in block.ranges):
                found.extend(block.rules)
        return sorted(found, key=lambda rule: rule.priority)

    def add_conditional_format(
        self,
        reference: str | RangeRef | Sequence[str | RangeRef],
        rule: ConditionalRule,
        *,
        dxf: Dxf | None = None,
        priority: int | None = None,
    ) -> ConditionalRule:
        """Apply a rule to a range, or to several ranges as one block.

        ``dxf`` is the formatting the rule paints; it is added to the
        workbook's table and the rule's ``dxfId`` set to point at it. A rule
        that paints nothing, such as a colour scale or data bar, needs none.

        The rule's compatibility formula is rebuilt for wherever it lands,
        which is not optional: see :mod:`pyofficeeditor.excel._conditional`.
        A ``cellIs`` or ``expression`` rule keeps the formula it was given,
        because there that formula is the condition.

        ``priority`` defaults to one past the highest already on the sheet,
        so a rule added later loses to one added earlier, which is what
        Excel's own "New Rule" does.
        """
        ranges = _as_ranges(reference)
        if not ranges:
            raise ValueError("a conditional format needs at least one range.")

        resolved = rule
        if dxf is not None and not dxf.is_empty:
            styles = self._workbook.styles
            if styles is None:
                raise ValueError(
                    "this workbook has no styles part, so a conditional format has "
                    "nowhere to record what it paints."
                )
            resolved = replace(resolved, dxf_id=styles.ensure_dxf(dxf))
        resolved = replace(resolved, priority=priority if priority is not None else self._next_priority())
        resolved = resolved.anchored_at(ranges[0].start)

        block = ConditionalFormatting(ranges=ranges, rules=(resolved,))
        insert_in_schema_order(self._root, block.write(), WORKSHEET_CHILD_ORDER)
        self._invalidate()
        return resolved

    def clear_conditional_formats(self, reference: str | RangeRef | None = None) -> int:
        """Remove conditional formatting, and report how many blocks went.

        With no argument every block on the sheet goes. With a range, only
        the blocks whose ranges all fall inside it: a block that also covers
        cells outside the range is left alone rather than silently narrowed,
        because narrowing it would change what the rest of the sheet shows.
        """
        block_range = (
            None
            if reference is None
            else (RangeRef.parse(reference) if isinstance(reference, str) else reference).normalized
        )
        removed = 0
        for element in list(self._root.children_named("conditionalFormatting")):
            if block_range is not None:
                covered = ConditionalFormatting.read(element).ranges
                if not covered or not all(block_range.contains(area) for area in covered):
                    continue
            self._root.remove(element)
            removed += 1
        if removed:
            self._invalidate()
        return removed

    def _next_priority(self) -> int:
        highest = 0
        for block in self.conditional_formats:
            for rule in block.rules:
                highest = max(highest, rule.priority)
        return highest + 1

    def merged_range_at(self, reference: CellRef) -> RangeRef | None:
        """The merged block covering a cell, if any.

        Worth asking before concluding a cell is empty: every cell of a merge
        but the top left reads as ``None``, because the value lives on the
        anchor.
        """
        for block in self.merged_ranges:
            if reference in block:
                return block
        return None

    def merge(self, reference: str | RangeRef) -> RangeRef:
        """Merge a block so it displays as one cell.

        The top-left cell keeps its value and every other cell in the block
        is cleared, which is what Excel does: the covered cells are not
        displayed, so data left in them would be invisible and misleading.
        Check :attr:`merged_ranges` or read the cells first if that matters.

        The covered cells are still written, carrying the anchor's style,
        because that is how a border renders across a merge.
        """
        block = (RangeRef.parse(reference) if isinstance(reference, str) else reference).normalized
        if block.is_single_cell:
            raise ValueError(f"{block.a1} is one cell; there is nothing to merge it with.")
        for existing in self.merged_ranges:
            if existing.intersects(block):
                raise ValueError(
                    f"{block.a1} overlaps the merged range {existing.a1}. Excel repairs a "
                    f"worksheet whose merges overlap rather than rendering it; unmerge "
                    f"{existing.a1} first."
                )

        anchor = CellRef(block.top, block.left)
        anchor_style = self.style_index(anchor)
        for cell in block.cells():
            if cell.sort_key == anchor.sort_key:
                continue
            element = self._ensure_cell(cell)
            self._drop_children(element, "v")
            self._drop_children(element, "is")
            self._drop_children(element, "f")
            element.unset("t")
            if anchor_style is not None:
                element.set("s", str(anchor_style))
        self._workbook.mark_values_changed()

        container = self._root.child("mergeCells")
        if container is None:
            container = Element.create("mergeCells")
            insert_in_schema_order(self._root, container, WORKSHEET_CHILD_ORDER)
        container.append(Element.create("mergeCell", {"ref": block.a1}))
        container.set("count", str(sum(1 for _ in container.children_named("mergeCell"))))
        self._invalidate()
        return block

    def unmerge(self, reference: str | RangeRef) -> RangeRef:
        """Split a merged block apart again.

        Accepts the block's own reference or any cell inside it, the way
        Excel's own command works on a selection.
        """
        wanted = (RangeRef.parse(reference) if isinstance(reference, str) else reference).normalized
        container = self._root.child("mergeCells")
        if container is not None:
            for entry in list(container.children_named("mergeCell")):
                raw = entry.get("ref")
                if raw is None:
                    continue
                try:
                    block = RangeRef.parse(raw).normalized
                except ValueError:
                    continue
                if block.a1 == wanted.a1 or (
                    wanted.is_single_cell and CellRef(wanted.top, wanted.left) in block
                ):
                    container.remove(entry)
                    remaining = sum(1 for _ in container.children_named("mergeCell"))
                    if remaining:
                        container.set("count", str(remaining))
                    else:
                        # Excel omits the element rather than writing count="0".
                        self._root.remove(container)
                    self._invalidate()
                    return block
        raise ValueError(f"{wanted.a1} is not a merged range on {self._name!r}.")

    # ------------------------------------------------------------------
    # Extent
    # ------------------------------------------------------------------

    @property
    def dimension(self) -> RangeRef | None:
        """The block the part records as used, or ``None`` if it has no
        ``dimension`` element."""
        element = self._root.child("dimension")
        if element is None:
            return None
        reference = element.get("ref")
        if reference is None:
            return None
        try:
            return RangeRef.parse(reference)
        except ValueError:
            return None

    @property
    def max_row(self) -> int:
        """The highest row number that has a cell, or 0 for an empty sheet.

        A row that is only hidden, sized or styled has an element and no
        cell, and does not count: rows a filter hid below the data would
        otherwise move where the data ends.
        """
        for number in sorted(self._rows, reverse=True):
            if next(self._rows[number].children_named("c"), None) is not None:
                return number
        return 0

    @property
    def max_column(self) -> int:
        """The highest column number that has a cell, or 0."""
        highest = 0
        for row in self._rows.values():
            for cell in row.children_named("c"):
                reference = cell.get("r")
                if reference is None:
                    continue
                try:
                    highest = max(highest, CellRef.parse(reference).column)
                except ValueError:
                    continue
        return highest

    @property
    def used_range(self) -> RangeRef | None:
        """The smallest block covering every cell the sheet carries.

        Computed from the cells, unlike :attr:`dimension`, which is what the
        file claims. They can disagree: a producer may leave a stale
        dimension behind, and Excel tolerates it.
        """
        rows = self.max_row
        columns = self.max_column
        if rows == 0 or columns == 0:
            return None
        return RangeRef(CellRef(1, 1), CellRef(rows, columns))

    # ------------------------------------------------------------------
    # Iteration
    # ------------------------------------------------------------------

    def rows(self) -> Iterator[list[Cell]]:
        """Every row that has cells, in order, as the cells it has.

        Absent cells are not padded; use :meth:`range` when a rectangle is
        wanted.
        """
        for number in sorted(self._rows):
            found: list[Cell] = []
            for element in self._rows[number].children_named("c"):
                reference = element.get("r")
                if reference is None:
                    continue
                try:
                    found.append(Cell(self, CellRef.parse(reference)))
                except ValueError:
                    continue
            yield found

    def values(self) -> Iterator[list[CellValue]]:
        """The used rectangle, row by row, padded with ``None``."""
        block = self.used_range
        if block is None:
            return
        for row in block.rows():
            yield [self.get_value(reference) for reference in row]

    # ------------------------------------------------------------------
    # Element plumbing
    # ------------------------------------------------------------------

    def _find_cell(self, reference: CellRef) -> Element | None:
        row = self._rows.get(reference.row)
        if row is None:
            return None
        wanted = reference.relative.a1
        for cell in row.children_named("c"):
            if (cell.get("r") or "") == wanted:
                return cell
        return None

    def _ensure_cell(self, reference: CellRef) -> Element:
        row = self._ensure_row(reference.row)
        wanted = reference.relative.a1
        before: Element | None = None
        for cell in row.children_named("c"):
            raw = cell.get("r")
            if raw == wanted:
                return cell
            if raw is None:
                continue
            try:
                column = CellRef.parse(raw).column
            except ValueError:
                continue
            if column > reference.column:
                before = cell
                break
        created = Element.create("c", {"r": wanted})
        if before is None:
            row.append(created)
        else:
            row.insert_before(before, created)
        self._widen_spans(row, reference.column)
        self._widen_dimension(reference)
        return created

    def _ensure_row(self, number: int) -> Element:
        existing = self._rows.get(number)
        if existing is not None:
            return existing
        created = Element.create("row", {"r": str(number)})
        # A row past the last, which is how a sheet gets written, is an
        # append; searching for its place each time made that quadratic.
        if number > self._highest_row:
            self._data.append(created)
            self._rows[number] = created
            self._highest_row = number
            return created
        following = min(n for n in self._rows if n > number)
        self._data.insert_before(self._rows[following], created)
        self._rows[number] = created
        return created

    def _remove_cell(self, reference: CellRef) -> None:
        row = self._rows.get(reference.row)
        if row is None:
            return
        element = self._find_cell(reference)
        if element is None:
            return
        row.remove(element)
        # A row with nothing left to say goes. One that is hidden, sized,
        # styled or in an outline keeps saying it: removing it would, say,
        # bring back a row a filter had hidden.
        if next(row.children_named("c"), None) is None and set(row.attributes) <= {"r", "spans"}:
            self._data.remove(row)
            del self._rows[reference.row]
            if reference.row == self._highest_row:
                self._highest_row = max(self._rows, default=0)

    @staticmethod
    def _drop_children(element: Element, name: str) -> None:
        existing = element.child(name)
        while existing is not None:
            element.remove(existing)
            existing = element.child(name)

    def _widen_spans(self, row: Element, column: int) -> None:
        """Keep a row's ``spans`` hint covering its cells.

        Excel writes ``spans="1:6"`` as a rendering hint. A stale one is
        tolerated, but keeping it right costs nothing and a wrong one has
        been known to confuse other readers.
        """
        current = row.get("spans")
        if current is None:
            return
        try:
            first, _, last = current.partition(":")
            low, high = int(first), int(last)
        except ValueError:
            return
        if low <= column <= high:
            return
        row.set("spans", f"{min(low, column)}:{max(high, column)}")

    def _widen_dimension(self, reference: CellRef) -> None:
        element = self._root.child("dimension")
        if element is None:
            element = Element.create("dimension", {"ref": reference.relative.a1})
            insert_in_schema_order(self._root, element, WORKSHEET_CHILD_ORDER)
            return
        current = element.get("ref")
        if current is None:
            element.set("ref", reference.relative.a1)
            return
        try:
            block = RangeRef.parse(current)
        except ValueError:
            element.set("ref", reference.relative.a1)
            return
        if reference in block:
            return
        element.set("ref", block.expanded(reference).a1)

    def _shared_master(self, index: str) -> tuple[CellRef, str] | None:
        """The cell and text of a shared-formula group's master."""
        if self._masters is None:
            masters: dict[str, tuple[CellRef, str]] = {}
            for row in self._data.children_named("row"):
                for cell in row.children_named("c"):
                    formula = cell.child("f")
                    if formula is None or (formula.get("t") or "") != "shared":
                        continue
                    key = formula.get("si")
                    text = formula.text
                    reference = cell.get("r")
                    if key is None or not text or reference is None:
                        continue
                    try:
                        masters.setdefault(key, (CellRef.parse(reference), text))
                    except ValueError:
                        continue
            self._masters = masters
        return self._masters.get(index)

    def rows_by_number(self) -> dict[int, Element]:
        """The ``<row>`` elements this sheet has, keyed by row number.

        A copy, so a caller may renumber the rows while walking it.
        """
        return dict(self._rows)

    def reindex_rows(self) -> None:
        """Rebuild the row lookup after the rows were renumbered."""
        found: dict[int, Element] = {}
        for row in self._data.children_named("row"):
            raw = row.get("r")
            if raw is None:
                continue
            try:
                found[int(raw)] = row
            except ValueError:
                continue
        self._rows = found
        self._highest_row = max(found, default=0)

    def invalidate(self) -> None:
        """Record that this sheet changed.

        Drops the shared-formula lookup, which the change may have
        invalidated, and tells the workbook so the file is saved with
        ``fullCalcOnLoad``.
        """
        self._invalidate()

    def _invalidate(self) -> None:
        self._masters = None
        self._workbook.mark_changed()

    def __repr__(self) -> str:
        return f"<Worksheet {self._name!r} {self._part_name}>"


class Cell:
    """One cell, addressed whether or not the file has it yet.

    A thin view over :class:`Worksheet`: every property here delegates to a
    method there, so nothing is only reachable through a cell.
    """

    __slots__ = ("_reference", "_sheet")

    def __init__(self, sheet: Worksheet, reference: CellRef) -> None:
        self._sheet = sheet
        self._reference = reference

    @property
    def sheet(self) -> Worksheet:
        return self._sheet

    @property
    def reference(self) -> CellRef:
        return self._reference

    @property
    def a1(self) -> str:
        return self._reference.a1

    @property
    def row(self) -> int:
        return self._reference.row

    @property
    def column(self) -> int:
        return self._reference.column

    @property
    def exists(self) -> bool:
        return self._sheet.has_cell(self._reference)

    @property
    def value(self) -> CellValue:
        return self._sheet.get_value(self._reference)

    @value.setter
    def value(self, value: CellValue) -> None:
        self._sheet.set_value(self._reference, value)

    @property
    def text(self) -> str:
        """What Excel shows in the cell; see :meth:`Worksheet.get_text`."""
        return self._sheet.get_text(self._reference)

    @property
    def rich_text(self) -> tuple[TextRun, ...] | None:
        """The text as runs, each in its font; see :meth:`Worksheet.get_rich_text`."""
        return self._sheet.get_rich_text(self._reference)

    @rich_text.setter
    def rich_text(self, runs: Sequence[TextRun | str]) -> None:
        self._sheet.set_rich_text(self._reference, runs)

    @property
    def comment(self) -> Comment | None:
        """The cell's note, or ``None``; see :meth:`Worksheet.set_comment`."""
        return self._sheet.comment(self._reference)

    @comment.setter
    def comment(self, text: str | None) -> None:
        """Set the note's text, with no author and hidden, or ``None`` to
        take it off."""
        if text is None:
            self._sheet.remove_comment(self._reference)
        else:
            self._sheet.set_comment(self._reference, text)

    @property
    def formula(self) -> str | None:
        return self._sheet.get_formula(self._reference)

    @formula.setter
    def formula(self, formula: str | None) -> None:
        self._sheet.set_formula(self._reference, formula)

    @property
    def style_index(self) -> int | None:
        return self._sheet.style_index(self._reference)

    @property
    def number_format(self) -> str:
        return self._sheet.number_format(self._reference)

    @property
    def format(self) -> CellFormat:
        """How the cell looks.  Derive from it rather than replacing it::

            cell.format = cell.format.with_font(bold=True)

        Assigning a bare :class:`~pyofficeeditor.excel.CellFormat` resets
        every aspect it does not mention, which is occasionally what you
        want and usually not.
        """
        return self._sheet.get_format(self._reference)

    @format.setter
    def format(self, wanted: CellFormat) -> None:
        self._sheet.set_format(self._reference, wanted)

    @property
    def style(self) -> str | None:
        """The name of the cell's style; see :meth:`Worksheet.get_cell_style`.
        Assigning a name gives the cell that style, as Excel does."""
        return self._sheet.get_cell_style(self._reference)

    @style.setter
    def style(self, name: str) -> None:
        self._sheet.set_cell_style(self._reference, name)

    @property
    def font(self) -> Font:
        return self.format.font

    @font.setter
    def font(self, font: Font) -> None:
        self._sheet.set_format(self._reference, replace(self.format, font=font))

    @property
    def fill(self) -> Fill:
        return self.format.fill

    @fill.setter
    def fill(self, fill: Fill | Color | str) -> None:
        """Set the background.  A color or a hex string means a solid fill."""
        self._sheet.set_format(self._reference, self.format.with_fill(fill))

    @property
    def border(self) -> Border:
        return self.format.border

    @border.setter
    def border(self, border: Border) -> None:
        self._sheet.set_format(self._reference, replace(self.format, border=border))

    @property
    def alignment(self) -> Alignment:
        return self.format.alignment

    @alignment.setter
    def alignment(self, alignment: Alignment) -> None:
        self._sheet.set_format(self._reference, replace(self.format, alignment=alignment))

    @property
    def merged_range(self) -> RangeRef | None:
        """The merged block this cell belongs to, if any."""
        return self._sheet.merged_range_at(self._reference)

    @property
    def is_merged(self) -> bool:
        return self.merged_range is not None

    @property
    def is_merge_anchor(self) -> bool:
        """Whether this is the top-left cell of a merge, the one that holds
        the value the whole block displays."""
        block = self.merged_range
        if block is None:
            return False
        return self._reference.sort_key == (block.top, block.left)

    def clear(self) -> None:
        self._sheet.clear_cell(self._reference)

    def __repr__(self) -> str:
        return f"<Cell {self._sheet.name}!{self.a1}>"


class Range:
    """A rectangular block of cells on one sheet."""

    __slots__ = ("_reference", "_sheet")

    def __init__(self, sheet: Worksheet, reference: RangeRef) -> None:
        self._sheet = sheet
        self._reference = reference.normalized

    @property
    def sheet(self) -> Worksheet:
        return self._sheet

    @property
    def reference(self) -> RangeRef:
        return self._reference

    @property
    def a1(self) -> str:
        return self._reference.a1

    def __iter__(self) -> Iterator[Cell]:
        for reference in self._reference.cells():
            yield Cell(self._sheet, reference)

    def __len__(self) -> int:
        return self._reference.size

    def rows(self) -> Iterator[list[Cell]]:
        for row in self._reference.rows():
            yield [Cell(self._sheet, reference) for reference in row]

    @property
    def values(self) -> list[list[CellValue]]:
        """The block's values, row by row."""
        return [
            [self._sheet.get_value(reference) for reference in row]
            for row in self._reference.rows()
        ]

    @values.setter
    def values(self, values: list[list[CellValue]]) -> None:
        """Fill the block from a rectangle of values.

        The rectangle must match the range's shape exactly, because silently
        filling part of a block is how half-written data gets saved.
        """
        rows = list(self._reference.rows())
        if len(values) != len(rows):
            raise ValueError(f"{self.a1} has {len(rows)} rows, the values have {len(values)}.")
        for target_row, source_row in zip(rows, values, strict=True):
            if len(source_row) != len(target_row):
                raise ValueError(
                    f"{self.a1} is {len(target_row)} columns wide, a row of values has "
                    f"{len(source_row)}."
                )
            for reference, value in zip(target_row, source_row, strict=True):
                self._sheet.set_value(reference, value)

    def set_format(self, wanted: CellFormat) -> None:
        """Give every cell in the block the same format.

        One entry is added to the workbook's tables however large the block
        is, because the cells all end up pointing at it.
        """
        for reference in self._reference.cells():
            self._sheet.set_format(reference, wanted)

    def apply_font(self, **changes: object) -> None:
        """Change some font aspects across the block, keeping the rest.

        Each cell keeps its own other formatting, so emboldening a row of
        differently coloured cells leaves the colours alone::

            sheet.range("A1:F1").apply_font(bold=True)
        """
        for reference in self._reference.cells():
            current = self._sheet.get_format(reference)
            self._sheet.set_format(reference, replace(current, font=replace(current.font, **changes)))  # type: ignore[arg-type]

    def apply_fill(self, fill: Fill | Color | str) -> None:
        """Give every cell in the block a background, keeping the rest."""
        for reference in self._reference.cells():
            current = self._sheet.get_format(reference)
            self._sheet.set_format(reference, current.with_fill(fill))

    def apply_border(self, border: Border) -> None:
        """Give every cell in the block the same edges, keeping the rest."""
        for reference in self._reference.cells():
            current = self._sheet.get_format(reference)
            self._sheet.set_format(reference, replace(current, border=border))

    def apply_alignment(self, **changes: object) -> None:
        """Change some alignment aspects across the block, keeping the rest."""
        for reference in self._reference.cells():
            current = self._sheet.get_format(reference)
            self._sheet.set_format(
                reference, replace(current, alignment=replace(current.alignment, **changes))  # type: ignore[arg-type]
            )

    def apply_number_format(self, code: str) -> None:
        """Give every cell in the block a number format, keeping the rest."""
        for reference in self._reference.cells():
            current = self._sheet.get_format(reference)
            self._sheet.set_format(reference, current.with_number_format(code))

    def apply_style(self, name: str) -> None:
        """Give every cell in the block a named style, each keeping its own
        formatting for what the style does not set."""
        for reference in self._reference.cells():
            self._sheet.set_cell_style(reference, name)

    def clear(self) -> None:
        """Remove every cell in the block."""
        for reference in self._reference.cells():
            self._sheet.clear_cell(reference)

    def __repr__(self) -> str:
        return f"<Range {self._sheet.name}!{self.a1}>"


__all__ = ["Cell", "Range", "Worksheet", "column_letter"]


def _as_cell(reference: str | CellRef) -> CellRef:
    """A cell as the parts that store one address it: ``C3``, not ``$C$3``."""
    return (CellRef.parse(reference) if isinstance(reference, str) else reference).relative


def _now() -> dt.datetime:
    return dt.datetime.now(dt.timezone.utc)


def _as_number(raw: str | None) -> int | None:
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None
