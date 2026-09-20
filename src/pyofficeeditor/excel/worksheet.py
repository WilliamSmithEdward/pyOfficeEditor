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

from collections.abc import Iterator, Sequence
from dataclasses import replace
from typing import TYPE_CHECKING, Literal

from pyofficeeditor._xml import Element, XmlDocument
from pyofficeeditor.excel._conditional import ConditionalFormatting, ConditionalRule
from pyofficeeditor.excel._dimensions import (
    COLUMN_ATTRIBUTES,
    Freeze,
    column_entry,
    isolate_column,
)
from pyofficeeditor.excel._dxf import Dxf
from pyofficeeditor.excel._formats import (
    Alignment,
    Border,
    CellFormat,
    Color,
    Fill,
    Font,
)
from pyofficeeditor.excel._formulas import quote_sheet_name, shared_formula_for
from pyofficeeditor.excel._names import PRINT_AREA, PRINT_TITLES, DefinedName
from pyofficeeditor.excel._pagesetup import (
    HeaderFooter,
    PageMargins,
    PageSetup,
    PrintOptions,
)
from pyofficeeditor.excel._protection import SheetProtection
from pyofficeeditor.excel._reference import CellRef, RangeRef, column_letter
from pyofficeeditor.excel._rowcol import (
    delete_columns,
    delete_rows,
    insert_columns,
    insert_rows,
)
from pyofficeeditor.excel._schema import (
    SHEET_PR_CHILD_ORDER,
    WORKSHEET_CHILD_ORDER,
    insert_in_schema_order,
)
from pyofficeeditor.excel._tables import (
    CT_TABLE,
    RT_TABLE,
    Table,
    TableStyle,
    build_table_part,
    unique_column_names,
)
from pyofficeeditor.excel._validation import DataValidation
from pyofficeeditor.excel._values import CellValue, read_value, write_value
from pyofficeeditor.exceptions import PackageError

#: What ``<sheet state=...>`` can say. A ``veryHidden`` sheet is not in
#: Excel's unhide list, so only code can bring it back.
SheetVisibility = Literal["visible", "hidden", "veryHidden"]

if TYPE_CHECKING:
    from pyofficeeditor.excel.workbook import Workbook


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

        self._rows: dict[int, Element] = {}
        for row in self._data.children_named("row"):
            raw = row.get("r")
            if raw is None:
                continue
            try:
                self._rows[int(raw)] = row
            except ValueError:
                continue
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
        """Record a new name on this object.

        A sheet's name lives in the workbook part, not in its own, so this
        only updates what the object reports. Call
        :meth:`Workbook.rename_sheet`, which validates the name, moves the
        entry, and repoints every formula and defined name that referred to
        the old one.
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
        self._invalidate()

    def get_formula(self, reference: CellRef) -> str | None:
        """A cell's formula, without the leading ``=``.

        A cell in a shared-formula group carries no text of its own, so its
        formula is derived from the group's master by shifting the relative
        references. That derivation is why this exists rather than callers
        reading ``<f>`` themselves.
        """
        element = self._find_cell(reference)
        if element is None:
            return None
        formula = element.child("f")
        if formula is None:
            return None
        text = formula.text
        if text:
            return text
        if (formula.get("t") or "") != "shared":
            return None
        index = formula.get("si")
        if index is None:
            return None
        master = self._shared_master(index)
        if master is None:
            return None
        master_cell, master_text = master
        return shared_formula_for(master_text, master_cell, reference)

    def set_formula(self, reference: CellRef, formula: str | None) -> None:
        """Put a formula in a cell, or remove the one it has.

        The cell's cached result goes too: a stale ``<v>`` beside a new
        ``<f>`` is the old answer, and Excel shows it until it recalculates.
        """
        element = self._ensure_cell(reference)
        self._drop_children(element, "f")

        if formula is None:
            self._invalidate()
            return

        self._drop_children(element, "v")
        element.unset("t")
        node = Element.create("f")
        node.set_text(formula[1:] if formula.startswith("=") else formula)
        element.insert(0, node)
        self._invalidate()

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

    def clear_cell(self, reference: CellRef) -> None:
        """Remove a cell, leaving the sheet as if it were never set."""
        self._remove_cell(reference)
        self._invalidate()

    # ------------------------------------------------------------------
    # Inserting rows and columns
    # ------------------------------------------------------------------

    def insert_rows(self, at: int, count: int = 1) -> None:
        """Insert blank rows, pushing everything at or below ``at`` down.

        Everything that records a cell address moves with the data: the
        cells themselves, every formula in the workbook that reads from this
        sheet, shared-formula groups, merged ranges, hyperlinks, tables, the
        sheet's filter and dimension, its page breaks, and the workbook's
        defined names.

        The insertion is refused when the sheet carries something that
        addresses cells and this library cannot move, such as conditional
        formatting or data validation. Moving everything else and leaving
        those behind produces a workbook that opens cleanly and points at
        the wrong cells, which is worse than not doing it.
        """
        insert_rows(self, at, count)

    def insert_columns(self, at: int, count: int = 1) -> None:
        """Insert blank columns, pushing everything at or right of ``at``
        over.  The same shifting and the same refusal as :meth:`insert_rows`."""
        insert_columns(self, at, count)

    def delete_rows(self, at: int, count: int = 1) -> None:
        """Delete rows, closing the gap behind them.

        Everything that referred to the deleted rows is repointed: a formula
        reading one of them becomes #REF!, while a range that only partly
        overlapped shrinks instead. Merges, tables, hyperlinks and defined
        names shrink or go the same way.

        Refused, like an insertion, when the sheet carries something that
        addresses cells and this library cannot move. Also refused when it
        would remove a table's header row, since a table's column names come
        from there.
        """
        delete_rows(self, at, count)

    def delete_columns(self, at: int, count: int = 1) -> None:
        """Delete columns, closing the gap.  The same repointing and the same
        refusals as :meth:`delete_rows`."""
        delete_columns(self, at, count)

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
        """Set a column's width, or ``None`` to restore the default.

        ``customWidth`` goes with it: a ``width`` without that flag is
        ignored by Excel, so the change would look like it never happened.
        """
        container = self._ensure_cols()
        entry = isolate_column(container, column)
        if width is None:
            entry.unset("width")
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
        entry = isolate_column(container, column)
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

        ``customHeight`` goes with it, for the same reason ``customWidth``
        does: without the flag Excel ignores the value.
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
        for entry in list(container.children_named("col")):
            if not any(entry.get(name) is not None for name in COLUMN_ATTRIBUTES):
                container.remove(entry)
        if next(container.children_named("col"), None) is None:
            self._root.remove(container)

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

    def _put_builtin(self, name: str, refers_to: str) -> None:
        """Define it, replacing whatever was there."""
        self._drop_builtin(name)
        self._workbook.add_defined_name(name, refers_to, scope=self._name, builtin=True)

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
        """The highest row number that has a cell, or 0 for an empty sheet."""
        return max(self._rows, default=0)

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
        later = [n for n in self._rows if n > number]
        if later:
            self._data.insert_before(self._rows[min(later)], created)
        else:
            self._data.append(created)
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
        if next(row.children_named("c"), None) is None:
            self._data.remove(row)
            del self._rows[reference.row]

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
        self._rows = {}
        for row in self._data.children_named("row"):
            raw = row.get("r")
            if raw is None:
                continue
            try:
                self._rows[int(raw)] = row
            except ValueError:
                continue

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

    def clear(self) -> None:
        """Remove every cell in the block."""
        for reference in self._reference.cells():
            self._sheet.clear_cell(reference)

    def __repr__(self) -> str:
        return f"<Range {self._sheet.name}!{self.a1}>"


__all__ = ["Cell", "Range", "Worksheet", "column_letter"]
