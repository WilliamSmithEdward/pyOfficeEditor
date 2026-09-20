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

from collections.abc import Iterator
from typing import TYPE_CHECKING

from pyofficeeditor._xml import Element, XmlDocument
from pyofficeeditor.excel._formulas import shared_formula_for
from pyofficeeditor.excel._reference import CellRef, RangeRef, column_letter
from pyofficeeditor.excel._schema import WORKSHEET_CHILD_ORDER, insert_in_schema_order
from pyofficeeditor.excel._values import CellValue, read_value, write_value

if TYPE_CHECKING:
    from pyofficeeditor.excel.workbook import Workbook


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

    def clear_cell(self, reference: CellRef) -> None:
        """Remove a cell, leaving the sheet as if it were never set."""
        self._remove_cell(reference)
        self._invalidate()

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

    def clear(self) -> None:
        """Remove every cell in the block."""
        for reference in self._reference.cells():
            self._sheet.clear_cell(reference)

    def __repr__(self) -> str:
        return f"<Range {self._sheet.name}!{self.a1}>"


__all__ = ["Cell", "Range", "Worksheet", "column_letter"]
