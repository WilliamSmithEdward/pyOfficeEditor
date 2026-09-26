"""Calculating a workbook: every formula, in the order its inputs allow,
with the results written back into the cells.

The workbook is read once into a grid per sheet. A formula cell starts
uncalculated; evaluating one that reads an uncalculated cell stops with
:class:`~.evaluator.PendingCellsError`, and the cells it named are calculated
first, on an explicit stack, before it is tried again. So a column of ten
thousand running totals, each reading the one above, calculates without
ten thousand nested Python calls.

A formula that cannot be calculated here, because it calls a function
this engine does not have, reads a workbook the file keeps no link to,
or will not parse, keeps the value the file cached for it, and so does
every formula that reads it: its value is only as good as that cache.
:class:`Calculation` lists both. A circular reference keeps its cached
values too, as Excel, with iteration off, leaves them.

Another workbook is read as Excel reads it while it is closed: from the
cells its link caches, loaded as sheets named ``[1]Sheet1`` beside the
workbook's own.

A what-if data table is calculated as Excel calculates one: each value it
tries is put in its input cell, and its formula calculated with it, by a
view of the workbook that calculates again only what reads that cell.
"""

from __future__ import annotations

import bisect
import dataclasses
import datetime as dt
import os
import typing
from collections.abc import Iterable, Iterator
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyofficeeditor._xml import Element

# Imported for what importing it does: registering every function.
from pyofficeeditor.excel._calc import functions as functions
from pyofficeeditor.excel._calc.evaluator import (
    CellKey,
    Context,
    PendingCellsError,
    TableShape,
    UnsupportedFormulaError,
)
from pyofficeeditor.excel._calc.lexer import FormulaSyntaxError
from pyofficeeditor.excel._calc.nodes import Call, NameReference, Node, Prefix, walk
from pyofficeeditor.excel._calc.parser import parse
from pyofficeeditor.excel._calc.values import EMPTY, Area, Array, Empty, Reference, Scalar, Value
from pyofficeeditor.excel._externals import read_external_books
from pyofficeeditor.excel._numfmt import BUILTIN_DISPLAY_CODES
from pyofficeeditor.excel._pivots import PivotReport, read_pivot_report
from pyofficeeditor.excel._reference import CellRef, RangeRef
from pyofficeeditor.excel._values import CellError, datetime_to_serial, read_value, write_cached

if TYPE_CHECKING:
    from pyofficeeditor.excel.workbook import Workbook
    from pyofficeeditor.excel.worksheet import Worksheet


@dataclass(frozen=True)
class _DataTable:
    """What a data table puts in its input cells: the values along the row
    above its block when ``along_row``, else those down the column left of
    it, into ``first``; or, for two variables, the row's into ``first`` and
    the column's into ``second``, trying the formula at the corner."""

    two_way: bool
    along_row: bool
    first: CellKey
    second: CellKey | None = None


@dataclass
class _Formula:
    key: CellKey
    text: str
    element: Element
    cached: Scalar
    node: Node | None = None
    #: What stops it being calculated, when something does.
    problem: str | None = None
    #: The block an array formula or a data table fills.
    area: Area | None = None
    table: _DataTable | None = None


@dataclass(frozen=True)
class _Override:
    """A value a data table puts in one of its input cells."""

    value: Scalar


@dataclass(frozen=True)
class _Part:
    """A cell of an array formula's block other than the one holding it."""

    master: CellKey
    row: int
    column: int


Content = Scalar | _Formula | _Part | _Override

#: The functions whose cells SUBTOTAL and AGGREGATE leave out.
_SUBTOTALS = frozenset({"SUBTOTAL", "AGGREGATE"})


@dataclass
class Calculation:
    """What calculating a workbook did."""

    #: How many formula cells now hold a result calculated here.
    calculated: int = 0
    #: Formula cells that kept their cached value, and why, by address.
    unsupported: dict[str, str] = field(default_factory=dict[str, str])
    #: Formula cells that read one of those, directly or not.
    dependent: list[str] = field(default_factory=list[str])
    #: Formula cells in a circular reference.
    circular: list[str] = field(default_factory=list[str])

    @property
    def complete(self) -> bool:
        """Whether every formula was calculated here."""
        return not (self.unsupported or self.dependent or self.circular)


def _scalar(value: object) -> Scalar:
    """A cell's stored value as the engine computes with it."""
    if value is None:
        return EMPTY
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, (str, CellError)):
        return value
    if isinstance(value, dt.datetime):
        # A cell of type d, which Excel does not write but may read.
        return datetime_to_serial(value)
    return EMPTY


def address(key: CellKey) -> str:
    sheet, row, column = key
    return f"{sheet}!{CellRef(row, column).a1}"


class Engine:
    """A workbook's formulas and values, for calculating."""

    def __init__(
        self,
        workbook: Workbook,
        *,
        today: dt.date | None = None,
        now: dt.datetime | None = None,
        name: str | None = None,
    ) -> None:
        self._workbook = workbook
        self._epoch_1904 = workbook.epoch_1904
        path = workbook.path
        self._name = name or (path.name if path is not None else "Book1")
        self._folder = str(path.parent) + os.sep if path is not None else None
        self.now = now or dt.datetime.now()
        self.today = today or self.now.date()
        self._order = [sheet.name for sheet in workbook.sheets]
        self._keys = {name.casefold(): name for name in self._order}
        self._grid: dict[str, dict[int, dict[int, Content]]] = {name: {} for name in self._order}
        self._rows: dict[str, list[int]] = {}
        self._used: dict[str, tuple[int, int]] = {}
        self._formulas: dict[CellKey, _Formula] = {}
        self._sheets: dict[str, Worksheet] = {sheet.name: sheet for sheet in workbook.sheets}
        self._results: dict[CellKey, Value] = {}
        self._failed: dict[CellKey, str] = {}
        self._tainted: set[CellKey] = set()
        self._circular: set[CellKey] = set()
        self._names: dict[tuple[str, str | None], Node | None] = {}
        self._tables: dict[str, TableShape] = {}
        #: The rows an active filter decides, below its header, per sheet.
        self._filters: dict[str, tuple[int, int]] = {}
        self._subtotals: dict[CellKey, bool] = {}
        self._pivots: dict[str, list[PivotReport]] = {}
        #: Each linked workbook's sheets, as sheets here are named for it,
        #: and its names, by the number formulas give the link.
        self._links: dict[str, list[str]] = {}
        self._link_names: dict[str, dict[str, str]] = {}
        self._reading_tainted = False
        for sheet in workbook.sheets:
            self._load(sheet)
            self._load_filter(sheet)
        self._load_tables()
        self._load_links()

    @property
    def epoch_1904(self) -> bool:
        return self._epoch_1904

    @property
    def name(self) -> str:
        return self._name

    @property
    def folder(self) -> str | None:
        return self._folder

    # ------------------------------------------------------------------
    # Loading
    # ------------------------------------------------------------------

    def _load(self, sheet: Worksheet) -> None:
        grid = self._grid[sheet.name]
        strings = self._workbook.shared_strings
        arrays: list[_Formula] = []
        for reference, element in sheet.cell_elements():
            try:
                stored = read_value(element, shared_strings=strings, styles=None, epoch_1904=self._epoch_1904)
            except (ValueError, IndexError):
                stored = None
            formula = element.child("f")
            if formula is None:
                value = _scalar(stored)
                if not isinstance(value, Empty):
                    grid.setdefault(reference.row, {})[reference.column] = value
                continue
            key = (sheet.name, reference.row, reference.column)
            kind = formula.get("t") or "normal"
            entry = _Formula(key, "", element, _scalar(stored))
            if kind == "dataTable":
                entry.table = _data_table(formula, sheet.name)
                if entry.table is None:
                    entry.problem = "a data table whose input cell is gone"
            else:
                text = sheet.formula_of(element, reference)
                if text is None:
                    entry.problem = "a formula with no text"
                else:
                    entry.text = text
                    try:
                        entry.node = parse(text)
                    except FormulaSyntaxError as error:
                        entry.problem = f"a formula that does not parse: {error}"
            if kind in ("array", "dataTable"):
                block = _block(formula.get("ref"), reference)
                entry.area = Area(sheet.name, block.top, block.left, block.bottom, block.right)
                arrays.append(entry)
            grid.setdefault(reference.row, {})[reference.column] = entry
            self._formulas[key] = entry
        for entry in arrays:
            assert entry.area is not None
            for row, column in entry.area.cells():
                if (row, column) != (entry.key[1], entry.key[2]):
                    grid.setdefault(row, {})[column] = _Part(entry.key, row - entry.area.top, column - entry.area.left)
        self._rows[sheet.name] = sorted(grid)
        last_row = self._rows[sheet.name][-1] if grid else 0
        last_column = max((max(cells) for cells in grid.values() if cells), default=0)
        self._used[sheet.name] = (last_row, last_column)

    def _load_filter(self, sheet: Worksheet) -> None:
        found = sheet.auto_filter
        if found is None or not found.filtering or not found.ref:
            return
        try:
            block = RangeRef.parse(found.ref)
        except ValueError:
            return
        self._filters[sheet.name] = (block.top + 1, block.bottom)

    def _load_links(self) -> None:
        """Each linked workbook's cells as its link caches them, as sheets
        named ``[1]Sheet1`` beside the workbook's own, and its names: what
        Excel calculates with while that workbook is closed."""
        for book in read_external_books(self._workbook.package, self._workbook.workbook_part):
            number = str(book.number)
            order: list[str] = []
            for sheet, cells in zip(book.sheets, book.cells, strict=True):
                key = f"[{number}]{sheet}"
                grid: dict[int, dict[int, Content]] = {}
                for (row, column), value in cells.items():
                    grid.setdefault(row, {})[column] = value
                self._grid[key] = grid
                self._rows[key] = sorted(grid)
                self._used[key] = (max(grid, default=0), max((max(line) for line in grid.values()), default=0))
                order.append(key)
            self._links[number] = order
            self._link_names[number] = {name.casefold(): refers for name, refers in book.names.items()}

    def _load_tables(self) -> None:
        for sheet in self._workbook.sheets:
            for table in sheet.tables:
                ref = table.ref
                self._tables[table.display_name.casefold()] = TableShape(
                    table.display_name,
                    sheet.name,
                    ref.top,
                    ref.left,
                    ref.bottom,
                    ref.right,
                    table.header_row_count,
                    table.totals_row_count,
                    tuple(table.column_names),
                )

    # ------------------------------------------------------------------
    # The book, as the evaluator reads it
    # ------------------------------------------------------------------

    def sheet_key(self, name: str) -> str | None:
        return self._keys.get(name.casefold())

    def sheet_order(self) -> list[str]:
        return list(self._order)

    def used(self, sheet: str) -> tuple[int, int]:
        return self._used.get(sheet, (0, 0))

    def cell(self, sheet: str, row: int, column: int) -> Scalar:
        content = self._grid[sheet].get(row, {}).get(column)
        if content is None:
            return EMPTY
        value, pending = self._resolve(content)
        if pending is not None:
            raise PendingCellsError([pending])
        return value

    def cells(self, area: Area) -> Iterator[tuple[int, int, Scalar]]:
        grid = self._grid[area.sheet]
        rows = self._rows[area.sheet]
        found: list[tuple[int, int, Scalar]] = []
        pending: list[CellKey] = []
        start = bisect.bisect_left(rows, area.top)
        stop = bisect.bisect_right(rows, area.bottom)
        narrow = area.width <= 64
        for row in rows[start:stop]:
            line = grid[row]
            if narrow:
                columns = [column for column in range(area.left, area.right + 1) if column in line]
            else:
                columns = sorted(column for column in line if area.left <= column <= area.right)
            for column in columns:
                value, waiting = self._resolve(line[column])
                if waiting is not None:
                    pending.append(waiting)
                elif not isinstance(value, Empty):
                    found.append((row, column, value))
        if pending:
            raise PendingCellsError(pending)
        return iter(found)

    def _resolve(self, content: Content) -> tuple[Scalar, CellKey | None]:
        """A cell's value, or the formula cell to calculate first."""
        if isinstance(content, _Override):
            return content.value, None
        if isinstance(content, _Formula):
            key = content.key
            if key not in self._results:
                return EMPTY, key
            self._note(key)
            result = self._results[key]
            if content.area is not None:
                assert isinstance(result, Array)
                return result.rows[0][0], None
            assert not isinstance(result, (Array, Reference))
            return result, None
        if isinstance(content, _Part):
            if content.master not in self._results:
                return EMPTY, content.master
            self._note(content.master)
            result = self._results[content.master]
            assert isinstance(result, Array)
            return result.rows[content.row][content.column], None
        return content, None

    def _note(self, key: CellKey) -> None:
        if key in self._failed or key in self._tainted:
            self._reading_tainted = True

    def defined_name(self, name: str, sheet: str | None) -> Node | None:
        cache_key = (name.casefold(), sheet)
        if cache_key not in self._names:
            found = None
            for entry in self._workbook.defined_names:
                if entry.name.casefold() == name.casefold() and entry.scope == sheet:
                    try:
                        found = parse(entry.refers_to)
                    except FormulaSyntaxError:
                        found = None
                    break
            self._names[cache_key] = found
        return self._names[cache_key]

    def table(self, name: str) -> TableShape | None:
        return self._tables.get(name.casefold())

    def table_at(self, sheet: str, row: int, column: int) -> TableShape | None:
        for table in self._tables.values():
            if table.sheet == sheet and table.top <= row <= table.bottom and table.left <= column <= table.right:
                return table
        return None

    def formula_text(self, sheet: str, row: int, column: int) -> str | None:
        content = self._formula_at(sheet, row, column)
        return content.text or None if content is not None else None

    def _formula_at(self, sheet: str, row: int, column: int) -> _Formula | None:
        content = self._grid[sheet].get(row, {}).get(column)
        if isinstance(content, _Part):
            return self._formulas.get(content.master)
        return content if isinstance(content, _Formula) else None

    def row_hidden(self, sheet: str, row: int) -> bool:
        """Whether a row is hidden; no row of a linked workbook is."""
        found = self._sheets.get(sheet)
        return found is not None and found.row_hidden(row)

    def row_filtered(self, sheet: str, row: int) -> bool:
        """A hidden row inside an active filter's range. The file records
        only that a row is hidden, not what hid it, so there every hidden
        row counts as one the filter left out."""
        span = self._filters.get(sheet)
        return span is not None and span[0] <= row <= span[1] and self.row_hidden(sheet, row)

    def external_sheets(self, book: str, first: str, last: str | None) -> list[str] | None:
        order = self._links.get(book)
        if order is None:
            return None
        names = [key[len(book) + 2 :].casefold() for key in order]
        try:
            start = names.index(first.casefold())
            end = start if last is None else names.index(last.casefold())
        except ValueError:
            return []
        low, high = sorted((start, end))
        return order[low : high + 1]

    def external_name(self, book: str, name: str) -> Node | None:
        refers = self._link_names.get(book, {}).get(name.casefold())
        if refers is None:
            return None
        try:
            return _in_book(parse(refers), book)
        except FormulaSyntaxError:
            return None

    def pivot_reports(self, sheet: str) -> list[PivotReport]:
        """The sheet's pivot tables, read once. Reading one Excel wrote in a
        way this does not follow is left to the formulas that need it."""
        found = self._pivots.get(sheet)
        if found is None and sheet not in self._sheets:
            # A linked workbook's pivot tables are not in its link.
            return []
        if found is None:
            styles = self._workbook.styles
            custom = {} if styles is None else styles.custom_formats

            def formats(format_id: int) -> str:
                return custom.get(format_id) or BUILTIN_DISPLAY_CODES.get(format_id, "General")

            package = self._workbook.package
            reports = (
                read_pivot_report(package, table.part_name, formats, epoch_1904=self._epoch_1904)
                for table in self._sheets[sheet].pivot_tables
            )
            found = self._pivots[sheet] = [report for report in reports if report is not None]
        return found

    def spill(self, sheet: str, row: int, column: int) -> Area | None:
        """A dynamic-array formula's block: an array formula whose cell
        carries cell metadata, ``cm``, which is how Excel marks a spill as
        against an array formula entered with Ctrl+Shift+Enter."""
        formula = self._grid[sheet].get(row, {}).get(column)
        if not isinstance(formula, _Formula) or formula.area is None or formula.element.get("cm") is None:
            return None
        return formula.area

    def subtotal(self, sheet: str, row: int, column: int) -> bool:
        formula = self._formula_at(sheet, row, column)
        if formula is None or formula.node is None:
            return False
        found = self._subtotals.get(formula.key)
        if found is None:
            found = any(isinstance(node, Call) and node.function in _SUBTOTALS for node in walk(formula.node))
            self._subtotals[formula.key] = found
        return found

    # ------------------------------------------------------------------
    # Calculating
    # ------------------------------------------------------------------

    @property
    def formula_keys(self) -> list[CellKey]:
        return list(self._formulas)

    def calculate(self, keys: Iterable[CellKey] | None = None) -> None:
        """Calculate formula cells, and whatever they read, until done."""
        for key in self._formulas if keys is None else keys:
            if key not in self._results:
                self._run(key)

    def _run(self, first: CellKey) -> None:
        stack = [first]
        waiting: set[CellKey] = set()
        while stack:
            key = stack[-1]
            if key in self._results:
                stack.pop()
                waiting.discard(key)
                continue
            self._reading_tainted = False
            try:
                value = self._evaluate(self._formulas[key])
            except PendingCellsError as need:
                fresh = list(dict.fromkeys(cell for cell in need.cells if cell not in self._results))
                looped = [cell for cell in fresh if cell in waiting or cell == key]
                if looped:
                    # Every cell on the stack from the first one asked for
                    # again up to this one is on the loop.
                    start = min(stack.index(cell) for cell in looped)
                    for cell in stack[start:]:
                        if cell not in self._results:
                            self._give_up(cell, "a circular reference")
                            self._circular.add(cell)
                    continue
                waiting.add(key)
                stack.extend(fresh)
                continue
            except UnsupportedFormulaError as reason:
                self._give_up(key, str(reason))
                stack.pop()
                waiting.discard(key)
                continue
            self._results[key] = value
            if self._reading_tainted:
                self._tainted.add(key)
            stack.pop()
            waiting.discard(key)

    def _give_up(self, key: CellKey, reason: str) -> None:
        """Keep a formula's cached value, as the value others read."""
        formula = self._formulas[key]
        self._failed[key] = reason
        if formula.area is None:
            self._results[key] = formula.cached
            return
        self._results[key] = self._cached_block(formula)

    def _cached_block(self, formula: _Formula) -> Array:
        assert formula.area is not None
        sheet = self._sheets[formula.key[0]]
        strings = self._workbook.shared_strings
        rows: list[list[Scalar]] = []
        for row in range(formula.area.top, formula.area.bottom + 1):
            line: list[Scalar] = []
            for column in range(formula.area.left, formula.area.right + 1):
                if (row, column) == formula.key[1:]:
                    line.append(formula.cached)
                    continue
                element = sheet.cell_element(CellRef(row, column)) if sheet.has_cell(CellRef(row, column)) else None
                stored = (
                    None
                    if element is None
                    else read_value(element, shared_strings=strings, styles=None, epoch_1904=self._epoch_1904)
                )
                line.append(_scalar(stored))
            rows.append(line)
        return Array(rows)

    def _evaluate(self, formula: _Formula) -> Value:
        if formula.table is not None:
            return self._tried(formula, formula.table)
        if formula.node is None:
            raise UnsupportedFormulaError(formula.problem or "a formula this engine cannot read")
        sheet, row, column = formula.key
        context = Context(self, sheet, row, column, array=formula.area is not None, today=self.today, now=self.now)
        value = context.formula(formula.node)
        if formula.area is None:
            return _cell_value(context.first(value))
        grid = context.array_of(value)
        area = formula.area
        return Array(
            [[_cell_value(grid.at(row, column)) for column in range(area.width)] for row in range(area.height)]
        )

    def _tried(self, formula: _Formula, table: _DataTable) -> Array:
        """A data table's block: its formula, or each of its formulas,
        calculated with each value it tries in its input cells, the values
        read as they are in the row above the block or the column left of
        it."""
        area = formula.area
        assert area is not None
        sheet = area.sheet
        what_if = _WhatIf(self, [table.first] if table.second is None else [table.first, table.second])
        rows: list[list[Scalar]] = []
        for row in range(area.top, area.bottom + 1):
            line: list[Scalar] = []
            for column in range(area.left, area.right + 1):
                if table.two_way:
                    target = (sheet, area.top - 1, area.left - 1)
                    values = [self.cell(sheet, area.top - 1, column), self.cell(sheet, row, area.left - 1)]
                elif table.along_row:
                    target = (sheet, row, area.left - 1)
                    values = [self.cell(sheet, area.top - 1, column)]
                else:
                    target = (sheet, area.top - 1, column)
                    values = [self.cell(sheet, row, area.left - 1)]
                line.append(what_if.value(target, values))
            rows.append(line)
        if what_if.tainted:
            self._reading_tainted = True
        return Array(rows)

    def evaluate(self, formula: str, sheet: str, row: int, column: int) -> Scalar:
        """What ``formula`` gives in a cell, calculating what it reads."""
        node = parse(formula)
        while True:
            context = Context(self, sheet, row, column, today=self.today, now=self.now)
            try:
                return _cell_value(context.first(context.formula(node)))
            except PendingCellsError as need:
                self.calculate(need.cells)

    def values(self, formula: str, sheet: str, row: int, column: int) -> list[Scalar]:
        """Every value ``formula`` gives in a cell, calculating what it
        reads: a range's cells that hold something, row by row, an array's
        items, or the one value."""
        node = parse(formula)
        while True:
            context = Context(self, sheet, row, column, today=self.today, now=self.now)
            try:
                value = context.formula(node)
                if isinstance(value, Reference):
                    return [found for area in value.areas for _, _, found in self.cells(area)]
                if isinstance(value, Array):
                    return [item for line in value.rows for item in line]
                return [value]
            except PendingCellsError as need:
                self.calculate(need.cells)

    # ------------------------------------------------------------------
    # Results
    # ------------------------------------------------------------------

    def report(self) -> Calculation:
        report = Calculation()
        for key in self._formulas:
            if key in self._circular:
                report.circular.append(address(key))
            elif key in self._failed:
                report.unsupported[address(key)] = self._failed[key]
            elif key in self._tainted:
                report.dependent.append(address(key))
            elif key in self._results:
                report.calculated += 1
        return report

    def write(self) -> None:
        """Put every result calculated here into its cell."""
        for key, formula in self._formulas.items():
            if key in self._failed or key in self._tainted or key not in self._results:
                continue
            result = self._results[key]
            if formula.area is None:
                write_cached(formula.element, _stored(result))
                continue
            assert isinstance(result, Array)
            sheet = self._sheets[key[0]]
            for row_offset, line in enumerate(result.rows):
                for column_offset, item in enumerate(line):
                    reference = CellRef(formula.area.top + row_offset, formula.area.left + column_offset)
                    element = formula.element if (reference.row, reference.column) == key[1:] else sheet.cell_element(reference)
                    write_cached(element, _stored(item))

    def result(self, key: CellKey) -> Value | None:
        return self._results.get(key)


class _WhatIf(Engine):
    """The workbook with values put in some of its cells, as a data table
    tries them. What reads those cells, however indirectly, is calculated
    again for each set of values; a formula that read none of it once
    never will, so it keeps what it came to the first time.

    It shares the workbook as the engine read it, and keeps results of its
    own; the rows holding its input cells are copied, so the values it
    puts there reach no one else.
    """

    def __init__(self, base: Engine, inputs: list[CellKey]) -> None:
        self.__dict__.update(base.__dict__)
        self._inputs = inputs
        self._grid = dict(base._grid)
        self._rows = dict(base._rows)
        self._used = dict(base._used)
        for sheet, row, column in inputs:
            rows = self._grid[sheet] = dict(self._grid[sheet])
            rows[row] = dict(rows.get(row, {}))
            if row not in self._rows[sheet]:
                self._rows[sheet] = sorted({*self._rows[sheet], row})
            last_row, last_column = self._used.get(sheet, (0, 0))
            self._used[sheet] = (max(last_row, row), max(last_column, column))
        self._results = {}
        self._failed = {}
        self._tainted = set()
        self._circular = set()
        #: Formulas that read nothing the table tries.
        self._independent: set[CellKey] = set()
        self._reading_tried = False
        #: Whether a value came from a cell that keeps its cached value.
        self.tainted = False

    def value(self, target: CellKey, values: list[Scalar]) -> Scalar:
        """What a cell comes to with ``values`` in the input cells."""
        for (sheet, row, column), value in zip(self._inputs, values, strict=True):
            self._grid[sheet][row][column] = _Override(value)
        keep = self._independent
        self._results = {key: found for key, found in self._results.items() if key in keep}
        self._failed = {key: reason for key, reason in self._failed.items() if key in keep}
        self._tainted &= keep
        self._circular &= keep
        while True:
            try:
                found = _cell_value(self.cell(*target))
                break
            except PendingCellsError as need:
                self.calculate(need.cells)
        if target in self._failed or target in self._tainted:
            self.tainted = True
        return found

    def _resolve(self, content: Content) -> tuple[Scalar, CellKey | None]:
        if isinstance(content, _Override):
            self._reading_tried = True
        found = super()._resolve(content)
        key = content.key if isinstance(content, _Formula) else content.master if isinstance(content, _Part) else None
        if key is not None and found[1] is None and key not in self._independent:
            self._reading_tried = True
        return found

    def _evaluate(self, formula: _Formula) -> Value:
        if formula.table is not None:
            raise UnsupportedFormulaError("a data table reading cells another data table tries")
        self._reading_tried = False
        value = super()._evaluate(formula)
        if not self._reading_tried:
            self._independent.add(formula.key)
        return value


#: Every kind of node a formula's tree is made of.
_NODES = typing.get_args(Node)


def _in_book(node: Node, book: str) -> Node:
    """A formula of a linked workbook, as one defining a name there, with
    each reference and name in it made to name that workbook."""
    changes: dict[str, object] = {}
    for item in dataclasses.fields(node):
        value = getattr(node, item.name)
        if isinstance(value, Prefix):
            if value.book is None:
                changes[item.name] = Prefix(value.sheet, value.last_sheet, book)
        elif isinstance(node, NameReference) and item.name == "prefix" and value is None:
            changes[item.name] = Prefix(book=book)
        elif isinstance(value, _NODES):
            changes[item.name] = _in_book(value, book)
        elif isinstance(value, tuple):
            parts = typing.cast("tuple[object, ...]", value)
            if parts and all(isinstance(part, _NODES) for part in parts):
                changes[item.name] = tuple(_in_book(typing.cast("Node", part), book) for part in parts)
    return dataclasses.replace(node, **changes) if changes else node


def _data_table(formula: Element, sheet: str) -> _DataTable | None:
    """A data table's input cells, from its ``<f t="dataTable">``, or
    ``None`` when one was deleted, which the file marks ``del1`` or
    ``del2``."""
    if any(formula.get(flag) in ("1", "true") for flag in ("del1", "del2")):
        return None
    two_way = formula.get("dt2D") in ("1", "true")
    try:
        first = CellRef.parse(formula.get("r1") or "")
        second = CellRef.parse(formula.get("r2") or "") if two_way else None
    except ValueError:
        return None
    return _DataTable(
        two_way,
        formula.get("dtr") in ("1", "true"),
        (sheet, first.row, first.column),
        None if second is None else (sheet, second.row, second.column),
    )


def _block(ref: str | None, master: CellRef) -> RangeRef:
    if ref:
        try:
            return RangeRef.parse(ref)
        except ValueError:
            pass
    return RangeRef(master, master)


def _cell_value(value: Scalar) -> Scalar:
    """A result as a cell holds it: a blank is 0."""
    if isinstance(value, Empty):
        return 0.0
    return value


def _stored(value: Value) -> str | float | bool | CellError | None:
    if isinstance(value, (Array, Reference)):  # pragma: no cover - results are scalars by now
        return None
    if isinstance(value, Empty):
        return 0.0
    return value


__all__ = ["Calculation", "Engine", "address"]
