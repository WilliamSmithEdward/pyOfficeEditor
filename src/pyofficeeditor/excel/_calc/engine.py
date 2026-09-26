"""Calculating a workbook: every formula, in the order its inputs allow,
with the results written back into the cells.

The workbook is read once into a grid per sheet. A formula cell starts
uncalculated; evaluating one that reads an uncalculated cell stops with
:class:`~.evaluator.PendingCellsError`, and the cells it named are calculated
first, on an explicit stack, before it is tried again. So a column of ten
thousand running totals, each reading the one above, calculates without
ten thousand nested Python calls.

A formula that cannot be calculated here, because it calls a function
this engine does not have, reads another workbook, or will not parse,
keeps the value the file cached for it, and so does every formula that
reads it: its value is only as good as that cache. :class:`Calculation`
lists both. A circular reference keeps its cached values too, as Excel,
with iteration off, leaves them.
"""

from __future__ import annotations

import bisect
import datetime as dt
import os
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
from pyofficeeditor.excel._calc.nodes import Call, Node, walk
from pyofficeeditor.excel._calc.parser import parse
from pyofficeeditor.excel._calc.values import EMPTY, Area, Array, Empty, Reference, Scalar, Value
from pyofficeeditor.excel._numfmt import BUILTIN_DISPLAY_CODES
from pyofficeeditor.excel._pivots import PivotReport, read_pivot_report
from pyofficeeditor.excel._reference import CellRef, RangeRef
from pyofficeeditor.excel._values import CellError, datetime_to_serial, read_value, write_cached

if TYPE_CHECKING:
    from pyofficeeditor.excel.workbook import Workbook
    from pyofficeeditor.excel.worksheet import Worksheet


@dataclass
class _Formula:
    key: CellKey
    text: str
    element: Element
    cached: Scalar
    node: Node | None = None
    #: What stops it being calculated, when something does.
    problem: str | None = None
    #: The block an array formula fills.
    area: Area | None = None


@dataclass(frozen=True)
class _Part:
    """A cell of an array formula's block other than the one holding it."""

    master: CellKey
    row: int
    column: int


Content = Scalar | _Formula | _Part

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
        self._reading_tainted = False
        for sheet in workbook.sheets:
            self._load(sheet)
            self._load_filter(sheet)
        self._load_tables()

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
                entry.problem = "a data table"
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
            if kind == "array":
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
        return self._sheets[sheet].row_hidden(row)

    def row_filtered(self, sheet: str, row: int) -> bool:
        """A hidden row inside an active filter's range. The file records
        only that a row is hidden, not what hid it, so there every hidden
        row counts as one the filter left out."""
        span = self._filters.get(sheet)
        return span is not None and span[0] <= row <= span[1] and self._sheets[sheet].row_hidden(row)

    def pivot_reports(self, sheet: str) -> list[PivotReport]:
        """The sheet's pivot tables, read once. Reading one Excel wrote in a
        way this does not follow is left to the formulas that need it."""
        found = self._pivots.get(sheet)
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

    def evaluate(self, formula: str, sheet: str, row: int, column: int) -> Scalar:
        """What ``formula`` gives in a cell, calculating what it reads."""
        node = parse(formula)
        while True:
            context = Context(self, sheet, row, column, today=self.today, now=self.now)
            try:
                return _cell_value(context.first(context.formula(node)))
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
