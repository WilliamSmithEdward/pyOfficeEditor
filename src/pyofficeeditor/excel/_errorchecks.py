"""Excel's error checking: what puts a green triangle on a cell.

Excel checks each cell it shows against ten rules and marks a cell a rule
catches. Each rule here was measured against ``Range.Errors`` in Excel,
on cells built for the purpose and, for the two that read text, on some
forty thousand strings, and each is named as a worksheet's
``<ignoredErrors>`` names it when a cell's error is ignored:

``evalError``
    A formula whose value is an error, but for ``#N/A`` from a formula
    that calls NA(), which is how a formula says it means it.
``twoDigitTextYear``
    Text that reads as a date with a year of one or two digits, as a
    constant or as a string in a formula. The text is read more loosely
    than DATEVALUE reads it: ``99 Jan`` and ``1 1/0`` are dates to it.
``numberStoredAsText``
    Text that is a number as VALUE reads one, and not a date or a time.
``formula``
    A formula unlike the two on either side of it, above and below or left
    and right, which are alike: formulas compared as R1C1 text, so a space
    or a pair of parentheses makes one unlike another.
``formulaRange``
    A formula reading a range one cell wide or high, of numbers and empty
    cells only, with a number next to either end of it that it leaves
    out. A formula calling SUMIF or LOOKUP is never caught.
``unlockedFormula``
    A formula in a cell that is not locked.
``emptyCellReference``
    A formula whose references, names and table references included,
    reach an empty cell. Excel leaves this rule off unless asked.
``listDataValidation``
    A value in a table that its data validation refuses.
``calculatedColumn``
    A cell of a table's calculated column holding another formula or a
    value.
``misleadingFormat``
    A formula that is one cell's reference, perhaps signed, showing a
    date in a format that is not a date's, or a number in one that is.

A rule reads what the file holds: a formula's value is the one Excel
cached for it, so a workbook changed since should be calculated first.
"""

from __future__ import annotations

import bisect
import datetime as dt
from collections.abc import Callable, Iterable, Iterator
from dataclasses import dataclass
from typing import TYPE_CHECKING, Literal

from pyofficeeditor.excel._calc.dates import days_in_month
from pyofficeeditor.excel._calc.evaluator import Context
from pyofficeeditor.excel._calc.lexer import FormulaSyntaxError
from pyofficeeditor.excel._calc.nodes import (
    AreaReference,
    AxisReference,
    Binary,
    Call,
    CellReference,
    Invoke,
    NameReference,
    Node,
    Paren,
    Postfix,
    Prefix,
    StructuredReference,
    Text,
    Unary,
    function_key,
    walk,
)
from pyofficeeditor.excel._calc.parser import parse
from pyofficeeditor.excel._calc.values import Area, Reference, Scalar, compare, plain_number, scalar_text
from pyofficeeditor.excel._formulas import translate_formula
from pyofficeeditor.excel._numfmt import parse as parse_format
from pyofficeeditor.excel._reference import MAX_COLUMN, MAX_ROW, AxisRef, CellRef, RangeRef
from pyofficeeditor.excel._tokens import TokenKind, tokenize
from pyofficeeditor.excel._validation import DataValidation, ValidationOperator
from pyofficeeditor.excel._values import CellError, CellValue, read_value
from pyofficeeditor.exceptions import UnsupportedFormulaError

if TYPE_CHECKING:
    from pyofficeeditor.excel._calc.engine import Engine
    from pyofficeeditor.excel._tables import Table
    from pyofficeeditor.excel.worksheet import Worksheet

ErrorRule = Literal[
    "evalError",
    "twoDigitTextYear",
    "numberStoredAsText",
    "formula",
    "formulaRange",
    "unlockedFormula",
    "emptyCellReference",
    "listDataValidation",
    "calculatedColumn",
    "misleadingFormat",
]

#: Every rule, in the order Excel numbers them.
ERROR_RULES: tuple[ErrorRule, ...] = (
    "evalError",
    "twoDigitTextYear",
    "numberStoredAsText",
    "formula",
    "formulaRange",
    "unlockedFormula",
    "emptyCellReference",
    "listDataValidation",
    "calculatedColumn",
    "misleadingFormat",
)

#: The rules Excel checks unless told otherwise: all but references to
#: empty cells.
DEFAULT_ERROR_RULES: frozenset[ErrorRule] = frozenset(ERROR_RULES) - {"emptyCellReference"}


@dataclass(frozen=True)
class ErrorCheck:
    """A rule that catches a cell, as Excel's green triangle marks it."""

    reference: CellRef
    rule: ErrorRule
    #: Whether the file records the error as ignored, which hides the
    #: triangle.
    ignored: bool = False


@dataclass(frozen=True)
class IgnoredError:
    """Cells whose errors under some rules the file records as ignored:
    one ``<ignoredError>`` of a sheet's ``<ignoredErrors>``."""

    ranges: tuple[RangeRef, ...]
    rules: frozenset[ErrorRule]


# ----------------------------------------------------------------------
# Text that reads as a date with a two-digit year
# ----------------------------------------------------------------------

_MONTH_NAMES = (
    "january", "february", "march", "april", "may", "june",
    "july", "august", "september", "october", "november", "december",
)  # fmt: skip


@dataclass(frozen=True)
class _DatePart:
    """A run of digits, or a month's name as its number."""

    digits: str = ""
    month: int | None = None

    @property
    def value(self) -> int:
        return int(self.digits)

    @property
    def two_digit(self) -> bool:
        return self.month is None and len(self.digits) <= 2


def _month_named(word: str) -> int | None:
    """A month named by three letters or more of its English name."""
    lowered = word.lower()
    if len(lowered) < 3:
        return None
    for number, name in enumerate(_MONTH_NAMES, start=1):
        if name.startswith(lowered):
            return number
    return None


def _date_parts(text: str) -> list[_DatePart] | None:
    """Text split as the check splits it: runs of digits or of letters,
    after any leading spaces, apart by spaces, by spaces and then ``/`` or
    ``-``, or by spaces, a comma and one space. Anything else, a space at
    the end included, is not a date."""
    parts: list[_DatePart] = []
    index = 0
    while index < len(text) and text[index] == " ":
        index += 1
    while True:
        start = index
        if index < len(text) and text[index].isdigit():
            while index < len(text) and text[index].isdigit():
                index += 1
            parts.append(_DatePart(text[start:index]))
        elif index < len(text) and text[index].isalpha():
            while index < len(text) and text[index].isalpha():
                index += 1
            month = _month_named(text[start:index])
            if month is None:
                return None
            parts.append(_DatePart(month=month))
        else:
            return None
        if index == len(text):
            return parts
        after = index
        while after < len(text) and text[after] == " ":
            after += 1
        if after < len(text) and text[after] in "/-":
            after += 1
        elif after < len(text) and text[after] == ",":
            if after + 1 >= len(text) or text[after + 1] != " ":
                return None
            after += 2
        elif after == index:
            return None
        index = after


def _is_day(month: int, day: int, year: int) -> bool:
    return 1 <= month <= 12 and 1 <= day <= days_in_month(year, month)


def _year(part: _DatePart) -> int:
    """A one- or two-digit year as Excel reads it: 30 is 1930, 29 2029."""
    value = part.value
    return value + (2000 if value < 30 else 1900)


def _wide_day(part: _DatePart) -> int | None:
    """The day after a month, which may have three or four digits without
    a leading zero and is then kept to its low byte: measured, 1/287/99 is
    January 31, 1999."""
    if len(part.digits) <= 2:
        return part.value
    if len(part.digits) <= 4 and not part.digits.startswith("0"):
        return part.value & 0xFF
    return None


def two_digit_year(text: str, today: dt.date) -> bool:
    """Whether Excel's check calls ``text`` a date with a two-digit year.

    Measured on 43,000 strings, 5,663 random ones among them kept back
    while the rules were found. Two or three parts, numbers or month
    names: a month, a day and a year in that order; a day and a year beside
    a month's name, either first; or a month and a year, the number after
    the month being no day of it this year. The year has one or two digits
    and the day falls in its month, February 29 only in a leap year. Text
    that is a number is not a date to it.
    """
    if plain_number(text) is not None:
        return False
    parts = _date_parts(text)
    if parts is None or not 2 <= len(parts) <= 3:
        return False
    shape = "".join("n" if part.month is None else "m" for part in parts)
    if shape == "nn":
        month, second = parts
        if not month.two_digit or not 1 <= month.value <= 12:
            return False
        return not _is_day(month.value, second.value, today.year) and second.two_digit
    if shape in ("mn", "nm"):
        name, number = (parts[0], parts[1]) if shape == "mn" else (parts[1], parts[0])
        assert name.month is not None
        if _is_day(name.month, number.value, today.year):
            return False
        # Measured: beside a month's name, a year of 0 is not one.
        return number.two_digit and number.value >= 1
    if shape in ("nnn", "mnn"):
        first, second, third = parts
        month = first.month
        if month is None:
            if not first.two_digit or not 1 <= first.value <= 12:
                return False
            month = first.value
        day = _wide_day(second)
        return day is not None and third.two_digit and _is_day(month, day, _year(third))
    if shape in ("nmn", "nnm"):
        name = parts[shape.index("m")]
        assert name.month is not None
        first, second = [part for part in parts if part.month is None]
        if not first.two_digit or not second.two_digit:
            return False
        return any(
            year.value >= 1 and _is_day(name.month, day.value, _year(year))
            for day, year in ((first, second), (second, first))
        )
    if shape in ("mmn", "nmm"):
        # Measured: two names count only when the second is January.
        names = [part for part in parts if part.month is not None]
        return names[1].month == 1 and parts[shape.index("n")].two_digit
    return False


# ----------------------------------------------------------------------
# Formulas compared as R1C1 text
# ----------------------------------------------------------------------


def r1c1_key(formula: str, row: int, column: int) -> str:
    """A formula as Excel compares formulas for consistency: its references
    in R1C1 form relative to the cell it is written in, everything else as
    written, spaces and parentheses included, but for case."""
    pieces: list[str] = []
    for token in tokenize(formula):
        if token.kind is TokenKind.REFERENCE and isinstance(token.value, CellRef):
            ref = token.value
            pieces.append(
                (f"R{ref.row}" if ref.absolute_row else f"R[{ref.row - row}]")
                + (f"C{ref.column}" if ref.absolute_column else f"C[{ref.column - column}]")
            )
        elif token.kind is TokenKind.AXIS and isinstance(token.value, AxisRef):
            axis = token.value
            letter, origin = ("R", row) if axis.is_row else ("C", column)
            low = f"{letter}{axis.low}" if axis.absolute_low else f"{letter}[{axis.low - origin}]"
            high = f"{letter}{axis.high}" if axis.absolute_high else f"{letter}[{axis.high - origin}]"
            pieces.append(f"{low}:{high}")
        elif token.kind is TokenKind.TEXT:
            pieces.append(token.raw.upper())
        else:
            pieces.append(token.raw)
    return "".join(pieces)


# ----------------------------------------------------------------------
# Formats
# ----------------------------------------------------------------------

FormatKind = Literal["date", "number", "neutral"]


def format_kind(code: str) -> FormatKind:
    """What a number under a format reads as, for the misleading-format
    rule: a date or a time, a number, or neither, as under ``@``, ``"x"``
    or ``;;;``, which show no digits. Measured on 26 formats each way."""
    parsed = parse_format(code)
    if parsed.is_date:
        return "date"
    numeric = [section for section in parsed.sections if section is not parsed.text]
    if numeric and (numeric[0].has_general or any(token.kind == "digit" for token in numeric[0].tokens)):
        return "number"
    return "neutral"


# ----------------------------------------------------------------------
# The checks
# ----------------------------------------------------------------------

#: Functions whose call anywhere in a formula spares it the omitted-cells
#: rule. Measured over some 110 functions: these two and no others.
_SPARES_RANGES = frozenset({"SUMIF", "LOOKUP"})


@dataclass
class _Cell:
    value: CellValue
    formula: str | None = None
    style: int | None = None
    #: The cell the formula is written in: this one, or for a cell of an
    #: array formula entered over several cells, the one holding it.
    anchor: tuple[int, int] | None = None


class _Sheet:
    """A sheet's cells as the checks read them: each cell's value, the one
    Excel cached for a formula, its formula and its format."""

    def __init__(self, sheet: Worksheet) -> None:
        self.name = sheet.name
        self.cells: dict[tuple[int, int], _Cell] = {}
        strings = sheet.workbook.shared_strings
        arrays: list[tuple[RangeRef, str, tuple[int, int]]] = []
        for reference, element in sheet.cell_elements():
            try:
                value = read_value(element, shared_strings=strings, styles=None)
            except (ValueError, KeyError, IndexError):
                value = None
            raw = element.get("s")
            style = int(raw) if raw is not None and raw.isdigit() else None
            formula = sheet.formula_of(element, reference)
            key = (reference.row, reference.column)
            self.cells[key] = _Cell(value, formula, style, key if formula is not None else None)
            written = element.child("f")
            span = None if written is None or written.get("t") != "array" else written.get("ref")
            # A dynamic array's cell carries cell metadata, and its spill
            # holds values only; an array entered over cells holds its
            # formula in each.
            if formula is not None and span and element.get("cm") is None:
                try:
                    arrays.append((RangeRef.parse(span).normalized, formula, key))
                except ValueError:
                    pass
        for block, formula, anchor in arrays:
            for row in range(block.top, block.bottom + 1):
                for column in range(block.left, block.right + 1):
                    member = self.cells.get((row, column))
                    if member is not None and member.formula is None:
                        member.formula = formula
                        member.anchor = anchor
        #: The columns of each row's cells that hold something.
        self._held: dict[int, list[int]] = {}
        #: The columns each block of 16 rows stores cells in, blank or not.
        self._spans: dict[int, tuple[int, int]] = {}
        self.last_row = self.last_column = 0
        for (row, column), cell in sorted(self.cells.items()):
            if cell.value is not None or cell.formula is not None:
                self._held.setdefault(row, []).append(column)
            block = (row - 1) // 16
            low, high = self._spans.get(block, (column, column))
            self._spans[block] = (min(low, column), max(high, column))
            self.last_row = max(self.last_row, row)
            self.last_column = max(self.last_column, column)
        self._stored_rows = sorted({row for row, _ in self.cells})
        self._held_rows = sorted(self._held)

    def held(self, top: int, left: int, bottom: int, right: int) -> Iterator[tuple[int, int]]:
        """The cells of a block that hold something, row by row."""
        start = bisect.bisect_left(self._held_rows, top)
        for index in range(start, len(self._held_rows)):
            row = self._held_rows[index]
            if row > bottom:
                return
            columns = self._held[row]
            for position in range(bisect.bisect_left(columns, left), bisect.bisect_right(columns, right)):
                yield row, columns[position]

    def reads_empty(self, top: int, left: int, bottom: int, right: int) -> bool:
        """Whether a formula reading a block reads an empty cell, as Excel's
        check finds one.

        Measured, it does not look at every cell. One cell alone, or a block
        with nothing in it, is judged whole. Otherwise an empty cell counts
        past the sheet's last used row or column, or in a row storing some
        cell, where the check sees only the columns that row's run of 16
        rows stores cells in: a block's empty cells in rows storing nothing
        go unseen.
        """
        if (top, left) == (bottom, right):
            cell = self.cells.get((top, left))
            return cell is None or (cell.value is None and cell.formula is None)
        if next(self.held(top, left, bottom, right), None) is None:
            return True
        if bottom > self.last_row or right > self.last_column:
            return True
        start = bisect.bisect_left(self._stored_rows, top)
        for index in range(start, len(self._stored_rows)):
            row = self._stored_rows[index]
            if row > bottom:
                break
            low, high = self._spans[(row - 1) // 16]
            first, last = max(left, low), min(right, high)
            if first > last:
                continue
            columns = self._held.get(row, [])
            held = bisect.bisect_right(columns, last) - bisect.bisect_left(columns, first)
            if held < last - first + 1:
                return True
        return False

    def number_constant(self, row: int, column: int) -> bool:
        cell = self.cells.get((row, column))
        return (
            cell is not None
            and cell.formula is None
            and isinstance(cell.value, (int, float))
            and not isinstance(cell.value, bool)
        )


class _Checker:
    def __init__(self, sheet: Worksheet, rules: frozenset[ErrorRule], today: dt.date) -> None:
        self.rules = rules
        self.today = today
        self.workbook = sheet.workbook
        self.styles = self.workbook.styles
        self.here = _Sheet(sheet)
        self._others: dict[str, _Sheet | None] = {}
        self._engine: Engine | None = None
        self._trees: dict[tuple[int, int], Node | None] = {}
        self._keys: dict[tuple[int, int], str] = {}
        self._tables: list[tuple[RangeRef, Table]] = [
            (data, table) for table in sheet.tables if (data := table.data_range) is not None
        ]
        self._validations = [found for found in sheet.data_validations if found.ranges] if self._tables else []

    def run(self) -> Iterator[tuple[tuple[int, int], ErrorRule]]:
        keys = set(self.here.cells)
        if "listDataValidation" in self.rules:
            keys.update(self._blanks_refused())
        for key in sorted(keys):
            cell = self.here.cells.get(key) or _Cell(None)
            for rule in ERROR_RULES:
                if rule in self.rules and self._catches(rule, key, cell):
                    yield key, rule

    def _blanks_refused(self) -> Iterator[tuple[int, int]]:
        """Cells of a table the file stores nothing for, under a validation
        that does not ignore a blank: measured, Excel flags them too."""
        for validation in self._validations:
            if validation.allow_blank or validation.kind == "none":
                continue
            for span in validation.ranges:
                for data, _ in self._tables:
                    top, bottom = max(span.top, data.top), min(span.bottom, data.bottom)
                    left, right = max(span.left, data.left), min(span.right, data.right)
                    for row in range(top, bottom + 1):
                        for column in range(left, right + 1):
                            if (row, column) not in self.here.cells:
                                yield row, column

    # -- the rules ------------------------------------------------------

    def _catches(self, rule: ErrorRule, key: tuple[int, int], cell: _Cell) -> bool:
        if rule == "calculatedColumn":
            return self._calculated_column(key, cell)
        if rule == "listDataValidation":
            return self._refused(key, cell)
        if cell.formula is None:
            if not isinstance(cell.value, str):
                return False
            if rule == "twoDigitTextYear":
                return two_digit_year(cell.value, self.today)
            return rule == "numberStoredAsText" and plain_number(cell.value) is not None
        if rule == "evalError":
            return self._eval_error(key, cell)
        if rule == "twoDigitTextYear":
            tree = self._tree(key)
            return tree is not None and any(
                isinstance(node, Text) and two_digit_year(node.value, self.today) for node in walk(tree)
            )
        if rule == "formula":
            return self._inconsistent(key)
        if rule == "formulaRange":
            return self._omitted(key)
        if rule == "unlockedFormula":
            return self.styles is not None and not self.styles.cell_format(cell.style).protection.locked
        if rule == "emptyCellReference":
            return self._empty_reference(key)
        if rule == "misleadingFormat":
            return self._misleading(key, cell)
        return False

    def _eval_error(self, key: tuple[int, int], cell: _Cell) -> bool:
        if not isinstance(cell.value, CellError):
            return False
        if cell.value.code != "#N/A":
            return True
        tree = self._tree(key)
        return tree is None or not any(
            isinstance(node, Call) and function_key(node.name) == "NA" for node in walk(tree)
        )

    def _inconsistent(self, key: tuple[int, int]) -> bool:
        row, column = key
        if self._column_formula(row, column) is not None:
            # Measured: a calculated column's cells answer to its own rule.
            return False
        mine = self._key(key)
        for before, after in (((row - 1, column), (row + 1, column)), ((row, column - 1), (row, column + 1))):
            first = self._key(before)
            if first is not None and first == self._key(after) and first != mine:
                return True
        return False

    def _omitted(self, key: tuple[int, int]) -> bool:
        tree = self._tree(key)
        if tree is None:
            return False
        nodes = list(walk(tree))
        if any(isinstance(node, Call) and function_key(node.name) in _SPARES_RANGES for node in nodes):
            return False
        read = [block for node in nodes if (block := self._block_read(node)) is not None]
        for node in nodes:
            if not isinstance(node, AreaReference) or not self._on_this_sheet(node.prefix):
                continue
            first, last = node.first, node.last
            top, bottom = sorted((first.row, last.row))
            left, right = sorted((first.column, last.column))
            if (top != bottom and left != right) or (top == bottom and left == right):
                continue
            held = self.here.held(top, left, bottom, right)
            if not all(self.here.number_constant(row, column) for row, column in held):
                continue
            # Measured: an end written absolute along the range is left alone.
            if left == right:
                upper, lower = (first, last) if first.row <= last.row else (last, first)
                ends = [(top - 1, left)] if not upper.absolute_row else []
                ends += [(bottom + 1, left)] if not lower.absolute_row else []
            else:
                start, end = (first, last) if first.column <= last.column else (last, first)
                ends = [(top, left - 1)] if not start.absolute_column else []
                ends += [(top, right + 1)] if not end.absolute_column else []
            for row, column in ends:
                # Measured: a cell the formula reads another way, not by a
                # name, is not left out.
                if self.here.number_constant(row, column) and not any(
                    block[0] <= row <= block[2] and block[1] <= column <= block[3] for block in read
                ):
                    return True
        return False

    def _block_read(self, node: Node) -> tuple[int, int, int, int] | None:
        """The block of this sheet a reference reads, as (top, left, bottom,
        right), or ``None`` for anything else."""
        if isinstance(node, CellReference) and self._on_this_sheet(node.prefix):
            return (node.ref.row, node.ref.column, node.ref.row, node.ref.column)
        if isinstance(node, AreaReference) and self._on_this_sheet(node.prefix):
            top, bottom = sorted((node.first.row, node.last.row))
            left, right = sorted((node.first.column, node.last.column))
            return (top, left, bottom, right)
        if isinstance(node, AxisReference) and self._on_this_sheet(node.prefix):
            axis = node.ref
            if axis.is_row:
                return (axis.low, 1, axis.high, MAX_COLUMN)
            return (1, axis.low, MAX_ROW, axis.high)
        return None

    def _empty_reference(self, key: tuple[int, int]) -> bool:
        tree = self._tree(key)
        if tree is None:
            return False
        context = Context(self._book(), self.here.name, key[0], key[1], today=self.today)
        for area in _areas(tree, context):
            target = self._sheet(area.sheet)
            if target is not None and target.reads_empty(area.top, area.left, area.bottom, area.right):
                return True
        return False

    def _misleading(self, key: tuple[int, int], cell: _Cell) -> bool:
        tree = self._tree(key)
        if isinstance(tree, Unary) and tree.op in ("+", "-"):
            tree = tree.operand
        if not isinstance(tree, CellReference):
            return False
        prefix = tree.prefix
        if prefix is None:
            source_sheet: _Sheet | None = self.here
        elif prefix.book is None and prefix.last_sheet is None and prefix.sheet is not None:
            source_sheet = self._sheet(prefix.sheet)
        else:
            return False
        if source_sheet is None:
            return False
        source = source_sheet.cells.get((tree.ref.row, tree.ref.column))
        if source is not None and (source.value is not None or source.formula is not None):
            value = source.value
            if not isinstance(value, (int, float)) or isinstance(value, bool):
                return False
        kind = format_kind(self._display_format(None if source is None else source.style))
        shown = parse_format(self._display_format(cell.style)).is_date
        return (kind == "date" and not shown) or (kind == "number" and shown)

    def _calculated_column(self, key: tuple[int, int], cell: _Cell) -> bool:
        row, column = key
        formula = self._column_formula(row, column)
        if formula is None or (cell.value is None and cell.formula is None):
            return False
        if cell.formula is None:
            return True
        data = self._table_data(row, column)
        assert data is not None
        return self._key(key) != r1c1_key(formula, data.top, column)

    def _refused(self, key: tuple[int, int], cell: _Cell) -> bool:
        row, column = key
        if self._table_data(row, column) is None:
            return False
        validation = next(
            (
                found
                for found in self._validations
                if any(span.top <= row <= span.bottom and span.left <= column <= span.right for span in found.ranges)
            ),
            None,
        )
        if validation is None or validation.kind == "none":
            return False
        if cell.value is None and cell.formula is None:
            return not validation.allow_blank
        return not _accepts(validation, cell.value, self._book(), self.here.name, row, column)

    # -- what the rules read --------------------------------------------

    def _tree(self, key: tuple[int, int]) -> Node | None:
        if key not in self._trees:
            formula = self.here.cells[key].formula
            try:
                self._trees[key] = None if formula is None else parse(formula)
            except (FormulaSyntaxError, ValueError, IndexError):
                self._trees[key] = None
        return self._trees[key]

    def _key(self, key: tuple[int, int]) -> str | None:
        """A formula cell's formula as R1C1 text from the cell it is written
        in, or ``None`` for a cell holding no formula."""
        cell = self.here.cells.get(key)
        if cell is None or cell.formula is None or cell.anchor is None:
            return None
        if key not in self._keys:
            self._keys[key] = r1c1_key(cell.formula, *cell.anchor)
        return self._keys[key]

    def _table_data(self, row: int, column: int) -> RangeRef | None:
        for data, _ in self._tables:
            if data.top <= row <= data.bottom and data.left <= column <= data.right:
                return data
        return None

    def _column_formula(self, row: int, column: int) -> str | None:
        """The calculated formula of the table column a cell is in."""
        for data, table in self._tables:
            if data.top <= row <= data.bottom and data.left <= column <= data.right:
                columns = table.columns
                index = column - data.left
                return columns[index].calculated_formula if index < len(columns) else None
        return None

    def _display_format(self, style: int | None) -> str:
        return "General" if self.styles is None else self.styles.display_format(style)

    def _on_this_sheet(self, prefix: Prefix | None) -> bool:
        if prefix is None:
            return True
        return (
            prefix.book is None
            and prefix.last_sheet is None
            and prefix.sheet is not None
            and prefix.sheet.casefold() == self.here.name.casefold()
        )

    def _sheet(self, name: str) -> _Sheet | None:
        key = name.casefold()
        if key == self.here.name.casefold():
            return self.here
        if key not in self._others:
            found = next((sheet for sheet in self.workbook.sheets if sheet.name.casefold() == key), None)
            self._others[key] = None if found is None else _Sheet(found)
        return self._others[key]

    def _book(self) -> Engine:
        """The calculation engine, which resolves names and table
        references to the cells they reach."""
        if self._engine is None:
            from pyofficeeditor.excel._calc.engine import Engine

            self._engine = Engine(self.workbook, today=self.today)
        return self._engine


_REFERENCES = (CellReference, AreaReference, AxisReference, NameReference, StructuredReference)


def _areas(node: Node, context: Context) -> Iterator[Area]:
    """The areas a formula reads by reference, each reference expression
    judged whole: measured, an intersection counts only where its ranges
    meet, and a range OFFSET moves to is not read."""
    if isinstance(node, _REFERENCES) or (isinstance(node, Binary) and node.op in (":", " ", ",")):
        try:
            value = context.evaluate(node)
        except Exception:
            # A reference that does not resolve, or a name whose formula
            # would need calculating first, reaches no cell to judge.
            return
        if isinstance(value, Reference):
            yield from value.areas
        return
    for child in _children(node):
        yield from _areas(child, context)


def _children(node: Node) -> tuple[Node, ...]:
    if isinstance(node, (Unary, Postfix)):
        return (node.operand,)
    if isinstance(node, Binary):
        return (node.left, node.right)
    if isinstance(node, Call):
        return node.args
    if isinstance(node, Invoke):
        return (node.target, *node.args)
    if isinstance(node, Paren):
        return (node.inner,)
    return ()


# ----------------------------------------------------------------------
# Data validation
# ----------------------------------------------------------------------

_NUMBER_KINDS = frozenset({"whole", "decimal", "date", "time"})


def _accepts(
    validation: DataValidation, value: CellValue, engine: Engine, sheet: str, row: int, column: int
) -> bool:
    """Whether a validation accepts a value that is not blank, as Excel
    judges it. Measured on 127 values: a whole number, a decimal, a date or
    a time must be a number, the whole number with no fraction, never text
    that reads as one; a text length counts the value's text; a typed list
    matches its items' text, spaces trimmed but case kept, where a list
    from cells matches as ``=`` does; a custom formula must give TRUE or a
    number other than 0. Its formulas are written for the validation's
    first cell and read here as they read from ``row`` and ``column``."""
    anchor = validation.ranges[0]
    down, across = row - anchor.top, column - anchor.left

    def result(formula: str | None) -> Scalar:
        if not formula:
            return CellError("#N/A")
        try:
            return engine.evaluate(translate_formula(formula, down, across), sheet, row, column)
        except (FormulaSyntaxError, ValueError, UnsupportedFormulaError):
            return CellError("#N/A")

    kind = validation.kind
    if kind == "custom":
        found = result(validation.formula1)
        return found is True or (isinstance(found, float) and found != 0.0)
    if isinstance(value, (bool, str)):
        scalar: Scalar = value
    elif isinstance(value, (int, float)):
        scalar = float(value)
    else:
        return False
    if kind == "list":
        return _listed(validation.formula1 or "", scalar, lambda text: _items(engine, text, down, across, sheet, row, column))
    if kind == "textLength":
        measure = float(len(scalar_text(scalar)))
    elif kind in _NUMBER_KINDS:
        if not isinstance(scalar, float):
            return False
        if kind == "whole" and not scalar.is_integer():
            return False
        measure = scalar
    else:
        return True
    return _holds(validation.operator or "between", measure, result(validation.formula1), result(validation.formula2))


def _items(engine: Engine, formula: str, down: int, across: int, sheet: str, row: int, column: int) -> list[Scalar]:
    try:
        return engine.values(translate_formula(formula, down, across), sheet, row, column)
    except (FormulaSyntaxError, ValueError, UnsupportedFormulaError):
        return []


def _listed(formula: str, value: Scalar, cells: Callable[[str], list[Scalar]]) -> bool:
    if formula.startswith('"') and formula.endswith('"') and len(formula) >= 2:
        wanted = scalar_text(value).strip(" ")
        return any(item.strip(" ") == wanted for item in formula[1:-1].split(","))
    return any(
        not isinstance(item, CellError) and compare(value, item) == 0 and _same_kind(value, item)
        for item in cells(formula)
    )


def _same_kind(left: Scalar, right: Scalar) -> bool:
    """Whether two values are both numbers, both text or both logicals."""
    return type(left) is type(right) or (isinstance(left, float) and isinstance(right, float))


def _holds(operator: ValidationOperator, measure: float, first: Scalar, second: Scalar) -> bool:
    low = first if isinstance(first, float) else None
    high = second if isinstance(second, float) else None
    if low is None:
        return False
    if operator in ("between", "notBetween"):
        if high is None:
            return False
        inside = low <= measure <= high
        return inside if operator == "between" else not inside
    if operator == "equal":
        return measure == low
    if operator == "notEqual":
        return measure != low
    if operator == "greaterThan":
        return measure > low
    if operator == "lessThan":
        return measure < low
    if operator == "greaterThanOrEqual":
        return measure >= low
    return measure <= low


# ----------------------------------------------------------------------
# Ignored errors
# ----------------------------------------------------------------------


def read_ignored_errors(sheet: Worksheet) -> list[IgnoredError]:
    """A sheet's ``<ignoredErrors>``, entry by entry."""
    found = sheet.document.root.child("ignoredErrors")
    if found is None:
        return []
    entries: list[IgnoredError] = []
    for element in found.children_named("ignoredError"):
        ranges: list[RangeRef] = []
        for piece in (element.get("sqref") or "").split():
            try:
                ranges.append(RangeRef.parse(piece).normalized)
            except ValueError:
                continue
        rules: frozenset[ErrorRule] = frozenset(rule for rule in ERROR_RULES if element.get(rule) in ("1", "true"))
        if ranges and rules:
            entries.append(IgnoredError(tuple(ranges), rules))
    return entries


def check_errors(
    sheet: Worksheet,
    rules: Iterable[ErrorRule] | None = None,
    *,
    include_ignored: bool = False,
    today: dt.date | None = None,
) -> list[ErrorCheck]:
    """The cells a sheet's error checking catches, by row, then column,
    then rule in Excel's order."""
    wanted = DEFAULT_ERROR_RULES if rules is None else frozenset(rules)
    unknown = wanted - set(ERROR_RULES)
    if unknown:
        raise ValueError(f"{sorted(unknown)} are not error-checking rules; they are {', '.join(ERROR_RULES)}.")
    ignored = read_ignored_errors(sheet)
    found: list[ErrorCheck] = []
    for (row, column), rule in _Checker(sheet, wanted, today or dt.date.today()).run():
        hidden = any(
            rule in entry.rules
            and any(span.top <= row <= span.bottom and span.left <= column <= span.right for span in entry.ranges)
            for entry in ignored
        )
        if hidden and not include_ignored:
            continue
        found.append(ErrorCheck(CellRef(row, column), rule, hidden))
    return found


__all__ = [
    "DEFAULT_ERROR_RULES",
    "ERROR_RULES",
    "ErrorCheck",
    "ErrorRule",
    "IgnoredError",
    "check_errors",
    "format_kind",
    "r1c1_key",
    "read_ignored_errors",
    "two_digit_year",
]
