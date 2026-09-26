"""The rules of Excel's error checking, and the cells a sheet tells it to
ignore them for.

Error checking marks a cell with a green triangle under one of ten rules,
named as a sheet's ``<ignoredErrors>`` names them. Ignore Error dismisses a
triangle for one rule, and the sheet keeps the dismissals as entries, each
naming some ranges and the rules their cells ignore. Measured in Excel, an
entry says all that a cell ignores: a later entry covering a cell replaces
what an earlier one said of it, and an entry naming no rule makes its
cells ignore nothing.

A change follows ``Range.Errors(i).Ignore`` one rule, one cell and one
row at a time, as Excel was measured to on 47 sequences of steps:

- A cell whose rules change leaves its entry. Alone in the entry, it takes
  the entry with it. Otherwise the range holding it splits into what lies
  above it, below it, to its left and to its right; the first of those
  takes the range's place and the rest follow the entry's last range.
- The cell then joins the entry holding its new rules, or a new one at the
  end. It widens the last range it makes a rectangle with, and that range
  merges with the range before it when the two stand one on the other,
  column for column; otherwise it is a range of its own at the end. A cell
  left ignoring nothing joins an entry naming no rule, as Excel keeps one,
  unless it was alone in its entry.

Excel's own saving loses the marks of a cell that was alone in its entry
and takes other rules, though Excel shows them ignored until it saves: all
of the sheet's marks when that entry came first, and otherwise the cell's,
unless another entry already held its new rules. Here they are kept.
"""

from __future__ import annotations

from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING, Literal

from pyofficeeditor._xml import Element
from pyofficeeditor.excel._reference import CellRef, RangeRef, column_letter
from pyofficeeditor.excel._schema import WORKSHEET_CHILD_ORDER, insert_in_schema_order

if TYPE_CHECKING:
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

#: Every rule, in the order Excel numbers them, which is also the order an
#: entry names them in.
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

_BY_NAME: dict[str, ErrorRule] = {rule: rule for rule in ERROR_RULES}

#: The namespace Excel writes ``misleadingFormat`` in, declared on the entry
#: itself, after the entry's other rules.
_X16R3 = "http://schemas.microsoft.com/office/spreadsheetml/2018/08/main"


@dataclass(frozen=True)
class IgnoredError:
    """One ``<ignoredError>`` of a sheet's ``<ignoredErrors>``: ranges, and
    the rules their cells ignore. It says all that a cell ignores, so a
    later entry covering the same cell replaces it, and an entry naming no
    rule makes its cells ignore nothing."""

    ranges: tuple[RangeRef, ...]
    rules: frozenset[ErrorRule]


def read_ignored_errors(sheet: Worksheet) -> list[IgnoredError]:
    """A sheet's ``<ignoredErrors>``, entry by entry, in the file's order."""
    return [IgnoredError(tuple(_range(block) for block in entry.blocks), entry.rules) for entry in _read(sheet)]


def ignored_rules(entries: Sequence[IgnoredError], row: int, column: int) -> frozenset[ErrorRule]:
    """The rules a cell ignores: those of the last entry covering it."""
    for entry in reversed(entries):
        for span in entry.ranges:
            if span.top <= row <= span.bottom and span.left <= column <= span.right:
                return entry.rules
    return frozenset()


def rule_list(rules: ErrorRule | Iterable[ErrorRule]) -> tuple[ErrorRule, ...]:
    """The rules asked for, each once, in the order given; an unknown one
    raises :class:`ValueError`."""
    asked: tuple[str, ...] = (rules,) if isinstance(rules, str) else tuple(rules)
    unknown = sorted(set(asked) - set(ERROR_RULES))
    if unknown:
        raise ValueError(f"{unknown} are not error-checking rules; they are {', '.join(ERROR_RULES)}.")
    return tuple(_BY_NAME[name] for name in dict.fromkeys(asked))


def set_ignored_errors(sheet: Worksheet, ranges: Sequence[RangeRef], rules: Sequence[ErrorRule], ignore: bool) -> None:
    """Ignore ``rules`` in the cells of ``ranges``, or stop ignoring them,
    one rule at a time and each range a row at a time, as Excel's
    ``Range.Errors(i).Ignore`` does a cell at a time."""
    ignores = _Ignores(_read(sheet))
    for rule in rules:
        for span in ranges:
            ignores.change_block((span.top, span.left, span.bottom, span.right), rule, ignore)
    ignores.write(sheet.document.root)


def clear_ignored_errors(sheet: Worksheet) -> bool:
    """Drop a sheet's ``<ignoredErrors>``, so no cell ignores any rule.
    Returns whether there was one."""
    found = sheet.document.root.child("ignoredErrors")
    if found is None:
        return False
    sheet.document.root.remove(found)
    return True


# ----------------------------------------------------------------------
# The entries as Excel changes them
# ----------------------------------------------------------------------

#: A block of cells: top row, left column, bottom row, right column.
_Block = tuple[int, int, int, int]


@dataclass(eq=False)
class _Entry:
    blocks: list[_Block]
    rules: frozenset[ErrorRule]
    #: The element read from the file, left as it is while the entry is.
    element: Element | None = None
    changed: bool = False


def _read(sheet: Worksheet) -> list[_Entry]:
    container = sheet.document.root.child("ignoredErrors")
    if container is None:
        return []
    entries: list[_Entry] = []
    for element in container.children_named("ignoredError"):
        blocks: list[_Block] = []
        for piece in (element.get("sqref") or "").split():
            try:
                span = RangeRef.parse(piece).normalized
            except ValueError:
                continue
            blocks.append((span.top, span.left, span.bottom, span.right))
        if blocks:
            rules: frozenset[ErrorRule] = frozenset(rule for rule in ERROR_RULES if element.get(rule) in ("1", "true"))
            entries.append(_Entry(blocks, rules, element))
    return entries


def _range(block: _Block) -> RangeRef:
    top, left, bottom, right = block
    return RangeRef(CellRef(top, left), CellRef(bottom, right))


def _holds(block: _Block, row: int, column: int) -> bool:
    return block[0] <= row <= block[2] and block[1] <= column <= block[3]


def _around(block: _Block, row: int, column: int) -> list[_Block]:
    """What is left of ``block`` without one cell of it: the rows above the
    cell, the rows below, then the cells to its left and to its right."""
    top, left, bottom, right = block
    pieces: list[_Block] = []
    if row > top:
        pieces.append((top, left, row - 1, right))
    if row < bottom:
        pieces.append((row + 1, left, bottom, right))
    if column > left:
        pieces.append((row, left, row, column - 1))
    if column < right:
        pieces.append((row, column + 1, row, right))
    return pieces


def _widen(block: _Block, row: int, column: int) -> _Block | None:
    """``block`` with the cell beside one of its ends, when the two make a
    rectangle: a row of cells widened along its row, or a column along its
    column."""
    top, left, bottom, right = block
    if top == bottom == row and column in (left - 1, right + 1):
        return (top, min(left, column), bottom, max(right, column))
    if left == right == column and row in (top - 1, bottom + 1):
        return (min(top, row), left, max(bottom, row), right)
    return None


def _stack(upper: _Block, lower: _Block) -> _Block | None:
    """The two blocks as one, when one stands on the other column for
    column."""
    if upper[1] != lower[1] or upper[3] != lower[3]:
        return None
    if upper[2] + 1 == lower[0] or lower[2] + 1 == upper[0]:
        return (min(upper[0], lower[0]), upper[1], max(upper[2], lower[2]), upper[3])
    return None


def _touches(block: _Block, area: _Block) -> bool:
    """Whether ``block`` shares or borders a cell of ``area``."""
    return not (
        block[3] < area[1] - 1 or block[1] > area[3] + 1 or block[2] < area[0] - 1 or block[0] > area[2] + 1
    )


def _meets(block: _Block, area: _Block) -> bool:
    return not (block[3] < area[1] or block[1] > area[3] or block[2] < area[0] or block[0] > area[2])


class _Ignores:
    """A sheet's entries, changed as Excel changes them."""

    def __init__(self, entries: list[_Entry], *, shortcuts: bool = True) -> None:
        self.entries = entries
        self.dropped: list[_Entry] = []
        #: Whether a block may be changed in strides rather than a cell at a
        #: time; both give the same entries.
        self.shortcuts = shortcuts

    def change_block(self, area: _Block, rule: ErrorRule, ignore: bool) -> None:
        """Every cell of ``area``, row by row, as :meth:`change_cell`."""
        fresh = frozenset((rule,)) if ignore else frozenset[ErrorRule]()
        if self.shortcuts and self._apart(area, fresh):
            # Cells nothing covers, beside nothing their entry holds, come to
            # one range at the entry's end, as they would a cell at a time.
            self._join_block(area, fresh)
            return
        top, left, bottom, right = area
        row = top
        while row <= bottom:
            column = left
            while column <= right:
                joined = self.change_cell(row, column, rule, ignore)
                if self.shortcuts:
                    column = self._stride(row, column, right, fresh, joined, rule, ignore)
                column += 1
            if self.shortcuts:
                row = self._descend(area, row, fresh)
            row += 1

    def change_cell(self, row: int, column: int, rule: ErrorRule, ignore: bool) -> bool:
        """As ``Range.Errors(rule).Ignore = ignore`` for one cell. Returns
        whether the cell was covered by no entry before."""
        holder = self._holder(row, column)
        if holder is None:
            self._join(row, column, frozenset((rule,)) if ignore else frozenset[ErrorRule]())
            return True
        entry = self.entries[holder]
        asked: frozenset[ErrorRule] = frozenset((rule,))
        rules = entry.rules | asked if ignore else entry.rules - asked
        if rules == entry.rules:
            return False
        for earlier in self.entries[:holder]:
            # An entry replaced for this cell by a later one, which Excel
            # never writes, lets it go too.
            if any(_holds(block, row, column) for block in earlier.blocks):
                self._split(earlier, row, column)
        if all(block == (row, column, row, column) for block in entry.blocks):
            self._drop(entry)
            if not rules:
                return False
        else:
            self._split(entry, row, column)
        self._join(row, column, rules)
        return False

    # -- the steps ------------------------------------------------------

    def _holder(self, row: int, column: int) -> int | None:
        for index in range(len(self.entries) - 1, -1, -1):
            if any(_holds(block, row, column) for block in self.entries[index].blocks):
                return index
        return None

    def _split(self, entry: _Entry, row: int, column: int) -> None:
        blocks = entry.blocks
        index = 0
        while index < len(blocks):
            if not _holds(blocks[index], row, column):
                index += 1
                continue
            pieces = _around(blocks[index], row, column)
            if pieces:
                blocks[index] = pieces[0]
                blocks.extend(pieces[1:])
                index += 1
            else:
                del blocks[index]
        entry.changed = True
        if not blocks:
            self._drop(entry)

    def _drop(self, entry: _Entry) -> None:
        self.entries.remove(entry)
        if entry.element is not None:
            self.dropped.append(entry)

    def _target(self, rules: frozenset[ErrorRule]) -> _Entry | None:
        return next((entry for entry in self.entries if entry.rules == rules), None)

    def _join(self, row: int, column: int, rules: frozenset[ErrorRule]) -> None:
        entry = self._target(rules)
        if entry is None:
            self.entries.append(_Entry([(row, column, row, column)], rules, changed=True))
            return
        blocks = entry.blocks
        entry.changed = True
        for index in range(len(blocks) - 1, -1, -1):
            grown = _widen(blocks[index], row, column)
            if grown is None:
                continue
            blocks[index] = grown
            if index > 0:
                merged = _stack(blocks[index - 1], grown)
                if merged is not None:
                    blocks[index - 1] = merged
                    del blocks[index]
            return
        blocks.append((row, column, row, column))

    # -- strides --------------------------------------------------------

    def _apart(self, area: _Block, fresh: frozenset[ErrorRule]) -> bool:
        """Whether no entry covers a cell of ``area`` and none of the blocks
        its cells would join borders it."""
        if any(_meets(block, area) for entry in self.entries for block in entry.blocks):
            return False
        target = self._target(fresh)
        return target is None or not any(_touches(block, area) for block in target.blocks)

    def _join_block(self, area: _Block, fresh: frozenset[ErrorRule]) -> None:
        entry = self._target(fresh)
        if entry is None:
            self.entries.append(_Entry([area], fresh, changed=True))
        else:
            entry.blocks.append(area)
            entry.changed = True

    def _stride(
        self, row: int, column: int, right: int, fresh: frozenset[ErrorRule], joined: bool, rule: ErrorRule, ignore: bool
    ) -> int:
        """The last column of the row handled once the cell at ``column``
        has been: past the cells after it that a cell at a time would
        change the same way."""
        if not joined:
            return self._skip(row, column, right, rule, ignore)
        entry = self._target(fresh)
        if entry is None:
            return column
        top, left, bottom, end = entry.blocks[-1]
        if not (top == bottom == row and end == column):
            return column
        # The cells after it widen the same row of cells while nothing
        # covers them, until it would stand on the block before it.
        limit = min(right, self._next_covered(row, column) - 1)
        if len(entry.blocks) > 1:
            before = entry.blocks[-2]
            if before[1] == left and row in (before[0] - 1, before[2] + 1) and before[3] > column:
                limit = min(limit, before[3] - 1)
        if limit <= column:
            return column
        entry.blocks[-1] = (row, left, row, limit)
        return limit

    def _descend(self, area: _Block, row: int, fresh: frozenset[ErrorRule]) -> int:
        """The last row of ``area`` handled once ``row`` has been: past the
        rows after it that a row at a time would add to the same block."""
        _, left, bottom, right = area
        entry = self._target(fresh)
        if entry is None or row >= bottom:
            return row
        last = entry.blocks[-1]
        if (last[1], last[2], last[3]) != (left, row, right):
            return row
        # The rows below, covered by nothing and beside nothing else their
        # entry holds, each add a row to the block as Excel adds them.
        band = (row + 1, left, bottom, right)
        if any(_meets(block, band) for other in self.entries for block in other.blocks):
            return row
        if any(_touches(block, band) for block in entry.blocks[:-1]):
            return row
        grown = (last[0], left, row + 1, right)
        if len(entry.blocks) > 1 and _stack(entry.blocks[-2], grown) is not None:
            return row
        entry.blocks[-1] = (last[0], left, bottom, right)
        return bottom

    def _skip(self, row: int, column: int, right: int, rule: ErrorRule, ignore: bool) -> int:
        """Past the cells after ``column`` that already ignore, or already
        do not ignore, ``rule`` as asked, under the same entry."""
        holder = self._holder(row, column)
        if holder is None:
            return column
        entry = self.entries[holder]
        if (rule in entry.rules) != ignore:
            return column
        block = next(block for block in entry.blocks if _holds(block, row, column))
        limit = min(right, block[3])
        for later in self.entries[holder + 1 :]:
            for other in later.blocks:
                if other[0] <= row <= other[2] and other[1] > column:
                    limit = min(limit, other[1] - 1)
        return max(column, limit)

    def _next_covered(self, row: int, column: int) -> int:
        """The first column after ``column`` in which some entry covers the
        row's cell."""
        found = 1 << 30
        for entry in self.entries:
            for top, left, bottom, right in entry.blocks:
                if top <= row <= bottom and right > column:
                    found = min(found, max(left, column + 1))
        return found

    # -- the markup -----------------------------------------------------

    def write(self, root: Element) -> None:
        """Bring the sheet's ``<ignoredErrors>`` in line with the entries,
        leaving every entry that did not change as the file had it."""
        container = root.child("ignoredErrors")
        for entry in self.dropped:
            if container is not None and entry.element is not None:
                container.remove(entry.element)
        if not self.entries:
            if container is not None and not any(True for _ in container.children_named("ignoredError")):
                root.remove(container)
            return
        if container is None:
            container = Element.create("ignoredErrors")
            insert_in_schema_order(root, container, WORKSHEET_CHILD_ORDER)
        for entry in self.entries:
            if entry.element is None:
                entry.element = _element(entry)
                container.append(entry.element)
            elif entry.changed:
                entry.element.set("sqref", _sqref(entry.blocks))


def _sqref(blocks: list[_Block]) -> str:
    return " ".join(_a1(block) for block in blocks)


def _a1(block: _Block) -> str:
    top, left, bottom, right = block
    first = f"{column_letter(left)}{top}"
    return first if (top, left) == (bottom, right) else f"{first}:{column_letter(right)}{bottom}"


def _element(entry: _Entry) -> Element:
    """An entry as Excel writes one: its ranges, its rules in Excel's order,
    and ``misleadingFormat`` last in a namespace declared on the entry."""
    element = Element.create("ignoredError", {"sqref": _sqref(entry.blocks)})
    for rule in ERROR_RULES:
        if rule in entry.rules and rule != "misleadingFormat":
            element.set(rule, "1")
    if "misleadingFormat" in entry.rules:
        element.set("xmlns:x16r3", _X16R3)
        element.set("x16r3:misleadingFormat", "1")
    return element


__all__ = [
    "ERROR_RULES",
    "ErrorRule",
    "IgnoredError",
    "clear_ignored_errors",
    "ignored_rules",
    "read_ignored_errors",
    "rule_list",
    "set_ignored_errors",
]
