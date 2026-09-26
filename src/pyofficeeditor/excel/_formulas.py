"""Shifting the references inside a formula.

A formula assigned to a range is stored once. Excel writes the text on the
group's first cell and leaves the rest pointing at it by index::

    <c r="D2"><f t="shared" ref="D2:D5" si="0">B2*C2</f><v>510</v></c>
    <c r="D3"><f t="shared" si="0"/><v>1445</v></c>

So D3's formula is not in the file. It is D2's, translated down one row:
``B3*C3``. Without that translation a cell-level reader reports an empty
formula for every follower, which is wrong rather than incomplete.

Finding the references to shift is the delicate part, because a formula is
not only references. Three things look like ``A1`` and are not:

- ``LOG10(x)`` contains ``G10``
- ``"A1"`` is a string literal
- ``'My Sheet A1'!B2`` has a reference inside a quoted sheet name

So quoted runs are skipped whole, and what remains is matched with a
boundary that rejects anything preceded by an identifier character or
followed by an opening parenthesis. A reference this module does not
recognise is left exactly as written, which is the safe direction: an
untranslated reference is visibly wrong in one cell, while a mangled
function name breaks the formula everywhere.

This is not a formula parser. The analyzer and linter named in the project's
goals will need one; this does the one job a shared formula requires.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from pyofficeeditor.excel._reference import MAX_COLUMN, MAX_ROW, AxisRef, CellRef, RangeRef
from pyofficeeditor.excel._tokens import Token, TokenKind, render, tokenize

#: A cell reference inside formula text.
#:
#: The lookbehind rejects a match that continues an identifier, so the
#: ``G10`` in ``LOG10`` is not a reference. The lookahead rejects one
#: followed by ``(``, which is a function call, or by a digit, which would
#: mean the row number was cut short. ``!`` is allowed before a reference
#: because ``Sheet1!A1`` really is one.
_REFERENCE = re.compile(
    r"""
    (?<![A-Za-z0-9_$.\[])      # not continuing an identifier or a table column
    (\$?)([A-Za-z]{1,3})       # optional $ then the column letters
    (\$?)([0-9]{1,7})          # optional $ then the row number
    (?![0-9(\[])               # not a longer number, a call, or a table ref
    """,
    re.VERBOSE,
)


def translate_formula(formula: str, rows: int, columns: int) -> str:
    """``formula`` as it reads ``rows`` down and ``columns`` across.

    Relative references move, absolute ones stay, and anything that is not a
    reference is untouched. A reference that would land off the sheet is left
    as written rather than raising, because Excel stores such a formula as an
    error in the cell rather than refusing the file.
    """
    if not formula or (rows == 0 and columns == 0):
        return formula

    out: list[str] = []
    for chunk, is_literal in _split_literals(formula):
        out.append(chunk if is_literal else _shift(chunk, rows, columns))
    return "".join(out)


def _shift(text: str, rows: int, columns: int) -> str:
    def replace(match: re.Match[str]) -> str:
        dollar_column, letters, dollar_row, digits = match.groups()
        try:
            reference = CellRef.parse(f"{dollar_column}{letters}{dollar_row}{digits}")
        except ValueError:
            # Not a reference Excel could store, such as ZZZ1 or row 0.
            return match.group(0)
        row = reference.row if reference.absolute_row else reference.row + rows
        column = reference.column if reference.absolute_column else reference.column + columns
        if not (1 <= row <= MAX_ROW and 1 <= column <= MAX_COLUMN):
            return match.group(0)
        return reference.translated(rows, columns).a1

    return _REFERENCE.sub(replace, text)


def _split_literals(formula: str) -> list[tuple[str, bool]]:
    """Break a formula into runs, marking which are literal text.

    A literal is a double-quoted string or a single-quoted sheet name. Excel
    escapes an embedded quote by doubling it, in both kinds, so a closing
    quote is one that is not itself followed by another.
    """
    parts: list[tuple[str, bool]] = []
    buffer: list[str] = []
    index = 0
    while index < len(formula):
        character = formula[index]
        if character in "\"'":
            if buffer:
                parts.append(("".join(buffer), False))
                buffer = []
            end = _end_of_literal(formula, index, character)
            parts.append((formula[index:end], True))
            index = end
            continue
        buffer.append(character)
        index += 1
    if buffer:
        parts.append(("".join(buffer), False))
    return parts


def _end_of_literal(formula: str, start: int, quote: str) -> int:
    index = start + 1
    while index < len(formula):
        if formula[index] == quote:
            if index + 1 < len(formula) and formula[index + 1] == quote:
                index += 2
                continue
            return index + 1
        index += 1
    # Unterminated. Treat the remainder as literal rather than shifting
    # references inside what the author meant as text.
    return len(formula)


#: Past ASCII, the characters Excel quotes a sheet's name for wherever they
#: stand in it. Measured a character at a time over Latin-1, general
#: punctuation, currency signs, letterlike symbols, the first arrows and
#: operators, CJK punctuation and the full-width forms; every other
#: character past ASCII reads as a letter, as every letter measured did.
_QUOTED_ANYWHERE: tuple[tuple[int, int], ...] = (
    (0x00A0, 0x00A0), (0x00A2, 0x00A3), (0x00A5, 0x00A6), (0x00A9, 0x00A9), (0x00AB, 0x00AC),
    (0x00AE, 0x00AE), (0x00BB, 0x00BB), (0x2011, 0x2012), (0x2017, 0x2017), (0x201A, 0x201B),
    (0x201E, 0x201F), (0x2022, 0x2024), (0x2031, 0x2031), (0x2034, 0x2034), (0x2036, 0x203A),
    (0x203C, 0x2043), (0x2045, 0x2051), (0x2053, 0x205E), (0x2065, 0x2069), (0x20B6, 0x20C0),
    (0x3018, 0x301C), (0x3030, 0x3030), (0x303D, 0x303D), (0xFF5F, 0xFF60),
)  # fmt: skip
#: Those it quotes a name for only when the name starts with them.
_QUOTED_FIRST: tuple[tuple[int, int], ...] = (
    (0x2000, 0x200F), (0x2028, 0x202F), (0x2044, 0x2044), (0x2052, 0x2052), (0x205F, 0x2064),
    (0x206A, 0x206F), (0x20A0, 0x20B5), (0x2100, 0x2101), (0x2104, 0x2104), (0x2106, 0x2106),
    (0x2108, 0x2108), (0x2114, 0x2114), (0x2117, 0x2118), (0x211E, 0x2120), (0x2123, 0x2123),
    (0x2125, 0x2125), (0x2127, 0x2127), (0x2129, 0x2129), (0x212E, 0x212E), (0x213A, 0x213B),
    (0x2140, 0x2144), (0x214A, 0x214D), (0x214F, 0x214F), (0x219A, 0x219F), (0x2201, 0x2201),
    (0x2204, 0x2206), (0x2209, 0x220A), (0x220C, 0x220E), (0x2210, 0x2210), (0x2212, 0x2214),
    (0x2216, 0x2219), (0x221B, 0x221C), (0x3004, 0x3004), (0x3020, 0x3020), (0x302A, 0x302F),
    (0x3036, 0x3037), (0x303E, 0x303F),
)  # fmt: skip

#: A name made of one to three letters and a number, which may read as a
#: cell in A1 notation.
_A1_NAME = re.compile(r"([A-Za-z]{1,3})([0-9]+)")
#: The highest number an R1C1 reference's row or column may carry for Excel
#: to quote a name that starts with one: measured, a column's is the row's,
#: not the sheet's last column.
_R1C1_LIMIT = MAX_ROW


def quote_sheet_name(name: str) -> str:
    """A sheet name as a formula must spell it: in apostrophes, with an
    apostrophe inside it doubled, when Excel would quote it.

    Measured in Excel, a name is quoted when it holds a character other
    than a letter, a digit, ``_`` and ``.``, or starts with a digit or
    ``.``; when it is ``TRUE`` or ``FALSE``; when it reads as a cell, as
    ``A1`` and ``ZZ99`` do; and when it starts as an R1C1 reference, as
    ``R2D2`` does. Past ASCII, a letter of any script is a letter, and some
    symbols are quoted, as ``©`` is anywhere and ``€`` at the start.
    """
    if _needs_quotes(name):
        return "'" + name.replace("'", "''") + "'"
    return name


def _needs_quotes(name: str) -> bool:
    if not name or not _may_start(name[0]) or not all(_may_follow(char) for char in name[1:]):
        return True
    return name.upper() in ("TRUE", "FALSE") or _reads_as_a1(name) or _starts_as_r1c1(name)


def _may_start(char: str) -> bool:
    if char.isascii():
        return char.isalpha() or char in "_\\"
    point = ord(char)
    return not _within(point, _QUOTED_ANYWHERE) and not _within(point, _QUOTED_FIRST)


def _may_follow(char: str) -> bool:
    if char.isascii():
        return char.isalnum() or char in "_.\\"
    return not _within(ord(char), _QUOTED_ANYWHERE)


def _within(point: int, spans: tuple[tuple[int, int], ...]) -> bool:
    return any(low <= point <= high for low, high in spans)


def _reads_as_a1(name: str) -> bool:
    """Whether the whole name is a cell: ``A1`` and ``A01`` are, ``A0``,
    ``XFE1`` and ``A1B`` are not. Measured, a name starting with ``LOG10``,
    a function's name, never is, though ``LOG100`` would be a cell."""
    match = _A1_NAME.fullmatch(name)
    if match is None or name.upper().startswith("LOG10"):
        return False
    column = 0
    for letter in match.group(1).upper():
        column = column * 26 + ord(letter) - ord("A") + 1
    return column <= MAX_COLUMN and 1 <= int(match.group(2)) <= MAX_ROW


def _starts_as_r1c1(name: str) -> bool:
    """Whether the name starts as an R1C1 reference, measured: ``R`` or
    ``C`` with a number in range, whatever follows, as ``R2D2`` and ``C1X``
    do; ``RC`` with one; or ``R``, ``C`` or ``RC`` alone. ``Rx``, ``RCX``
    and ``R0`` do not."""
    rest = name.upper()
    if rest.startswith("R"):
        number = _leading_number(rest[1:])
        if number:
            return 1 <= int(number) <= _R1C1_LIMIT
        rest = rest[1:]
        if not rest:
            return True
        if not rest.startswith("C"):
            return False
    if rest.startswith("C"):
        number = _leading_number(rest[1:])
        if number:
            return 1 <= int(number) <= _R1C1_LIMIT
        return rest == "C"
    return False


def _leading_number(text: str) -> str:
    match = re.match(r"[0-9]*", text)
    return "" if match is None else match.group()


def rename_sheet_in_formula(formula: str, old: str, new: str) -> str:
    """``formula`` with references to one sheet repointed at another.

    ``old`` and ``new`` are spelled as a formula stores a name. A sheet is
    named bare or quoted, alone or as an end of a 3D reference's span of
    sheets, and each is repointed; a name inside a span is not, since the
    span still runs between the same two ends. Measured in Excel, a span is
    quoted whole when either end needs quotes, so renaming ``Mar`` to
    ``Q1 End`` turns ``Jan:Mar!B3`` into ``'Jan:Q1 End'!B3``, and each name
    is quoted as it needs rather than as the old one was.

    Text is left alone, and so is another workbook's sheet of the same
    name: ``[1]Data!A1`` is not this workbook's ``Data``.
    """
    if old == new:
        return formula
    tokens = tokenize(formula)
    changed = False
    for index, token in enumerate(tokens):
        if token.kind is not TokenKind.SHEET or (index and tokens[index - 1].raw.endswith("]")):
            continue
        names = _qualifier_names(token.raw)
        if names is None or old not in names:
            continue
        token.raw = _qualifier([new if name == old else name for name in names]) + "!"
        changed = True
    return render(tokens) if changed else formula


def _qualifier_names(raw: str) -> list[str] | None:
    """The names a qualifier such as ``'Jan 1:Mar'!`` gives, as stored, or
    ``None`` for another workbook's."""
    body = raw[:-1]
    if len(body) >= 2 and body[0] == body[-1] == "'":
        body = body[1:-1].replace("''", "'")
    if "[" in body:
        return None
    return body.split(":")


def _qualifier(names: list[str]) -> str:
    """Names as a qualifier spells them before its ``!``: measured, a span
    is quoted whole when either end would be quoted alone."""
    joined = ":".join(names)
    if any(_needs_quotes(name) for name in names):
        return "'" + joined.replace("'", "''") + "'"
    return joined


@dataclass(frozen=True)
class Shift:
    """How far, and from where, an insertion moves things.

    One value rather than four loose arguments, because it travels through
    every function that has to move something and a transposed pair would be
    a silent corruption.
    """

    rows_at: int | None = None
    row_count: int = 0
    columns_at: int | None = None
    column_count: int = 0

    @classmethod
    def rows(cls, at: int, count: int) -> Shift:
        return cls(rows_at=at, row_count=count)

    @classmethod
    def columns(cls, at: int, count: int) -> Shift:
        return cls(columns_at=at, column_count=count)

    @property
    def is_empty(self) -> bool:
        return not self.row_count and not self.column_count

    def moved(self, reference: CellRef) -> CellRef | None:
        """Where a reference lands, or ``None`` if it would leave the sheet."""
        row = reference.row
        column = reference.column
        if self.rows_at is not None and self.row_count and row >= self.rows_at:
            row += self.row_count
        if self.columns_at is not None and self.column_count and column >= self.columns_at:
            column += self.column_count
        if row == reference.row and column == reference.column:
            return reference
        if not (1 <= row <= MAX_ROW and 1 <= column <= MAX_COLUMN):
            return None
        return CellRef(row, column, reference.absolute_row, reference.absolute_column)

    def moved_axis(self, span: AxisRef) -> AxisRef | None:
        """Where a whole-row or whole-column reference lands.

        ``None`` when an end would leave the sheet. The insertion is on the
        other axis for half of these, in which case nothing moves.
        """
        at = self.rows_at if span.is_row else self.columns_at
        count = self.row_count if span.is_row else self.column_count
        limit = MAX_ROW if span.is_row else MAX_COLUMN
        if at is None or not count:
            return span
        low = span.low + count if span.low >= at else span.low
        high = span.high + count if span.high >= at else span.high
        if low == span.low and high == span.high:
            return span
        if high > limit:
            return None
        return span.with_span(low, high)


def shift_formula(
    formula: str,
    shift: Shift,
    *,
    formula_sheet: str,
    target_sheet: str,
) -> str:
    """Move the references a row or column insertion pushed along.

    Only references addressing ``target_sheet`` move. A bare reference
    addresses the sheet its formula lives on, so ``formula_sheet`` decides
    whether it counts; a qualified one says which sheet outright. That
    distinction is the whole reason this goes through the tokenizer: a
    formula on ``Summary`` reading ``Data!A5`` has to move when rows are
    inserted into ``Data`` and stay put when they are inserted into
    ``Summary``.

    A reference at or after the insertion point moves by the count. One
    before it does not, which is what makes a range spanning the insertion
    grow rather than slide.
    """
    if shift.is_empty:
        return formula
    if formula_sheet != target_sheet and "!" not in formula:
        # Nothing here can address the edited sheet.
        return formula

    tokens = tokenize(formula)
    changed = False
    for token in tokens:
        if token.kind is TokenKind.AXIS and isinstance(token.value, AxisRef):
            addresses = token.sheet if token.sheet is not None else formula_sheet
            if addresses != target_sheet:
                continue
            moved_span = shift.moved_axis(token.value)
            if moved_span is None or moved_span == token.value:
                continue
            token.raw = moved_span.a1
            token.value = moved_span
            changed = True
            continue

        if token.kind is not TokenKind.REFERENCE or not isinstance(token.value, CellRef):
            continue
        addresses = token.sheet if token.sheet is not None else formula_sheet
        if addresses != target_sheet:
            continue

        moved = shift.moved(token.value)
        if moved is None:
            # Pushed off the sheet. Excel turns such a formula into #REF!;
            # leaving it as written is the conservative half of that, and
            # the caller refuses the insertion before it can happen.
            continue
        if moved == token.value:
            continue
        token.raw = moved.a1
        token.value = moved
        changed = True

    return render(tokens) if changed else formula


#: What Excel puts in a formula whose target no longer exists.
REF_ERROR = "#REF!"


@dataclass(frozen=True)
class Deletion:
    """Which rows or columns are being removed, and what that does to a
    reference.

    Deletion is not the inverse of insertion. A reference to something that
    is gone becomes ``#REF!``, and a *range* that only partly overlaps the
    deletion shrinks instead, so the two ends have to be decided together.
    """

    rows_at: int | None = None
    row_count: int = 0
    columns_at: int | None = None
    column_count: int = 0

    @classmethod
    def rows(cls, at: int, count: int) -> Deletion:
        return cls(rows_at=at, row_count=count)

    @classmethod
    def columns(cls, at: int, count: int) -> Deletion:
        return cls(columns_at=at, column_count=count)

    @property
    def is_empty(self) -> bool:
        return not self.row_count and not self.column_count

    def covers_row(self, row: int) -> bool:
        return (
            self.rows_at is not None
            and bool(self.row_count)
            and self.rows_at <= row < self.rows_at + self.row_count
        )

    def covers_column(self, column: int) -> bool:
        return (
            self.columns_at is not None
            and bool(self.column_count)
            and self.columns_at <= column < self.columns_at + self.column_count
        )

    @staticmethod
    def _survivor(value: int, at: int | None, count: int) -> int | None:
        """Where an index lands, or ``None`` when it is one of the deleted."""
        if at is None or not count:
            return value
        if value < at:
            return value
        if value < at + count:
            return None
        return value - count

    def moved(self, reference: CellRef) -> CellRef | None:
        """Where a single cell lands, or ``None`` if it is deleted."""
        row = self._survivor(reference.row, self.rows_at, self.row_count)
        column = self._survivor(reference.column, self.columns_at, self.column_count)
        if row is None or column is None:
            return None
        return CellRef(row, column, reference.absolute_row, reference.absolute_column)

    def moved_range(self, block: RangeRef) -> RangeRef | None:
        """Where a block lands, or ``None`` when nothing of it survives.

        A block straddling the deletion shrinks: its deleted rows come out
        and the rest closes up. One wholly inside it is gone.
        """
        rows = self._surviving_span(
            block.top, block.bottom, self.rows_at, self.row_count
        )
        columns = self._surviving_span(
            block.left, block.right, self.columns_at, self.column_count
        )
        if rows is None or columns is None:
            return None
        top, bottom = rows
        left, right = columns
        return RangeRef(
            CellRef(top, left, block.start.absolute_row, block.start.absolute_column),
            CellRef(bottom, right, block.end.absolute_row, block.end.absolute_column),
        )

    def moved_axis(self, span: AxisRef) -> AxisRef | None:
        """Where a whole-axis reference lands, or ``None`` when it is gone.

        Excel shrinks these the same way it shrinks a range: deleting rows
        2 to 4 turns ``1:6`` into ``1:3`` and ``2:4`` into ``#REF!``.
        """
        at = self.rows_at if span.is_row else self.columns_at
        count = self.row_count if span.is_row else self.column_count
        surviving = self._surviving_span(span.low, span.high, at, count)
        if surviving is None:
            return None
        low, high = surviving
        if low == span.low and high == span.high:
            return span
        return span.with_span(low, high)

    @classmethod
    def _surviving_span(
        cls, low: int, high: int, at: int | None, count: int
    ) -> tuple[int, int] | None:
        """One axis of a block after the deletion, or ``None`` if it is gone.

        An end that was itself deleted collapses onto the edge of what
        remains, which is how a range shrinks rather than breaking.
        """
        if at is None or not count:
            return (low, high)
        if at <= low and high < at + count:
            return None
        new_low = cls._survivor(low, at, count)
        if new_low is None:
            new_low = at
        new_high = cls._survivor(high, at, count)
        if new_high is None:
            new_high = at - 1
        if new_high < new_low or new_high < 1:
            return None
        return (new_low, new_high)


def delete_in_formula(
    formula: str,
    deletion: Deletion,
    *,
    formula_sheet: str,
    target_sheet: str,
) -> str:
    """Rewrite a formula for a deletion, turning what is gone into ``#REF!``.

    A range is handled as one thing, because ``SUM(A1:A10)`` over a deletion
    of rows 3 to 4 becomes ``SUM(A1:A8)`` rather than an error: only a range
    with nothing left becomes ``#REF!``. Single references become ``#REF!``
    the moment their own cell goes.
    """
    if deletion.is_empty:
        return formula
    if formula_sheet != target_sheet and "!" not in formula:
        return formula

    tokens = tokenize(formula)
    rebuilt: list[Token] = []
    index = 0
    changed = False

    while index < len(tokens):
        token = tokens[index]

        if token.kind is TokenKind.AXIS and isinstance(token.value, AxisRef):
            addresses = token.sheet if token.sheet is not None else formula_sheet
            if addresses != target_sheet:
                rebuilt.append(token)
            else:
                moved_span = deletion.moved_axis(token.value)
                if moved_span is None:
                    rebuilt.append(Token(TokenKind.TEXT, REF_ERROR))
                    changed = True
                elif moved_span != token.value:
                    rebuilt.append(Token(TokenKind.AXIS, moved_span.a1, moved_span))
                    changed = True
                else:
                    rebuilt.append(token)
            index += 1
            continue

        if token.kind is not TokenKind.REFERENCE or not isinstance(token.value, CellRef):
            rebuilt.append(token)
            index += 1
            continue

        addresses = token.sheet if token.sheet is not None else formula_sheet
        separator = tokens[index + 1] if index + 1 < len(tokens) else None
        far = tokens[index + 2] if index + 2 < len(tokens) else None
        is_range = (
            separator is not None
            and separator.kind is TokenKind.TEXT
            and separator.raw.strip() == ":"
            and far is not None
            and far.kind is TokenKind.REFERENCE
            and isinstance(far.value, CellRef)
        )

        if addresses != target_sheet:
            rebuilt.append(token)
            index += 3 if is_range else 1
            if is_range and separator is not None and far is not None:
                rebuilt.extend((separator, far))
            continue

        if is_range and far is not None and isinstance(far.value, CellRef):
            block = RangeRef(token.value, far.value)
            moved_block = deletion.moved_range(block)
            if moved_block is None:
                rebuilt.append(Token(TokenKind.TEXT, REF_ERROR))
            else:
                # A range that shrank to a single cell keeps its range shape:
                # Excel writes SUM(A1:A1), not SUM(A1), and a formula whose
                # argument must be a range would break if it were collapsed.
                rebuilt.append(Token(TokenKind.REFERENCE, moved_block.start.a1, moved_block.start))
                rebuilt.append(Token(TokenKind.TEXT, ":"))
                rebuilt.append(Token(TokenKind.REFERENCE, moved_block.end.a1, moved_block.end))
            changed = changed or (moved_block is None or moved_block.a1 != block.a1)
            index += 3
            continue

        moved = deletion.moved(token.value)
        if moved is None:
            rebuilt.append(Token(TokenKind.TEXT, REF_ERROR))
            changed = True
        elif moved != token.value:
            rebuilt.append(Token(TokenKind.REFERENCE, moved.a1, moved))
            changed = True
        else:
            rebuilt.append(token)
        index += 1

    return render(rebuilt) if changed else formula


def sorted_formula(formula: str, rows: int) -> str:
    """``formula`` as Excel's sort writes it when it moves the formula's row
    ``rows`` down, or up when negative.

    Measured in Excel's sort: a reference to the formula's own sheet moves
    as a copy moves it, a relative row following the formula, and one moved
    off the sheet becomes ``#REF!``, a range with it whole, so ``=B1+B2``
    moved three rows up is ``=#REF!+#REF!``. A reference that names a sheet
    stays as it is written, even one naming the formula's own sheet, and so
    does a 3D range: ``=S!B3`` on sheet S is still ``=S!B3``.

    :func:`translate_formula` reads a shared formula's followers instead,
    and leaves a reference that would leave the sheet as it is written.
    """
    if not formula or rows == 0:
        return formula
    tokens = tokenize(formula)
    rebuilt: list[Token] = []
    index = 0
    changed = False
    while index < len(tokens):
        token = tokens[index]
        if (
            token.kind is TokenKind.AXIS
            and isinstance(token.value, AxisRef)
            and token.value.is_row
            and token.sheet is None
        ):
            span = _sorted_rows(token.value, rows)
            if span is None:
                rebuilt.append(Token(TokenKind.TEXT, REF_ERROR))
            else:
                rebuilt.append(Token(TokenKind.AXIS, span.a1, span))
            changed = True
            index += 1
            continue
        if token.kind is not TokenKind.REFERENCE or not isinstance(token.value, CellRef):
            rebuilt.append(token)
            index += 1
            continue
        separator = tokens[index + 1] if index + 1 < len(tokens) else None
        far = tokens[index + 2] if index + 2 < len(tokens) else None
        ends: list[CellRef] = [token.value]
        if (
            separator is not None
            and separator.kind is TokenKind.TEXT
            and separator.raw.strip() == ":"
            and far is not None
            and far.kind is TokenKind.REFERENCE
            and isinstance(far.value, CellRef)
        ):
            ends.append(far.value)
        width = 2 * len(ends) - 1
        if token.sheet is not None:
            rebuilt.extend(tokens[index : index + width])
            index += width
            continue
        moved = [_sorted_cell(end, rows) for end in ends]
        placed = [end for end in moved if end is not None]
        if len(placed) < len(moved):
            rebuilt.append(Token(TokenKind.TEXT, REF_ERROR))
        else:
            for position, end in enumerate(_in_order(placed)):
                if position:
                    rebuilt.append(Token(TokenKind.TEXT, ":"))
                rebuilt.append(Token(TokenKind.REFERENCE, end.a1, end))
        changed = True
        index += width
    return render(rebuilt) if changed else formula


def _sorted_cell(reference: CellRef, rows: int) -> CellRef | None:
    row = reference.row if reference.absolute_row else reference.row + rows
    if not 1 <= row <= MAX_ROW:
        return None
    return CellRef(row, reference.column, reference.absolute_row, reference.absolute_column)


def _in_order(ends: list[CellRef]) -> list[CellRef]:
    """A range's ends with the top row first, as a moved range that turned
    over is written. Measured: the rows change places and the ``$`` markers
    stay where they were, so ``B$4:B5`` moved two rows up is ``B$3:B4``."""
    if len(ends) < 2:
        return ends
    start, end = ends
    if start.row <= end.row:
        return ends
    return [
        CellRef(end.row, start.column, start.absolute_row, start.absolute_column),
        CellRef(start.row, end.column, end.absolute_row, end.absolute_column),
    ]


def _sorted_rows(span: AxisRef, rows: int) -> AxisRef | None:
    """Whole rows as :func:`_in_order` moves a range: ``$4:5`` moved two
    rows up is ``$3:4``."""
    low = span.low if span.absolute_low else span.low + rows
    high = span.high if span.absolute_high else span.high + rows
    if not (1 <= low <= MAX_ROW and 1 <= high <= MAX_ROW):
        return None
    return span.with_span(min(low, high), max(low, high))


def shift_range(block: RangeRef, shift: Shift) -> RangeRef:
    """The same shift, applied to a stored range such as a merge or a table.

    Each corner moves on its own, so a block spanning the insertion point
    grows and one entirely after it slides. A corner that would leave the
    sheet is clamped to its edge rather than dropped, since a range has to
    keep two corners.
    """

    def move(reference: CellRef) -> CellRef:
        moved = shift.moved(reference)
        if moved is not None:
            return moved
        return CellRef(
            min(MAX_ROW, reference.row),
            min(MAX_COLUMN, reference.column),
            reference.absolute_row,
            reference.absolute_column,
        )

    return RangeRef(move(block.start), move(block.end))


def shared_formula_for(master: str, master_cell: CellRef, target: CellRef) -> str:
    """A follower's formula, derived from its group's master.

    The delta is the target's position relative to the cell that carries the
    text, which is what a shared formula's ``si`` group means.
    """
    return translate_formula(
        master,
        rows=target.row - master_cell.row,
        columns=target.column - master_cell.column,
    )


__all__ = [
    "REF_ERROR",
    "Deletion",
    "Shift",
    "delete_in_formula",
    "quote_sheet_name",
    "rename_sheet_in_formula",
    "shared_formula_for",
    "shift_formula",
    "shift_range",
    "sorted_formula",
    "translate_formula",
]
