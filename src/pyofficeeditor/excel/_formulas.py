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

from pyofficeeditor.excel._reference import MAX_COLUMN, MAX_ROW, CellRef, RangeRef
from pyofficeeditor.excel._tokens import TokenKind, render, tokenize

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


#: How a sheet name is written inside a formula.  It is quoted with
#: apostrophes when it is not a bare identifier, and an apostrophe inside it
#: is doubled.
_BARE_SHEET_NAME = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.\\]*$")


def quote_sheet_name(name: str) -> str:
    """A sheet name as a formula must spell it.

    Excel quotes a name that is not a bare identifier, and doubles any
    apostrophe inside it. A name that looks like a cell reference has to be
    quoted too, or ``=A1!B2`` would read as a reference to column A.
    """
    if _BARE_SHEET_NAME.match(name) and not _looks_like_a_reference(name):
        return name
    return "'" + name.replace("'", "''") + "'"


def _looks_like_a_reference(name: str) -> bool:
    try:
        CellRef.parse(name)
    except ValueError:
        return False
    return True


def rename_sheet_in_formula(formula: str, old: str, new: str) -> str:
    """``formula`` with references to one sheet repointed at another.

    A sheet reference is a name followed by ``!``, and the name may be bare
    or apostrophe-quoted. Both spellings are rewritten, and the result is
    quoted according to the new name rather than the old one, because
    renaming ``Data`` to ``Q1 Data`` turns a bare reference into a quoted
    one.

    Text is left alone: a string literal that happens to contain the sheet's
    name is not a reference to it.
    """
    if old == new:
        return formula

    replacement = quote_sheet_name(new) + "!"
    bare = quote_sheet_name(old)
    quoted = "'" + old.replace("'", "''") + "'"

    out: list[str] = []
    parts = _split_literals(formula)
    index = 0
    while index < len(parts):
        chunk, is_literal = parts[index]
        if is_literal:
            # A quoted sheet name is a literal run, and it is a reference
            # only when a '!' follows it.
            follows = parts[index + 1][0] if index + 1 < len(parts) else ""
            if chunk == quoted and follows.startswith("!"):
                out.append(replacement)
                parts[index + 1] = (follows[1:], parts[index + 1][1])
                index += 1
                continue
            out.append(chunk)
            index += 1
            continue
        out.append(_rename_bare(chunk, bare, replacement))
        index += 1
    return "".join(out)


def _rename_bare(text: str, bare: str, replacement: str) -> str:
    """Rewrite ``Name!`` where the name is unquoted.

    The boundary check keeps ``MyData!A1`` from matching a rename of
    ``Data``, which a plain substring replacement would corrupt.
    """
    if "!" not in text:
        return text
    pattern = re.compile(r"(?<![A-Za-z0-9_.'])" + re.escape(bare) + r"!")
    return pattern.sub(replacement, text)


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
    "Shift",
    "quote_sheet_name",
    "rename_sheet_in_formula",
    "shared_formula_for",
    "shift_formula",
    "shift_range",
    "translate_formula",
]
