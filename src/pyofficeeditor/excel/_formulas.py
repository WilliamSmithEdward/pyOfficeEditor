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

from pyofficeeditor.excel._reference import MAX_COLUMN, MAX_ROW, CellRef

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


__all__ = ["shared_formula_for", "translate_formula"]
