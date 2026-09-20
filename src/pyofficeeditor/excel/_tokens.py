"""Breaking a formula into the pieces a transform may touch.

Two jobs need to change references inside formula text: shifting them when
rows or columns are inserted, and repointing them when a sheet is renamed.
Both did it with a regex over the whole string, and both needed the same
three exceptions to avoid corrupting a formula:

- ``LOG10(x)`` contains ``G10``, which is not a reference
- ``"A1"`` is a string, and so is the ``A1`` inside it
- ``'My Sheet A1'!B2`` has a reference inside a quoted sheet name

So the formula is tokenized once, here, and the transforms work on tokens.
That also settles a question a regex cannot answer: *which sheet* a
reference addresses. A bare ``A1`` means the sheet the formula lives on; an
``A1`` after ``Data!`` means that sheet. Shifting the rows of one sheet must
move the second and leave the first alone, and nothing but the token stream
knows the difference.

This is not a formula parser. It knows strings, sheet qualifiers and cell
references, and calls everything else text. The analyzer and linter named in
the project's goals will need a real grammar; this needs exactly enough to
avoid breaking a formula it edits.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from enum import Enum

from pyofficeeditor.excel._reference import CellRef

#: A cell reference. The lookbehind rejects a match that continues an
#: identifier, so the ``G10`` in ``LOG10`` is not one, and the lookahead
#: rejects a function call and a longer number.
_REFERENCE = re.compile(
    r"""
    (\$?)([A-Za-z]{1,3})       # optional $ then the column letters
    (\$?)([0-9]{1,7})          # optional $ then the row number
    (?![0-9(\[])               # not a longer number, a call, or a table ref
    """,
    re.VERBOSE,
)
#: An unquoted sheet name in front of ``!``.
_BARE_SHEET = re.compile(r"[A-Za-z_\\][A-Za-z0-9_.\\]*(?=!)")
#: Characters that may not sit just before a reference, because a reference
#: never continues an identifier.
_IDENTIFIER = set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_.$[")


class TokenKind(Enum):
    TEXT = "text"
    STRING = "string"
    SHEET = "sheet"
    REFERENCE = "reference"


@dataclass
class Token:
    """One piece of a formula.

    ``raw`` is always the exact source text, so a token nothing changed
    serializes back to the bytes it came from.
    """

    kind: TokenKind
    raw: str
    #: For a reference, where it points. For a sheet qualifier, the sheet's
    #: name with any quoting removed.
    value: CellRef | str | None = None
    #: For a reference, the sheet it addresses: the qualifier in front of it,
    #: or ``None`` when it is bare and so means the formula's own sheet.
    sheet: str | None = None


def tokenize(formula: str) -> list[Token]:
    """Break a formula into tokens.

    ``"".join(token.raw for token in tokenize(f)) == f`` for every input,
    which is what makes a transform that rewrites some tokens safe.
    """
    tokens: list[Token] = []
    text: list[str] = []
    position = 0

    def flush() -> None:
        if text:
            tokens.append(Token(TokenKind.TEXT, "".join(text)))
            text.clear()

    while position < len(formula):
        character = formula[position]

        if character == '"':
            flush()
            end = _end_of_quoted(formula, position, '"')
            tokens.append(Token(TokenKind.STRING, formula[position:end]))
            position = end
            continue

        if character == "'":
            end = _end_of_quoted(formula, position, "'")
            quoted = formula[position:end]
            if end < len(formula) and formula[end] == "!":
                flush()
                name = quoted[1:-1].replace("''", "'") if quoted.endswith("'") else quoted[1:]
                tokens.append(Token(TokenKind.SHEET, quoted + "!", value=name))
                position = end + 1
            else:
                flush()
                tokens.append(Token(TokenKind.STRING, quoted))
                position = end
            continue

        bare = _BARE_SHEET.match(formula, position)
        if bare and not _continues_identifier(formula, position):
            flush()
            tokens.append(Token(TokenKind.SHEET, bare.group(0) + "!", value=bare.group(0)))
            position = bare.end() + 1
            continue

        reference = _REFERENCE.match(formula, position)
        if reference and not _continues_identifier(formula, position):
            try:
                parsed = CellRef.parse(reference.group(0))
            except ValueError:
                parsed = None
            if parsed is not None:
                flush()
                tokens.append(Token(TokenKind.REFERENCE, reference.group(0), value=parsed))
                position = reference.end()
                continue

        text.append(character)
        position += 1

    flush()
    _attach_sheets(tokens)
    return tokens


def _continues_identifier(formula: str, position: int) -> bool:
    """Whether the character before ``position`` makes this a continuation.

    ``LOG10``'s ``G10`` is rejected here, and so is the ``A1`` in ``MyA1``.
    """
    return position > 0 and formula[position - 1] in _IDENTIFIER


def _end_of_quoted(formula: str, start: int, quote: str) -> int:
    """The index just past a quoted run, doubling being the escape."""
    position = start + 1
    while position < len(formula):
        if formula[position] == quote:
            if position + 1 < len(formula) and formula[position + 1] == quote:
                position += 2
                continue
            return position + 1
        position += 1
    # Unterminated: treat the rest as quoted rather than finding references
    # inside what the author meant as text.
    return len(formula)


def _attach_sheets(tokens: list[Token]) -> None:
    """Record which sheet each reference addresses.

    A qualifier applies to the reference after it, and to the far end of a
    range: in ``Data!A1:A5`` both endpoints are on ``Data``, so shifting that
    sheet's rows has to move both.
    """
    pending: str | None = None
    index = 0
    while index < len(tokens):
        token = tokens[index]

        if token.kind is TokenKind.SHEET:
            pending = token.value if isinstance(token.value, str) else None
            index += 1
            continue

        if token.kind is TokenKind.REFERENCE:
            token.sheet = pending
            separator = tokens[index + 1] if index + 1 < len(tokens) else None
            far = tokens[index + 2] if index + 2 < len(tokens) else None
            if (
                separator is not None
                and separator.kind is TokenKind.TEXT
                and separator.raw.strip() == ":"
                and far is not None
                and far.kind is TokenKind.REFERENCE
            ):
                # Both ends of a range are on the qualifier's sheet.
                far.sheet = pending
                index += 3
            else:
                index += 1
            pending = None
            continue

        if token.kind is TokenKind.TEXT and token.raw.strip() == "":
            index += 1
            continue

        pending = None
        index += 1


def render(tokens: list[Token]) -> str:
    """Put a token stream back together."""
    return "".join(token.raw for token in tokens)


__all__ = ["Token", "TokenKind", "render", "tokenize"]
