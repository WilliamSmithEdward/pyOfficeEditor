"""Text in one cell in more than one font.

A string of runs rather than one ``<t>``, each run with its own font::

    <si><r><t>bold</t></r>
        <r><rPr><sz val="11"/><color theme="1"/><rFont val="Aptos Narrow"/>
               <family val="2"/><scheme val="minor"/></rPr>
           <t xml:space="preserve"> plain</t></r></si>

Measured, Excel writes it in a way the markup would not lead anyone to
guess. The first run carries no font of its own: Excel gives the cell the
first run's font instead, so ``bold`` above is bold because the cell is.
Every later run carries a font written out in full, size, colour, typeface,
family and scheme, even where it is only the workbook's default. A line
break inside a run is stored as CRLF. Text written here is written the same
way.

Excel reads runs other writers leave less complete by rules measured
apart from that, in a workbook whose default font differs from its cells'.
Runs with no font show in the cell's font until a run has one. After that,
a run with no font shows in the workbook's default font, the first in the
styles part, and not in the cell's. A run's font takes what it leaves
unsaid from that default font too: a colour left unsaid is automatic.
``<rPr/>`` with nothing in it counts as no font. Text is read back here
with each run's font as Excel shows it.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, replace

from pyofficeeditor._xml import Element, local_name
from pyofficeeditor.excel._formats import Font
from pyofficeeditor.excel._sharedstrings import needs_space_preserved
from pyofficeeditor.excel._xstring import decode, encode_text


@dataclass(frozen=True)
class TextRun:
    """Part of a cell's text, and the font it shows in: ``None`` for the
    cell's own font."""

    text: str
    font: Font | None = None


def read_runs(container: Element) -> list[tuple[str, Font | None]]:
    """The runs of a shared string or an inline one, each with the font it
    carries, ``None`` where it carries none. Plain text is one run, and a
    run whose ``<rPr>`` is empty carries none."""
    direct = container.child("t")
    if direct is not None:
        return [(decode(direct.text), None)]
    runs: list[tuple[str, Font | None]] = []
    for run in container.children_named("r"):
        text = "".join(decode(piece.text) for piece in run.children_named("t"))
        properties = run.child("rPr")
        if properties is None or next(properties.elements(), None) is None:
            runs.append((text, None))
        else:
            runs.append((text, Font.read(properties)))
    return runs


def shown(runs: Sequence[tuple[str, Font | None]], cell: Font, default: Font) -> list[Font]:
    """The font each run shows in: the cell's for runs with none until one
    has a font, the workbook's ``default`` for runs with none after that,
    and otherwise the run's own with what it leaves unsaid filled in."""
    fonts: list[Font] = []
    formatted = False
    for _, font in runs:
        if font is None:
            fonts.append(default if formatted else cell)
            continue
        formatted = True
        fonts.append(filled(font, default))
    return fonts


def filled(font: Font, default: Font) -> Font:
    """A run's font with the typeface and size it leaves unsaid taken from
    the workbook's default font.

    A run that names no typeface shows in the default font's, with its
    family, character set and scheme. A colour left unsaid stays unsaid,
    which is automatic, as it is for a cell.
    """
    unnamed = font.name is None
    return replace(
        font,
        name=default.name if unnamed else font.name,
        size=default.size if font.size is None else font.size,
        family=default.family if unnamed and font.family is None else font.family,
        charset=default.charset if unnamed and font.charset is None else font.charset,
        scheme=default.scheme if unnamed and font.scheme is None else font.scheme,
    )


def completed(font: Font | None, base: Font) -> Font:
    """A font asked for with what it leaves unsaid taken from the cell's,
    which is how runs are written here, and not how Excel reads them.

    A typeface named differently leaves the cell's family, character set
    and scheme behind with its own name: a theme scheme would otherwise
    put the theme's typeface back over it. On and off settings, bold and
    the like, are the run's own.
    """
    if font is None:
        return base
    renamed = font.name is not None and font.name != base.name
    return Font(
        name=font.name or base.name,
        size=font.size if font.size is not None else base.size,
        bold=font.bold,
        italic=font.italic,
        strike=font.strike,
        underline=font.underline,
        script=font.script,
        color=font.color if font.color is not None else base.color,
        family=font.family if renamed or font.family is not None else base.family,
        charset=font.charset if renamed or font.charset is not None else base.charset,
        scheme=font.scheme if renamed or font.scheme is not None else base.scheme,
        condense=font.condense,
        extend=font.extend,
        outline=font.outline,
        shadow=font.shadow,
    )


def run_properties(font: Font) -> Element:
    """A font as a run carries it: a font's own children, in its own
    order, with the typeface in ``rFont``."""
    written = font.write()
    properties = Element.create("rPr")
    for child in list(written.elements()):
        written.remove(child)
        if local_name(child.name) == "name":
            properties.append(Element.create("rFont", {"val": child.get("val") or ""}))
        else:
            properties.append(child)
    return properties


def rich_entry(runs: Sequence[tuple[str, Font | None]]) -> Element:
    """A shared string of runs, a run with no font written without one."""
    entry = Element.create("si")
    for text, font in runs:
        run = Element.create("r")
        if font is not None:
            run.append(run_properties(font))
        piece = Element.create("t")
        if needs_space_preserved(text):
            piece.set("xml:space", "preserve")
        piece.set_text(encode_text(text))
        run.append(piece)
        entry.append(run)
    return entry


__all__ = [
    "TextRun",
    "completed",
    "filled",
    "read_runs",
    "rich_entry",
    "run_properties",
    "shown",
]
