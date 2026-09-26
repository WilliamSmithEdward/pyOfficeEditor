"""Links to other workbooks, and what Excel keeps of them.

A formula reads another workbook as ``[1]Sheet1!A1``, the number being
the link's place among the workbook part's ``<externalReferences>``, and
names one of its names as ``[1]!Rate``. Each link is a part of its own,
whose ``<externalBook>`` lists the linked workbook's sheets and names and
caches the value of every cell a formula read there, as it stood when the
workbook was last saved with the link up to date. That cache is what
Excel calculates with while the linked workbook is closed.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyofficeeditor.excel._reference import CellRef
from pyofficeeditor.excel._values import CellError
from pyofficeeditor.excel._xstring import decode

if TYPE_CHECKING:
    from pyofficeeditor._xml import Element
    from pyofficeeditor.opc import OpcPackage

RT_EXTERNAL_LINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"

Cached = str | float | bool | CellError


@dataclass(frozen=True)
class ExternalBook:
    """A linked workbook as its link caches it: its sheets in order, each
    sheet's cells by row and column, and its workbook-wide names with what
    each refers to, written as in that workbook."""

    number: int
    sheets: tuple[str, ...]
    cells: tuple[dict[tuple[int, int], Cached], ...]
    names: dict[str, str]


def read_external_books(package: OpcPackage, workbook_part: str) -> list[ExternalBook]:
    """The links a workbook lists, numbered as its formulas number them.
    A link that is not to a workbook, such as DDE or OLE, keeps its place
    and is left out."""
    references = package.xml(workbook_part).root.child("externalReferences")
    if references is None:
        return []
    relationships = {relationship.id: relationship for relationship in package.relationships(workbook_part)}
    found: list[ExternalBook] = []
    for number, reference in enumerate(references.children_named("externalReference"), start=1):
        relationship = relationships.get(reference.get("r:id") or reference.get("id") or "")
        if relationship is None or relationship.is_external or not package.has_part(relationship.target_part):
            continue
        book = package.xml(relationship.target_part).root.child("externalBook")
        if book is not None:
            found.append(_book(number, book))
    return found


def _children(parent: Element, container: str, name: str) -> list[Element]:
    found = parent.child(container)
    return [] if found is None else list(found.children_named(name))


def _book(number: int, book: Element) -> ExternalBook:
    sheets = tuple(decode(entry.get("val") or "") for entry in _children(book, "sheetNames", "sheetName"))
    cells: list[dict[tuple[int, int], Cached]] = [{} for _ in sheets]
    for sheet in _children(book, "sheetDataSet", "sheetData"):
        try:
            index = int(sheet.get("sheetId") or "")
        except ValueError:
            continue
        if not 0 <= index < len(cells):
            continue
        for row in sheet.children_named("row"):
            for cell in row.children_named("cell"):
                value = _value(cell)
                if value is None:
                    continue
                try:
                    reference = CellRef.parse(cell.get("r") or "")
                except ValueError:
                    continue
                cells[index][(reference.row, reference.column)] = value
    names = {
        decode(entry.get("name") or ""): decode(entry.get("refersTo") or "")
        for entry in _children(book, "definedNames", "definedName")
        if entry.get("sheetId") is None
    }
    return ExternalBook(number, sheets, tuple(cells), names)


def _value(cell: Element) -> Cached | None:
    """A cached cell's value: a number unless its type says text (str), a
    logical (b) or an error (e)."""
    found = cell.child("v")
    if found is None:
        return None
    raw = decode(found.text)
    kind = cell.get("t") or "n"
    if kind == "str":
        return raw
    if kind == "b":
        return raw in ("1", "true")
    if kind == "e":
        return CellError(raw)
    try:
        return float(raw)
    except ValueError:
        return None


__all__ = ["RT_EXTERNAL_LINK", "ExternalBook", "read_external_books"]
