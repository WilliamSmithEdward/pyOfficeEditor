"""Defined names, and the naming rules they share with tables.

A defined name gives a formula a label::

    <definedNames>
      <definedName name="TotalUnits">Tabled!$B$2:$B$4</definedName>
      <definedName name="LocalRegion" localSheetId="0">Tabled!$A$2</definedName>
    </definedNames>

Two things about that are easy to get wrong.

**Scope is an index, not a name.** ``localSheetId="0"`` means the first sheet
in tab order, so a name scoped to a sheet breaks when sheets are reordered
unless the index moves with them. A name with no ``localSheetId`` is
workbook-wide, and the two can coexist: a workbook may have one ``Totals``
for the whole book and another for one sheet, and the sheet-scoped one wins
on that sheet.

**A name is a formula label, so it has to be usable in one.** No spaces, no
operator characters, and nothing that reads as a cell reference, or Excel
could not tell ``Q1`` the name from ``Q1`` the cell. Table names follow the
same rules, which is why the check lives here and both use it.

Excel also keeps built-in names here, spelled with an ``_xlnm.`` prefix:
``_xlnm.Print_Area``, ``_xlnm.Print_Titles``, ``_xlnm._FilterDatabase``.
They are read like any other and are not something a caller may create, since
the prefix is reserved.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from pyofficeeditor._xml import Element
from pyofficeeditor.excel._reference import CellRef
from pyofficeeditor.excel._xstring import decode, encode_attribute, encode_text

#: The longest name Excel accepts, for a defined name or a table.
MAX_NAME_LENGTH = 255
#: A name has to start with a letter, an underscore or a backslash, and carry
#: nothing a formula would read as an operator.
_VALID_NAME = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.\\]*$")
#: The prefix Excel reserves for its own names.
BUILTIN_NAME_PREFIX = "_xlnm."

#: The built-in name holding a sheet's print area.
PRINT_AREA = "_xlnm.Print_Area"

#: And the rows and columns it repeats on every page.
PRINT_TITLES = "_xlnm.Print_Titles"

#: The range a sheet's autofilter covers, which Excel defines, hidden,
#: whenever a filter goes on and keeps after it comes off.
FILTER_DATABASE = "_xlnm._FilterDatabase"

#: Every built-in name this library writes. Excel has more, and one it does
#: not know is refused rather than written under the reserved prefix.
BUILTIN_NAMES = frozenset({PRINT_AREA, PRINT_TITLES, FILTER_DATABASE})


def check_name(
    name: str,
    *,
    taken: set[str] | None = None,
    what: str = "name",
    unique_within: str = "this workbook",
) -> None:
    """Refuse a name Excel would refuse, for a defined name or a table.

    Both are labels a formula uses, so both follow the same rules and both
    make Excel reject the file rather than repair the name. ``unique_within``
    names where the collision would be, because that differs: a table name is
    unique across the workbook, while a defined name may repeat at a
    different scope.
    """
    if not name:
        raise ValueError(f"a {what} cannot be empty.")
    if len(name) > MAX_NAME_LENGTH:
        raise ValueError(
            f"{name!r} is {len(name)} characters; Excel allows at most {MAX_NAME_LENGTH}."
        )
    if name.startswith(BUILTIN_NAME_PREFIX):
        raise ValueError(
            f"{name!r} uses the {BUILTIN_NAME_PREFIX!r} prefix, which Excel reserves for its "
            f"own names such as Print_Area."
        )
    if not _VALID_NAME.match(name):
        raise ValueError(
            f"{name!r} is not a usable {what}. It must start with a letter, an underscore "
            f"or a backslash and contain no spaces or operator characters, because "
            f"formulas refer to it by name."
        )
    try:
        CellRef.parse(name)
    except ValueError:
        pass
    else:
        raise ValueError(
            f"{name!r} reads as a cell reference, so a formula could not tell it from "
            f"the cell."
        )
    if name.upper() in ("C", "R"):
        raise ValueError(f"{name!r} is reserved by R1C1 reference notation.")
    if taken and name.casefold() in {existing.casefold() for existing in taken}:
        raise ValueError(
            f"{unique_within} already has a {what} called {name!r}; they are unique "
            f"regardless of case."
        )


@dataclass(frozen=True)
class DefinedName:
    """One entry of ``<definedNames>``.

    ``scope`` is the sheet's *name* rather than the index the file stores,
    because an index silently means a different sheet once sheets move.
    """

    name: str
    refers_to: str
    scope: str | None = None
    comment: str | None = None
    hidden: bool = False

    @property
    def is_builtin(self) -> bool:
        """Whether this is one of Excel's own, such as ``Print_Area``."""
        return self.name.startswith(BUILTIN_NAME_PREFIX)

    @property
    def is_workbook_scoped(self) -> bool:
        return self.scope is None

    def __str__(self) -> str:
        where = "workbook" if self.scope is None else self.scope
        return f"{self.name} ({where}) = {self.refers_to}"


def read_defined_name(element: Element, sheet_order: list[str]) -> DefinedName | None:
    """Read one entry, resolving ``localSheetId`` to a sheet name."""
    raw_name = element.get("name")
    if raw_name is None:
        return None
    name = decode(raw_name)
    scope: str | None = None
    raw = element.get("localSheetId")
    if raw is not None:
        try:
            index = int(raw)
        except ValueError:
            index = -1
        if 0 <= index < len(sheet_order):
            scope = sheet_order[index]
    comment = element.get("comment")
    return DefinedName(
        name=name,
        refers_to=decode(element.text),
        scope=scope,
        comment=None if comment is None else decode(comment),
        hidden=element.get("hidden") in ("1", "true"),
    )


def write_defined_name(entry: DefinedName, sheet_order: list[str]) -> Element:
    """Build one entry, turning the scope's sheet name back into an index."""
    element = Element.create("definedName", {"name": encode_attribute(entry.name)})
    if entry.scope is not None:
        try:
            element.set("localSheetId", str(sheet_order.index(entry.scope)))
        except ValueError:
            raise ValueError(
                f"{entry.name!r} is scoped to a sheet called {entry.scope!r}, which this "
                f"workbook does not have."
            ) from None
    if entry.comment is not None:
        element.set("comment", encode_attribute(entry.comment))
    if entry.hidden:
        element.set("hidden", "1")
    element.set_text(encode_text(entry.refers_to))
    return element


__all__ = [
    "BUILTIN_NAMES",
    "BUILTIN_NAME_PREFIX",
    "FILTER_DATABASE",
    "MAX_NAME_LENGTH",
    "PRINT_AREA",
    "PRINT_TITLES",
    "DefinedName",
    "check_name",
    "read_defined_name",
    "write_defined_name",
]
