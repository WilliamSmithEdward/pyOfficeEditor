"""A workbook: the sheets, and the parts they share.

Sheets are resolved the way Excel resolves them, through ``<sheets>`` in
``xl/workbook.xml``::

    <sheets>
      <sheet name="Data"  sheetId="1" r:id="rId1"/>
      <sheet name="Notes" sheetId="2" r:id="rId2"/>
    </sheets>

and not by indexing the worksheet relationships, whose order is the rels
part's own. Excel really does reorder them: in a workbook it authored here,
``rId2`` pointing at ``sheet2.xml`` is listed before ``rId1`` pointing at
``sheet1.xml``, so taking the first worksheet relationship gets the second
sheet.

**Changing a cell invalidates cached results.** A formula cell stores the
value it last evaluated to, so setting ``B2`` leaves ``D2``'s cached ``510``
behind, and Excel shows that stale number until something makes it
recalculate. So a workbook this library modified is saved with
``fullCalcOnLoad`` set and the ``calcChain`` part dropped, which is what
Excel itself does when it cannot trust the cache. Excel rebuilds both.
"""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

from pyofficeeditor._xml import Element
from pyofficeeditor.excel._formulas import rename_sheet_in_formula
from pyofficeeditor.excel._names import (
    DefinedName,
    check_name,
    read_defined_name,
    write_defined_name,
)
from pyofficeeditor.excel._schema import WORKBOOK_CHILD_ORDER, insert_in_schema_order
from pyofficeeditor.excel._sharedstrings import (
    CT_SHARED_STRINGS,
    RT_SHARED_STRINGS,
    SharedStrings,
)
from pyofficeeditor.excel._styles import Styles
from pyofficeeditor.excel._tables import Table, check_table_name
from pyofficeeditor.excel.worksheet import Worksheet
from pyofficeeditor.exceptions import PackageError, UnsupportedFormatError
from pyofficeeditor.opc import OpcPackage

RT_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
RT_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
RT_CALC_CHAIN = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"

CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

NS_SPREADSHEETML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_OFFICE_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

#: The longest sheet name Excel accepts.
MAX_SHEET_NAME_LENGTH = 31
#: Characters a sheet name may not contain, because they mean something in a
#: formula reference: a colon separates a range, brackets delimit a table
#: column, and the rest are path or wildcard characters.
FORBIDDEN_SHEET_NAME_CHARS = frozenset(r":\/?*[]")
#: Excel reserves this name for its change-tracking sheet.
RESERVED_SHEET_NAMES = frozenset({"history"})

#: A new worksheet part, in the order ``CT_Worksheet`` requires.  Excel adds
#: revision-tracking namespaces of its own; none is required, and leaving
#: them out keeps a sheet this library creates honest about who made it.
_NEW_WORKSHEET = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    f'<worksheet xmlns="{NS_SPREADSHEETML}" xmlns:r="{NS_OFFICE_RELATIONSHIPS}">'
    '<dimension ref="A1"/>'
    '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
    '<sheetFormatPr defaultRowHeight="14.5"/>'
    "<sheetData/>"
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    "</worksheet>"
).encode()


def check_sheet_name(name: str, *, taken: set[str] | None = None) -> None:
    """Refuse a sheet name Excel would refuse.

    Excel rejects the file rather than repairing the name, so the check
    happens here instead of at save. Uniqueness is case-insensitive: Excel
    will not have both ``Data`` and ``data``.
    """
    if not name:
        raise ValueError("a sheet name cannot be empty.")
    if len(name) > MAX_SHEET_NAME_LENGTH:
        raise ValueError(
            f"{name!r} is {len(name)} characters; Excel allows at most {MAX_SHEET_NAME_LENGTH}."
        )
    bad = sorted(FORBIDDEN_SHEET_NAME_CHARS & set(name))
    if bad:
        raise ValueError(
            f"{name!r} contains {''.join(bad)!r}, which a sheet name may not: "
            f"those characters mean something in a formula reference."
        )
    if name.startswith("'") or name.endswith("'"):
        raise ValueError(
            f"{name!r} starts or ends with an apostrophe, which is how a formula quotes a "
            f"sheet name, so Excel does not allow it."
        )
    if name.lower() in RESERVED_SHEET_NAMES:
        raise ValueError(f"{name!r} is reserved by Excel for change tracking.")
    if taken and name.casefold() in {existing.casefold() for existing in taken}:
        raise ValueError(
            f"the workbook already has a sheet called {name!r}; names are unique "
            f"regardless of case."
        )

#: Extensions whose package this class understands.  ``.xlsb`` is a ZIP too,
#: but its parts are binary records rather than XML, so it is refused by
#: name rather than misread.
SUPPORTED_SUFFIXES = frozenset({".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam"})
BINARY_SUFFIXES = frozenset({".xlsb"})
LEGACY_SUFFIXES = frozenset({".xls", ".xlt", ".xla"})


class Workbook:
    """An Excel workbook, opened from a file or from bytes.

    ::

        with Workbook.open("book.xlsx") as book:
            sheet = book["Data"]
            sheet["B2"].value = 200
            book.save()
    """

    def __init__(self, package: OpcPackage) -> None:
        self._package = package
        self._workbook_part = package.main_document_part()
        self._document = package.xml(self._workbook_part)
        self._sheets: dict[str, Worksheet] = {}
        self._order: list[str] = []
        self._shared_strings: SharedStrings | None = None
        self._shared_strings_part: str | None = None
        self._styles: Styles | None = None
        self._changed = False
        self._load_sheet_index()

    # ------------------------------------------------------------------
    # Opening
    # ------------------------------------------------------------------

    @classmethod
    def open(cls, path: str | Path) -> Workbook:
        """Open a workbook from disk."""
        resolved = Path(path)
        _check_suffix(resolved)
        return cls(OpcPackage.open(resolved))

    @classmethod
    def from_bytes(cls, data: bytes) -> Workbook:
        """Open a workbook already in memory."""
        return cls(OpcPackage.from_bytes(data))

    @property
    def package(self) -> OpcPackage:
        """The package underneath, for anything this class does not cover."""
        return self._package

    @property
    def path(self) -> Path | None:
        return self._package.path

    # ------------------------------------------------------------------
    # Sheets
    # ------------------------------------------------------------------

    def _load_sheet_index(self) -> None:
        container = self._document.root.child("sheets")
        if container is None:
            raise PackageError(
                f"{self._workbook_part} has no <sheets> element, so the workbook declares "
                f"no worksheets. It is not a workbook this library can read."
            )
        relationships = self._package.relationships(self._workbook_part)
        for entry in container.children_named("sheet"):
            name = entry.get("name")
            relationship_id = entry.get("r:id") or entry.get("id")
            if name is None or relationship_id is None:
                continue
            target = relationships.by_id(relationship_id).target_part
            if not self._package.has_part(target):
                raise PackageError(
                    f"sheet {name!r} points at {target!r} through {relationship_id}, "
                    f"but the package has no such part."
                )
            self._order.append(name)
            self._sheets[name] = Worksheet(self, name, target, self._package.xml(target))

    @property
    def sheet_names(self) -> list[str]:
        """Sheet names in the order Excel shows their tabs."""
        return list(self._order)

    @property
    def sheets(self) -> list[Worksheet]:
        """The sheets, in tab order."""
        return [self._sheets[name] for name in self._order]

    def sheet(self, name: str) -> Worksheet:
        """One sheet, by name."""
        try:
            return self._sheets[name]
        except KeyError:
            raise KeyError(
                f"no sheet named {name!r}. The workbook has: {', '.join(self._order)}"
            ) from None

    def __getitem__(self, key: str | int) -> Worksheet:
        """``book["Data"]`` by name, or ``book[0]`` by tab position."""
        if isinstance(key, int):
            try:
                return self._sheets[self._order[key]]
            except IndexError:
                raise IndexError(
                    f"sheet {key} is out of range; the workbook has {len(self._order)}"
                ) from None
        return self.sheet(key)

    def __contains__(self, name: str) -> bool:
        return name in self._sheets

    def __iter__(self) -> Iterator[Worksheet]:
        return iter(self.sheets)

    def __len__(self) -> int:
        return len(self._order)

    @property
    def active(self) -> Worksheet:
        """The sheet Excel will show when the workbook opens."""
        view = self._document.root.child("bookViews")
        index = 0
        if view is not None:
            first = view.child("workbookView")
            if first is not None:
                raw = first.get("activeTab")
                if raw is not None:
                    try:
                        index = int(raw)
                    except ValueError:
                        index = 0
        if not 0 <= index < len(self._order):
            index = 0
        return self._sheets[self._order[index]]

    @active.setter
    def active(self, sheet: Worksheet | str) -> None:
        """Choose the sheet Excel shows when the workbook opens."""
        name = sheet if isinstance(sheet, str) else sheet.name
        if name not in self._sheets:
            raise KeyError(f"no sheet named {name!r}. The workbook has: {', '.join(self._order)}")
        self._set_active_tab(self._order.index(name))
        self.mark_changed()

    # ------------------------------------------------------------------
    # Adding, removing, renaming and reordering sheets
    # ------------------------------------------------------------------

    def add_sheet(self, name: str, index: int | None = None) -> Worksheet:
        """Add an empty worksheet, by default after the existing ones.

        Four things have to line up or Excel will not show the sheet: the
        part itself, its content-type override, a relationship from the
        workbook part, and an entry in ``<sheets>``. Missing any one of them
        produces a file that opens with the sheet silently absent, or does
        not open at all.
        """
        check_sheet_name(name, taken=set(self._order))
        position = len(self._order) if index is None else max(0, min(index, len(self._order)))

        part_name = self._free_worksheet_part_name()
        self._package.write(part_name, _NEW_WORKSHEET, content_type=CT_WORKSHEET)
        relationship = self._package.relationships(self._workbook_part).add_part(
            RT_WORKSHEET, part_name
        )

        entry = Element.create(
            "sheet",
            {"name": name, "sheetId": str(self._free_sheet_id()), "r:id": relationship.id},
        )
        container = self._document.root.require("sheets")
        existing = list(container.children_named("sheet"))
        if position < len(existing):
            container.insert_before(existing[position], entry)
        else:
            container.append(entry)

        sheet = Worksheet(self, name, part_name, self._package.xml(part_name))
        previous_order = list(self._order)
        self._order.insert(position, name)
        self._sheets[name] = sheet
        self._remap_defined_name_scopes(previous_order)
        self.mark_changed()
        return sheet

    def remove_sheet(self, name: str) -> None:
        """Delete a worksheet, its part, its relationship and its entry.

        A workbook must keep at least one sheet, so removing the last is
        refused. Formulas elsewhere that referenced the sheet are left as
        they are: Excel turns them into ``#REF!`` itself when it opens the
        file, and rewriting them here would be guessing at what the author
        wanted instead.
        """
        sheet = self.sheet(name)
        if len(self._order) == 1:
            raise ValueError(
                f"{name!r} is the only sheet; a workbook must have at least one, so Excel "
                f"would refuse the file."
            )

        container = self._document.root.require("sheets")
        for entry in list(container.children_named("sheet")):
            if entry.get("name") == name:
                relationship_id = entry.get("r:id") or entry.get("id")
                container.remove(entry)
                if relationship_id is not None:
                    self._package.relationships(self._workbook_part).remove(relationship_id)
                break

        self._package.remove_part(sheet.part_name)
        previous_order = list(self._order)
        position = self._order.index(name)
        del self._order[position]
        del self._sheets[name]
        self._remap_defined_name_scopes(previous_order)
        self._clamp_active_tab()
        self.mark_changed()

    def rename_sheet(self, old: str, new: str) -> Worksheet:
        """Rename a worksheet, repointing every reference to it.

        A sheet's name appears in more places than its ``<sheets>`` entry:
        every formula that reads from it, and every defined name scoped to
        it. Renaming only the entry leaves those pointing at a sheet that no
        longer exists, which Excel reports as ``#REF!``.

        The rewrite is textual over formula bodies, with string literals
        skipped, so a formula containing the sheet's name as text is left
        alone.
        """
        sheet = self.sheet(old)
        if old == new:
            return sheet
        check_sheet_name(new, taken={n for n in self._order if n != old})

        container = self._document.root.require("sheets")
        for entry in container.children_named("sheet"):
            if entry.get("name") == old:
                entry.set("name", new)
                break

        for other in self._sheets.values():
            _rename_in_formulas(other, old, new)
        self._rename_in_defined_names(old, new)

        position = self._order.index(old)
        self._order[position] = new
        del self._sheets[old]
        self._sheets[new] = sheet
        sheet.rename(new)
        self.mark_changed()
        return sheet

    def move_sheet(self, name: str, index: int) -> None:
        """Move a sheet's tab to a new position.

        The active tab is recorded as an index, so it is adjusted to keep
        pointing at whichever sheet was active before the move.
        """
        if name not in self._sheets:
            raise KeyError(f"no sheet named {name!r}. The workbook has: {', '.join(self._order)}")
        target = max(0, min(index, len(self._order) - 1))
        current = self._order.index(name)
        if current == target:
            return

        active = self.active.name
        container = self._document.root.require("sheets")
        entries = list(container.children_named("sheet"))
        moving = entries[current]
        container.remove(moving)

        remaining = list(container.children_named("sheet"))
        if target < len(remaining):
            container.insert_before(remaining[target], moving)
        else:
            container.append(moving)

        previous_order = list(self._order)
        self._order.insert(target, self._order.pop(current))
        self._remap_defined_name_scopes(previous_order)
        self._set_active_tab(self._order.index(active))
        self.mark_changed()

    def _free_worksheet_part_name(self) -> str:
        number = 1
        while self._package.has_part(f"xl/worksheets/sheet{number}.xml"):
            number += 1
        return f"xl/worksheets/sheet{number}.xml"

    def _free_sheet_id(self) -> int:
        used: set[int] = set()
        container = self._document.root.child("sheets")
        if container is not None:
            for entry in container.children_named("sheet"):
                raw = entry.get("sheetId")
                if raw is None:
                    continue
                try:
                    used.add(int(raw))
                except ValueError:
                    continue
        candidate = 1
        while candidate in used:
            candidate += 1
        return candidate

    def _rename_in_defined_names(self, old: str, new: str) -> None:
        container = self._document.root.child("definedNames")
        if container is None:
            return
        for entry in container.children_named("definedName"):
            text = entry.text
            if text:
                entry.set_text(rename_sheet_in_formula(text, old, new))

    def _workbook_view(self) -> Element | None:
        views = self._document.root.child("bookViews")
        return None if views is None else views.child("workbookView")

    def _set_active_tab(self, index: int) -> None:
        view = self._workbook_view()
        if view is None:
            return
        view.set("activeTab", str(index))

    def _clamp_active_tab(self) -> None:
        view = self._workbook_view()
        if view is None:
            return
        raw = view.get("activeTab")
        if raw is None:
            return
        try:
            current = int(raw)
        except ValueError:
            return
        if current >= len(self._order):
            view.set("activeTab", str(max(0, len(self._order) - 1)))

    # ------------------------------------------------------------------
    # Defined names
    # ------------------------------------------------------------------

    @property
    def defined_names(self) -> list[DefinedName]:
        """Every defined name, Excel's own built-ins included.

        Scope comes back as a sheet's *name* rather than the index the file
        stores, because that index means a different sheet as soon as sheets
        are reordered.
        """
        container = self._document.root.child("definedNames")
        if container is None:
            return []
        found: list[DefinedName] = []
        for element in container.children_named("definedName"):
            entry = read_defined_name(element, self._order)
            if entry is not None:
                found.append(entry)
        return found

    def defined_name(self, name: str, *, scope: str | None = None) -> DefinedName:
        """One defined name.

        A workbook-wide name and a sheet-scoped one can share a name, so the
        scope is part of the lookup. With no scope given, a workbook-wide
        name is preferred and a sheet-scoped one is returned only if it is
        the only match.
        """
        matches = [e for e in self.defined_names if e.name.casefold() == name.casefold()]
        if not matches:
            available = ", ".join(e.name for e in self.defined_names) or "none"
            raise KeyError(f"no defined name {name!r} in this workbook. It has: {available}")
        if scope is not None:
            for entry in matches:
                if entry.scope == scope:
                    return entry
            raise KeyError(f"no defined name {name!r} scoped to {scope!r}.")
        for entry in matches:
            if entry.is_workbook_scoped:
                return entry
        if len(matches) == 1:
            return matches[0]
        scopes = ", ".join(str(e.scope) for e in matches)
        raise KeyError(
            f"{name!r} is defined on several sheets ({scopes}) and not workbook-wide; "
            f"say which scope you mean."
        )

    def add_defined_name(
        self,
        name: str,
        refers_to: str,
        *,
        scope: str | None = None,
        comment: str | None = None,
        hidden: bool = False,
    ) -> DefinedName:
        """Define a name for a formula or a range.

        ``refers_to`` is a formula without the leading ``=``, normally a
        fully qualified range such as ``Data!$A$1:$B$4``. An unqualified
        reference is resolved by Excel against whichever sheet is active,
        which is rarely what anyone means.
        """
        if scope is not None and scope not in self._sheets:
            raise KeyError(f"no sheet named {scope!r}. The workbook has: {', '.join(self._order)}")
        taken = {e.name for e in self.defined_names if e.scope == scope}
        check_name(
            name,
            taken=taken,
            what="defined name",
            # A name may repeat at a different scope, so the message says
            # where the collision actually is.
            unique_within="this workbook" if scope is None else f"the sheet {scope!r}",
        )

        entry = DefinedName(
            name=name,
            refers_to=refers_to[1:] if refers_to.startswith("=") else refers_to,
            scope=scope,
            comment=comment,
            hidden=hidden,
        )
        container = self._document.root.child("definedNames")
        if container is None:
            container = Element.create("definedNames")
            insert_in_schema_order(self._document.root, container, WORKBOOK_CHILD_ORDER)
        element = write_defined_name(entry, self._order)
        # Excel keeps them in name order. Nothing requires it, but a diff
        # against a workbook Excel later rewrites is quieter this way.
        for existing in container.children_named("definedName"):
            if (existing.get("name") or "").casefold() > name.casefold():
                container.insert_before(existing, element)
                break
        else:
            container.append(element)
        self.mark_changed()
        return entry

    def _remap_defined_name_scopes(self, previous_order: list[str]) -> None:
        """Repoint sheet-scoped names after the sheet order changed.

        ``localSheetId`` is a position, so adding, removing or moving a sheet
        silently rescopes every name after it. Names are resolved through the
        order as it was, then rewritten against the order as it is; a name
        scoped to a sheet that is gone goes with it, which is what Excel
        does.
        """
        container = self._document.root.child("definedNames")
        if container is None:
            return
        for element in list(container.children_named("definedName")):
            raw = element.get("localSheetId")
            if raw is None:
                continue
            try:
                old_index = int(raw)
            except ValueError:
                continue
            if not 0 <= old_index < len(previous_order):
                container.remove(element)
                continue
            sheet_name = previous_order[old_index]
            if sheet_name not in self._order:
                container.remove(element)
                continue
            element.set("localSheetId", str(self._order.index(sheet_name)))
        if next(container.children_named("definedName"), None) is None:
            self._document.root.remove(container)

    def remove_defined_name(self, name: str, *, scope: str | None = None) -> None:
        """Remove a defined name.  Formulas using it will read ``#NAME?``."""
        target = self.defined_name(name, scope=scope)
        container = self._document.root.child("definedNames")
        if container is None:
            return
        for element in list(container.children_named("definedName")):
            entry = read_defined_name(element, self._order)
            if entry is None:
                continue
            if entry.name == target.name and entry.scope == target.scope:
                container.remove(element)
                break
        if next(container.children_named("definedName"), None) is None:
            self._document.root.remove(container)
        self.mark_changed()

    # ------------------------------------------------------------------
    # Tables, whose names and ids are workbook-wide
    # ------------------------------------------------------------------

    @property
    def tables(self) -> list[Table]:
        """Every table in the workbook, sheet by sheet in tab order."""
        return [table for sheet in self.sheets for table in sheet.tables]

    @property
    def table_names(self) -> list[str]:
        return [table.name for table in self.tables]

    def table(self, name: str) -> Table:
        """One table, by name, from wherever in the workbook it lives.

        Names are unique across the workbook, not per sheet, so there is no
        ambiguity to resolve.
        """
        for table in self.tables:
            if table.name.casefold() == name.casefold():
                return table
        available = ", ".join(self.table_names) or "none"
        raise KeyError(f"no table named {name!r} in this workbook. It has: {available}")

    def check_new_table_name(self, name: str) -> None:
        """Refuse a table name Excel would refuse, or one already in use."""
        check_table_name(name, taken=set(self.table_names))

    def next_table_id(self) -> int:
        """An unused table id.

        Ids are workbook-wide rather than per sheet, so allocating from one
        sheet's tables alone would collide.
        """
        used = {table.id for table in self.tables}
        candidate = 1
        while candidate in used:
            candidate += 1
        return candidate

    def free_table_part_name(self) -> str:
        number = 1
        while self._package.has_part(f"xl/tables/table{number}.xml"):
            number += 1
        return f"xl/tables/table{number}.xml"

    # ------------------------------------------------------------------
    # Shared parts
    # ------------------------------------------------------------------

    @property
    def epoch_1904(self) -> bool:
        """Whether the workbook uses the 1904 date system.

        Old Mac workbooks do. Reading a date without checking is four years
        and a day wrong.
        """
        properties = self._document.root.child("workbookPr")
        if properties is None:
            return False
        raw = properties.get("date1904") or properties.get("date1904Compat")
        return raw in ("1", "true", "on")

    @property
    def styles(self) -> Styles | None:
        """The workbook's number formats, or ``None`` if it has no styles
        part."""
        if self._styles is not None:
            return self._styles
        part = self._related_part(RT_STYLES)
        if part is None:
            return None
        self._styles = Styles(self._package.xml(part))
        return self._styles

    @property
    def shared_strings(self) -> SharedStrings | None:
        """The workbook's string table, or ``None`` if it has no such part."""
        if self._shared_strings is not None:
            return self._shared_strings
        part = self._related_part(RT_SHARED_STRINGS)
        if part is None:
            return None
        self._shared_strings_part = part
        self._shared_strings = SharedStrings(self._package.xml(part))
        return self._shared_strings

    def ensure_shared_strings(self) -> SharedStrings:
        """The string table, created if the workbook has none yet.

        A workbook with no text has no ``sharedStrings`` part. Writing the
        first string needs one, and :meth:`_flush` then adds the part, its
        content-type override and its relationship, without which Excel
        ignores the table and every text cell shows as empty.
        """
        existing = self.shared_strings
        if existing is not None:
            return existing
        created = SharedStrings.empty()
        self._shared_strings = created
        self._shared_strings_part = "xl/sharedStrings.xml"
        return created

    def _related_part(self, relationship_type: str) -> str | None:
        found = self._package.relationships(self._workbook_part).by_type(relationship_type)
        if not found:
            return None
        target = found[0].target_part
        return target if self._package.has_part(target) else None

    # ------------------------------------------------------------------
    # Saving
    # ------------------------------------------------------------------

    def mark_changed(self) -> None:
        """Record that a cell changed, and act on it now.

        The two consequences that matter for correctness are applied here
        rather than at save, so the package is consistent the moment the edit
        lands and a caller reading a part sees the same thing Excel will.
        Both are idempotent, so the cost after the first edit is a lookup.

        The shared string table's ``count`` is not done here: it is a walk
        over every cell in the workbook, which would turn one write into an
        O(n) operation. It is advisory, and :meth:`_flush` sets it.
        """
        already = self._changed
        self._changed = True
        if not already:
            self._force_recalculation()
            self._drop_calc_chain()

    @property
    def is_modified(self) -> bool:
        return self._changed

    def _flush(self) -> None:
        """Write the parts this class owns back into the package."""
        if not self._changed:
            return

        if self._shared_strings is not None and self._shared_strings_part is not None:
            self._attach_shared_strings(self._shared_strings_part)
            self._count_string_references()

        # Both are idempotent and normally already done by mark_changed;
        # repeated here because a caller may have edited the package
        # directly, underneath this class.
        self._force_recalculation()
        self._drop_calc_chain()

    def _attach_shared_strings(self, part_name: str) -> None:
        """Make sure a newly created shared string table is in the package.

        A table that came out of the file is already a part and already
        related; one this library created needs the part, the content-type
        override and the relationship, or Excel ignores it and every text
        cell shows as empty.
        """
        assert self._shared_strings is not None
        if self._package.has_part(part_name):
            return
        self._package.write(
            part_name,
            self._shared_strings.document.to_bytes(),
            content_type=CT_SHARED_STRINGS,
        )
        relationships = self._package.relationships(self._workbook_part)
        if not relationships.by_type(RT_SHARED_STRINGS):
            relationships.add_part(RT_SHARED_STRINGS, part_name)

    def _count_string_references(self) -> None:
        """Set ``count`` on the table to the number of cells using it."""
        assert self._shared_strings is not None
        total = 0
        for sheet in self._sheets.values():
            data = sheet.document.root.child("sheetData")
            if data is None:
                continue
            for row in data.children_named("row"):
                for cell in row.children_named("c"):
                    if (cell.get("t") or "") == "s":
                        total += 1
        self._shared_strings.set_reference_count(total)

    def _force_recalculation(self) -> None:
        """Tell Excel to recalculate every formula when it opens the file.

        A formula cell caches its last result. Change one of its inputs and
        that cache is wrong, and Excel will show the stale number unless it
        is told otherwise.
        """
        properties = self._document.root.child("calcPr")
        if properties is None:
            # Not appended: CT_Workbook is a sequence, and a workbook that
            # ends with extLst would then carry calcPr after an element it
            # has to precede, which Excel refuses rather than repairs.
            properties = Element.create("calcPr", {"calcId": "0"})
            insert_in_schema_order(self._document.root, properties, WORKBOOK_CHILD_ORDER)
        properties.set("fullCalcOnLoad", "1")

    def _drop_calc_chain(self) -> None:
        """Remove the calculation-order cache.

        It lists the formula cells in dependency order. Once a cell has
        changed the order can be wrong, and a wrong calcChain is one of the
        things that makes Excel offer to repair a file. Excel rebuilds it.
        """
        part = self._related_part(RT_CALC_CHAIN)
        if part is None:
            return
        relationships = self._package.relationships(self._workbook_part)
        for relationship in relationships.by_type(RT_CALC_CHAIN):
            relationships.remove(relationship.id)
        self._package.remove_part(part)

    def to_bytes(self) -> bytes:
        """The whole workbook.  Unchanged, this reproduces the input bytes."""
        self._flush()
        return self._package.to_bytes()

    def save(self, path: str | Path | None = None) -> Path:
        """Write the workbook out, defaulting to where it was opened from."""
        if path is not None:
            _check_suffix(Path(path))
        self._flush()
        return self._package.save(path)

    # ------------------------------------------------------------------
    # Context manager
    # ------------------------------------------------------------------

    def close(self) -> None:
        """Drop the in-memory state.  Does not save."""
        self._package.close()

    def __enter__(self) -> Workbook:
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    def __repr__(self) -> str:
        where = self.path.name if self.path is not None else "<bytes>"
        return f"Workbook({where!r}, sheets={self._order})"


def _rename_in_formulas(sheet: Worksheet, old: str, new: str) -> None:
    """Repoint every ``<f>`` in one sheet from one sheet name to another.

    This works on the element rather than through the cell API because a
    shared formula's followers carry no text: rewriting the master's text is
    enough, and going through ``Cell.formula`` would write each follower its
    own copy and break the group.
    """
    data = sheet.document.root.child("sheetData")
    if data is None:
        return
    for row in data.children_named("row"):
        for cell in row.children_named("c"):
            formula = cell.child("f")
            if formula is None:
                continue
            text = formula.text
            if not text:
                continue
            updated = rename_sheet_in_formula(text, old, new)
            if updated != text:
                formula.set_text(updated)


def _check_suffix(path: Path) -> None:
    suffix = path.suffix.lower()
    if suffix in SUPPORTED_SUFFIXES:
        return
    if suffix in BINARY_SUFFIXES:
        raise UnsupportedFormatError(
            f"{path.name} is a binary workbook. Its worksheets are BIFF12 records rather "
            f"than XML, so this class cannot read them; support for .xlsb is a separate job."
        )
    if suffix in LEGACY_SUFFIXES:
        raise UnsupportedFormatError(
            f"{path.name} is a legacy BIFF8 workbook, which is a compound file rather than "
            f"an OPC package. Support for .xls is a separate job."
        )
    raise UnsupportedFormatError(
        f"{path.name} is not a workbook this library opens. Supported: "
        f"{', '.join(sorted(SUPPORTED_SUFFIXES))}"
    )


__all__ = ["BINARY_SUFFIXES", "LEGACY_SUFFIXES", "SUPPORTED_SUFFIXES", "Workbook"]
