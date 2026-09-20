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
from pyofficeeditor.excel._sharedstrings import (
    CT_SHARED_STRINGS,
    RT_SHARED_STRINGS,
    SharedStrings,
)
from pyofficeeditor.excel._styles import Styles
from pyofficeeditor.excel.worksheet import Worksheet
from pyofficeeditor.exceptions import PackageError, UnsupportedFormatError
from pyofficeeditor.opc import OpcPackage

RT_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
RT_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
RT_CALC_CHAIN = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"

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
            properties = Element.create("calcPr", {"calcId": "0", "fullCalcOnLoad": "1"})
            self._document.root.append(properties)
            return
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
