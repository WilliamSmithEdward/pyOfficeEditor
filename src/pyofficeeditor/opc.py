"""The Open Packaging Conventions layer: parts, content types, relationships.

Every modern Office file is an OPC package, so Excel, Word and PowerPoint
all sit on this module.  It owns three things and nothing above them:

- **Parts.** Named byte streams, held in a :class:`~pyofficeeditor._zip.ZipArchive`
  so an untouched part keeps its stored bytes (see that module for why the
  standard library's writer cannot).
- **Content types.** ``[Content_Types].xml``, which maps a part to its type
  either by file extension (``Default``) or by exact name (``Override``).
  A part with no resolvable content type is invisible to Office, which is
  the usual reason a hand-built package refuses to open.
- **Relationships.** The ``.rels`` part beside each part, naming what it
  points at.  Office navigates a package by relationship, not by path, so
  adding a worksheet means adding a relationship, not just a file.

Part names are stored without a leading slash, the way ZIP entry names are.
Content types and relationship overrides use the OPC spelling with the
leading slash; conversion happens at this boundary so no layer above has to
think about it.

An XML part fetched through :meth:`OpcPackage.xml` is cached, so two callers
edit the same tree, and it is flushed back into the archive on save only if
something changed.  A part nobody modified is never re-serialized and never
re-deflated.
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from pathlib import Path

from pyofficeeditor._xml import Element, XmlDocument
from pyofficeeditor._zip import ZipArchive
from pyofficeeditor.exceptions import PackageError

CONTENT_TYPES_PART = "[Content_Types].xml"
ROOT_RELS_PART = "_rels/.rels"

NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships"

#: The relationship type of a package's primary document part.
RT_OFFICE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"

_RID = re.compile(r"^rId(\d+)$")


def normalize_part_name(name: str) -> str:
    """A part name as this library stores it: no leading slash, forward
    slashes, no ``.`` or ``..`` segments left in it."""
    cleaned = name.replace("\\", "/").lstrip("/")
    if "./" not in cleaned and not cleaned.endswith("/."):
        return cleaned
    return _collapse(cleaned)


def _collapse(path: str) -> str:
    out: list[str] = []
    for segment in path.split("/"):
        if segment in ("", "."):
            continue
        if segment == "..":
            if not out:
                raise PackageError(f"the part name {path!r} climbs above the package root.")
            out.pop()
        else:
            out.append(segment)
    return "/".join(out)


def rels_part_for(part_name: str) -> str:
    """The name of the relationship part belonging to ``part_name``.

    The package root, spelled as the empty string, owns ``_rels/.rels``.
    """
    if not part_name:
        return ROOT_RELS_PART
    name = normalize_part_name(part_name)
    folder, _, leaf = name.rpartition("/")
    return f"{folder}/_rels/{leaf}.rels" if folder else f"_rels/{leaf}.rels"


def resolve_target(source_part: str, target: str) -> str:
    """Resolve a relationship target against the part that declares it.

    Targets are relative to the source part's *folder*, which is why a
    workbook's ``worksheets/sheet1.xml`` lands at ``xl/worksheets/sheet1.xml``.
    A target with a leading slash is already absolute.
    """
    if target.startswith("/"):
        return normalize_part_name(target)
    folder = source_part.rpartition("/")[0] if source_part else ""
    return _collapse(f"{folder}/{target}" if folder else target)


def relative_target(source_part: str, target_part: str) -> str:
    """The target string to write into ``source_part``'s relationships so
    that it points at ``target_part``."""
    source_folder = source_part.rpartition("/")[0] if source_part else ""
    if not source_folder:
        return target_part
    source_segments = source_folder.split("/")
    target_segments = target_part.split("/")
    shared = 0
    while (
        shared < len(source_segments)
        and shared < len(target_segments) - 1
        and source_segments[shared] == target_segments[shared]
    ):
        shared += 1
    ups = [".."] * (len(source_segments) - shared)
    return "/".join([*ups, *target_segments[shared:]])


class ContentTypes:
    """``[Content_Types].xml``, edited in place."""

    def __init__(self, document: XmlDocument) -> None:
        self._document = document
        root = document.root
        self._defaults: dict[str, Element] = {}
        self._overrides: dict[str, Element] = {}
        for entry in root.children_named("Default"):
            extension = (entry.get("Extension") or "").lower()
            if extension:
                self._defaults[extension] = entry
        for entry in root.children_named("Override"):
            part = entry.get("PartName") or ""
            if part:
                self._overrides[normalize_part_name(part)] = entry

    @property
    def defaults(self) -> dict[str, str]:
        return {ext: el.get("ContentType") or "" for ext, el in self._defaults.items()}

    @property
    def overrides(self) -> dict[str, str]:
        return {part: el.get("ContentType") or "" for part, el in self._overrides.items()}

    def of(self, part_name: str) -> str | None:
        """The content type Office will resolve for a part, or ``None`` if
        the package declares none, which makes the part invisible."""
        name = normalize_part_name(part_name)
        override = self._overrides.get(name)
        if override is not None:
            return override.get("ContentType")
        extension = name.rpartition(".")[2].lower()
        default = self._defaults.get(extension)
        return default.get("ContentType") if default is not None else None

    def set_override(self, part_name: str, content_type: str) -> None:
        name = normalize_part_name(part_name)
        existing = self._overrides.get(name)
        if existing is not None:
            existing.set("ContentType", content_type)
            return
        entry = Element.create("Override", {"PartName": "/" + name, "ContentType": content_type})
        self._document.root.append(entry)
        self._overrides[name] = entry

    def remove_override(self, part_name: str) -> bool:
        name = normalize_part_name(part_name)
        entry = self._overrides.pop(name, None)
        if entry is None:
            return False
        self._document.root.remove(entry)
        return True

    def set_default(self, extension: str, content_type: str) -> None:
        key = extension.lower().lstrip(".")
        existing = self._defaults.get(key)
        if existing is not None:
            existing.set("ContentType", content_type)
            return
        entry = Element.create("Default", {"Extension": key, "ContentType": content_type})
        # Every Default goes before the first Override, and this is not
        # cosmetic: ``CT_Types`` declares them in that order, and Excel
        # refuses a package whose Defaults come after an Override rather
        # than repairing it. Appending would be the obvious simplification
        # and would break every file this adds an extension to.
        anchor = next(self._document.root.children_named("Override"), None)
        if anchor is None:
            self._document.root.append(entry)
        else:
            self._document.root.insert_before(anchor, entry)
        self._defaults[key] = entry


class Relationship:
    """One relationship of one part."""

    __slots__ = ("_element", "_source")

    def __init__(self, element: Element, source_part: str) -> None:
        self._element = element
        self._source = source_part

    @property
    def id(self) -> str:
        return self._element.get("Id") or ""

    @property
    def type(self) -> str:
        return self._element.get("Type") or ""

    @property
    def target(self) -> str:
        """The target exactly as written in the part."""
        return self._element.get("Target") or ""

    @target.setter
    def target(self, value: str) -> None:
        self._element.set("Target", value)

    @property
    def is_external(self) -> bool:
        return (self._element.get("TargetMode") or "Internal") == "External"

    @property
    def target_part(self) -> str:
        """The target resolved to a part name.  Raises for an external
        target, which names no part in this package."""
        if self.is_external:
            raise PackageError(
                f"relationship {self.id} of {self._source or 'the package root'!r} is external "
                f"({self.target!r}); it names no part in this package."
            )
        return resolve_target(self._source, self.target)

    def __repr__(self) -> str:
        return f"Relationship({self.id!r}, {self.type.rpartition('/')[2]!r}, {self.target!r})"


class Relationships:
    """The relationships declared by one part, or by the package root."""

    def __init__(self, document: XmlDocument, source_part: str) -> None:
        self._document = document
        self._source = source_part

    @property
    def source_part(self) -> str:
        return self._source

    def __iter__(self) -> Iterator[Relationship]:
        for element in self._document.root.children_named("Relationship"):
            yield Relationship(element, self._source)

    def __len__(self) -> int:
        return sum(1 for _ in self)

    def by_id(self, relationship_id: str) -> Relationship:
        for relationship in self:
            if relationship.id == relationship_id:
                return relationship
        raise PackageError(
            f"{self._source or 'the package root'!r} has no relationship {relationship_id!r}."
        )

    def by_type(self, relationship_type: str) -> list[Relationship]:
        return [r for r in self if r.type == relationship_type]

    def one(self, relationship_type: str) -> Relationship:
        """The single relationship of a type, or raise.

        Used where the format allows exactly one, such as a package's
        primary document part.
        """
        found = self.by_type(relationship_type)
        if len(found) != 1:
            short = relationship_type.rpartition("/")[2]
            raise PackageError(
                f"{self._source or 'the package root'!r} declares {len(found)} "
                f"{short!r} relationships; the format allows exactly one."
            )
        return found[0]

    def next_id(self) -> str:
        """The lowest unused ``rIdN``."""
        used = {r.id for r in self}
        highest = 0
        for identifier in used:
            match = _RID.match(identifier)
            if match:
                highest = max(highest, int(match.group(1)))
        candidate = highest + 1
        while f"rId{candidate}" in used:
            candidate += 1
        return f"rId{candidate}"

    def add(
        self,
        relationship_type: str,
        target: str,
        *,
        external: bool = False,
        relationship_id: str | None = None,
    ) -> Relationship:
        identifier = relationship_id or self.next_id()
        if any(r.id == identifier for r in self):
            raise PackageError(
                f"{self._source or 'the package root'!r} already has a relationship {identifier!r}."
            )
        attributes = {"Id": identifier, "Type": relationship_type, "Target": target}
        if external:
            attributes["TargetMode"] = "External"
        element = Element.create("Relationship", attributes)
        self._document.root.append(element)
        return Relationship(element, self._source)

    def add_part(self, relationship_type: str, target_part: str, *, relationship_id: str | None = None) -> Relationship:
        """Add an internal relationship to a part, writing the target
        relative to this part's folder the way Office does."""
        return self.add(
            relationship_type,
            relative_target(self._source, normalize_part_name(target_part)),
            relationship_id=relationship_id,
        )

    def remove(self, relationship_id: str) -> None:
        for element in self._document.root.children_named("Relationship"):
            if (element.get("Id") or "") == relationship_id:
                self._document.root.remove(element)
                return
        raise PackageError(
            f"{self._source or 'the package root'!r} has no relationship {relationship_id!r}."
        )


class OpcPackage:
    """An OPC package held in memory.

    Use it as a context manager, or call :meth:`save` yourself::

        with OpcPackage.open("book.xlsx") as package:
            document = package.xml("xl/workbook.xml")
            ...
            package.save()
    """

    def __init__(self, archive: ZipArchive, path: Path | None = None) -> None:
        self._archive = archive
        self._path = path
        self._documents: dict[str, XmlDocument] = {}
        self._content_types: ContentTypes | None = None
        self._relationships: dict[str, Relationships] = {}
        if CONTENT_TYPES_PART not in archive:
            raise PackageError(
                f"not an OPC package: {CONTENT_TYPES_PART} is missing. "
                f"An Office file always has one."
            )

    # ------------------------------------------------------------------
    # Opening
    # ------------------------------------------------------------------

    @classmethod
    def from_bytes(cls, data: bytes, path: Path | None = None) -> OpcPackage:
        return cls(ZipArchive.from_bytes(data), path)

    @classmethod
    def open(cls, path: str | Path) -> OpcPackage:
        resolved = Path(path)
        if not resolved.is_file():
            raise PackageError(f"no such file: {resolved}")
        return cls.from_bytes(resolved.read_bytes(), resolved)

    @property
    def path(self) -> Path | None:
        """Where the package was opened from, if it came from disk."""
        return self._path

    # ------------------------------------------------------------------
    # Parts
    # ------------------------------------------------------------------

    def part_names(self) -> list[str]:
        """Every part in the package, in archive order."""
        return self._archive.names()

    def has_part(self, name: str) -> bool:
        return normalize_part_name(name) in self._archive

    def read(self, name: str) -> bytes:
        """A part's bytes.  A part fetched through :meth:`xml` and modified
        reads back as its current state."""
        part = normalize_part_name(name)
        document = self._documents.get(part)
        if document is not None and document.is_modified:
            return document.to_bytes()
        if part not in self._archive:
            raise PackageError(f"the package has no part {part!r}.")
        return self._archive.read(part)

    def write(self, name: str, data: bytes, *, content_type: str | None = None) -> None:
        """Replace or create a part.

        ``content_type`` adds the ``Override`` entry Office needs when the
        part's extension does not already resolve to the right type.  A new
        part with no resolvable content type is refused rather than written,
        because Office would silently ignore it.
        """
        part = normalize_part_name(name)
        self._documents.pop(part, None)
        self._relationships.pop(part, None)
        if content_type is not None:
            self.content_types.set_override(part, content_type)
        elif self.content_types.of(part) is None:
            raise PackageError(
                f"the package declares no content type for {part!r}. Pass content_type= "
                f"so an Override is written, or Office will ignore the part."
            )
        self._archive.write(part, data)

    def remove_part(self, name: str, *, remove_relationships: bool = True) -> None:
        """Delete a part, its content-type override, and by default its
        own ``.rels`` part."""
        part = normalize_part_name(name)
        if part not in self._archive:
            raise PackageError(f"the package has no part {part!r}.")
        self._archive.remove(part)
        self._documents.pop(part, None)
        self._relationships.pop(part, None)
        self.content_types.remove_override(part)
        if remove_relationships:
            rels = rels_part_for(part)
            if rels in self._archive:
                self._archive.remove(rels)
                self._documents.pop(rels, None)

    def xml(self, name: str) -> XmlDocument:
        """A part parsed as XML, cached so every caller edits one tree."""
        part = normalize_part_name(name)
        cached = self._documents.get(part)
        if cached is not None:
            return cached
        if part not in self._archive:
            raise PackageError(f"the package has no part {part!r}.")
        document = XmlDocument.parse(self._archive.read(part))
        self._documents[part] = document
        return document

    # ------------------------------------------------------------------
    # Content types and relationships
    # ------------------------------------------------------------------

    @property
    def content_types(self) -> ContentTypes:
        if self._content_types is None:
            self._content_types = ContentTypes(self.xml(CONTENT_TYPES_PART))
        return self._content_types

    def relationships(self, source_part: str = "") -> Relationships:
        """The relationships of a part, or of the package root.

        A part with no ``.rels`` yet gets an empty one, created in memory
        and written on save only if something is added to it.
        """
        part = normalize_part_name(source_part)
        cached = self._relationships.get(part)
        if cached is not None:
            return cached
        rels_part = rels_part_for(part)
        if rels_part in self._archive:
            document = self.xml(rels_part)
        else:
            document = XmlDocument(Element.create("Relationships", {"xmlns": NS_RELATIONSHIPS}))
            self._documents[rels_part] = document
        relationships = Relationships(document, part)
        self._relationships[part] = relationships
        return relationships

    def main_document_part(self) -> str:
        """The package's primary document part, found the way Office finds
        it: the single ``officeDocument`` relationship of the root."""
        return self.relationships().one(RT_OFFICE_DOCUMENT).target_part

    # ------------------------------------------------------------------
    # Saving
    # ------------------------------------------------------------------

    def _flush(self) -> None:
        """Write every modified XML part back into the archive."""
        # Resolving a content type parses [Content_Types].xml and so adds to
        # the cache; touch it first and iterate over a snapshot, because a
        # part created during the flush must not disturb the walk.
        content_types = self.content_types
        for part, document in list(self._documents.items()):
            if part not in self._archive:
                # A relationship part created on demand; only persist it if
                # it ended up with something in it.
                if not list(document.root.elements()):
                    continue
                if content_types.of(part) is None:
                    raise PackageError(
                        f"cannot write {part!r}: the package declares no content type for it."
                    )
                self._archive.write(part, document.to_bytes())
            elif document.is_modified:
                self._archive.write(part, document.to_bytes())

    def to_bytes(self) -> bytes:
        """The whole package.  With nothing modified this reproduces the
        bytes the package was opened from."""
        self._flush()
        return self._archive.to_bytes()

    def save(self, path: str | Path | None = None) -> Path:
        """Write the package out.  Defaults to where it was opened from."""
        destination = Path(path) if path is not None else self._path
        if destination is None:
            raise PackageError("this package was not opened from a file; pass a path to save it.")
        data = self.to_bytes()
        destination.write_bytes(data)
        self._path = destination
        return destination

    # ------------------------------------------------------------------
    # Context manager
    # ------------------------------------------------------------------

    def close(self) -> None:
        """Drop the in-memory state.  Does not save."""
        self._documents.clear()
        self._relationships.clear()
        self._content_types = None

    def __enter__(self) -> OpcPackage:
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    def __repr__(self) -> str:
        where = self._path.name if self._path is not None else "<bytes>"
        return f"OpcPackage({where!r}, {len(self._archive)} parts)"


__all__ = [
    "CONTENT_TYPES_PART",
    "NS_CONTENT_TYPES",
    "NS_RELATIONSHIPS",
    "ROOT_RELS_PART",
    "RT_OFFICE_DOCUMENT",
    "ContentTypes",
    "OpcPackage",
    "Relationship",
    "Relationships",
    "normalize_part_name",
    "relative_target",
    "rels_part_for",
    "resolve_target",
]
