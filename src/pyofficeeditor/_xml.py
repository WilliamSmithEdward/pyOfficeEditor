"""XML that remembers the bytes it came from.

Every OOXML part this library touches is edited in place: change one cell
and the other ten thousand rows must come back exactly as Excel wrote them.
No general-purpose XML library preserves that, and the ways they fail are
not cosmetic.  Excel writes a worksheet as::

    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet
    xmlns="...main" xmlns:r="...relationships" xmlns:mc="...compatibility"
    mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="..." xmlns:xr="..." ...>

with a CRLF after the declaration, no trailing newline, and ``mc:Ignorable``
naming prefixes that are declared *after* it.  Round-tripping that through
:mod:`xml.etree.ElementTree` renames the prefixes to ``ns0``/``ns1``,
reorders the declarations, and drops the CRLF, so every save rewrites the
whole file and a one-cell diff becomes unreadable.

So parts are parsed into a tree where each node keeps the exact source text
it was cut from.  Serializing a node that was not modified copies that text.
Serializing a node that was modified rebuilds it and recurses, and its
untouched children still copy.  Editing one cell therefore rewrites that
cell's row and memcpies the rest.

An edited element's own start tag is rebuilt with single-space separators
between attributes, which is what Excel writes anyway.  Untouched elements
are byte-identical.

Security posture.  Office documents are untrusted input, so the parser is
deliberately less capable than a conforming XML processor:

- ``<!DOCTYPE`` is refused outright.  No DTD means no external entity
  resolution and no entity-expansion amplification.
- Only the five predefined entities and numeric character references are
  recognised.  Any other ``&name;`` raises rather than resolving.
- Nesting deeper than :data:`MAX_DEPTH` raises instead of exhausting the
  interpreter stack.

None of the three restricts anything OOXML is allowed to contain.
"""

from __future__ import annotations

import re
from collections.abc import Iterator, Mapping

from pyofficeeditor.exceptions import UnsupportedFormatError, XmlError

#: Deepest element nesting the parser will follow.  Real OOXML parts nest
#: around fifteen deep; the limit exists so a hostile document cannot
#: exhaust the stack.
MAX_DEPTH = 256

_NAME_START = re.compile(r"[^\s/>]")
_NAME = re.compile(r"[^\s/>=]+")

_BOM_UTF8 = b"\xef\xbb\xbf"
_BOM_UTF16_LE = b"\xff\xfe"
_BOM_UTF16_BE = b"\xfe\xff"

_PREDEFINED = {"amp": "&", "lt": "<", "gt": ">", "quot": '"', "apos": "'"}
_ENTITY = re.compile(r"&(#[0-9]+|#x[0-9A-Fa-f]+|[A-Za-z_][\w.\-]*);")
_DECLARED_ENCODING = re.compile(r"""encoding\s*=\s*['"]([^'"]+)['"]""")


def _detect_encoding(data: bytes) -> tuple[bytes, str, bytes]:
    """Split a part into (byte-order mark, codec name, remaining bytes).

    A byte-order mark decides on its own.  Without one, a UTF-16 part still
    announces itself: its first character is ``<``, so the bytes begin
    ``3C 00`` or ``00 3C``.  Everything else is read as UTF-8, and the
    declaration is checked afterwards against what was detected.

    Office does write UTF-16: a workbook's DataMashup custom XML part is
    UTF-16LE with a byte-order mark (measured on the Power Query fixture),
    even though its worksheets are UTF-8.
    """
    if data.startswith(_BOM_UTF8):
        return _BOM_UTF8, "utf-8", data[len(_BOM_UTF8) :]
    if data.startswith(_BOM_UTF16_LE):
        return _BOM_UTF16_LE, "utf-16-le", data[len(_BOM_UTF16_LE) :]
    if data.startswith(_BOM_UTF16_BE):
        return _BOM_UTF16_BE, "utf-16-be", data[len(_BOM_UTF16_BE) :]
    if data[:2] == b"<\x00":
        return b"", "utf-16-le", data
    if data[:2] == b"\x00<":
        return b"", "utf-16-be", data
    return b"", "utf-8", data


def _declared_encoding(text: str) -> str | None:
    """The encoding named in the XML declaration, if there is one."""
    if not text.startswith("<?xml"):
        return None
    end = text.find("?>")
    match = _DECLARED_ENCODING.search(text[: end + 2] if end >= 0 else text[:200])
    return match.group(1).lower() if match else None


def _family(encoding: str) -> str:
    """``'utf-8'`` or ``'utf-16'``, so a declaration and a detection can be
    compared without arguing about endianness or punctuation."""
    normalized = encoding.replace("_", "-").replace(" ", "")
    if normalized in ("utf-8", "utf8"):
        return "utf-8"
    if normalized in ("utf-16", "utf16", "utf-16-le", "utf-16le", "utf-16-be", "utf-16be", "unicode"):
        return "utf-16"
    return normalized


def decode_entities(text: str) -> str:
    """Resolve the predefined entities and numeric character references.

    Raises :class:`XmlError` for any other entity, because resolving one
    would mean processing a DTD.
    """
    if "&" not in text:
        return text

    def replace(match: re.Match[str]) -> str:
        body = match.group(1)
        if body.startswith(("#x", "#X")):
            return chr(int(body[2:], 16))
        if body.startswith("#"):
            return chr(int(body[1:], 10))
        try:
            return _PREDEFINED[body]
        except KeyError:
            raise XmlError(
                f"unknown entity '&{body};'. Only the five predefined entities and "
                f"numeric character references are resolved; this document would need a DTD."
            ) from None

    return _ENTITY.sub(replace, text)


def escape_text(value: str) -> str:
    """Escape a string for element content."""
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def escape_attribute(value: str, quote: str = '"') -> str:
    """Escape a string for an attribute value delimited by ``quote``."""
    out = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return out.replace(quote, "&quot;" if quote == '"' else "&apos;")


def local_name(name: str) -> str:
    """The part of a qualified name after the prefix."""
    _, _, local = name.rpartition(":")
    return local


def _matches(name: str, query: str) -> bool:
    """Name matching as callers mean it.

    A query carrying a prefix must match the qualified name exactly.  A
    bare query matches the local name, so ``child("sheetData")`` finds the
    element whether the producer wrote ``<sheetData>`` or
    ``<x:sheetData>``.
    """
    if ":" in query:
        return name == query
    return local_name(name) == query


class Node:
    """Anything that can sit inside an element."""

    __slots__ = ("_parent",)

    def __init__(self) -> None:
        self._parent: Element | None = None

    @property
    def parent(self) -> Element | None:
        return self._parent

    def to_xml(self) -> str:  # pragma: no cover - abstract
        raise NotImplementedError


class Raw(Node):
    """A run of source this library keeps but does not model: text,
    a comment, a processing instruction, a CDATA section."""

    __slots__ = ("kind", "raw")

    def __init__(self, raw: str, kind: str = "text") -> None:
        super().__init__()
        self.raw = raw
        self.kind = kind

    @property
    def text(self) -> str:
        """The decoded text, for a text node.  Other kinds decode to ''."""
        return decode_entities(self.raw) if self.kind == "text" else ""

    def to_xml(self) -> str:
        return self.raw

    def __repr__(self) -> str:
        head = self.raw if len(self.raw) <= 24 else self.raw[:21] + "..."
        return f"Raw({self.kind}, {head!r})"


class Attribute:
    """One attribute, with the value as written kept beside the decoded one."""

    __slots__ = ("name", "quote", "raw_value")

    def __init__(self, name: str, raw_value: str, quote: str = '"') -> None:
        self.name = name
        self.raw_value = raw_value
        self.quote = quote

    @property
    def value(self) -> str:
        return decode_entities(self.raw_value)

    def to_xml(self) -> str:
        return f"{self.name}={self.quote}{self.raw_value}{self.quote}"

    def __repr__(self) -> str:
        return f"Attribute({self.name!r}, {self.value!r})"


class Element(Node):
    """An element, addressable by name and able to reproduce its source."""

    __slots__ = ("_attrs", "_children", "_dirty", "_self_closing", "_source", "name")

    def __init__(self, name: str, *, self_closing: bool = True) -> None:
        super().__init__()
        self.name = name
        self._attrs: list[Attribute] = []
        self._children: list[Node] = []
        self._source: str | None = None
        self._dirty: bool = True
        self._self_closing: bool = self_closing

    # -- construction ---------------------------------------------------

    @classmethod
    def create(cls, name: str, attributes: Mapping[str, str] | None = None) -> Element:
        """A new element, not tied to any source."""
        element = cls(name)
        if attributes:
            for key, value in attributes.items():
                element.set(key, value)
        return element

    @classmethod
    def from_source(
        cls,
        name: str,
        attributes: list[Attribute],
        children: list[Node],
        *,
        self_closing: bool,
        source: str,
    ) -> Element:
        """An element that knows the exact text it came from.

        This is how the parser builds a tree, and it is what makes scoped
        rewriting possible: the element starts out clean, so serializing it
        copies ``source`` until something marks it modified.  ``source``
        must be the text that produced the other arguments, and the caller
        owns that invariant; :meth:`XmlDocument.parse` is the supported way
        in.
        """
        element = cls(name, self_closing=self_closing)
        element._attrs = attributes
        element._children = children
        for child in children:
            child._parent = element
        element._source = source
        element._dirty = False
        return element

    # -- dirty tracking -------------------------------------------------

    def _touch(self) -> None:
        """Mark this element and every ancestor as needing to be rebuilt.

        An ancestor's source text contains this element's, so it cannot be
        copied verbatim any more.  Its other children still can.
        """
        node: Element | None = self
        while node is not None and not node._dirty:
            node._dirty = True
            node = node._parent
        # The starting element may already have been dirty; make sure.
        self._dirty = True

    @property
    def is_modified(self) -> bool:
        return self._dirty

    @property
    def self_closing(self) -> bool:
        """Whether an element with no children writes ``<a/>`` or ``<a></a>``.

        Both are the same element; keeping the one the producer chose is
        what lets an edited element's siblings stay byte-identical.
        """
        return self._self_closing

    @self_closing.setter
    def self_closing(self, value: bool) -> None:
        if value != self._self_closing:
            self._self_closing = value
            self._touch()

    # -- attributes -----------------------------------------------------

    @property
    def attributes(self) -> dict[str, str]:
        """The attributes as a dict, in source order, values decoded."""
        return {a.name: a.value for a in self._attrs}

    def has(self, name: str) -> bool:
        return any(_matches(a.name, name) for a in self._attrs)

    def get(self, name: str, default: str | None = None) -> str | None:
        for a in self._attrs:
            if _matches(a.name, name):
                return a.value
        return default

    def set(self, name: str, value: str) -> None:
        """Set an attribute, keeping its position if it already exists.

        Setting an attribute to the value it already has changes nothing and
        does not mark the element modified, so an idempotent write does not
        cost a part its original bytes.
        """
        for a in self._attrs:
            if _matches(a.name, name):
                escaped = escape_attribute(value, a.quote)
                if escaped == a.raw_value:
                    return
                a.raw_value = escaped
                self._touch()
                return
        self._attrs.append(Attribute(name, escape_attribute(value)))
        self._touch()

    def unset(self, name: str) -> bool:
        """Remove an attribute.  Returns whether it was there."""
        for index, a in enumerate(self._attrs):
            if _matches(a.name, name):
                del self._attrs[index]
                self._touch()
                return True
        return False

    # -- children -------------------------------------------------------

    @property
    def children(self) -> tuple[Node, ...]:
        return tuple(self._children)

    def elements(self) -> Iterator[Element]:
        """Direct child elements."""
        for child in self._children:
            if isinstance(child, Element):
                yield child

    def children_named(self, name: str) -> Iterator[Element]:
        """Direct child elements matching ``name``."""
        for child in self._children:
            if isinstance(child, Element) and _matches(child.name, name):
                yield child

    def child(self, name: str) -> Element | None:
        """The first direct child element matching ``name``."""
        return next(self.children_named(name), None)

    def require(self, name: str) -> Element:
        """The first direct child matching ``name``, or raise."""
        found = self.child(name)
        if found is None:
            raise XmlError(f"<{self.name}> has no <{name}> child.")
        return found

    def descendants(self, name: str | None = None) -> Iterator[Element]:
        """Every element beneath this one, document order, self excluded."""
        for child in self._children:
            if isinstance(child, Element):
                if name is None or _matches(child.name, name):
                    yield child
                yield from child.descendants(name)

    def append(self, node: Node) -> None:
        self._adopt(node)
        self._children.append(node)
        self._self_closing = False
        self._touch()

    def insert(self, index: int, node: Node) -> None:
        self._adopt(node)
        self._children.insert(index, node)
        self._self_closing = False
        self._touch()

    def insert_before(self, reference: Node, node: Node) -> None:
        """Insert ``node`` immediately before ``reference``."""
        try:
            index = self._children.index(reference)
        except ValueError:
            raise XmlError(f"the reference node is not a child of <{self.name}>.") from None
        self.insert(index, node)

    def remove(self, node: Node) -> None:
        try:
            self._children.remove(node)
        except ValueError:
            raise XmlError(f"the node is not a child of <{self.name}>.") from None
        node._parent = None
        self._touch()

    def clear(self) -> None:
        """Drop every child, leaving the attributes alone."""
        for child in self._children:
            child._parent = None
        self._children.clear()
        self._touch()

    def _adopt(self, node: Node) -> None:
        if node is self:
            raise XmlError("an element cannot contain itself.")
        if node._parent is not None:
            node._parent.remove(node)
        node._parent = self

    # -- text -----------------------------------------------------------

    @property
    def text(self) -> str:
        """The decoded text of the direct text children, concatenated.

        Whitespace is returned exactly as stored, which is what
        ``xml:space="preserve"`` parts such as a shared string require.
        """
        return "".join(c.text for c in self._children if isinstance(c, Raw) and c.kind == "text")

    def set_text(self, value: str) -> None:
        """Replace every child with a single text node."""
        self.clear()
        if value:
            self.append(Raw(escape_text(value), "text"))
        else:
            # An element with no children keeps whichever empty form it had.
            self._touch()

    # -- serialization --------------------------------------------------

    def to_xml(self) -> str:
        if not self._dirty and self._source is not None:
            return self._source
        start = "<" + self.name
        for a in self._attrs:
            start += " " + a.to_xml()
        if not self._children:
            return start + ("/>" if self._self_closing else f"></{self.name}>")
        body = "".join(c.to_xml() for c in self._children)
        return f"{start}>{body}</{self.name}>"

    def __repr__(self) -> str:
        return f"<Element {self.name} attrs={len(self._attrs)} children={len(self._children)}>"


class XmlDocument:
    """A parsed XML part.

    ``prolog`` and ``epilog`` are the source before and after the root
    element, kept verbatim: the declaration, its line ending, and any
    comment or processing instruction outside the root.
    """

    __slots__ = ("bom", "encoding", "epilog", "prolog", "root")

    def __init__(
        self,
        root: Element,
        *,
        prolog: str = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n',
        epilog: str = "",
        encoding: str = "utf-8",
        bom: bytes = b"",
    ) -> None:
        self.root = root
        self.prolog = prolog
        self.epilog = epilog
        self.encoding = encoding
        self.bom = bom

    @classmethod
    def parse(cls, data: bytes) -> XmlDocument:
        """Parse a part.  Raises :class:`XmlError` on anything malformed."""
        bom, encoding, body = _detect_encoding(data)
        try:
            text = body.decode(encoding)
        except UnicodeDecodeError as exc:
            raise XmlError(f"the part is not valid {encoding}: {exc}") from exc

        declared = _declared_encoding(text)
        if declared is not None and _family(declared) != _family(encoding):
            raise UnsupportedFormatError(
                f"the part declares encoding {declared!r}, which is neither UTF-8 nor UTF-16. "
                f"Those are the two encodings every XML processor must support, and the two "
                f"Office writes."
            )

        parser = _Parser(text)
        prolog, root, epilog = parser.parse_document()
        return cls(root, prolog=prolog, epilog=epilog, encoding=encoding, bom=bom)

    @property
    def is_modified(self) -> bool:
        return self.root.is_modified

    def to_text(self) -> str:
        return self.prolog + self.root.to_xml() + self.epilog

    def to_bytes(self) -> bytes:
        """Serialize.  With nothing modified this reproduces the input."""
        return self.bom + self.to_text().encode(self.encoding)

    def __repr__(self) -> str:
        return f"XmlDocument(<{self.root.name}>, modified={self.is_modified})"


class _Parser:
    """A recursive-descent XML scanner that records every node's source span."""

    __slots__ = ("pos", "text")

    def __init__(self, text: str) -> None:
        self.text = text
        self.pos = 0

    # -- helpers --------------------------------------------------------

    def _fail(self, message: str, at: int | None = None) -> XmlError:
        where = self.pos if at is None else at
        line = self.text.count("\n", 0, where) + 1
        column = where - (self.text.rfind("\n", 0, where) + 1) + 1
        excerpt = self.text[where : where + 40].replace("\n", "\\n")
        return XmlError(f"{message} at line {line} column {column} (byte {where}): {excerpt!r}")

    def _skip_space(self) -> None:
        while self.pos < len(self.text) and self.text[self.pos] in " \t\r\n":
            self.pos += 1

    # -- document -------------------------------------------------------

    def parse_document(self) -> tuple[str, Element, str]:
        prolog_start = self.pos
        while True:
            index = self.text.find("<", self.pos)
            if index < 0:
                raise self._fail("the part contains no root element", prolog_start)
            self.pos = index
            if self.text.startswith("<?", index):
                self.pos = self._consume_until("?>", "an unterminated processing instruction")
            elif self.text.startswith("<!--", index):
                self.pos = self._consume_until("-->", "an unterminated comment")
            elif self.text.startswith("<!DOCTYPE", index) or self.text.startswith("<!doctype", index):
                raise self._fail(
                    "a DOCTYPE declaration is refused: this library does not process DTDs, "
                    "so no external entity can be resolved and no entity expansion can occur",
                    index,
                )
            elif self.text.startswith("<!", index):
                raise self._fail("an unexpected declaration outside the root element", index)
            else:
                break

        prolog = self.text[prolog_start : self.pos]
        root = self._parse_element(depth=0)
        epilog = self.text[self.pos :]
        return prolog, root, epilog

    def _consume_until(self, terminator: str, what: str) -> int:
        end = self.text.find(terminator, self.pos)
        if end < 0:
            raise self._fail(what)
        return end + len(terminator)

    # -- elements -------------------------------------------------------

    def _parse_element(self, depth: int) -> Element:
        if depth > MAX_DEPTH:
            raise self._fail(f"element nesting deeper than {MAX_DEPTH}, which this parser refuses")

        start = self.pos
        if self.text[self.pos] != "<":
            raise self._fail("expected an element")
        self.pos += 1
        name = self._parse_name()
        attributes = self._parse_attributes(name)

        if self.text.startswith("/>", self.pos):
            self.pos += 2
            self_closing = True
            children: list[Node] = []
        elif self.text.startswith(">", self.pos):
            self.pos += 1
            self_closing = False
            children = self._parse_children(name, depth)
        else:
            raise self._fail(f"unterminated start tag for <{name}>")

        return Element.from_source(
            name,
            attributes,
            children,
            self_closing=self_closing,
            source=self.text[start : self.pos],
        )

    def _parse_name(self) -> str:
        match = _NAME.match(self.text, self.pos)
        if not match or not _NAME_START.match(self.text, self.pos):
            raise self._fail("expected an element or attribute name")
        self.pos = match.end()
        return match.group(0)

    def _parse_attributes(self, owner: str) -> list[Attribute]:
        attributes: list[Attribute] = []
        while True:
            self._skip_space()
            if self.pos >= len(self.text):
                raise self._fail(f"unterminated start tag for <{owner}>")
            if self.text[self.pos] in ">/":
                return attributes
            name = self._parse_name()
            self._skip_space()
            if self.pos >= len(self.text) or self.text[self.pos] != "=":
                raise self._fail(f"attribute {name!r} of <{owner}> has no value")
            self.pos += 1
            self._skip_space()
            if self.pos >= len(self.text) or self.text[self.pos] not in "\"'":
                raise self._fail(f"attribute {name!r} of <{owner}> is not quoted")
            quote = self.text[self.pos]
            self.pos += 1
            end = self.text.find(quote, self.pos)
            if end < 0:
                raise self._fail(f"unterminated value for attribute {name!r}")
            attributes.append(Attribute(name, self.text[self.pos : end], quote))
            self.pos = end + 1

    def _parse_children(self, owner: str, depth: int) -> list[Node]:
        children: list[Node] = []
        while True:
            index = self.text.find("<", self.pos)
            if index < 0:
                raise self._fail(f"<{owner}> is never closed", self.pos)
            if index > self.pos:
                children.append(Raw(self.text[self.pos : index], "text"))
                self.pos = index

            if self.text.startswith("</", index):
                self.pos = index + 2
                closing = self._parse_name()
                if closing != owner:
                    raise self._fail(f"</{closing}> closes <{owner}>", index)
                self._skip_space()
                if self.pos >= len(self.text) or self.text[self.pos] != ">":
                    raise self._fail(f"unterminated end tag for <{owner}>")
                self.pos += 1
                return children

            if self.text.startswith("<!--", index):
                self.pos = self._consume_until("-->", "an unterminated comment")
                children.append(Raw(self.text[index : self.pos], "comment"))
            elif self.text.startswith("<![CDATA[", index):
                self.pos = self._consume_until("]]>", "an unterminated CDATA section")
                children.append(Raw(self.text[index : self.pos], "cdata"))
            elif self.text.startswith("<?", index):
                self.pos = self._consume_until("?>", "an unterminated processing instruction")
                children.append(Raw(self.text[index : self.pos], "pi"))
            elif self.text.startswith("<!", index):
                raise self._fail("an unexpected declaration inside an element", index)
            else:
                children.append(self._parse_element(depth + 1))


__all__ = [
    "MAX_DEPTH",
    "Attribute",
    "Element",
    "Node",
    "Raw",
    "XmlDocument",
    "decode_entities",
    "escape_attribute",
    "escape_text",
    "local_name",
]
