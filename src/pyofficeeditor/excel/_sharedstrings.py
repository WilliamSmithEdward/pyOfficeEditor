"""The shared string table.

A text cell does not hold its text. It holds an index::

    <c r="A2" t="s"><v>6</v></c>

and ``6`` indexes ``<si>`` elements in ``xl/sharedStrings.xml``. The table is
per workbook, so a string typed on the second sheet and a string typed on the
first share one entry.

Two details decide whether text survives a round trip.

**Whitespace needs declaring.** Excel writes ``<t xml:space="preserve">``
whenever the text has leading or trailing spaces (measured: the fixture's
``"  padded  "`` carries it). Without the attribute a reader is entitled to
strip, and Excel does.

**An entry can be rich text.** Instead of one ``<t>``, an ``<si>`` may hold
``<r>`` runs each with their own ``<rPr>`` formatting. Reading concatenates
the runs, which is the string the cell displays. Rewriting such an entry
would throw the formatting away, so this module never rewrites one: a
changed string gets a new entry instead. Writing one is
:mod:`pyofficeeditor.excel._richtext`'s business.
"""

from __future__ import annotations

from pyofficeeditor._xml import Element, XmlDocument
from pyofficeeditor.excel._xstring import decode, encode_text

NS_SPREADSHEETML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

CT_SHARED_STRINGS = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
)
RT_SHARED_STRINGS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
)


def needs_space_preserved(text: str) -> bool:
    """Whether a ``<t>`` holding this text must declare ``xml:space``.

    Leading or trailing white space survives only when the attribute is
    present. A line break between words needs nothing, and Excel writes
    none for one, measured.
    """
    return text != text.strip(" \t\r\n")


class SharedStrings:
    """``xl/sharedStrings.xml``, read and appended to.

    Entries are never reordered or removed, because every text cell in the
    workbook addresses them by position. A string that is no longer used
    just sits there, which is what Excel does too until it rewrites the
    file.
    """

    def __init__(self, document: XmlDocument) -> None:
        self._document = document
        self._root = document.root
        self._entries: list[Element] = list(self._root.children_named("si"))
        self._index_of: dict[str, int] = {}
        for index, entry in enumerate(self._entries):
            # Rich text is left out: plain text reusing it would show its
            # fonts. First occurrence wins: if a workbook carries the same
            # text twice, reusing the earlier index matches Excel.
            if entry.child("t") is not None:
                self._index_of.setdefault(entry_text(entry), index)

    @classmethod
    def empty(cls) -> SharedStrings:
        """A new, empty table, for a workbook that has no text yet."""
        root = Element.create("sst", {"xmlns": NS_SPREADSHEETML, "count": "0", "uniqueCount": "0"})
        return cls(XmlDocument(root))

    @property
    def document(self) -> XmlDocument:
        return self._document

    def __len__(self) -> int:
        return len(self._entries)

    def __iter__(self):  # type: ignore[no-untyped-def]
        for index in range(len(self._entries)):
            yield self[index]

    def __getitem__(self, index: int) -> str:
        """The string at an index, rich-text runs concatenated."""
        if not 0 <= index < len(self._entries):
            raise IndexError(
                f"shared string {index} is outside 0..{len(self._entries) - 1}; "
                f"the workbook's table has {len(self._entries)} entries."
            )
        return entry_text(self._entries[index])

    def is_rich_text(self, index: int) -> bool:
        """Whether the entry carries formatted runs rather than plain text."""
        if not 0 <= index < len(self._entries):
            raise IndexError(f"shared string {index} is out of range")
        return self._entries[index].child("r") is not None

    def entry(self, index: int) -> Element:
        """The ``<si>`` at an index, runs and all."""
        if not 0 <= index < len(self._entries):
            raise IndexError(f"shared string {index} is out of range")
        return self._entries[index]

    def index_for_entry(self, entry: Element) -> int:
        """The index for a whole ``<si>``, such as one of rich text runs,
        appending it unless the table has one written the same already."""
        wanted = entry.to_xml()
        for index, existing in enumerate(self._entries):
            if existing.child("r") is not None and existing.to_xml() == wanted:
                return index
        self._root.append(entry)
        index = len(self._entries)
        self._entries.append(entry)
        self._refresh_counts()
        return index

    def index_for(self, text: str) -> int:
        """The index for a string, appending an entry if it is new."""
        existing = self._index_of.get(text)
        if existing is not None:
            return existing

        entry = Element.create("si")
        run = Element.create("t")
        if needs_space_preserved(text):
            run.set("xml:space", "preserve")
        run.set_text(encode_text(text))
        entry.append(run)
        self._root.append(entry)

        index = len(self._entries)
        self._entries.append(entry)
        self._index_of[text] = index
        self._refresh_counts()
        return index

    def _refresh_counts(self) -> None:
        """Keep ``uniqueCount`` honest.

        ``count`` is the number of cells referring to the table, which this
        class cannot know on its own; the worksheet layer sets it through
        :meth:`set_reference_count`. ``uniqueCount`` is the number of
        entries, which it can.
        """
        self._root.set("uniqueCount", str(len(self._entries)))
        if self._root.get("count") is None:
            self._root.set("count", str(len(self._entries)))

    def set_reference_count(self, count: int) -> None:
        """Record how many cells point at the table."""
        self._root.set("count", str(count))

    def __repr__(self) -> str:
        return f"SharedStrings({len(self._entries)} entries)"


def entry_text(entry: Element) -> str:
    """The display text of one ``<si>``, or of an inline ``<is>``, which
    has the same shape.

    Plain text lives in a single ``<t>``. Rich text is a sequence of ``<r>``
    runs, each with its own ``<t>``, and the displayed string is those runs
    joined. A phonetic hint (``<rPh>``, used for Japanese furigana) is not
    part of the string and is skipped. Each ``<t>`` spells characters XML
    cannot hold as ``_xHHHH_``, and is read as the characters it spells.
    """
    direct = entry.child("t")
    if direct is not None:
        return decode(direct.text)
    return "".join(decode(run.text) for section in entry.children_named("r") for run in section.children_named("t"))


__all__ = [
    "CT_SHARED_STRINGS",
    "NS_SPREADSHEETML",
    "RT_SHARED_STRINGS",
    "SharedStrings",
    "entry_text",
    "needs_space_preserved",
]
