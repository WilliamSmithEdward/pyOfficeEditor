"""The shared string table.

The gates are the two details that lose text silently: ``xml:space`` on
whitespace, and rich-text runs that must be read but never rewritten.
"""

from __future__ import annotations

import pytest

from pyofficeeditor._xml import XmlDocument
from pyofficeeditor.excel._sharedstrings import SharedStrings, needs_space_preserved
from pyofficeeditor.opc import OpcPackage


def table(xml: bytes) -> SharedStrings:
    return SharedStrings(XmlDocument.parse(xml))


class TestNeedsSpacePreserved:
    @pytest.mark.parametrize("text", ["  padded  ", " leading", "trailing ", "a\n", "\r\nb", " ", "\t x"])
    def test_whitespace_that_must_be_declared(self, text: str) -> None:
        assert needs_space_preserved(text) is True

    @pytest.mark.parametrize("text", ["plain", "", "two words", "a b c", "punctuation!", "a\nb", "a\r\nb"])
    def test_text_that_does_not_need_it(self, text: str) -> None:
        """A line break between words is kept without it, and Excel writes
        ``<t>head\\r\\ntwo</t>`` for one, measured."""
        assert needs_space_preserved(text) is False


class TestReading:
    def test_plain_entries(self) -> None:
        strings = table(b"<sst><si><t>one</t></si><si><t>two</t></si></sst>")
        assert len(strings) == 2
        assert strings[0] == "one"
        assert strings[1] == "two"
        assert list(strings) == ["one", "two"]

    def test_whitespace_is_returned_exactly(self) -> None:
        strings = table(b'<sst><si><t xml:space="preserve">  padded  </t></si></sst>')
        assert strings[0] == "  padded  "

    def test_entities_are_decoded(self) -> None:
        strings = table(b"<sst><si><t>a &amp; b &lt; c</t></si></sst>")
        assert strings[0] == "a & b < c"

    def test_rich_text_runs_are_concatenated(self) -> None:
        strings = table(
            b"<sst><si>"
            b"<r><rPr><b/></rPr><t>bold</t></r>"
            b"<r><t> then </t></r>"
            b"<r><rPr><i/></rPr><t>italic</t></r>"
            b"</si></sst>"
        )
        assert strings[0] == "bold then italic"
        assert strings.is_rich_text(0) is True

    def test_plain_text_is_not_rich_text(self) -> None:
        strings = table(b"<sst><si><t>plain</t></si></sst>")
        assert strings.is_rich_text(0) is False

    def test_a_phonetic_hint_is_not_part_of_the_string(self) -> None:
        """Furigana sits beside the text and is not displayed as part of it."""
        strings = table(
            b"<sst><si><t>\xe6\x9c\xb1</t><rPh sb=\"0\" eb=\"1\"><t>hint</t></rPh></si></sst>"
        )
        assert strings[0] == "朱"

    def test_an_empty_entry_reads_as_empty(self) -> None:
        assert table(b"<sst><si><t></t></si></sst>")[0] == ""

    @pytest.mark.parametrize("index", [-1, 1, 99])
    def test_an_index_out_of_range_says_how_big_the_table_is(self, index: int) -> None:
        strings = table(b"<sst><si><t>only</t></si></sst>")
        with pytest.raises(IndexError, match="1 entries"):
            strings[index]


class TestWriting:
    def test_an_existing_string_is_reused(self) -> None:
        strings = table(b"<sst><si><t>one</t></si><si><t>two</t></si></sst>")
        assert strings.index_for("two") == 1
        assert len(strings) == 2

    def test_a_new_string_is_appended(self) -> None:
        strings = table(b'<sst count="1" uniqueCount="1"><si><t>one</t></si></sst>')
        assert strings.index_for("three") == 1
        assert len(strings) == 2
        assert strings[1] == "three"

    def test_a_new_string_round_trips_through_xml(self) -> None:
        document = XmlDocument.parse(b'<sst count="0" uniqueCount="0"></sst>')
        strings = SharedStrings(document)
        strings.index_for("hello")
        again = SharedStrings(XmlDocument.parse(document.to_bytes()))
        assert again[0] == "hello"

    def test_whitespace_gets_its_attribute(self) -> None:
        document = XmlDocument.parse(b"<sst></sst>")
        strings = SharedStrings(document)
        strings.index_for("  padded  ")
        assert b'xml:space="preserve"' in document.to_bytes()
        assert SharedStrings(XmlDocument.parse(document.to_bytes()))[0] == "  padded  "

    def test_plain_text_does_not_get_the_attribute(self) -> None:
        document = XmlDocument.parse(b"<sst></sst>")
        SharedStrings(document).index_for("plain")
        assert b"xml:space" not in document.to_bytes()

    def test_special_characters_are_escaped_and_come_back(self) -> None:
        document = XmlDocument.parse(b"<sst></sst>")
        strings = SharedStrings(document)
        strings.index_for('a & b < c > d " e')
        raw = document.to_bytes()
        assert b"&amp;" in raw and b"&lt;" in raw
        assert SharedStrings(XmlDocument.parse(raw))[0] == 'a & b < c > d " e'

    def test_asking_twice_appends_once(self) -> None:
        strings = table(b"<sst></sst>")
        first = strings.index_for("same")
        second = strings.index_for("same")
        assert first == second
        assert len(strings) == 1

    def test_unique_count_is_maintained(self) -> None:
        document = XmlDocument.parse(b'<sst count="0" uniqueCount="0"></sst>')
        strings = SharedStrings(document)
        strings.index_for("a")
        strings.index_for("b")
        assert b'uniqueCount="2"' in document.to_bytes()

    def test_the_reference_count_is_the_worksheet_layers_to_set(self) -> None:
        document = XmlDocument.parse(b'<sst count="0" uniqueCount="0"></sst>')
        strings = SharedStrings(document)
        strings.index_for("a")
        strings.set_reference_count(17)
        assert b'count="17"' in document.to_bytes()

    def test_a_duplicate_entry_in_the_file_resolves_to_the_first(self) -> None:
        strings = table(b"<sst><si><t>dup</t></si><si><t>dup</t></si></sst>")
        assert strings.index_for("dup") == 0, "the earlier index is still valid"
        assert len(strings) == 2, "nothing was removed"

    def test_rich_text_is_never_rewritten_or_reused(self) -> None:
        """Changing a rich-text entry in place would discard its runs, and
        pointing plain text at one would show its fonts, so a string that
        happens to match one gets its own plain entry."""
        strings = table(b"<sst><si><r><rPr><b/></rPr><t>bold</t></r></si></sst>")
        assert strings[0] == "bold"
        assert strings.index_for("bold") == 1
        assert strings.is_rich_text(0) is True, "the runs are untouched"
        assert strings.is_rich_text(1) is False

    def test_empty_starts_usable(self) -> None:
        strings = SharedStrings.empty()
        assert len(strings) == 0
        assert strings.index_for("first") == 0
        assert b"spreadsheetml" in strings.document.to_bytes()


class TestAgainstRealExcelOutput:
    @pytest.fixture()
    def strings(self, live_sample_xlsx: object) -> SharedStrings:
        package = OpcPackage.open(str(live_sample_xlsx))
        return SharedStrings(package.xml("xl/sharedStrings.xml"))

    def test_the_table_excel_wrote(self, strings: SharedStrings) -> None:
        assert len(strings) == 14
        assert strings[0] == "Region"
        assert strings[6] == "North"

    def test_the_string_that_needed_escaping(self, strings: SharedStrings) -> None:
        assert strings[10] == 'ampersand & angle < bracket > quote " done'

    def test_the_string_that_needed_space_preserved(self, strings: SharedStrings) -> None:
        assert strings[11] == "  padded  "

    def test_a_string_from_the_second_sheet_is_in_the_same_table(self, strings: SharedStrings) -> None:
        """The table is per workbook, not per sheet."""
        assert strings[13] == "Second sheet"

    def test_appending_leaves_the_existing_entries_byte_identical(self, live_sample_xlsx: object) -> None:
        """Only two things may change: the root's counts, and a new entry at
        the end. Every existing ``<si>`` keeps its bytes."""
        package = OpcPackage.open(str(live_sample_xlsx))
        original = package.read("xl/sharedStrings.xml").decode()
        SharedStrings(package.xml("xl/sharedStrings.xml")).index_for("brand new")
        edited = package.read("xl/sharedStrings.xml").decode()

        entries = original[original.index("<si>") : original.rindex("</sst>")]
        assert entries in edited, "the existing entries are a verbatim run"
        assert edited.endswith("<si><t>brand new</t></si></sst>")
        assert 'uniqueCount="15"' in edited, "the count follows the table"
        assert 'uniqueCount="14"' in original
