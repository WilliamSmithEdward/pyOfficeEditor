"""XML that reproduces the bytes it came from.

Two gates. Fidelity: every XML part of a real Office package parses and
serializes back byte for byte, and an edit rewrites only the element that
changed. Safety: a document is untrusted input, so hostile constructs are
refused rather than resolved.
"""

from __future__ import annotations

import pytest

from pyofficeeditor._xml import (
    MAX_DEPTH,
    Element,
    Raw,
    XmlDocument,
    decode_entities,
    escape_attribute,
    escape_text,
    local_name,
)
from pyofficeeditor._zip import ZipArchive
from pyofficeeditor.exceptions import UnsupportedFormatError, XmlError

SHEET_PART = "xl/worksheets/sheet1.xml"


def _xml_parts(data: bytes) -> list[tuple[str, bytes]]:
    archive = ZipArchive.from_bytes(data)
    return [
        (name, archive.read(name))
        for name in archive.names()
        if name.endswith((".xml", ".rels"))
    ]


@pytest.fixture()
def sheet(minimal_xlsm_bytes: bytes) -> bytes:
    return ZipArchive.from_bytes(minimal_xlsm_bytes).read(SHEET_PART)


class TestRoundTrip:
    def test_every_part_of_every_authored_package(
        self, excel_authored_packages: list[tuple[str, bytes]], openpyxl_xlsx_bytes: bytes
    ) -> None:
        packages = [*excel_authored_packages, ("openpyxl.xlsx", openpyxl_xlsx_bytes)]
        checked = 0
        for package_name, data in packages:
            for part_name, raw in _xml_parts(data):
                assert XmlDocument.parse(raw).to_bytes() == raw, f"{package_name}:{part_name}"
                checked += 1
        assert checked > 30, "the fixtures should cover a useful number of parts"

    def test_an_unparsed_document_reports_unmodified(self, sheet: bytes) -> None:
        assert not XmlDocument.parse(sheet).is_modified

    def test_the_utf16_part_office_writes(self, powerquery_xlsx_bytes: bytes) -> None:
        """A workbook's DataMashup part is UTF-16LE with a byte-order mark,
        even though its worksheets are UTF-8.  Both must round-trip."""
        archive = ZipArchive.from_bytes(powerquery_xlsx_bytes)
        utf16 = [
            (name, raw)
            for name, raw in ((n, archive.read(n)) for n in archive.names() if n.endswith(".xml"))
            if raw.startswith((b"\xff\xfe", b"\xfe\xff"))
        ]
        assert utf16, "the Power Query fixture should carry a UTF-16 part"
        for name, raw in utf16:
            document = XmlDocument.parse(raw)
            assert document.encoding.startswith("utf-16"), name
            assert document.bom in (b"\xff\xfe", b"\xfe\xff"), name
            assert document.to_bytes() == raw, name

    @pytest.mark.parametrize(
        ("label", "payload"),
        [
            ("utf-16le with a mark", b"\xff\xfe" + '<?xml version="1.0" encoding="utf-16"?><r a="x"/>'.encode("utf-16-le")),
            ("utf-16be with a mark", b"\xfe\xff" + '<?xml version="1.0" encoding="utf-16"?><r a="x"/>'.encode("utf-16-be")),
            ("utf-16le with no mark", '<?xml version="1.0" encoding="utf-16"?><r a="x"/>'.encode("utf-16-le")),
        ],
    )
    def test_utf16_round_trips_and_reads(self, label: str, payload: bytes) -> None:
        document = XmlDocument.parse(payload)
        assert document.root.get("a") == "x", label
        assert document.to_bytes() == payload, label

    @pytest.mark.parametrize(
        "payload",
        [
            b"<r/>",
            b"<r></r>",
            b'<?xml version="1.0"?><r/>',
            b'<?xml version="1.0"?>\r\n<r/>',
            b"<r>text</r>",
            b"<r><!-- comment --><a/><?pi here?></r>",
            b"<r><![CDATA[ raw < & > ]]></r>",
            b"<r a='single' b=\"double\"/>",
            b"<r>  leading and trailing  </r>",
            b'<?xml version="1.0"?>\n<!-- before -->\n<r/>\n<!-- after -->\n',
            b"\xef\xbb\xbf<r/>",
        ],
    )
    def test_shapes_that_must_survive(self, payload: bytes) -> None:
        assert XmlDocument.parse(payload).to_bytes() == payload


class TestDetailsGeneralLibrariesDestroy:
    def test_the_declaration_and_its_crlf_are_kept(self, sheet: bytes) -> None:
        assert XmlDocument.parse(sheet).prolog.endswith("?>\r\n")

    def test_no_trailing_newline_is_added(self, sheet: bytes) -> None:
        assert XmlDocument.parse(sheet).to_bytes().endswith(b"</worksheet>")

    def test_prefixes_are_not_renamed(self, sheet: bytes) -> None:
        assert XmlDocument.parse(sheet).root.name == "worksheet"

    def test_mc_ignorable_survives(self, sheet: bytes) -> None:
        """It names prefixes declared after it, so any reordering breaks it."""
        assert XmlDocument.parse(sheet).root.get("mc:Ignorable") == "x14ac xr xr2 xr3"

    def test_attribute_order_is_the_sources(self, sheet: bytes) -> None:
        order = list(XmlDocument.parse(sheet).root.attributes)
        assert order[:4] == ["xmlns", "xmlns:r", "xmlns:mc", "mc:Ignorable"]

    def test_the_revision_guid_survives(self, sheet: bytes) -> None:
        assert (XmlDocument.parse(sheet).root.get("xr:uid") or "").startswith("{96814566")

    def test_the_quote_character_is_kept_per_attribute(self) -> None:
        document = XmlDocument.parse(b"<r a='one' b=\"two\"/>")
        assert document.to_bytes() == b"<r a='one' b=\"two\"/>"
        document.root.set("a", "three")
        assert document.to_bytes() == b"<r a='three' b=\"two\"/>"

    def test_each_empty_tag_keeps_its_own_spelling(self) -> None:
        document = XmlDocument.parse(b"<r><a/><b></b></r>")
        document.root.require("a").set("x", "1")
        document.root.require("b").set("y", "2")
        assert document.to_bytes() == b'<r><a x="1"/><b y="2"></b></r>'


class TestScopedRewriting:
    def test_an_edit_touches_only_its_own_element(self, sheet: bytes) -> None:
        document = XmlDocument.parse(sheet)
        document.root.require("dimension").set("ref", "A1:C3")
        out = document.to_bytes()
        assert b'<dimension ref="A1:C3"/>' in out
        assert b'<pageMargins left="0.7" right="0.7"' in out, "the sibling is verbatim"
        assert out.split(b"><")[0] == sheet.split(b"><")[0], "the root start tag is verbatim"

    def test_an_edit_survives_a_reparse(self, sheet: bytes) -> None:
        document = XmlDocument.parse(sheet)
        document.root.require("dimension").set("ref", "A1:C3")
        assert XmlDocument.parse(document.to_bytes()).root.require("dimension").get("ref") == "A1:C3"

    def test_editing_a_child_marks_its_ancestors(self) -> None:
        document = XmlDocument.parse(b"<r><a><b><c/></b></a></r>")
        deep = next(document.root.descendants("c"))
        deep.set("x", "1")
        assert deep.is_modified
        assert document.root.is_modified, "an ancestor's source contains the child's"
        assert document.to_bytes() == b'<r><a><b><c x="1"/></b></a></r>'

    def test_a_clean_sibling_still_copies_when_the_parent_is_dirty(self) -> None:
        document = XmlDocument.parse(b"<r><keep a='1'   b='2'/><edit/></r>")
        document.root.require("edit").set("z", "9")
        # The odd spacing inside <keep> proves it was copied, not rebuilt.
        assert b"<keep a='1'   b='2'/>" in document.to_bytes()


class TestAttributes:
    def test_get_set_unset(self) -> None:
        element = Element.create("c", {"r": "A1"})
        assert element.get("r") == "A1"
        assert element.get("t") is None
        assert element.get("t", "n") == "n"
        element.set("t", "s")
        assert element.attributes == {"r": "A1", "t": "s"}
        assert element.unset("t") is True
        assert element.unset("t") is False
        assert element.has("r") and not element.has("t")

    def test_setting_an_existing_attribute_keeps_its_position(self) -> None:
        document = XmlDocument.parse(b'<c r="A1" s="1" t="s"/>')
        document.root.set("s", "7")
        assert document.to_bytes() == b'<c r="A1" s="7" t="s"/>'

    def test_values_are_escaped_on_write_and_decoded_on_read(self) -> None:
        element = Element.create("c")
        element.set("v", 'a & b < c > d " e')
        assert element.get("v") == 'a & b < c > d " e'
        assert "&amp;" in element.to_xml() and "&quot;" in element.to_xml()
        assert XmlDocument.parse(b"<r>" + element.to_xml().encode() + b"</r>").root.require("c").get(
            "v"
        ) == 'a & b < c > d " e'

    def test_a_single_quoted_value_escapes_the_apostrophe(self) -> None:
        document = XmlDocument.parse(b"<r a='x'/>")
        document.root.set("a", "it's")
        assert document.to_bytes() == b"<r a='it&apos;s'/>"
        assert XmlDocument.parse(document.to_bytes()).root.get("a") == "it's"


class TestNameMatching:
    def test_a_bare_query_ignores_the_producers_prefix(self) -> None:
        document = XmlDocument.parse(
            b'<x:worksheet xmlns:x="u"><x:sheetData><x:row r="1"/></x:sheetData></x:worksheet>'
        )
        assert document.root.child("sheetData") is not None
        assert next(document.root.descendants("row")).get("r") == "1"

    def test_a_prefixed_query_is_exact(self) -> None:
        document = XmlDocument.parse(b'<x:worksheet xmlns:x="u"><x:sheetData/></x:worksheet>')
        assert document.root.child("x:sheetData") is not None
        assert document.root.child("y:sheetData") is None

    def test_local_name_helper(self) -> None:
        assert local_name("x:sheetData") == "sheetData"
        assert local_name("sheetData") == "sheetData"

    def test_require_names_what_is_missing(self) -> None:
        document = XmlDocument.parse(b"<r/>")
        with pytest.raises(XmlError, match="no <sheetData> child"):
            document.root.require("sheetData")


class TestChildren:
    def test_traversal_scope(self) -> None:
        document = XmlDocument.parse(b'<r><a><c r="1"/></a><b><c r="2"/><c r="3"/></b></r>')
        assert [c.get("r") for c in document.root.descendants("c")] == ["1", "2", "3"]
        assert list(document.root.children_named("c")) == [], "direct children only"
        assert [e.name for e in document.root.elements()] == ["a", "b"]

    def test_append_insert_remove(self) -> None:
        document = XmlDocument.parse(b"<r><a/><b/><c/></r>")
        document.root.remove(document.root.require("b"))
        assert document.to_bytes() == b"<r><a/><c/></r>"
        document.root.insert_before(document.root.require("c"), Element.create("b2"))
        assert document.to_bytes() == b"<r><a/><b2/><c/></r>"
        document.root.insert(0, Element.create("first"))
        assert document.to_bytes() == b"<r><first/><a/><b2/><c/></r>"

    def test_appending_to_an_empty_element_opens_it(self) -> None:
        document = XmlDocument.parse(b"<r><a/></r>")
        document.root.require("a").append(Element.create("b"))
        assert document.to_bytes() == b"<r><a><b/></a></r>"

    def test_reparenting_detaches_from_the_old_parent(self) -> None:
        document = XmlDocument.parse(b"<r><p><x/></p><q/></r>")
        moved = document.root.require("p").require("x")
        document.root.require("q").append(moved)
        assert document.to_bytes() == b"<r><p></p><q><x/></q></r>"
        assert moved.parent is document.root.require("q")

    def test_clear_keeps_attributes(self) -> None:
        document = XmlDocument.parse(b'<r a="1"><x/><y/></r>')
        document.root.clear()
        assert document.to_bytes() == b'<r a="1"></r>'

    def test_removing_a_stranger_raises(self) -> None:
        document = XmlDocument.parse(b"<r><a/></r>")
        with pytest.raises(XmlError, match="not a child"):
            document.root.remove(Element.create("stranger"))

    def test_an_element_cannot_contain_itself(self) -> None:
        element = Element.create("r")
        with pytest.raises(XmlError, match="cannot contain itself"):
            element.append(element)


class TestText:
    def test_whitespace_is_exact(self) -> None:
        document = XmlDocument.parse(b'<si><t xml:space="preserve">  two  spaces  </t></si>')
        assert document.root.require("t").text == "  two  spaces  "

    def test_entities_decode_but_their_source_is_kept(self) -> None:
        document = XmlDocument.parse(b"<t>a &amp; b &lt; c &#65; &#x42;</t>")
        assert document.root.text == "a & b < c A B"
        assert document.to_bytes() == b"<t>a &amp; b &lt; c &#65; &#x42;</t>"

    def test_set_text_escapes(self) -> None:
        document = XmlDocument.parse(b"<t>old</t>")
        document.root.set_text("new & <shiny>")
        assert document.to_bytes() == b"<t>new &amp; &lt;shiny&gt;</t>"
        assert XmlDocument.parse(document.to_bytes()).root.text == "new & <shiny>"

    def test_set_text_replaces_every_child(self) -> None:
        document = XmlDocument.parse(b"<t><a/>keep<b/></t>")
        document.root.set_text("only")
        assert document.to_bytes() == b"<t>only</t>"

    def test_text_of_mixed_content_concatenates_direct_text_only(self) -> None:
        document = XmlDocument.parse(b"<p>one<r>inner</r>two</p>")
        assert document.root.text == "onetwo"

    def test_a_comment_is_not_text(self) -> None:
        document = XmlDocument.parse(b"<t><!-- note -->real</t>")
        assert document.root.text == "real"


class TestBuilding:
    def test_composing_a_cell_from_nothing(self) -> None:
        row = Element.create("row", {"r": "1"})
        cell = Element.create("c", {"r": "A1", "t": "n"})
        value = Element.create("v")
        value.set_text("42")
        cell.append(value)
        row.append(cell)
        assert row.to_xml() == '<row r="1"><c r="A1" t="n"><v>42</v></c></row>'

    def test_a_document_built_from_nothing_gets_a_declaration(self) -> None:
        document = XmlDocument(Element.create("Types", {"xmlns": "u"}))
        assert document.to_bytes().startswith(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n')

    def test_raw_nodes_pass_through(self) -> None:
        element = Element.create("r")
        element.append(Raw("<!-- kept -->", "comment"))
        assert element.to_xml() == "<r><!-- kept --></r>"


class TestHelpers:
    def test_escape_helpers(self) -> None:
        assert escape_text("a & b < c > d") == "a &amp; b &lt; c &gt; d"
        assert escape_attribute('say "hi"') == "say &quot;hi&quot;"
        assert escape_attribute("it's", "'") == "it&apos;s"

    def test_decode_entities_passes_plain_text_through(self) -> None:
        assert decode_entities("no entities here") == "no entities here"

    def test_decode_entities_rejects_the_unknown(self) -> None:
        with pytest.raises(XmlError, match="unknown entity"):
            decode_entities("&xxe;")


class TestUntrustedInput:
    def test_doctype_is_refused(self) -> None:
        payload = b'<?xml version="1.0"?><!DOCTYPE r [<!ENTITY x "boom">]><r>&x;</r>'
        with pytest.raises(XmlError, match="DOCTYPE"):
            XmlDocument.parse(payload)

    def test_a_billion_laughs_never_gets_the_chance(self) -> None:
        payload = (
            b'<?xml version="1.0"?>'
            b'<!DOCTYPE lolz [<!ENTITY lol "lol">'
            b'<!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">]>'
            b"<lolz>&lol2;</lolz>"
        )
        with pytest.raises(XmlError, match="DOCTYPE"):
            XmlDocument.parse(payload)

    def test_an_unknown_entity_raises_when_decoded(self) -> None:
        document = XmlDocument.parse(b"<r>&xxe;</r>")
        with pytest.raises(XmlError, match="unknown entity"):
            _ = document.root.text

    def test_nesting_past_the_limit_is_refused(self) -> None:
        payload = b"<r>" + b"<a>" * (MAX_DEPTH + 40) + b"</a>" * (MAX_DEPTH + 40) + b"</r>"
        with pytest.raises(XmlError, match="nesting deeper"):
            XmlDocument.parse(payload)

    def test_nesting_inside_the_limit_is_fine(self) -> None:
        payload = b"<r>" + b"<a>" * 100 + b"</a>" * 100 + b"</r>"
        assert XmlDocument.parse(payload).to_bytes() == payload

    def test_an_unsupported_declaration_is_refused_by_name(self) -> None:
        with pytest.raises(UnsupportedFormatError, match="neither UTF-8 nor UTF-16"):
            XmlDocument.parse(b'<?xml version="1.0" encoding="windows-1252"?><r/>')

    def test_invalid_utf8_is_reported(self) -> None:
        with pytest.raises(XmlError, match="not valid utf-8"):
            XmlDocument.parse(b"<r>\x80\x81 broken</r>")

    @pytest.mark.parametrize(
        ("label", "payload"),
        [
            ("no root", b"<!-- only a comment -->"),
            ("empty input", b""),
            ("unclosed element", b"<r><a></r>"),
            ("mismatched end tag", b"<r></q>"),
            ("never closed", b"<r><a>"),
            ("unquoted attribute", b"<r a=1/>"),
            ("valueless attribute", b"<r a/>"),
            ("unterminated attribute", b'<r a="x/>'),
            ("truncated start tag", b"<r "),
            ("unterminated comment", b"<r><!-- oops</r>"),
            ("unterminated cdata", b"<r><![CDATA[ oops </r>"),
            ("stray declaration", b"<r><!ATTLIST x></r>"),
        ],
    )
    def test_malformed_documents_are_refused(self, label: str, payload: bytes) -> None:
        with pytest.raises(XmlError):
            XmlDocument.parse(payload)

    def test_the_error_names_a_line_and_column(self) -> None:
        with pytest.raises(XmlError, match=r"line 3 column 4"):
            XmlDocument.parse(b"<r>\n  <a>\n   </b>\n</r>")
