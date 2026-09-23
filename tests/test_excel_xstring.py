"""Text as SpreadsheetML stores it: ``_xHHHH_`` escapes.

Every expectation here is what Excel wrote or read when measured: each
character below U+0020 and the others XML refuses, written by Excel into a
cell and into attributes, and entries written by hand that Excel then read.
"""

from __future__ import annotations

import pytest

from pyofficeeditor.excel._xstring import decode, encode_attribute, encode_text

#: What Excel wrote into a cell's text for "a" & ChrW(code) & "b".
WRITTEN_IN_TEXT = {
    **{code: f"a_x{code:04X}_b" for code in (*range(0x01, 0x09), *range(0x0B, 0x20))},
    0x09: "a\tb",
    0x0A: "a\r\nb",
    0x7F: "a\x7fb",
    0x80: "a\x80b",
    0x85: "a\x85b",
    0x9F: "a\x9fb",
    0xA0: "a\xa0b",
    0xD800: "a_xD800_b",
    0xDFFF: "a_xDFFF_b",
    0xFDD0: "a\ufdd0b",
    0xFFFE: "a_xFFFE_b",
    0xFFFF: "a_xFFFF_b",
}

#: What Excel wrote into a cell for text that already reads as an escape,
#: or nearly does.
LITERALS_IN_TEXT = {
    "_x0041_": "_x005F_x0041_",
    "_x00e9_": "_x005F_x00e9_",
    "_x00E9_": "_x005F_x00E9_",
    "_x005f_": "_x005F_x005f_",
    "a_x005F_b": "a_x005F_x005F_b",
    "__x0041_": "__x005F_x0041_",
    "_x0041__x0042_": "_x005F_x0041__x005F_x0042_",
    "_x0001_\x01": "_x005F_x0001__x0001_",
    "_X0041_": "_X0041_",
    "_x004_": "_x004_",
    "x_x0041": "x_x0041",
}

#: What Excel wrote into attributes: a validation's messages, a hyperlink's
#: tip, a table's column, a sheet's name and a defined name's comment.
WRITTEN_IN_ATTRIBUTES = {
    "er\rror": "er_x000d_ror",
    "ti\x01t": "ti_x0001_t",
    "in\nput\x01 _x0041_": "in_x000a_put_x0001_ _x005f_x0041_",
    "tip\x01\nx": "tip_x0001__x000a_x",
    "head\ntwo": "head_x000a_two",
    "_x0041_": "_x005f_x0041_",
    "a_x0041_b": "a_x005f_x0041_b",
    "c\n_x0041_": "c_x000a__x005f_x0041_",
}

#: Entries written by hand, and the characters Excel read from them.
READ_BY_EXCEL = {
    "a_x00e9_b": "a\xe9b",
    "a_x00E9_b": "a\xe9b",
    "a_X00E9_b": "a_X00E9_b",
    "a_x00E9b": "a_x00E9b",
    "a_x0000_b": "a\x00b",
    "a_xD83D__xDE00_b": "a\U0001f600b",
    "a_x000a_b": "a\nb",
    "a_x000A_b": "a\nb",
    "a_x005f_b": "a_b",
    "a_x0001_b": "a\x01b",
    "a_xFFFF_b": "a\uffffb",
    "a_xD800_b": "a\ud800b",
    "a_x00E9__x00E9_b": "a\xe9\xe9b",
    "a__x00E9_b": "a_\xe9b",
}


class TestWritingText:
    @pytest.mark.parametrize("code", sorted(WRITTEN_IN_TEXT))
    def test_each_character_as_excel_wrote_it(self, code: int) -> None:
        assert encode_text(f"a{chr(code)}b") == WRITTEN_IN_TEXT[code]

    @pytest.mark.parametrize("text", sorted(LITERALS_IN_TEXT))
    def test_what_would_read_as_an_escape_is_escaped_itself(self, text: str) -> None:
        assert encode_text(text) == LITERALS_IN_TEXT[text]

    def test_a_carriage_return_before_a_line_feed_is_still_a_character(self) -> None:
        """Measured: ``"crlf" & vbCrLf & "x"`` is stored ``crlf_x000D_`` and a
        line break, since the line break alone is what XML keeps."""
        assert encode_text("crlf\r\nx") == "crlf_x000D_\r\nx"

    def test_plain_text_is_left_alone(self) -> None:
        assert encode_text("plain text, 100% ordinary") == "plain text, 100% ordinary"


class TestWritingAttributes:
    @pytest.mark.parametrize("text", sorted(WRITTEN_IN_ATTRIBUTES))
    def test_as_excel_wrote_them(self, text: str) -> None:
        assert encode_attribute(text) == WRITTEN_IN_ATTRIBUTES[text]

    def test_a_tab_is_escaped_too(self) -> None:
        """An attribute reads a tab written as it is as a space, measured."""
        assert encode_attribute("a\tb") == "a_x0009_b"


class TestReading:
    @pytest.mark.parametrize("stored", sorted(READ_BY_EXCEL))
    def test_as_excel_read_it(self, stored: str) -> None:
        assert decode(stored) == READ_BY_EXCEL[stored]

    @pytest.mark.parametrize("code", sorted(WRITTEN_IN_TEXT))
    def test_what_excel_wrote_reads_back(self, code: int) -> None:
        # An XML reader has already turned the CRLF into one line feed.
        stored = WRITTEN_IN_TEXT[code].replace("\r\n", "\n")
        assert decode(stored) == f"a{chr(code)}b"

    @pytest.mark.parametrize("text", sorted(LITERALS_IN_TEXT))
    def test_literals_read_back(self, text: str) -> None:
        assert decode(LITERALS_IN_TEXT[text]) == text

    @pytest.mark.parametrize("text", sorted(WRITTEN_IN_ATTRIBUTES))
    def test_attributes_read_back(self, text: str) -> None:
        assert decode(WRITTEN_IN_ATTRIBUTES[text]) == text


class TestRoundTrips:
    TEXTS = (
        "",
        "plain",
        "\x00\x01\x02\t\n\r\x0b\x1f",
        "_x0041__X0041__x00e9__x005F_",
        "emoji \U0001f600 and a lone \ud800",
        "\ufffe\uffff\u2028",
        "__x0041___x0042___",
        "crlf\r\nand cr\rand lf\n",
    )

    @pytest.mark.parametrize("text", TEXTS)
    def test_text(self, text: str) -> None:
        assert decode(encode_text(text).replace("\r\n", "\n")) == text

    @pytest.mark.parametrize("text", TEXTS)
    def test_attribute(self, text: str) -> None:
        encoded = encode_attribute(text)
        assert not any(character in encoded for character in "\t\n\r")
        assert decode(encoded) == text
