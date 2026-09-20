"""Differential formats, and the ``dxfs`` table they live in.

Every XML fragment quoted here was taken verbatim from a workbook real Excel
wrote, built by driving Excel through pyvbaharness. They are not what the
schema permits, they are what Excel emits, which is the only thing a reader
has to survive.
"""

from __future__ import annotations

import pytest

from pyofficeeditor._xml import XmlDocument
from pyofficeeditor.excel._dxf import Dxf, DxfFont, write_partial_border
from pyofficeeditor.excel._formats import Border, Color, Side
from pyofficeeditor.excel._styles import Styles

#: Excel's own output for a "greater than 50, pink fill with dark red bold
#: text" rule. Note bgColor, and the explicit ``i val="0"``.
EXCEL_FONT_AND_FILL = (
    b'<dxf><font><b/><i val="0"/><color rgb="FF9C0006"/></font>'
    b'<fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>'
)
#: Excel's output for a rule that sets a number format and one border edge.
EXCEL_NUMFMT_AND_BORDER = (
    b'<dxf><numFmt numFmtId="14" formatCode="0.00%"/>'
    b'<border><left style="thin"><color rgb="FFFF0000"/></left></border></dxf>'
)
#: A plain fill, which is what most rules carry.
EXCEL_FILL_ONLY = b'<dxf><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>'


def parse(raw: bytes) -> Dxf:
    return Dxf.read(XmlDocument.parse(raw).root)


def written(dxf: Dxf) -> str:
    return dxf.write().to_xml()


class TestTheFillIsBackground:
    """The trap that matters most: a dxf puts its colour in ``bgColor``
    while a cell fill uses ``fgColor``. Getting it backwards gives a rule
    that matches and highlights nothing."""

    def test_read(self) -> None:
        dxf = parse(EXCEL_FILL_ONLY)
        assert dxf.fill is not None
        assert dxf.fill.background == Color(rgb="FFC6EFCE")
        assert dxf.fill.foreground is None

    def test_written_as_background(self) -> None:
        assert "bgColor" in written(Dxf.of(fill="C6EFCE"))
        assert "fgColor" not in written(Dxf.of(fill="C6EFCE"))

    def test_no_pattern_type(self) -> None:
        """Excel writes no patternType on a dxf fill, unlike a cell fill
        where ``solid`` is what makes the colour appear."""
        assert "patternType" not in written(Dxf.of(fill="C6EFCE"))

    def test_a_six_digit_colour_gains_full_alpha(self) -> None:
        dxf = Dxf.of(fill="FF0000")
        assert dxf.fill is not None
        assert dxf.fill.background == Color(rgb="FFFF0000")


class TestFlagsAreThreeValued:
    def test_absent_means_inherit(self) -> None:
        dxf = parse(EXCEL_FILL_ONLY)
        assert dxf.font is None, "no font element at all"

    def test_present_means_on(self) -> None:
        font = parse(EXCEL_FONT_AND_FILL).font
        assert font is not None
        assert font.bold is True

    def test_explicit_zero_means_off(self) -> None:
        """Not the same as absent. ``<i val="0"/>`` clears italic on a
        matching cell; no ``<i>`` leaves whatever the cell had."""
        font = parse(EXCEL_FONT_AND_FILL).font
        assert font is not None
        assert font.italic is False
        assert font.strike is None, "never mentioned, so inherited"

    def test_off_round_trips_as_val_zero(self) -> None:
        assert '<i val="0"/>' in written(Dxf(font=DxfFont(italic=False)))

    def test_on_round_trips_bare(self) -> None:
        assert "<i/>" in written(Dxf(font=DxfFont(italic=True)))

    def test_inherit_writes_nothing(self) -> None:
        assert "<i" not in written(Dxf(font=DxfFont(bold=True)))

    def test_underline_without_val_is_single(self) -> None:
        dxf = parse(b"<dxf><font><u/></font></dxf>")
        assert dxf.font is not None
        assert dxf.font.underline == "single"


class TestNumberFormatIsInline:
    def test_the_code_is_what_counts(self) -> None:
        """numFmtId 14 is the built-in ``m/d/yyyy``, and it sits next to a
        percentage. Resolving the id against the workbook's table would
        report a date for a cell the file formats as percent."""
        assert parse(EXCEL_NUMFMT_AND_BORDER).number_format == "0.00%"

    def test_the_id_is_carried_but_not_compared(self) -> None:
        from_excel = parse(EXCEL_NUMFMT_AND_BORDER)
        assert from_excel.number_format_id == 14
        mine = Dxf(number_format="0.00%", border=from_excel.border)
        assert mine == from_excel, "the id must not split two equal formats"

    def test_an_id_is_always_written(self) -> None:
        """``CT_NumFmt`` requires the attribute. Leaving it out produces a
        dxf that does not validate."""
        assert 'numFmtId="164"' in written(Dxf(number_format="0.00%"))

    def test_a_read_id_survives_the_round_trip(self) -> None:
        assert 'numFmtId="14"' in written(parse(EXCEL_NUMFMT_AND_BORDER))


class TestBorderIsPartial:
    def test_only_the_sides_that_were_set(self) -> None:
        """A cell border always writes all five. A dxf border that did the
        same would say the rule clears every edge it never mentioned."""
        border = Border(left=Side(style="thin", color=Color(rgb="FFFF0000")))
        assert write_partial_border(border).to_xml() == (
            '<border><left style="thin"><color rgb="FFFF0000"/></left></border>'
        )

    def test_a_cell_border_writes_every_side(self) -> None:
        """The contrast, so the difference is not silently lost later."""
        border = Border(left=Side(style="thin"))
        for present in ("<right", "<top", "<bottom", "<diagonal"):
            assert present in border.write().to_xml()

    def test_read_through_the_cell_border_type(self) -> None:
        dxf = parse(EXCEL_NUMFMT_AND_BORDER)
        assert dxf.border is not None
        assert dxf.border.left.style == "thin"
        assert dxf.border.right.is_empty


class TestRoundTrip:
    @pytest.mark.parametrize(
        "raw", [EXCEL_FONT_AND_FILL, EXCEL_NUMFMT_AND_BORDER, EXCEL_FILL_ONLY]
    )
    def test_excel_output_comes_back_byte_for_byte(self, raw: bytes) -> None:
        assert written(parse(raw)) == raw.decode()

    @pytest.mark.parametrize(
        "raw", [EXCEL_FONT_AND_FILL, EXCEL_NUMFMT_AND_BORDER, EXCEL_FILL_ONLY]
    )
    def test_parsing_is_stable(self, raw: bytes) -> None:
        """Reading what was written gives the same value, or ``ensure_dxf``
        would append a duplicate every time the same rule is added."""
        once = parse(raw)
        assert parse(written(once).encode()) == once

    def test_an_empty_dxf(self) -> None:
        assert Dxf().is_empty
        assert written(Dxf()) == "<dxf/>"


class TestOf:
    def test_the_common_case(self) -> None:
        dxf = Dxf.of(fill="FFC7CE", color="9C0006", bold=True)
        assert dxf.fill is not None and dxf.fill.background == Color(rgb="FFFFC7CE")
        assert dxf.font is not None and dxf.font.bold is True
        assert dxf.font.color == Color(rgb="FF9C0006")

    def test_nothing_at_all(self) -> None:
        assert Dxf.of().is_empty

    def test_a_font_that_only_clears(self) -> None:
        dxf = Dxf.of(bold=False)
        assert dxf.font is not None and dxf.font.bold is False
        assert not dxf.is_empty, "clearing bold is a change, not an absence"


class TestTheDxfsTable:
    def sheet(self, body: bytes = b"") -> tuple[Styles, XmlDocument]:
        """The styles, and the document they were built over, so a test can
        assert on what actually lands in the part."""
        document = XmlDocument.parse(b'<styleSheet xmlns="x">' + body + b"</styleSheet>")
        return Styles(document), document

    def styles(self, body: bytes = b"") -> Styles:
        return self.sheet(body)[0]

    def test_an_empty_workbook_has_none(self) -> None:
        assert self.styles().dxf_count() == 0

    def test_appending(self) -> None:
        styles = self.styles()
        assert styles.ensure_dxf(Dxf.of(fill="FF0000")) == 0
        assert styles.ensure_dxf(Dxf.of(fill="00FF00")) == 1
        assert styles.dxf_count() == 2

    def test_the_same_format_is_reused(self) -> None:
        styles = self.styles()
        first = styles.ensure_dxf(Dxf.of(fill="FF0000", bold=True))
        again = styles.ensure_dxf(Dxf.of(fill="FF0000", bold=True))
        assert first == again
        assert styles.dxf_count() == 1

    def test_index_zero_is_not_reserved(self) -> None:
        """Unlike fills, where 0 and 1 are ``none`` and ``gray125``, because
        nothing refers to a dxf unless a rule names it."""
        styles = self.styles()
        assert styles.ensure_dxf(Dxf.of(fill="FF0000")) == 0

    def test_reading_one_excel_wrote(self) -> None:
        styles = self.styles(b'<dxfs count="1">' + EXCEL_FILL_ONLY + b"</dxfs>")
        dxf = styles.dxf(0)
        assert dxf.fill is not None
        assert dxf.fill.background == Color(rgb="FFC6EFCE")

    def test_an_existing_entry_is_matched_not_duplicated(self) -> None:
        styles = self.styles(b'<dxfs count="1">' + EXCEL_FILL_ONLY + b"</dxfs>")
        assert styles.ensure_dxf(Dxf.of(fill="C6EFCE")) == 0
        assert styles.dxf_count() == 1

    def test_the_count_attribute_follows(self) -> None:
        styles, document = self.sheet()
        styles.ensure_dxf(Dxf.of(fill="FF0000"))
        styles.ensure_dxf(Dxf.of(fill="00FF00"))
        assert '<dxfs count="2">' in document.to_bytes().decode()

    def test_a_dangling_index_formats_nothing(self) -> None:
        """A rule whose dxfId points past the table formats nothing, which
        is what Excel shows. Refusing to read the workbook over it would be
        worse than reporting what it does."""
        assert self.styles().dxf(7) == Dxf()
        assert self.styles().dxf(-1) == Dxf()

    def test_the_table_lands_in_schema_order(self) -> None:
        """``dxfs`` sits after ``cellStyles`` and before ``tableStyles``.
        Appending instead would give a stylesheet Excel repairs."""
        styles, document = self.sheet(b"<cellStyles/><tableStyles/>")
        styles.ensure_dxf(Dxf.of(fill="FF0000"))
        rendered = document.to_bytes().decode()
        assert rendered.index("<cellStyles") < rendered.index("<dxfs")
        assert rendered.index("<dxfs") < rendered.index("<tableStyles")
