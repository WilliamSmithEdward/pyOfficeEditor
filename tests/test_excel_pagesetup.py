"""How a sheet prints.

The XML quoted here is Excel's own. Margins in inches, fit-to-page living on
``sheetPr``, and the header's three boxes coded into one string are all
measured rather than taken from the schema.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyofficeeditor._xml import XmlDocument
from pyofficeeditor.excel import Workbook, Worksheet
from pyofficeeditor.excel._pagesetup import (
    HeaderFooter,
    HeaderFooterText,
    PageMargins,
    PageSetup,
    PrintOptions,
)

EXCEL_MARGINS = (
    b'<pageMargins left="0.25" right="0.25" top="1" bottom="1" header="0.5" footer="0.5"/>'
)
EXCEL_SETUP = (
    b'<pageSetup paperSize="9" firstPageNumber="3" fitToHeight="2" '
    b'orientation="landscape" useFirstPageNumber="1" r:id="rId1"/>'
)
EXCEL_OPTIONS = b'<printOptions horizontalCentered="1" headings="1" gridLines="1"/>'
EXCEL_HEADER = (
    b"<headerFooter><oddHeader>&amp;Lleft head&amp;C&amp;Bbold centre&amp;B</oddHeader>"
    b"<oddFooter>&amp;RPage &amp;P of &amp;N</oddFooter></headerFooter>"
)


class TestRoundTrip:
    def round_trip(self, raw: bytes, value: object) -> None:
        assert value.write().to_xml() == raw.decode()  # type: ignore[attr-defined]

    def test_margins(self) -> None:
        self.round_trip(EXCEL_MARGINS, PageMargins.read(XmlDocument.parse(EXCEL_MARGINS).root))

    def test_setup(self) -> None:
        self.round_trip(EXCEL_SETUP, PageSetup.read(XmlDocument.parse(EXCEL_SETUP).root))

    def test_options(self) -> None:
        self.round_trip(EXCEL_OPTIONS, PrintOptions.read(XmlDocument.parse(EXCEL_OPTIONS).root))

    def test_header(self) -> None:
        self.round_trip(EXCEL_HEADER, HeaderFooter.read(XmlDocument.parse(EXCEL_HEADER).root))

    def test_the_printer_settings_relationship_survives(self) -> None:
        """This library never makes one, and a sheet Excel set up keeps it."""
        found = PageSetup.read(XmlDocument.parse(EXCEL_SETUP).root)
        assert found.printer_settings_id == "rId1"


class TestMarginsAreInches:
    def test_reading(self) -> None:
        found = PageMargins.read(XmlDocument.parse(EXCEL_MARGINS).root)
        assert found.left == 0.25
        assert found.top == 1.0

    def test_points_convert(self) -> None:
        """The object model takes points; the file holds inches."""
        assert PageMargins.points(left=18).left == 0.25

    def test_all_six_are_always_written(self) -> None:
        """``CT_PageMargins`` requires them, so none is dropped as a
        default."""
        rendered = PageMargins().write().to_xml()
        for name in ("left", "right", "top", "bottom", "header", "footer"):
            assert f'{name}="' in rendered


class TestDefaultsAreAbsent:
    def test_a_default_orientation_is_unwritten(self) -> None:
        assert "orientation" not in PageSetup().write().to_xml()

    def test_one_is_the_default_for_fitting(self) -> None:
        rendered = PageSetup(fit_to_width=1, fit_to_height=2).write().to_xml()
        assert "fitToWidth" not in rendered
        assert 'fitToHeight="2"' in rendered

    def test_reading_an_absent_orientation(self) -> None:
        assert PageSetup.read(XmlDocument.parse(b"<pageSetup/>").root).orientation == "default"


class TestPaper:
    def test_a_named_size(self) -> None:
        assert PageSetup.on("A4").paper_size == 9

    def test_reading_the_name_back(self) -> None:
        assert PageSetup.read(XmlDocument.parse(EXCEL_SETUP).root).paper == "A4"

    def test_an_unknown_name(self) -> None:
        with pytest.raises(ValueError, match="not a paper size"):
            PageSetup.on("foolscap")

    def test_an_unknown_number_passes_through(self) -> None:
        """Excel has dozens; only the common ones are named here."""
        assert PageSetup(paper_size=256).paper is None
        assert 'paperSize="256"' in PageSetup(paper_size=256).write().to_xml()


class TestHeadersAreOneString:
    def test_splitting(self) -> None:
        found = HeaderFooter.read(XmlDocument.parse(EXCEL_HEADER).root)
        assert found.odd_header.left == "left head"
        assert found.odd_header.center == "&Bbold centre&B", "the codes are carried as text"
        assert found.odd_header.right == ""

    def test_joining(self) -> None:
        text = HeaderFooterText(left="a", center="b", right="c")
        assert text.render() == "&La&Cb&Rc"

    def test_only_the_sections_that_have_text(self) -> None:
        assert HeaderFooterText(right="c").render() == "&Rc"

    def test_unmarked_text_is_the_centre(self) -> None:
        """Which is where Excel puts a header with no marker at all."""
        assert HeaderFooterText.parse("plain").center == "plain"

    def test_an_empty_one(self) -> None:
        assert HeaderFooterText.parse("") == HeaderFooterText()
        assert HeaderFooter().is_empty

    def test_a_section_repeated_is_joined(self) -> None:
        assert HeaderFooterText.parse("&Lone&Ltwo").left == "onetwo"

    def test_the_alternates_are_only_written_when_used(self) -> None:
        plain = HeaderFooter(odd_header=HeaderFooterText(left="a"))
        assert "evenHeader" not in plain.write().to_xml()
        assert "differentOddEven" not in plain.write().to_xml()


class TestOnAWorksheet:
    @pytest.fixture()
    def sheet(self, live_sample_xlsx: Path) -> Worksheet:
        return Workbook.open(live_sample_xlsx)["Data"]

    def test_margins_default_to_excels(self, sheet: Worksheet) -> None:
        assert sheet.page_margins.left == 0.7

    def test_setting_margins(self, sheet: Worksheet) -> None:
        sheet.page_margins = PageMargins(left=0.25, right=0.25)
        assert sheet.page_margins.left == 0.25

    def test_setting_the_setup(self, sheet: Worksheet) -> None:
        sheet.page_setup = PageSetup.on("A4", orientation="landscape")
        found = sheet.page_setup
        assert found.orientation == "landscape"
        assert found.paper == "A4"

    def test_print_options(self, sheet: Worksheet) -> None:
        sheet.print_options = PrintOptions(gridlines=True, headings=True)
        assert sheet.print_options.gridlines is True

    def test_empty_print_options_remove_the_element(self, sheet: Worksheet) -> None:
        sheet.print_options = PrintOptions(gridlines=True)
        sheet.print_options = PrintOptions()
        assert b"printOptions" not in sheet.document.to_bytes()

    def test_a_header(self, sheet: Worksheet) -> None:
        sheet.header_footer = HeaderFooter(
            odd_header=HeaderFooterText(center="Report"),
            odd_footer=HeaderFooterText(right="Page &P of &N"),
        )
        found = sheet.header_footer
        assert found.odd_header.center == "Report"
        assert found.odd_footer.right == "Page &P of &N"

    def test_fit_to_page_is_off_by_default(self, sheet: Worksheet) -> None:
        assert sheet.fit_to_page is False

    def test_turning_fit_to_page_on(self, sheet: Worksheet) -> None:
        """The numbers live on pageSetup and do nothing without this."""
        sheet.page_setup = PageSetup(fit_to_width=1, fit_to_height=2)
        sheet.fit_to_page = True
        assert sheet.fit_to_page is True
        assert 'fitToPage="1"' in sheet.document.to_bytes().decode()

    def test_turning_it_off_again(self, sheet: Worksheet) -> None:
        sheet.fit_to_page = True
        sheet.fit_to_page = False
        assert sheet.fit_to_page is False
        assert b"fitToPage" not in sheet.document.to_bytes()

    def test_everything_lands_in_schema_order(self, sheet: Worksheet) -> None:
        sheet.print_options = PrintOptions(gridlines=True)
        sheet.page_margins = PageMargins(left=1)
        sheet.page_setup = PageSetup(orientation="landscape")
        sheet.header_footer = HeaderFooter(odd_header=HeaderFooterText(center="x"))
        rendered = sheet.document.to_bytes().decode()
        order = [
            rendered.index(f"<{name}")
            for name in ("printOptions", "pageMargins", "pageSetup", "headerFooter")
        ]
        assert order == sorted(order)

    def test_it_survives_a_save(self, sheet: Worksheet) -> None:
        sheet.page_setup = PageSetup.on("A4", orientation="landscape")
        sheet.header_footer = HeaderFooter(odd_header=HeaderFooterText(left="L"))
        reopened = Workbook.from_bytes(sheet.workbook.to_bytes())["Data"]
        assert reopened.page_setup.paper == "A4"
        assert reopened.header_footer.odd_header.left == "L"


class TestPrintAreaIsADefinedName:
    @pytest.fixture()
    def sheet(self, live_sample_xlsx: Path) -> Worksheet:
        return Workbook.open(live_sample_xlsx)["Data"]

    def test_none_by_default(self, sheet: Worksheet) -> None:
        assert sheet.print_area == ()

    def test_setting_one(self, sheet: Worksheet) -> None:
        sheet.print_area = "A1:D20"
        assert [block.a1 for block in sheet.print_area] == ["$A$1:$D$20"]

    def test_it_is_stored_as_the_builtin_name(self, sheet: Worksheet) -> None:
        """Which is why it already moves when rows are inserted."""
        sheet.print_area = "A1:D20"
        assert sheet.workbook.defined_name("_xlnm.Print_Area", scope="Data") is not None

    def test_it_is_written_absolute(self, sheet: Worksheet) -> None:
        """A relative one would move with whichever cell was selected."""
        sheet.print_area = "A1:D20"
        name = sheet.workbook.defined_name("_xlnm.Print_Area", scope="Data")
        assert name.refers_to == "Data!$A$1:$D$20"

    def test_several_areas(self, sheet: Worksheet) -> None:
        sheet.print_area = ["A1:B5", "D1:E5"]
        assert [block.a1 for block in sheet.print_area] == ["$A$1:$B$5", "$D$1:$E$5"]

    def test_replacing_one(self, sheet: Worksheet) -> None:
        sheet.print_area = "A1:B5"
        sheet.print_area = "C1:D5"
        assert [block.a1 for block in sheet.print_area] == ["$C$1:$D$5"]

    def test_clearing_it(self, sheet: Worksheet) -> None:
        sheet.print_area = "A1:B5"
        sheet.print_area = None
        assert sheet.print_area == ()

    def test_clearing_one_that_was_never_set(self, sheet: Worksheet) -> None:
        sheet.print_area = None
        assert sheet.print_area == ()

    def test_an_empty_area(self, sheet: Worksheet) -> None:
        with pytest.raises(ValueError, match="at least one range"):
            sheet.print_area = ""

    def test_it_moves_when_rows_are_inserted(self, sheet: Worksheet) -> None:
        sheet.print_area = "A1:D20"
        sheet.insert_rows(3, 2)
        assert [block.a1 for block in sheet.print_area] == ["$A$1:$D$22"]

    def test_titles(self, sheet: Worksheet) -> None:
        sheet.print_titles = "$1:$1"
        assert sheet.print_titles == "Data!$1:$1"

    def test_titles_for_rows_and_columns(self, sheet: Worksheet) -> None:
        sheet.print_titles = "$A:$A,$1:$1"
        assert sheet.print_titles == "Data!$A:$A,Data!$1:$1"

    def test_clearing_titles(self, sheet: Worksheet) -> None:
        sheet.print_titles = "$1:$1"
        sheet.print_titles = None
        assert sheet.print_titles is None


class TestReadingWhatExcelWrote:
    def test_everything_together(self, live_settings_xlsx: Path) -> None:
        sheet = Workbook.open(live_settings_xlsx)["S"]
        assert sheet.page_margins == PageMargins(0.25, 0.25, 1.0, 1.0, 0.5, 0.5)
        assert sheet.page_setup.orientation == "landscape"
        assert sheet.page_setup.paper == "A4"
        assert sheet.page_setup.fit_to_height == 2
        assert sheet.page_setup.first_page_number == 3
        assert sheet.fit_to_page is True
        assert sheet.print_options == PrintOptions(
            horizontal_centered=True, headings=True, gridlines=True
        )
        assert sheet.header_footer.odd_header.left == "left head"
        assert sheet.header_footer.odd_footer.right == "Page &P of &N"
        assert [block.a1 for block in sheet.print_area] == ["$A$1:$D$20"]
        assert sheet.print_titles == "S!$A:$A,S!$1:$1"
