"""Sheet protection, tab colour, view settings and visibility.

The XML quoted here is Excel's own. The protection flags in particular are
measured rather than inferred: every one of them names a lock, and which way
absence falls is not the same for all of them.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyofficeeditor._xml import XmlDocument
from pyofficeeditor.excel import Color, Workbook, Worksheet
from pyofficeeditor.excel._protection import (
    DEFAULT_SPIN_COUNT,
    SheetProtection,
    hash_password,
)

#: Excel protecting with a password, drawing objects and scenarios.
EXCEL_WITH_PASSWORD = (
    b'<sheetProtection algorithmName="SHA-512" '
    b'hashValue="1TlAxgwmvIZdHyvYaSX4qaGInAb3qzNhsaQG3oBu1vdYP9GO5t4v/tgDehQd6J234'
    b'kstSzr/MhVpPkn4B4UJIQ==" saltValue="E4xlJdthqYle67vvctFHQg==" '
    b'spinCount="100000" sheet="1" objects="1" scenarios="1"/>'
)

#: Excel protecting while allowing formatting and sorting. Note the zeros.
EXCEL_ALLOWING = b'<sheetProtection sheet="1" formatCells="0" sort="0"/>'

#: The password behind EXCEL_WITH_PASSWORD, and the salt it used.
PASSWORD = "secret"
SALT = "E4xlJdthqYle67vvctFHQg=="


def parse(raw: bytes) -> SheetProtection:
    return SheetProtection.read(XmlDocument.parse(raw).root)


class TestEveryFlagNamesALock:
    """``formatCells="0"`` means formatting *is* allowed. Measured: calling
    ``Protect(AllowFormattingCells:=True, AllowSorting:=True)`` produced
    exactly that, so the attribute asserts the restriction."""

    def test_zero_means_allowed(self) -> None:
        found = parse(EXCEL_ALLOWING)
        assert found.allow_format_cells is True
        assert found.allow_sort is True

    def test_absent_means_blocked_for_most(self) -> None:
        found = parse(EXCEL_ALLOWING)
        assert found.allow_insert_rows is False
        assert found.allow_delete_columns is False
        assert found.allow_pivot_tables is False

    def test_absent_means_allowed_for_the_two_selection_flags(self) -> None:
        """Which is why the direction cannot be read off the name: an
        ordinary protected sheet still lets you select cells."""
        found = parse(EXCEL_ALLOWING)
        assert found.allow_select_locked_cells is True
        assert found.allow_select_unlocked_cells is True

    def test_writing_an_allowance_writes_a_zero(self) -> None:
        rendered = SheetProtection(allow_sort=True).write().to_xml()
        assert 'sort="0"' in rendered

    def test_a_default_block_writes_nothing(self) -> None:
        rendered = SheetProtection().write().to_xml()
        assert "insertRows" not in rendered

    def test_forbidding_selection_writes_a_one(self) -> None:
        rendered = SheetProtection(allow_select_locked_cells=False).write().to_xml()
        assert 'selectLockedCells="1"' in rendered


class TestRoundTrip:
    @pytest.mark.parametrize("raw", [EXCEL_WITH_PASSWORD, EXCEL_ALLOWING])
    def test_excel_output_comes_back_byte_for_byte(self, raw: bytes) -> None:
        assert parse(raw).write().to_xml() == raw.decode()

    @pytest.mark.parametrize("raw", [EXCEL_WITH_PASSWORD, EXCEL_ALLOWING])
    def test_parsing_is_stable(self, raw: bytes) -> None:
        once = parse(raw)
        assert parse(once.write().to_xml().encode()) == once


class TestPasswords:
    def test_the_hash_reproduces_excels(self) -> None:
        """SHA-512 over the salt then the password in UTF-16LE, spun 100000
        times over the previous digest and the round number. Checked against
        a workbook Excel protected with this password and salt."""
        import base64

        digest, salt = hash_password(PASSWORD, salt=base64.b64decode(SALT))
        assert salt == SALT
        assert digest == parse(EXCEL_WITH_PASSWORD).hash_value

    def test_matching(self) -> None:
        found = parse(EXCEL_WITH_PASSWORD)
        assert found.matches(PASSWORD) is True
        assert found.matches("wrong") is False

    def test_a_protection_with_no_password_matches_nothing(self) -> None:
        assert parse(EXCEL_ALLOWING).matches("") is False
        assert parse(EXCEL_ALLOWING).has_password is False

    def test_setting_one(self) -> None:
        guarded = SheetProtection().with_password("hunter2")
        assert guarded.has_password
        assert guarded.algorithm_name == "SHA-512"
        assert guarded.spin_count == DEFAULT_SPIN_COUNT
        assert guarded.matches("hunter2")

    def test_the_salt_is_random_so_two_look_different(self) -> None:
        first = SheetProtection().with_password("same")
        second = SheetProtection().with_password("same")
        assert first.salt_value != second.salt_value
        assert first.hash_value != second.hash_value
        assert first.matches("same") and second.matches("same")

    def test_an_empty_password_is_refused(self) -> None:
        with pytest.raises(ValueError, match="protects nothing"):
            SheetProtection().with_password("")

    def test_an_unknown_algorithm_is_refused(self) -> None:
        with pytest.raises(ValueError, match="not an algorithm"):
            hash_password("x", algorithm="SHA-3")  # type: ignore[arg-type]


class TestOnAWorksheet:
    @pytest.fixture()
    def sheet(self, live_sample_xlsx: Path) -> Worksheet:
        return Workbook.open(live_sample_xlsx)["Data"]

    def test_a_clean_sheet_is_unprotected(self, sheet: Worksheet) -> None:
        assert sheet.protection is None

    def test_protecting(self, sheet: Worksheet) -> None:
        sheet.protect()
        found = sheet.protection
        assert found is not None and found.contents is True

    def test_protecting_with_a_password(self, sheet: Worksheet) -> None:
        sheet.protect(password="hunter2")
        found = sheet.protection
        assert found is not None and found.matches("hunter2")

    def test_protecting_while_allowing_something(self, sheet: Worksheet) -> None:
        sheet.protect(SheetProtection(allow_sort=True, allow_autofilter=True))
        found = sheet.protection
        assert found is not None
        assert found.allow_sort is True
        assert found.allow_insert_rows is False

    def test_protecting_twice_replaces(self, sheet: Worksheet) -> None:
        sheet.protect()
        sheet.protect(SheetProtection(allow_sort=True))
        rendered = sheet.document.to_bytes().decode()
        assert rendered.count("<sheetProtection") == 1

    def test_unprotecting(self, sheet: Worksheet) -> None:
        sheet.protect()
        assert sheet.unprotect() is True
        assert sheet.protection is None
        assert b"sheetProtection" not in sheet.document.to_bytes()

    def test_unprotecting_what_was_never_protected(self, sheet: Worksheet) -> None:
        assert sheet.unprotect() is False

    def test_it_lands_in_schema_order(self, sheet: Worksheet) -> None:
        sheet.protect()
        rendered = sheet.document.to_bytes().decode()
        assert rendered.index("</sheetData>") < rendered.index("<sheetProtection")
        assert rendered.index("<sheetProtection") < rendered.index("<mergeCells")


class TestAppearance:
    @pytest.fixture()
    def sheet(self, live_sample_xlsx: Path) -> Worksheet:
        return Workbook.open(live_sample_xlsx)["Data"]

    def test_the_tab_has_no_colour_by_default(self, sheet: Worksheet) -> None:
        assert sheet.tab_color is None

    def test_setting_one(self, sheet: Worksheet) -> None:
        sheet.tab_color = "FF0000"
        assert sheet.tab_color == Color(rgb="FFFF0000")
        assert "<tabColor" in sheet.document.to_bytes().decode()

    def test_clearing_it(self, sheet: Worksheet) -> None:
        sheet.tab_color = "FF0000"
        sheet.tab_color = None
        assert sheet.tab_color is None

    def test_gridlines_and_headings_are_on_by_default(self, sheet: Worksheet) -> None:
        assert sheet.show_gridlines is True
        assert sheet.show_headings is True

    def test_turning_them_off(self, sheet: Worksheet) -> None:
        sheet.show_gridlines = False
        sheet.show_headings = False
        assert sheet.show_gridlines is False
        assert sheet.show_headings is False
        assert 'showGridLines="0"' in sheet.document.to_bytes().decode()

    def test_turning_them_back_on_removes_the_attribute(self, sheet: Worksheet) -> None:
        """The default is on, so saying so explicitly is noise."""
        sheet.show_gridlines = False
        sheet.show_gridlines = True
        assert "showGridLines" not in sheet.document.to_bytes().decode()

    def test_zoom_defaults_to_a_hundred(self, sheet: Worksheet) -> None:
        assert sheet.zoom == 100

    def test_setting_zoom_writes_both_attributes(self, sheet: Worksheet) -> None:
        """Excel writes zoomScaleNormal too, and restores from it when
        switching back from page-break preview."""
        sheet.zoom = 85
        rendered = sheet.document.to_bytes().decode()
        assert 'zoomScale="85"' in rendered
        assert 'zoomScaleNormal="85"' in rendered

    def test_an_impossible_zoom(self, sheet: Worksheet) -> None:
        with pytest.raises(ValueError, match="outside 10 to 400"):
            sheet.zoom = 5


class TestVisibility:
    @pytest.fixture()
    def book(self, live_sample_xlsx: Path) -> Workbook:
        return Workbook.open(live_sample_xlsx)

    def test_sheets_start_visible(self, book: Workbook) -> None:
        assert book["Data"].visible == "visible"

    def test_hiding_one(self, book: Workbook) -> None:
        book["Notes"].visible = "hidden"
        assert book["Notes"].visible == "hidden"
        assert 'state="hidden"' in book.package.xml(book.workbook_part).to_bytes().decode()

    def test_very_hidden(self, book: Workbook) -> None:
        book["Notes"].visible = "veryHidden"
        assert book["Notes"].visible == "veryHidden"

    def test_showing_again_removes_the_attribute(self, book: Workbook) -> None:
        book["Notes"].visible = "hidden"
        book["Notes"].visible = "visible"
        assert "state=" not in book.package.xml(book.workbook_part).to_bytes().decode()

    def test_hiding_the_last_visible_sheet_is_refused(self, book: Workbook) -> None:
        """Excel refuses to open a workbook in which every sheet is hidden,
        so it is refused here rather than written."""
        book["Notes"].visible = "hidden"
        with pytest.raises(ValueError, match="only visible sheet"):
            book["Data"].visible = "hidden"

    def test_an_invented_state(self, book: Workbook) -> None:
        with pytest.raises(ValueError, match="not a visibility"):
            book["Notes"].visible = "sortOfHidden"  # type: ignore[assignment]

    def test_it_survives_a_save(self, book: Workbook) -> None:
        book["Notes"].visible = "veryHidden"
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Notes"].visible == "veryHidden"


class TestReadingWhatExcelWrote:
    def test_everything_together(self, live_settings_xlsx: Path) -> None:
        book = Workbook.open(live_settings_xlsx)
        sheet = book["S"]
        assert sheet.tab_color == Color(rgb="FFFF0000")
        assert sheet.show_gridlines is False
        assert sheet.show_headings is False
        assert sheet.zoom == 85

        guarded = sheet.protection
        assert guarded is not None
        assert guarded.contents and guarded.objects and guarded.scenarios
        assert guarded.matches("secret")

        allowing = book["Plain"].protection
        assert allowing is not None
        assert allowing.allow_format_cells is True
        assert allowing.allow_sort is True
        assert allowing.has_password is False

        assert book["Hidden"].visible == "hidden"
        assert book["VeryHidden"].visible == "veryHidden"
