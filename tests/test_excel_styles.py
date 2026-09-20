"""Number formats, the only thing that tells a date from a number.

The gate is the fixture: real Excel wrote ``2026-01-15`` as ``46037`` with
``s="2"``, and following that index has to arrive at "this is a date".
"""

from __future__ import annotations

import pytest

from pyofficeeditor.excel._styles import (
    BUILTIN_NUMBER_FORMATS,
    FIRST_CUSTOM_NUMBER_FORMAT,
    Styles,
    is_date_format,
)
from pyofficeeditor.opc import OpcPackage


class TestIsDateFormat:
    @pytest.mark.parametrize(
        "code",
        [
            "yyyy-mm-dd",
            r"yyyy\-mm\-dd",
            "mm-dd-yy",
            "d-mmm-yy",
            "d-mmm",
            "mmm-yy",
            "h:mm AM/PM",
            "h:mm:ss",
            "m/d/yy h:mm",
            "mm:ss",
            "[h]:mm:ss",
            "mmss.0",
            "yyyy",
            "dddd",
            "[$-409]d/mmm/yyyy",
            'yyyy"年"m"月"',
        ],
    )
    def test_date_and_time_codes(self, code: str) -> None:
        assert is_date_format(code) is True, code

    @pytest.mark.parametrize(
        "code",
        [
            "General",
            "",
            "0",
            "0.00",
            "#,##0",
            "#,##0.00",
            "0%",
            "0.00%",
            "0.00E+00",
            "##0.0E+0",
            "@",
            "# ?/?",
            "[Red]0.00",
            "[Blue]#,##0",
            "[>100]0.00;[<=100]0.0",
            '#,##0 "days"',
            '0.0 "hours"',
            '#,##0" m"',
            r"0\d",
            "[$-409]#,##0.00",
            '"y"0',
        ],
    )
    def test_codes_that_are_not_dates(self, code: str) -> None:
        assert is_date_format(code) is False, code

    def test_a_quoted_literal_does_not_make_a_date(self) -> None:
        """``"days"`` holds d, a, y, s. All literal."""
        assert is_date_format('#,##0 "days"') is False

    def test_an_escape_does_not_make_a_date(self) -> None:
        assert is_date_format(r"0\y") is False

    def test_a_colour_does_not_make_a_date_but_elapsed_time_does(self) -> None:
        assert is_date_format("[Red]0") is False
        assert is_date_format("[h]") is True
        assert is_date_format("[mm]:ss") is True

    def test_only_the_first_section_decides(self) -> None:
        assert is_date_format('0.00;[Red]"yyyy"') is False
        assert is_date_format("yyyy;0.00") is True

    def test_a_semicolon_inside_a_literal_is_not_a_separator(self) -> None:
        assert is_date_format('"a;b"0') is False
        assert is_date_format('"a;b"yyyy') is True

    def test_an_unterminated_bracket_does_not_crash(self) -> None:
        assert is_date_format("[Red") is False

    def test_an_unterminated_quote_does_not_crash(self) -> None:
        assert is_date_format('0"unclosed') is False

    def test_every_builtin_classifies_without_raising(self) -> None:
        for format_id, code in BUILTIN_NUMBER_FORMATS.items():
            result = is_date_format(code)
            expected = format_id in {*range(14, 23), *range(45, 48)}
            assert result is expected, f"builtin {format_id} {code!r}"


class TestAgainstRealExcelOutput:
    @pytest.fixture()
    def styles(self, live_sample_xlsx: object) -> Styles:
        package = OpcPackage.open(str(live_sample_xlsx))
        return Styles(package.xml("xl/styles.xml"))

    def test_the_custom_format_excel_wrote(self, styles: Styles) -> None:
        assert styles.custom_formats == {164: r"yyyy\-mm\-dd"}

    def test_the_date_column_resolves_to_a_date(self, styles: Styles) -> None:
        """F2 carries s="2". cellXfs[2] has numFmtId 164, which is the
        custom yyyy-mm-dd. Three hops, and the answer must be yes."""
        assert styles.number_format_id(2) == 164
        assert styles.number_format(2) == r"yyyy\-mm\-dd"
        assert styles.is_date(2) is True

    def test_the_general_styles_are_not_dates(self, styles: Styles) -> None:
        assert styles.number_format_id(0) == 0
        assert styles.number_format(0) == "General"
        assert styles.is_date(0) is False
        assert styles.is_date(1) is False, "s=1 is the bold header, not a date"

    def test_a_cell_with_no_style_is_general(self, styles: Styles) -> None:
        assert styles.number_format_id(None) == 0
        assert styles.is_date(None) is False

    def test_a_style_index_past_the_table_agrees_with_excel(self, styles: Styles) -> None:
        """Excel renders such a cell with the default format rather than
        refusing the file, so reading must not raise either."""
        assert styles.number_format_id(9999) == 0
        assert styles.is_date(9999) is False

    def test_the_cell_format_count_matches_the_part(self, styles: Styles) -> None:
        assert styles.cell_format_count == 4


class TestWriting:
    @pytest.fixture()
    def package(self, live_sample_xlsx: object) -> OpcPackage:
        return OpcPackage.open(str(live_sample_xlsx))

    def test_an_existing_format_is_reused(self, package: OpcPackage) -> None:
        styles = Styles(package.xml("xl/styles.xml"))
        assert styles.ensure_number_format(r"yyyy\-mm\-dd") == 2
        assert styles.cell_format_count == 4, "nothing was added"
        assert not package.xml("xl/styles.xml").is_modified

    def test_a_builtin_is_recognised_rather_than_duplicated(self, package: OpcPackage) -> None:
        styles = Styles(package.xml("xl/styles.xml"))
        index = styles.ensure_number_format("h:mm:ss")
        assert styles.number_format_id(index) == 21, "the builtin id, not a new custom one"
        assert 164 not in {k for k in styles.custom_formats if k != 164}

    def test_a_new_format_is_added_and_resolves(self, package: OpcPackage) -> None:
        styles = Styles(package.xml("xl/styles.xml"))
        before = styles.cell_format_count
        index = styles.ensure_number_format("yyyy-mm-dd hh:mm:ss")
        assert index == before
        assert styles.is_date(index) is True
        assert styles.number_format(index) == "yyyy-mm-dd hh:mm:ss"
        assert styles.number_format_id(index) >= FIRST_CUSTOM_NUMBER_FORMAT

    def test_a_new_format_survives_a_save(self, package: OpcPackage) -> None:
        styles = Styles(package.xml("xl/styles.xml"))
        index = styles.ensure_number_format("0.000%")
        reopened = OpcPackage.from_bytes(package.to_bytes())
        again = Styles(reopened.xml("xl/styles.xml"))
        assert again.number_format(index) == "0.000%"
        assert again.is_date(index) is False

    def test_asking_twice_adds_once(self, package: OpcPackage) -> None:
        styles = Styles(package.xml("xl/styles.xml"))
        first = styles.ensure_number_format("dd/mm/yyyy")
        count = styles.cell_format_count
        second = styles.ensure_number_format("dd/mm/yyyy")
        assert first == second
        assert styles.cell_format_count == count

    def test_adding_a_format_leaves_the_rest_of_styles_alone(self, package: OpcPackage) -> None:
        original = package.read("xl/styles.xml").decode()
        Styles(package.xml("xl/styles.xml")).ensure_number_format("dd/mm/yyyy")
        edited = package.read("xl/styles.xml").decode()
        for section in ("<fonts", "<fills", "<borders", "<cellStyles"):
            start = original.index(section)
            end = original.index(">", start)
            assert original[start:end] in edited, f"{section} was disturbed"
        assert "<fonts count=" in edited

    def test_the_counts_are_maintained(self, package: OpcPackage) -> None:
        styles = Styles(package.xml("xl/styles.xml"))
        styles.ensure_number_format("dd/mm/yyyy")
        text = package.read("xl/styles.xml").decode()
        assert '<numFmts count="2">' in text
        assert '<cellXfs count="5">' in text
