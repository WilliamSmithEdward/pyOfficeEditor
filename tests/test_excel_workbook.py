"""The Excel surface: workbooks, worksheets, cells and ranges.

Every read here has a wrong answer that a naive implementation gives
instead, and the tests name it: an index for a string, a serial for a date,
an empty formula for a shared one, a text match for an error. The fixtures
are bytes real Excel wrote, because those wrong answers all look plausible.
"""

from __future__ import annotations

import datetime as dt
import zipfile
from pathlib import Path

import pytest

from pyofficeeditor.excel import CellError, Workbook
from pyofficeeditor.excel._values import (
    DateOutOfRangeError,
    datetime_to_serial,
    format_number,
    serial_to_datetime,
)
from pyofficeeditor.exceptions import UnsupportedFormatError


@pytest.fixture()
def book(live_sample_xlsx: Path) -> Workbook:
    return Workbook.open(live_sample_xlsx)


class TestOpening:
    def test_sheets_are_in_tab_order_not_relationship_order(self, book: Workbook) -> None:
        """Excel listed rId2 before rId1 in the rels part. Tab order comes
        from <sheets>, so Data must be first."""
        assert book.sheet_names == ["Data", "Notes"]
        assert book[0].name == "Data"
        assert book[1].name == "Notes"

    def test_lookup_by_name_and_membership(self, book: Workbook) -> None:
        assert book["Data"].name == "Data"
        assert book.sheet("Notes").name == "Notes"
        assert "Data" in book
        assert "Absent" not in book
        assert len(book) == 2
        assert [s.name for s in book] == ["Data", "Notes"]

    def test_a_missing_sheet_names_what_is_there(self, book: Workbook) -> None:
        with pytest.raises(KeyError, match="Data, Notes"):
            book.sheet("Nope")

    def test_an_out_of_range_index(self, book: Workbook) -> None:
        with pytest.raises(IndexError, match="workbook has 2"):
            book[5]

    def test_the_active_sheet(self, book: Workbook) -> None:
        assert book.active.name == "Data"

    def test_the_date_system(self, book: Workbook) -> None:
        assert book.epoch_1904 is False

    def test_shared_parts_are_found(self, book: Workbook) -> None:
        assert book.shared_strings is not None
        assert len(book.shared_strings) == 14
        assert book.styles is not None

    @pytest.mark.parametrize(
        ("name", "expected"),
        [
            ("book.xlsb", "binary workbook"),
            ("book.xls", "legacy BIFF8"),
            ("book.docx", "not a workbook"),
            ("book.txt", "not a workbook"),
        ],
    )
    def test_formats_this_class_does_not_open_are_named(
        self, tmp_path: Path, name: str, expected: str
    ) -> None:
        target = tmp_path / name
        target.write_bytes(b"not really")
        with pytest.raises(UnsupportedFormatError, match=expected):
            Workbook.open(target)


class TestReadingValues:
    def test_a_shared_string_resolves_to_its_text(self, book: Workbook) -> None:
        """The cell holds ``<v>6</v>``. The wrong answer is 6."""
        assert book["Data"]["A2"].value == "North"
        assert book["Data"]["A1"].value == "Region"

    def test_a_number_is_a_number(self, book: Workbook) -> None:
        assert book["Data"]["B2"].value == 120
        assert isinstance(book["Data"]["B2"].value, int)
        assert book["Data"]["C2"].value == 4.25

    def test_a_date_resolves_through_its_number_format(self, book: Workbook) -> None:
        """The cell holds ``<v>46037</v>``. The wrong answer is 46037."""
        assert book["Data"]["F2"].value == dt.date(2026, 1, 15)
        assert book["Data"]["F5"].value == dt.date(2026, 3, 9)

    def test_a_boolean_is_a_boolean(self, book: Workbook) -> None:
        assert book["Data"]["E2"].value is True
        assert book["Data"]["E3"].value is False

    def test_an_error_is_not_a_string(self, book: Workbook) -> None:
        """A cell holding the text ``#DIV/0!`` and a cell holding the error
        are different cells."""
        value = book["Data"]["B8"].value
        assert isinstance(value, CellError)
        assert value.code == "#DIV/0!"
        assert value.is_known
        assert str(value) == "#DIV/0!"
        assert value != "#DIV/0!"

    def test_escaped_text_comes_back_unescaped(self, book: Workbook) -> None:
        assert book["Data"]["A8"].value == 'ampersand & angle < bracket > quote " done'

    def test_whitespace_survives(self, book: Workbook) -> None:
        assert book["Data"]["A9"].value == "  padded  "

    def test_an_absent_cell_is_none_and_stays_absent(self, book: Workbook) -> None:
        cell = book["Data"]["Z99"]
        assert cell.value is None
        assert cell.exists is False
        assert book["Data"].document.is_modified is False, "reading created nothing"

    def test_values_from_the_second_sheet(self, book: Workbook) -> None:
        assert book["Notes"]["A1"].value == "Second sheet"

    def test_cell_by_row_and_column(self, book: Workbook) -> None:
        assert book["Data"].cell(row=2, column=1).value == "North"
        assert book["Data"].cell(row=2, column=2).value == 120


class TestReadingFormulas:
    def test_a_plain_formula(self, book: Workbook) -> None:
        assert book["Data"]["D6"].formula == "SUM(D2:D5)"

    def test_the_master_of_a_shared_group(self, book: Workbook) -> None:
        assert book["Data"]["D2"].formula == "B2*C2"

    def test_a_follower_is_translated_from_the_master(self, book: Workbook) -> None:
        """D3, D4 and D5 carry ``<f t="shared" si="0"/>`` with no text at
        all. The wrong answer is an empty formula."""
        sheet = book["Data"]
        assert sheet["D3"].formula == "B3*C3"
        assert sheet["D4"].formula == "B4*C4"
        assert sheet["D5"].formula == "B5*C5"

    def test_the_translation_agrees_with_excels_own_arithmetic(self, book: Workbook) -> None:
        """Each follower's cached value equals its translated formula applied
        to the row's inputs, which is the check that the shift is right."""
        sheet = book["Data"]
        for row in (2, 3, 4, 5):
            units = sheet.cell(row, 2).value
            price = sheet.cell(row, 3).value
            total = sheet.cell(row, 4).value
            assert isinstance(units, (int, float))
            assert isinstance(price, (int, float))
            assert total == pytest.approx(units * price), f"row {row}"
            assert sheet.cell(row, 4).formula == f"B{row}*C{row}"

    def test_a_cell_with_no_formula(self, book: Workbook) -> None:
        assert book["Data"]["A1"].formula is None
        assert book["Data"]["Z99"].formula is None

    def test_a_formula_cells_value_is_its_cached_result(self, book: Workbook) -> None:
        """The library does not evaluate formulas, so ``value`` is whatever
        Excel last computed. That is a documented contract, not an accident:
        change an input and the cache is stale until Excel reopens the file,
        which is why saving sets fullCalcOnLoad."""
        assert book["Data"]["D2"].value == 510
        book["Data"]["B2"].value = 200
        assert book["Data"]["D2"].value == 510, "still the cache; Excel recalculates on open"


class TestWritingValues:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            ("text", "text"),
            ("", ""),
            ("  spaced  ", "  spaced  "),
            ('quote " and & amp', 'quote " and & amp'),
            (0, 0),
            (42, 42),
            (-7, -7),
            (3.5, 3.5),
            (0.1, 0.1),
            (1e20, 1e20),
            (True, True),
            (False, False),
            (dt.date(2026, 7, 4), dt.date(2026, 7, 4)),
            (dt.datetime(2026, 7, 4, 13, 45), dt.datetime(2026, 7, 4, 13, 45)),
            (dt.time(9, 30), dt.time(9, 30)),
            (CellError("#N/A"), CellError("#N/A")),
            (None, None),
        ],
    )
    def test_a_value_round_trips_through_a_save(
        self, book: Workbook, value: object, expected: object
    ) -> None:
        book["Data"]["J1"].value = value  # type: ignore[assignment]
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"]["J1"].value == expected

    def test_writing_creates_the_cell_and_the_row(self, book: Workbook) -> None:
        sheet = book["Data"]
        assert sheet["Z50"].exists is False
        sheet["Z50"].value = "far away"
        assert sheet["Z50"].exists is True
        assert Workbook.from_bytes(book.to_bytes())["Data"]["Z50"].value == "far away"

    def test_setitem_shorthand(self, book: Workbook) -> None:
        book["Data"]["J2"] = "via setitem"
        assert book["Data"]["J2"].value == "via setitem"

    def test_a_new_string_joins_the_shared_table(self, book: Workbook) -> None:
        before = len(book.shared_strings or [])
        book["Data"]["J3"].value = "never seen before"
        assert len(book.shared_strings or []) == before + 1
        raw = book.to_bytes()
        assert b"never seen before" in zipfile.ZipFile(__import__("io").BytesIO(raw)).read(
            "xl/sharedStrings.xml"
        )

    def test_an_existing_string_reuses_its_entry(self, book: Workbook) -> None:
        before = len(book.shared_strings or [])
        book["Data"]["J4"].value = "North"
        assert len(book.shared_strings or []) == before

    def test_a_date_gets_a_date_number_format(self, book: Workbook) -> None:
        book["Data"]["J5"].value = dt.date(2026, 7, 4)
        reopened = Workbook.from_bytes(book.to_bytes())
        cell = reopened["Data"]["J5"]
        assert cell.value == dt.date(2026, 7, 4)
        styles = reopened.styles
        assert styles is not None and styles.is_date(cell.style_index)

    def test_writing_a_date_twice_adds_one_style(self, book: Workbook) -> None:
        styles = book.styles
        assert styles is not None
        book["Data"]["J6"].value = dt.date(2026, 7, 4)
        after_first = styles.cell_format_count
        book["Data"]["J7"].value = dt.date(2026, 8, 1)
        assert styles.cell_format_count == after_first

    def test_overwriting_a_value_replaces_it(self, book: Workbook) -> None:
        book["Data"]["B2"].value = 200
        assert book["Data"]["B2"].value == 200
        book["Data"]["B2"].value = "now text"
        assert book["Data"]["B2"].value == "now text"
        assert Workbook.from_bytes(book.to_bytes())["Data"]["B2"].value == "now text"

    def test_writing_over_a_formula_removes_it(self, book: Workbook) -> None:
        """A cell keeping a stale <f> beside a new <v> shows the formula's
        old result as soon as Excel recalculates."""
        book["Data"]["D6"].value = 99
        assert book["Data"]["D6"].formula is None
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"]["D6"].value == 99
        assert reopened["Data"]["D6"].formula is None

    def test_clearing_a_cell(self, book: Workbook) -> None:
        book["Data"]["C2"].clear()
        assert book["Data"]["C2"].value is None
        assert book["Data"]["C2"].exists is False
        assert Workbook.from_bytes(book.to_bytes())["Data"]["C2"].value is None

    def test_an_unsupported_type_is_refused(self, book: Workbook) -> None:
        with pytest.raises(TypeError, match="not something a worksheet cell can hold"):
            book["Data"]["J8"].value = [1, 2, 3]  # type: ignore[assignment]

    def test_a_timezone_aware_datetime_is_refused(self, book: Workbook) -> None:
        aware = dt.datetime(2026, 7, 4, 12, 0, tzinfo=dt.timezone.utc)
        with pytest.raises(ValueError, match="no Excel serial"):
            book["Data"]["J9"].value = aware


class TestWritingFormulas:
    def test_setting_a_formula(self, book: Workbook) -> None:
        book["Data"]["J10"].formula = "SUM(B2:B5)"
        assert book["Data"]["J10"].formula == "SUM(B2:B5)"
        assert Workbook.from_bytes(book.to_bytes())["Data"]["J10"].formula == "SUM(B2:B5)"

    def test_a_leading_equals_is_accepted_and_dropped(self, book: Workbook) -> None:
        book["Data"]["J11"].formula = "=SUM(B2:B5)"
        assert book["Data"]["J11"].formula == "SUM(B2:B5)"

    def test_a_new_formula_has_no_cached_value(self, book: Workbook) -> None:
        book["Data"]["J12"].formula = "1+1"
        assert book["Data"]["J12"].value is None, "Excel computes it on open"

    def test_changing_a_formula_drops_its_stale_result(self, book: Workbook) -> None:
        assert book["Data"]["D6"].value == 3265
        book["Data"]["D6"].formula = "SUM(B2:B5)"
        assert book["Data"]["D6"].value is None

    def test_removing_a_formula(self, book: Workbook) -> None:
        book["Data"]["D6"].formula = None
        assert book["Data"]["D6"].formula is None

    def test_overwriting_a_shared_follower_gives_it_its_own_text(self, book: Workbook) -> None:
        sheet = book["Data"]
        assert sheet["D3"].formula == "B3*C3"
        sheet["D3"].formula = "B3*C3*2"
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"]["D3"].formula == "B3*C3*2"
        assert reopened["Data"]["D4"].formula == "B4*C4", "its siblings still translate"


class TestStructure:
    def test_dimension_and_used_range(self, book: Workbook) -> None:
        sheet = book["Data"]
        assert sheet.dimension is not None
        assert sheet.dimension.a1 == "A1:F11"
        assert sheet.used_range is not None
        assert sheet.max_row == 11
        assert sheet.max_column == 6

    def test_dimension_grows_with_a_write(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet["Z50"].value = "far away"
        assert sheet.dimension is not None
        assert sheet.dimension.a1 == "A1:Z50"
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"].dimension is not None
        assert reopened["Data"].dimension.a1 == "A1:Z50"

    def test_rows_and_cells_stay_in_order(self, book: Workbook) -> None:
        """Excel refuses a worksheet whose rows or cells are out of order, so
        a cell written into the middle has to be placed, not appended."""
        sheet = book["Data"]
        sheet["A20"].value = "row twenty"
        sheet["A15"].value = "row fifteen"
        sheet["B3"].value = "inserted mid row"

        data = sheet.document.root.require("sheetData")
        numbers = [int(r.get("r") or 0) for r in data.children_named("row")]
        assert numbers == sorted(numbers), f"rows out of order: {numbers}"
        for row in data.children_named("row"):
            from pyofficeeditor.excel import CellRef

            columns = [
                CellRef.parse(c.get("r") or "A1").column for c in row.children_named("c")
            ]
            assert columns == sorted(columns), f"row {row.get('r')} out of order: {columns}"

    def test_a_row_past_the_last_is_appended_even_after_the_last_goes(self, book: Workbook) -> None:
        """A row past the highest is appended rather than placed, which is
        what keeps writing a sheet top to bottom linear. Emptying the last
        row brings the highest back down, so the next row past it still
        lands in order."""
        sheet = book["Data"]
        sheet["A30"].value = "thirty"
        sheet["A25"].value = "twenty-five"
        sheet["A30"].clear()
        sheet["A28"].value = "twenty-eight"
        sheet["A40"].value = "forty"
        data = sheet.document.root.require("sheetData")
        numbers = [int(r.get("r") or 0) for r in data.children_named("row")]
        assert numbers == sorted(numbers)
        assert numbers[-3:] == [25, 28, 40]
        assert sheet.max_row == 40

    def test_iterating_rows(self, book: Workbook) -> None:
        rows = list(book["Data"].rows())
        assert len(rows) == 9, "nine rows carry cells in the fixture"
        assert [c.a1 for c in rows[0]] == ["A1", "B1", "C1", "D1", "E1", "F1"]

    def test_values_pads_the_rectangle(self, book: Workbook) -> None:
        grid = list(book["Data"].values())
        assert len(grid) == 11, "the used rectangle is eleven rows tall"
        assert all(len(row) == 6 for row in grid)
        assert grid[0][0] == "Region"
        assert grid[6][0] is None, "row 7 is empty in the fixture"


class TestRanges:
    def test_reading_a_block(self, book: Workbook) -> None:
        block = book["Data"].range("A1:C2")
        assert block.a1 == "A1:C2"
        assert len(block) == 6
        assert block.values == [["Region", "Units", "Price"], ["North", 120, 4.25]]

    def test_iterating_a_block(self, book: Workbook) -> None:
        assert [c.a1 for c in book["Data"].range("A1:B2")] == ["A1", "B1", "A2", "B2"]

    def test_writing_a_block(self, book: Workbook) -> None:
        book["Data"].range("J1:K2").values = [["a", "b"], [1, 2]]
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"].range("J1:K2").values == [["a", "b"], [1, 2]]

    def test_a_mismatched_shape_is_refused(self, book: Workbook) -> None:
        block = book["Data"].range("J1:K2")
        with pytest.raises(ValueError, match="has 2 rows"):
            block.values = [["only one row"]]
        with pytest.raises(ValueError, match="columns wide"):
            block.values = [["a"], ["b"]]

    def test_clearing_a_block(self, book: Workbook) -> None:
        book["Data"].range("A2:A5").clear()
        assert book["Data"].range("A2:A5").values == [[None], [None], [None], [None]]

    def test_a_reversed_range_normalizes(self, book: Workbook) -> None:
        assert book["Data"].range("C2:A1").a1 == "A1:C2"


class TestSaving:
    def test_a_no_op_save_reproduces_the_input(self, live_sample_xlsx: Path) -> None:
        original = live_sample_xlsx.read_bytes()
        assert Workbook.from_bytes(original).to_bytes() == original

    def test_reading_alone_is_not_a_change(self, book: Workbook, live_sample_xlsx: Path) -> None:
        for sheet in book:
            list(sheet.values())
            for row in sheet.rows():
                for cell in row:
                    _ = cell.value, cell.formula
        assert book.is_modified is False
        assert book.to_bytes() == live_sample_xlsx.read_bytes()

    def test_a_change_is_recorded(self, book: Workbook) -> None:
        assert book.is_modified is False
        book["Data"]["J1"].value = 1
        assert book.is_modified is True

    def test_saving_to_a_path(self, book: Workbook, tmp_path: Path) -> None:
        book["Data"]["J1"].value = "saved"
        target = tmp_path / "out.xlsx"
        assert book.save(target) == target
        assert Workbook.open(target)["Data"]["J1"].value == "saved"

    def test_saving_to_an_unsupported_suffix_is_refused(self, book: Workbook, tmp_path: Path) -> None:
        with pytest.raises(UnsupportedFormatError):
            book.save(tmp_path / "out.xlsb")

    def test_a_change_forces_recalculation_on_open(self, book: Workbook) -> None:
        """A formula cell caches its result, so changing an input leaves a
        wrong number in the file until Excel recalculates."""
        book["Data"]["B2"].value = 200
        text = book.package.read("xl/workbook.xml").decode()
        assert 'fullCalcOnLoad="1"' in text

    def test_a_change_drops_the_calculation_chain(self, book: Workbook) -> None:
        assert book.package.has_part("xl/calcChain.xml"), "the fixture has one"
        book["Data"]["B2"].value = 200
        reopened = Workbook.from_bytes(book.to_bytes())
        assert not reopened.package.has_part("xl/calcChain.xml")
        assert not reopened.package.relationships("xl/workbook.xml").by_type(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
        ), "its relationship went too, or Excel would follow a dangling one"

    def test_an_untouched_workbook_keeps_its_calculation_chain(self, book: Workbook) -> None:
        assert Workbook.from_bytes(book.to_bytes()).package.has_part("xl/calcChain.xml")

    def test_editing_one_cell_leaves_other_parts_stored(
        self, book: Workbook, live_sample_xlsx: Path
    ) -> None:
        from pyofficeeditor._zip import ZipArchive

        book["Data"]["B2"].value = 200
        before = ZipArchive.from_bytes(live_sample_xlsx.read_bytes())
        after = ZipArchive.from_bytes(book.to_bytes())
        expected_changes = {
            # The cell that changed.
            "xl/worksheets/sheet1.xml",
            # Gained fullCalcOnLoad, and lost the calcChain relationship.
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            # calcChain went, so its content-type override went with it.
            "[Content_Types].xml",
        }
        for name in after.names():
            if name in expected_changes:
                continue
            assert after.member(name).stored == before.member(name).stored, name
        assert "xl/calcChain.xml" not in after.names()

    def test_the_string_reference_count_is_set_at_save(self, book: Workbook) -> None:
        """``uniqueCount`` follows the table as entries are added.  ``count``
        is the number of cells pointing at it, which takes a walk over the
        whole workbook, so it is computed once at save rather than on every
        write."""
        book["Data"]["J1"].value = "one more string cell"
        assert 'uniqueCount="15"' in book.package.read("xl/sharedStrings.xml").decode()
        book.to_bytes()
        text = book.package.read("xl/sharedStrings.xml").decode()
        assert 'uniqueCount="15"' in text
        assert 'count="15"' in text, "fourteen string cells plus the new one"


class TestSerialDates:
    @pytest.mark.parametrize(
        ("serial", "expected"),
        [
            (1, dt.date(1900, 1, 1)),
            (2, dt.date(1900, 1, 2)),
            (59, dt.date(1900, 2, 28)),
            (61, dt.date(1900, 3, 1)),
            (36526, dt.date(2000, 1, 1)),
            (46037, dt.date(2026, 1, 15)),
        ],
    )
    def test_the_1900_system(self, serial: int, expected: dt.date) -> None:
        assert serial_to_datetime(serial).date() == expected

    def test_the_phantom_leap_day_is_refused(self) -> None:
        """Excel reserves serial 60 for 29 February 1900, which never
        happened. Reporting it as 1 March would be a silent lie."""
        with pytest.raises(DateOutOfRangeError, match="never existed"):
            serial_to_datetime(60)

    def test_the_fractional_part_is_the_time(self) -> None:
        assert serial_to_datetime(46037.5) == dt.datetime(2026, 1, 15, 12, 0)

    def test_the_1904_system(self) -> None:
        assert serial_to_datetime(0, epoch_1904=True).date() == dt.date(1904, 1, 1)
        assert serial_to_datetime(1, epoch_1904=True).date() == dt.date(1904, 1, 2)

    @pytest.mark.parametrize(
        "value",
        [
            dt.date(1900, 1, 1),
            dt.date(1900, 3, 1),
            dt.date(2000, 2, 29),
            dt.date(2026, 1, 15),
            dt.datetime(2026, 1, 15, 6, 30),
            dt.datetime(9999, 12, 31),
        ],
    )
    def test_dates_round_trip(self, value: dt.date | dt.datetime) -> None:
        serial = datetime_to_serial(value)
        back = serial_to_datetime(serial)
        expected = value if isinstance(value, dt.datetime) else dt.datetime(value.year, value.month, value.day)
        assert back == expected

    def test_a_date_before_the_epoch_is_refused(self) -> None:
        with pytest.raises(DateOutOfRangeError, match="1900"):
            datetime_to_serial(dt.date(1899, 12, 31))

    def test_a_bare_time_is_a_fraction_of_a_day(self) -> None:
        assert datetime_to_serial(dt.time(12, 0)) == 0.5
        assert datetime_to_serial(dt.time(6, 0)) == 0.25
        assert datetime_to_serial(dt.time(0, 0)) == 0.0


class TestNumberFormatting:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            (0, "0"),
            (42, "42"),
            (-7, "-7"),
            (4.25, "4.25"),
            (3.0, "3"),
            (0.1, "0.1"),
            (1 / 3, "0.3333333333333333"),
        ],
    )
    def test_numbers_are_written_as_excel_writes_them(self, value: object, expected: str) -> None:
        assert format_number(value) == expected  # type: ignore[arg-type]

    def test_a_float_round_trips_exactly(self) -> None:
        for value in (0.1, 1 / 3, 1e-10, 123456789.123456789):
            assert float(format_number(value)) == value

    def test_nan_and_infinity_are_refused(self) -> None:
        for value in (float("nan"), float("inf"), float("-inf")):
            with pytest.raises(ValueError, match="no worksheet representation"):
                format_number(value)
