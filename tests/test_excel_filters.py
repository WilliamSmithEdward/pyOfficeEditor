"""The autofilter on a sheet, through the worksheet.

``filters.xlsx`` was built by Excel with one criterion kind per sheet, and
``filters_answers.json`` beside it records what Excel answered for each and
how many rows it hid. What each criterion keeps is held to Excel case by
case in ``test_excel_filter_semantics.py``; this module holds what the
worksheet does with the rules: reading every kind, re-applying Excel's own
filters to Excel's own data, and writing criteria so the file opens the way
Excel itself would have left it.

That last part carries weight because a filter is stored twice, the
criteria in ``<autoFilter>`` and a ``hidden`` flag on each row, and Excel
does not re-apply the first when the workbook opens. It shows the rows the
flags say.
"""

from __future__ import annotations

import datetime as dt
import json
import shutil
import time
from pathlib import Path

import pytest

from pyofficeeditor._xml import Element, XmlDocument
from pyofficeeditor.excel import (
    AutoFilter,
    Comparison,
    CustomFilter,
    DateGroup,
    DynamicFilter,
    FilterColumn,
    OpaqueCriterion,
    Top10Filter,
    ValueFilter,
    Workbook,
    Worksheet,
    criteria,
)
from pyofficeeditor.excel._filters import FilterCell, keeps

#: Every sheet in the fixture, and the criterion kind it carries.
SHEETS = {
    "Values": ValueFilter,
    "Compare": CustomFilter,
    "Between": CustomFilter,
    "TopTen": Top10Filter,
    "Blanks": ValueFilter,
    "Dynamic": DynamicFilter,
    "Numbers": ValueFilter,
    "Dates": ValueFilter,
}

#: Any day will do where no relative date period is involved.
TODAY = dt.date(2026, 3, 10)


@pytest.fixture(scope="module")
def book(live_filters_xlsx: Path) -> Workbook:
    return Workbook.open(live_filters_xlsx)


@pytest.fixture(scope="module")
def answers(live_filters_answers: Path) -> dict[str, dict[str, object]]:
    return json.loads(live_filters_answers.read_text(encoding="utf-8"))


@pytest.fixture
def filters_copy(tmp_path: Path, live_filters_xlsx: Path) -> Workbook:
    target = tmp_path / "filters.xlsx"
    shutil.copy(live_filters_xlsx, target)
    return Workbook.open(target)


@pytest.fixture
def data(tmp_path: Path, live_sample_xlsx: Path) -> Worksheet:
    """sample.xlsx's Data sheet: numbers in B with blanks in B6, B7 and B9
    and #DIV/0! in B8, formulas with cached results in D, dates in F."""
    target = tmp_path / "sample.xlsx"
    shutil.copy(live_sample_xlsx, target)
    return Workbook.open(target)["Data"]


def hidden(sheet: Worksheet, rows: range) -> set[int]:
    return {row for row in rows if sheet.row_hidden(row)}


def only_criterion(sheet: Worksheet) -> object:
    found = sheet.auto_filter
    assert found is not None, f"{sheet.name} should carry an autofilter"
    assert len(found.columns) == 1
    return found.columns[0].criterion


class TestReadingEachKind:
    def test_every_sheet_has_one(self, book: Workbook) -> None:
        for name in SHEETS:
            found = book[name].auto_filter
            assert found is not None, name
            assert found.ref == "A1:C21", name

    def test_each_kind_reads_as_its_own_type(self, book: Workbook) -> None:
        for name, expected in SHEETS.items():
            assert isinstance(only_criterion(book[name]), expected), name

    def test_a_value_list(self, book: Workbook) -> None:
        assert only_criterion(book["Values"]) == ValueFilter(("r2", "r3", "r4"))

    def test_blanks_are_an_attribute_not_a_value(self, book: Workbook) -> None:
        assert only_criterion(book["Blanks"]) == ValueFilter(blank=True)

    def test_one_comparison(self, book: Workbook) -> None:
        assert only_criterion(book["Compare"]) == CustomFilter((Comparison("greaterThanOrEqual", "10"),))

    def test_two_comparisons_joined_by_and(self, book: Workbook) -> None:
        assert only_criterion(book["Between"]) == CustomFilter(
            (Comparison("greaterThanOrEqual", "5"), Comparison("lessThanOrEqual", "40")), require_all=True
        )

    def test_top_ten_carries_both_numbers(self, book: Workbook) -> None:
        """The 3 that was asked for, and the 47 Excel ranked to."""
        assert only_criterion(book["TopTen"]) == Top10Filter(3, threshold=47)

    def test_a_dynamic_filter_carries_its_aggregate(self, book: Workbook) -> None:
        assert only_criterion(book["Dynamic"]) == DynamicFilter("aboveAverage", 28)

    def test_a_date_through_the_object_model_keeps_its_text(self, book: Workbook) -> None:
        assert only_criterion(book["Dates"]) == ValueFilter(("1/5/2026", "3/5/2026"))


class TestReapplyingExcelsOwnFilters:
    """Excel's criteria over Excel's data, applied again here, have to hide
    the rows Excel hid. The flags are cleared first, so nothing survives
    from the file."""

    def test_the_fixture_really_does_hide_rows(self, book: Workbook, answers: dict[str, dict[str, object]]) -> None:
        for name in SHEETS:
            assert len(hidden(book[name], range(2, 22))) == answers[name]["hidden"], name
            assert hidden(book[name], range(2, 22)), name

    @pytest.mark.parametrize("name", sorted(SHEETS))
    def test_the_same_rows_are_hidden(self, filters_copy: Workbook, name: str) -> None:
        sheet = filters_copy[name]
        excel = hidden(sheet, range(1, 23))
        for row in range(1, 23):
            if sheet.row_hidden(row):
                sheet.set_row_hidden(row, False)
        outcome = sheet.apply_auto_filter(today=TODAY)
        assert hidden(sheet, range(1, 23)) == excel
        assert set(outcome.hidden) == excel
        assert not outcome.undecided

    def test_a_stale_threshold_is_worked_out_again(self, filters_copy: Workbook) -> None:
        """As Excel's ApplyFilter does, and what it stores follows."""
        sheet = filters_copy["TopTen"]
        element = sheet.document.root.require("autoFilter").require("filterColumn").require("top10")
        element.set("filterVal", "1")
        sheet.apply_auto_filter(today=TODAY)
        assert only_criterion(sheet) == Top10Filter(3, threshold=47)


class TestWriting:
    def test_a_value_list_hides_what_it_excludes(self, data: Worksheet) -> None:
        outcome = data.set_auto_filter("A1:D5", [FilterColumn(0, ValueFilter(("North", "West")))])
        assert hidden(data, range(1, 6)) == {3, 4}
        assert outcome.hidden == (3, 4)
        assert outcome.shown == (2, 5)

    def test_numbers_blanks_and_errors_compare_as_excel_compares_them(self, data: Worksheet) -> None:
        """Only numbers pass an ordering against a number: blanks and the
        error do not."""
        data.set_auto_filter("A1:B9", [FilterColumn(1, criteria("<100"))])
        assert hidden(data, range(2, 10)) == {2, 3, 5, 6, 7, 8, 9}

    def test_not_blank_is_its_own_spelling(self, data: Worksheet) -> None:
        """``<>`` alone keeps everything but the blanks, the error included."""
        data.set_auto_filter("A1:B9", [FilterColumn(1, criteria("<>"))])
        assert hidden(data, range(2, 10)) == {6, 7, 9}

    def test_a_value_is_matched_as_the_text_the_cell_shows(self, data: Worksheet) -> None:
        cell_format = data["C4"].format
        data["C4"].format = type(cell_format)(number_format="0.00")
        data.set_auto_filter("A1:C5", [FilterColumn(2, ValueFilter(("5.50",)))])
        assert hidden(data, range(2, 6)) == {2, 3, 5}
        data.set_auto_filter("A1:C5", [FilterColumn(2, ValueFilter(("5.5",)))])
        assert hidden(data, range(2, 6)) == {2, 3, 4, 5}

    def test_wildcards_match_text(self, filters_copy: Workbook) -> None:
        """Excel stores begins-with as equal and a wildcard."""
        sheet = filters_copy["Values"]
        sheet.set_auto_filter("A1:C21", [FilterColumn(0, criteria("=r1*"))])
        kept = {row for row in range(2, 22) if not sheet.row_hidden(row)}
        assert kept == {10, 11, 12, 13, 14, 16, 17, 18, 19}

    def test_dates_are_numbers_to_a_top_ten_and_a_comparison(self, filters_copy: Workbook) -> None:
        sheet = filters_copy["Dates"]
        sheet.set_auto_filter("A1:C21", [FilterColumn(2, Top10Filter(3))], today=TODAY)
        assert {row for row in range(2, 22) if not sheet.row_hidden(row)} == {11, 12, 13}
        sheet.set_auto_filter("A1:C21", [FilterColumn(2, criteria(">=10/1/2026"))], today=TODAY)
        assert {row for row in range(2, 22) if not sheet.row_hidden(row)} == {11, 12, 13}

    def test_a_date_period_is_measured_from_today(self, filters_copy: Workbook) -> None:
        sheet = filters_copy["Dates"]
        sheet.set_auto_filter("A1:C21", [FilterColumn(2, DynamicFilter("thisMonth"))], today=TODAY)
        assert {row for row in range(2, 22) if not sheet.row_hidden(row)} == {4, 16}
        stored = only_criterion(sheet)
        assert stored == DynamicFilter("thisMonth", 46082, 46113)

    def test_a_date_group(self, filters_copy: Workbook) -> None:
        sheet = filters_copy["Dates"]
        group = DateGroup("month", 2026, 3)
        sheet.set_auto_filter("A1:C21", [FilterColumn(2, ValueFilter(date_groups=(group,)))])
        assert {row for row in range(2, 22) if not sheet.row_hidden(row)} == {4, 16}
        assert only_criterion(sheet) == ValueFilter(date_groups=(group,))

    def test_apply_false_touches_no_row(self, data: Worksheet) -> None:
        data.set_auto_filter("A1:D5", [FilterColumn(0, ValueFilter(("Nonesuch",)))], apply=False)
        assert not hidden(data, range(1, 12))

    def test_it_survives_a_round_trip(self, data: Worksheet, tmp_path: Path) -> None:
        wanted = [
            FilterColumn(0, ValueFilter(("North", "South"), blank=True)),
            FilterColumn(1, CustomFilter((Comparison("greaterThan", 5), Comparison("lessThan", 900)), require_all=True)),
        ]
        data.set_auto_filter("A1:D5", wanted)
        data.workbook.save()
        with Workbook.open(tmp_path / "sample.xlsx") as reopened:
            found = reopened["Data"].auto_filter
            assert found is not None
            assert found.columns[0] == wanted[0]
            # A number is stored as the text Excel stores.
            assert found.columns[1].criterion == CustomFilter(
                (Comparison("greaterThan", "5"), Comparison("lessThan", "900")), require_all=True
            )
            assert hidden(reopened["Data"], range(1, 6)) == hidden(data, range(1, 6))


class TestTheRange:
    def test_it_stops_at_the_last_row_holding_data(self, data: Worksheet) -> None:
        """Row 11 holds the merged heading; nothing below it does."""
        data.set_auto_filter("A1:D60")
        found = data.auto_filter
        assert found is not None
        assert found.ref == "A1:D11"

    def test_a_whole_column_range_is_quick_and_adds_no_rows(self, data: Worksheet) -> None:
        rows_before = max(data.rows_by_number())
        started = time.perf_counter()
        data.set_auto_filter("A1:D1048576", [FilterColumn(0, ValueFilter(("North",)))])
        data.clear_auto_filter()
        assert time.perf_counter() - started < 5
        assert max(data.rows_by_number()) == rows_before

    def test_rows_a_filter_hides_do_not_move_the_end_of_the_data(self, data: Worksheet) -> None:
        """``max_row`` is the last row with a cell. A row hidden below the
        data has an element and nothing in it."""
        data.set_row_hidden(40, True)
        assert data.max_row == 11
        used = data.used_range
        assert used is not None and used.a1 == "A1:F11"

    def test_arrows_alone_touch_no_row(self, data: Worksheet) -> None:
        """A row hidden by hand stays hidden when the arrows go on, as in
        Excel."""
        data.set_row_hidden(3, True)
        outcome = data.set_auto_filter("A1:D5")
        assert hidden(data, range(1, 6)) == {3}
        assert outcome.hidden == outcome.shown == ()

    def test_a_row_hidden_by_hand_that_the_criteria_keep_is_shown(self, data: Worksheet) -> None:
        data.set_row_hidden(2, True)
        data.set_auto_filter("A1:D5", [FilterColumn(0, ValueFilter(("North", "East")))])
        assert hidden(data, range(1, 6)) == {3, 5}

    def test_filtering_another_range_takes_the_old_filter_off_first(self, filters_copy: Workbook) -> None:
        """Rows 5 to 21 were hidden by the old filter; none of them is in
        the new range's criteria, so all of them show."""
        sheet = filters_copy["Values"]
        sheet.set_auto_filter("A1:C10")
        assert not hidden(sheet, range(1, 22))
        found = sheet.auto_filter
        assert found is not None
        assert found.ref == "A1:C10"

    def test_clearing_shows_every_row_and_keeps_the_filter_database(self, filters_copy: Workbook) -> None:
        sheet = filters_copy["Values"]
        sheet.set_auto_filter("A1:C21", [FilterColumn(0, ValueFilter(("r2",)))])
        sheet.clear_auto_filter()
        assert sheet.auto_filter is None
        assert not hidden(sheet, range(1, 22))
        assert filters_copy.defined_name("_xlnm._FilterDatabase", scope="Values").hidden

    def test_clearing_can_leave_the_rows_hidden(self, data: Worksheet) -> None:
        data.set_auto_filter("A1:D5", [FilterColumn(0, ValueFilter(("North",)))])
        data.clear_auto_filter(show_rows=False)
        assert hidden(data, range(1, 6)) == {3, 4, 5}

    def test_clearing_nothing_is_not_an_error(self, data: Worksheet) -> None:
        data.clear_auto_filter()
        assert data.auto_filter is None

    def test_showing_a_collapsed_group_opens_it(self, data: Worksheet) -> None:
        """Rows 3 and 4 in a collapsed group, summarised by row 5: a filter
        that keeps them shows them, and the group is no longer collapsed."""
        for row in (3, 4):
            element = data.rows_by_number()[row]
            element.set("outlineLevel", "1")
            element.set("hidden", "1")
        data.rows_by_number()[5].set("collapsed", "1")
        data.set_auto_filter("A1:D5", [FilterColumn(0, ValueFilter(("South", "East", "West")))])
        assert hidden(data, range(1, 6)) == {2}
        assert data.rows_by_number()[5].get("collapsed") is None


class TestWhatExcelWritesAlongside:
    def test_filter_mode_is_set_while_criteria_hide_rows(self, data: Worksheet) -> None:
        properties = data.document.root.child("sheetPr")
        assert properties is None or properties.get("filterMode") is None
        data.set_auto_filter("A1:D5", [FilterColumn(0, ValueFilter(("North",)))])
        assert data.document.root.require("sheetPr").get("filterMode") == "1"
        data.clear_auto_filter()
        properties = data.document.root.child("sheetPr")
        assert properties is None or properties.get("filterMode") is None

    def test_the_filter_database_name_follows_the_range(self, data: Worksheet) -> None:
        data.set_auto_filter("A1:D5")
        name = data.workbook.defined_name("_xlnm._FilterDatabase", scope="Data")
        assert name.refers_to == "Data!$A$1:$D$5"
        assert name.hidden


class TestRefusals:
    """Each refused before anything changed, so the sheet is as it was."""

    def unchanged(self, sheet: Worksheet, before: AutoFilter | None, rows: set[int]) -> None:
        assert sheet.auto_filter == before
        assert hidden(sheet, range(1, 23)) == rows

    def test_a_column_outside_the_range(self, filters_copy: Workbook) -> None:
        sheet = filters_copy["Values"]
        before, rows = sheet.auto_filter, hidden(sheet, range(1, 23))
        with pytest.raises(ValueError, match="outside"):
            sheet.set_auto_filter("A1:C21", [FilterColumn(3, ValueFilter(("x",)))])
        self.unchanged(sheet, before, rows)

    def test_a_column_given_twice(self, filters_copy: Workbook) -> None:
        sheet = filters_copy["Values"]
        before, rows = sheet.auto_filter, hidden(sheet, range(1, 23))
        with pytest.raises(ValueError, match="twice"):
            sheet.set_auto_filter("A1:C21", [FilterColumn(0, ValueFilter(("x",))), FilterColumn(0, ValueFilter(("y",)))])
        self.unchanged(sheet, before, rows)

    def test_a_top_ten_over_an_error(self, data: Worksheet) -> None:
        with pytest.raises(ValueError, match="error"):
            data.set_auto_filter("A1:B9", [FilterColumn(1, Top10Filter(2))])
        assert data.auto_filter is None

    def test_a_range_overlapping_a_table(self, tmp_path: Path, live_structures_xlsx: Path) -> None:
        target = tmp_path / "structures.xlsx"
        shutil.copy(live_structures_xlsx, target)
        sheet = Workbook.open(target)["Tabled"]
        for reference in ("A1:C5", "C1:D5", "A5:B5"):
            with pytest.raises(ValueError, match="table"):
                sheet.set_auto_filter(reference)
        assert sheet.auto_filter is None

    def test_a_table_overlapping_the_filter(self, data: Worksheet) -> None:
        data.set_auto_filter("A1:D5")
        with pytest.raises(ValueError, match="autofilter"):
            data.add_table("Sales", "C1:E5")


class TestKeepingWhatIsNotModelled:
    def add_column(self, sheet: Worksheet, markup: str) -> None:
        element = sheet.document.root.require("autoFilter")
        element.append(XmlDocument.parse(markup.encode()).root)

    def test_a_colour_filter_survives_a_rewrite_of_its_neighbours(self, filters_copy: Workbook) -> None:
        sheet = filters_copy["Values"]
        self.add_column(sheet, '<filterColumn colId="1"><colorFilter dxfId="0"/></filterColumn>')
        found = sheet.auto_filter
        assert found is not None
        colour = found.column(1)
        assert colour is not None and isinstance(colour.criterion, OpaqueCriterion)
        assert colour.criterion.kind == "colorFilter"
        sheet.set_auto_filter(found.ref, [*found.columns, FilterColumn(2, ValueFilter(("1/5/2026",)))], today=TODAY)
        again = sheet.auto_filter
        assert again is not None
        assert again.column(1) == colour
        markup = sheet.document.root.require("autoFilter").to_xml()
        assert '<filterColumn colId="1"><colorFilter dxfId="0"/></filterColumn>' in markup

    def test_a_row_only_an_unevaluated_criterion_could_hide_is_left_alone(self, filters_copy: Workbook) -> None:
        sheet = filters_copy["Values"]
        self.add_column(sheet, '<filterColumn colId="1"><iconFilter iconSet="3Arrows" iconId="0"/></filterColumn>')
        found = sheet.auto_filter
        assert found is not None
        sheet.set_row_hidden(3, True)
        outcome = sheet.apply_auto_filter(today=TODAY)
        assert 3 in outcome.undecided
        assert sheet.row_hidden(3), "the icon filter might hide it, so it stays hidden"
        assert not sheet.row_hidden(2)
        assert outcome.unevaluated == (1,)
        assert all(sheet.row_hidden(row) for row in range(5, 22)), "the value list still hides these"

    def test_an_unchanged_column_keeps_its_markup(self, filters_copy: Workbook) -> None:
        """Excel's own spelling of a date group, an extension, is read and
        not rewritten when the column did not change."""
        sheet = filters_copy["Dates"]
        element = sheet.document.root.require("autoFilter")
        element.remove(element.require("filterColumn"))
        extension = (
            '<filterColumn colId="2"><extLst><ext uri="{1AD28BCE-077C-4C59-8B6E-1921CE8616D4}" '
            'xmlns:xlrd2="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2">'
            '<xlrd2:filterColumn><xlrd2:filters><xlrd2:dateGroupItem year="2026" month="3" '
            'dateTimeGrouping="month"/></xlrd2:filters></xlrd2:filterColumn></ext></extLst></filterColumn>'
        )
        self.add_column(sheet, extension)
        found = sheet.auto_filter
        assert found is not None
        assert found.columns[0].criterion == ValueFilter(date_groups=(DateGroup("month", 2026, 3),))
        sheet.set_auto_filter(found.ref, [*found.columns, FilterColumn(0, ValueFilter(("r4",)))])
        assert extension in sheet.document.root.require("autoFilter").to_xml()
        assert {row for row in range(2, 22) if not sheet.row_hidden(row)} == {4}

    def test_columns_go_before_the_sort_state(self, filters_copy: Workbook) -> None:
        """Excel refuses a workbook with a filter column after it."""
        sheet = filters_copy["Values"]
        element = sheet.document.root.require("autoFilter")
        element.append(Element.create("sortState", {"ref": "A2:C21"}))
        sheet.set_auto_filter("A1:C21", [FilterColumn(0, ValueFilter(("r2",))), FilterColumn(1, criteria(">5"))])
        names = [child.name for child in element.elements()]
        assert names == ["filterColumn", "filterColumn", "sortState"]


class TestFormulas:
    """A formula's cached result is what Excel computed last. Once a value
    changes it may be out of date, and a row decided on it is left alone."""

    def test_a_cached_result_is_filtered_on(self, data: Worksheet) -> None:
        data.set_auto_filter("A1:D5", [FilterColumn(3, criteria(">600"))])
        assert hidden(data, range(1, 6)) == {2, 4}

    def test_after_an_edit_a_formula_row_is_left_as_it_was(self, data: Worksheet) -> None:
        data["B2"].value = 1
        outcome = data.set_auto_filter("A1:D5", [FilterColumn(3, criteria(">600"))])
        assert outcome.undecided == (2, 3, 4, 5)
        assert not hidden(data, range(1, 6))


class TestColumnsMoving:
    """Excel renumbers a filter column as columns come and go, and drops
    one whose column is deleted."""

    def colid(self, sheet: Worksheet) -> list[int]:
        found = sheet.auto_filter
        assert found is not None
        return [entry.column for entry in found.columns]

    def test_inserting_inside_the_range_renumbers(self, filters_copy: Workbook) -> None:
        sheet = filters_copy["Compare"]
        sheet.insert_columns(2)
        assert self.colid(sheet) == [2]
        found = sheet.auto_filter
        assert found is not None and found.ref == "A1:D21"

    def test_deleting_to_the_left_renumbers(self, filters_copy: Workbook) -> None:
        sheet = filters_copy["Compare"]
        sheet.delete_columns(1)
        assert self.colid(sheet) == [0]

    def test_deleting_the_filtered_column_drops_its_criterion(self, filters_copy: Workbook) -> None:
        sheet = filters_copy["Compare"]
        assert hidden(sheet, range(1, 22)) == {8, 15}
        sheet.delete_columns(2)
        assert self.colid(sheet) == []
        assert not hidden(sheet, range(1, 22)), "the rows only it hid show"


NAMES = ("a", "b", "c", "d", "e", "f", "g")
AMOUNTS = (50, 150, 75, 300, 100, 101, 20)


def fill(sheet: Worksheet, left: int, width: int = 2) -> None:
    """The block Excel's table measurements used, from column ``left``:
    names a to g, their amounts and a running number, under a header."""
    columns = (("Name", NAMES), ("Amount", AMOUNTS), ("When", tuple(range(1, 8))))
    for offset, (header, values) in enumerate(columns[:width]):
        sheet.cell(1, left + offset).value = header
        for row, value in enumerate(values, start=2):
            sheet.cell(row, left + offset).value = value


class TestTableFilters:
    """A table's own filter, on the sheet filter's machinery. What Excel
    does around it was measured: it writes no ``filterMode`` and no filter
    database, Clear keeps the dropdowns, unticking Filter Button drops them
    and shows every row, and a filter decides the rows of its own range
    whatever another filter did to them."""

    def sales(self, book: Workbook, *, width: int = 2, totals: bool = False) -> Worksheet:
        sheet = book.add_sheet("Tabled")
        fill(sheet, 1, width)
        if totals:
            sheet.cell(9, 1).value = "Total"
        sheet.add_table("Sales", f"A1:{'ABC'[width - 1]}{9 if totals else 8}", totals_row=totals)
        return sheet

    @pytest.mark.parametrize("name", sorted(SHEETS))
    def test_excels_criteria_hide_what_they_hid_on_the_sheet(self, filters_copy: Workbook, name: str) -> None:
        sheet = filters_copy[name]
        found = sheet.auto_filter
        assert found is not None
        excel = hidden(sheet, range(1, 23))
        sheet.clear_auto_filter()
        sheet.add_table("Checked", "A1:C21")
        outcome = sheet.set_table_filter("Checked", found.columns, today=TODAY)
        assert hidden(sheet, range(1, 23)) == excel
        assert set(outcome.hidden) == excel

    def test_reading_one(self, live_structures_xlsx: Path) -> None:
        """Excel's own: the totals row is left out of the filter's range."""
        table = Workbook.open(live_structures_xlsx)["Tabled"].table("SalesTable")
        assert table.ref.a1 == "A1:C5"
        assert table.auto_filter == AutoFilter("A1:C4")

    def test_the_totals_row_is_left_out(self, filters_copy: Workbook) -> None:
        sheet = self.sales(filters_copy, totals=True)
        sheet.set_table_filter("Sales", [FilterColumn(1, criteria(">100"))])
        assert hidden(sheet, range(1, 10)) == {2, 4, 6, 8}
        found = sheet.table("Sales").auto_filter
        assert found is not None and found.ref == "A1:B8"

    def test_no_filter_mode_and_no_filter_database(self, filters_copy: Workbook) -> None:
        sheet = self.sales(filters_copy)
        sheet.set_table_filter("Sales", [FilterColumn(1, criteria(">100"))])
        properties = sheet.document.root.child("sheetPr")
        assert properties is None or properties.get("filterMode") is None
        assert not [
            name for name in filters_copy.defined_names
            if name.name == "_xlnm._FilterDatabase" and name.scope == "Tabled"
        ]

    def test_no_columns_is_excels_clear(self, filters_copy: Workbook) -> None:
        sheet = self.sales(filters_copy)
        sheet.set_table_filter("Sales", [FilterColumn(1, criteria(">100"))])
        outcome = sheet.set_table_filter("Sales")
        assert not hidden(sheet, range(1, 10))
        assert outcome.shown == (2, 4, 6, 8)
        assert sheet.table("Sales").auto_filter == AutoFilter("A1:B8"), "the dropdowns stay"

    def test_clearing_turns_the_dropdowns_off_and_shows_every_row(self, filters_copy: Workbook) -> None:
        sheet = self.sales(filters_copy)
        sheet.set_table_filter("Sales", [FilterColumn(1, criteria(">100"))])
        sheet.clear_table_filter("Sales")
        assert sheet.table("Sales").auto_filter is None
        assert not hidden(sheet, range(1, 10))
        sheet.set_table_filter("Sales", [FilterColumn(1, criteria(">100"))])
        assert hidden(sheet, range(1, 10)) == {2, 4, 6, 8}, "and filtering turns them on again"

    def test_clearing_can_leave_the_rows_hidden(self, filters_copy: Workbook) -> None:
        sheet = self.sales(filters_copy)
        sheet.set_table_filter("Sales", [FilterColumn(1, criteria(">100"))])
        sheet.clear_table_filter("Sales", show_rows=False)
        assert sheet.table("Sales").auto_filter is None
        assert hidden(sheet, range(1, 10)) == {2, 4, 6, 8}

    def test_each_filter_decides_the_rows_of_its_own_range(self, filters_copy: Workbook) -> None:
        """Measured: the sheet's filter hides b and d, then the table beside
        it hides d and f, and b shows again because the table's filter came
        last and decided every row of its range. Taking the sheet's filter
        off then shows the table's rows too."""
        sheet = filters_copy.add_sheet("Both")
        fill(sheet, 1)
        fill(sheet, 5)
        sheet.add_table("Beside", "E1:F8")
        sheet.set_auto_filter("A1:B8", [FilterColumn(0, criteria("<>b", "<>d"))])
        assert hidden(sheet, range(1, 10)) == {3, 5}
        sheet.set_table_filter("Beside", [FilterColumn(0, criteria("<>d", "<>f"))])
        assert hidden(sheet, range(1, 10)) == {5, 7}
        sheet.clear_auto_filter()
        assert not hidden(sheet, range(1, 10))
        sheet.apply_table_filter("Beside")
        assert hidden(sheet, range(1, 10)) == {5, 7}

    def test_a_headerless_table_is_refused(self, filters_copy: Workbook) -> None:
        """Excel turns the header row back on to filter; this library does
        not add rows to a table."""
        sheet = self.sales(filters_copy)
        root = sheet.table("Sales").document.root
        root.set("headerRowCount", "0")
        root.remove(root.require("autoFilter"))
        with pytest.raises(ValueError, match="header row"):
            sheet.set_table_filter("Sales", [FilterColumn(1, criteria(">100"))])
        assert sheet.table("Sales").auto_filter is None
        assert not hidden(sheet, range(1, 10))

    def test_a_column_outside_the_table_is_refused(self, filters_copy: Workbook) -> None:
        sheet = self.sales(filters_copy)
        with pytest.raises(ValueError, match="outside"):
            sheet.set_table_filter("Sales", [FilterColumn(2, criteria(">100"))])
        assert sheet.table("Sales").auto_filter == AutoFilter("A1:B8")

    def test_applying_again_works_a_top_ten_out_again(self, filters_copy: Workbook) -> None:
        sheet = self.sales(filters_copy)
        sheet.set_table_filter("Sales", [FilterColumn(1, Top10Filter(2))])
        assert hidden(sheet, range(1, 10)) == {2, 4, 6, 7, 8}
        sheet.cell(2, 2).value = 1000
        sheet.apply_table_filter("Sales")
        assert hidden(sheet, range(1, 10)) == {3, 4, 6, 7, 8}
        found = sheet.table("Sales").auto_filter
        assert found is not None and found.columns[0].criterion == Top10Filter(2, threshold=300)

    def test_applying_again_with_the_dropdowns_off_is_refused(self, filters_copy: Workbook) -> None:
        sheet = self.sales(filters_copy)
        sheet.clear_table_filter("Sales")
        with pytest.raises(ValueError, match="dropdowns"):
            sheet.apply_table_filter("Sales")

    def test_deleting_a_filtered_column_applies_the_filter_again(self, filters_copy: Workbook) -> None:
        """Measured: criteria on the names and the amounts hide rows 2, 3
        and 8, and deleting the names column shows row 3."""
        sheet = self.sales(filters_copy, width=3)
        sheet.set_table_filter("Sales", [FilterColumn(0, criteria("<>b")), FilterColumn(1, criteria(">60"))])
        assert hidden(sheet, range(1, 10)) == {2, 3, 8}
        sheet.delete_columns(1)
        assert hidden(sheet, range(1, 10)) == {2, 8}
        found = sheet.table("Sales").auto_filter
        assert found is not None and [entry.column for entry in found.columns] == [0]

    def test_deleting_the_last_criterion_shows_every_row(self, filters_copy: Workbook) -> None:
        sheet = self.sales(filters_copy, width=3)
        sheet.set_table_filter("Sales", [FilterColumn(0, criteria("<>b"))])
        assert hidden(sheet, range(1, 10)) == {3}
        sheet.delete_columns(1)
        assert not hidden(sheet, range(1, 10))

    def test_deleting_another_column_leaves_the_rows(self, filters_copy: Workbook) -> None:
        sheet = self.sales(filters_copy, width=3)
        sheet.set_table_filter("Sales", [FilterColumn(0, criteria("<>b")), FilterColumn(1, criteria(">60"))])
        sheet.delete_columns(3)
        assert hidden(sheet, range(1, 10)) == {2, 3, 8}

    def test_it_survives_a_round_trip(self, filters_copy: Workbook, tmp_path: Path) -> None:
        sheet = self.sales(filters_copy, totals=True)
        sheet.set_table_filter("Sales", [FilterColumn(1, criteria(">100"))])
        filters_copy.save()
        with Workbook.open(tmp_path / "filters.xlsx") as reopened:
            again = reopened["Tabled"]
            # A number is stored as the text Excel stores.
            stored = CustomFilter((Comparison("greaterThan", "100"),))
            assert again.table("Sales").auto_filter == AutoFilter("A1:B8", (FilterColumn(1, stored),))
            assert hidden(again, range(1, 10)) == {2, 4, 6, 8}


class TestClearingACell:
    def test_a_hidden_row_stays_hidden_when_its_last_cell_goes(self, data: Worksheet) -> None:
        data.set_auto_filter("A1:D5", [FilterColumn(0, ValueFilter(("North",)))])
        for letter in "ABCDEF":
            data[f"{letter}3"].clear()
        assert data.row_hidden(3)


class TestModels:
    """What Excel refuses to open, or silently ignores, is refused here."""

    @pytest.mark.parametrize(
        "make",
        [
            lambda: Comparison("between", "5"),  # type: ignore[arg-type]
            lambda: Comparison("equal", float("nan")),
            lambda: Comparison("equal", object()),  # type: ignore[arg-type]
            lambda: CustomFilter(()),
            lambda: CustomFilter((Comparison(), Comparison(), Comparison())),
            lambda: ValueFilter(("",)),
            lambda: ValueFilter(),
            lambda: ValueFilter("North"),  # type: ignore[arg-type]
            lambda: ValueFilter((5,)),  # type: ignore[arg-type]
            lambda: Top10Filter(0),
            lambda: Top10Filter(501),
            lambda: Top10Filter(101, percent=True),
            lambda: Top10Filter(2.5),  # type: ignore[arg-type]
            lambda: DynamicFilter("someday"),  # type: ignore[arg-type]
            lambda: DateGroup("day", 2026, 2, 30),
            lambda: DateGroup("month", 2026),
            lambda: DateGroup("year", 2026, 3),
            lambda: FilterColumn(-1),
        ],
    )
    def test_refused(self, make: object) -> None:
        with pytest.raises((ValueError, TypeError)):
            make()  # type: ignore[operator]

    def test_a_date_group_from_a_moment(self) -> None:
        assert DateGroup.of(dt.datetime(2026, 1, 5, 10, 30), "hour") == DateGroup("hour", 2026, 1, 5, 10)
        with pytest.raises(ValueError, match="datetime"):
            DateGroup.of(dt.date(2026, 1, 5), "minute")


class TestCriteriaStrings:
    """Excel's own ``Criteria1`` spellings, stored as Excel stores them."""

    def test_an_equality_is_a_value_list(self) -> None:
        assert criteria("=North") == ValueFilter(("North",))
        assert criteria("North") == ValueFilter(("North",))
        assert criteria("= North ") == ValueFilter(("North",))

    def test_blanks_and_everything_else(self) -> None:
        assert criteria("=") == ValueFilter(blank=True)
        assert criteria("<>") == CustomFilter((Comparison("notEqual", " "),))

    def test_wildcards_and_escapes(self) -> None:
        assert criteria("=a*") == CustomFilter((Comparison("equal", "a*"),))
        assert criteria("=a~*b") == ValueFilter(("a*b",))
        assert criteria("<>*an*") == CustomFilter((Comparison("notEqual", "*an*"),))

    def test_orderings_are_typed(self) -> None:
        assert criteria(">=10") == CustomFilter((Comparison("greaterThanOrEqual", 10),))
        assert criteria(">M") == CustomFilter((Comparison("greaterThan", "M"),))
        assert criteria(">=1/5/2026") == CustomFilter((Comparison("greaterThanOrEqual", dt.date(2026, 1, 5)),))

    def test_two_criteria(self) -> None:
        assert criteria(">5", "<100") == CustomFilter(
            (Comparison("greaterThan", 5), Comparison("lessThan", 100)), require_all=True
        )
        assert criteria("=", "=5", require_all=False) == ValueFilter(("5",), blank=True)


class TestTheEvaluatorDirectly:
    def test_unknown_cells_decide_nothing(self) -> None:
        cells = [FilterCell(5.0, "5"), FilterCell(5.0, "5", unknown=True)]
        assert keeps(ValueFilter(("5",)), cells, today=TODAY) == [True, None]

    def test_an_ignored_criterion_keeps_everything(self) -> None:
        ignored = OpaqueCriterion("<filters/>", ignored=True)
        assert keeps(ignored, [FilterCell("a", "a")], today=TODAY) == [True]

    def test_an_average_adds_in_order_on_every_python(self) -> None:
        """1e16 + 1 rounds back to 1e16, so the mean is 0.0625 and 0.25 is
        above it. Python's ``sum`` compensates from 3.12 on and makes the
        mean 0.3125, which would hide that row on some interpreters only."""
        values = [1e16, 1.0, -1e16, 0.25]
        cells = [FilterCell(value, str(value)) for value in values]
        assert keeps(DynamicFilter("aboveAverage", 0.0), cells, today=TODAY) == [True, True, False, True]
