"""Tables, which Excel's object model calls ListObjects.

A table is four things wired together: its own part, a content-type override,
a relationship from the sheet, and a ``<tablePart>`` entry. Excel ignores a
table whose wiring is incomplete, or refuses the file, so each piece is
asserted separately.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from pyofficeeditor._xml import XmlDocument
from pyofficeeditor.excel import RangeRef, Workbook
from pyofficeeditor.excel._tables import (
    CT_TABLE,
    MAX_TABLE_NAME_LENGTH,
    RT_TABLE,
    TableStyle,
    build_table_part,
    check_table_name,
    unique_column_names,
)


@pytest.fixture()
def book(live_structures_xlsx: Path) -> Workbook:
    return Workbook.open(live_structures_xlsx)


class TestTableNameRules:
    @pytest.mark.parametrize(
        "name", ["Table1", "SalesData", "_private", "a", "With.Dots", "x" * MAX_TABLE_NAME_LENGTH]
    )
    def test_names_excel_accepts(self, name: str) -> None:
        check_table_name(name)

    @pytest.mark.parametrize(
        ("name", "expected"),
        [
            ("", "cannot be empty"),
            ("x" * (MAX_TABLE_NAME_LENGTH + 1), "at most 255"),
            ("Sales Table", "no spaces"),
            ("1stTable", "must start with"),
            ("has-dash", "no spaces or operator"),
            ("has!bang", "no spaces or operator"),
            ("A1", "reads as a cell reference"),
            ("XFD1048576", "reads as a cell reference"),
            ("C", "R1C1"),
            ("R", "R1C1"),
        ],
    )
    def test_names_excel_refuses(self, name: str, expected: str) -> None:
        with pytest.raises(ValueError, match=expected):
            check_table_name(name)

    def test_uniqueness_ignores_case(self) -> None:
        with pytest.raises(ValueError, match="already has a table"):
            check_table_name("salestable", taken={"SalesTable"})
        check_table_name("other", taken={"SalesTable"})


class TestColumnNames:
    @pytest.mark.parametrize(
        ("given", "expected"),
        [
            (["A", "B"], ["A", "B"]),
            (["", ""], ["Column1", "Column2"]),
            (["A", ""], ["A", "Column2"]),
            (["dup", "dup"], ["dup", "dup2"]),
            (["dup", "dup", "dup"], ["dup", "dup2", "dup3"]),
            (["DUP", "dup"], ["DUP", "dup2"]),
            (["  spaced  "], ["spaced"]),
            (["dup", "dup2", "dup"], ["dup", "dup2", "dup3"]),
        ],
    )
    def test_blanks_are_filled_and_duplicates_disambiguated(
        self, given: list[str], expected: list[str]
    ) -> None:
        """Excel invents these rather than refusing the table, so doing the
        same keeps a caller from writing a file Excel then rewrites."""
        assert unique_column_names(given) == expected


class TestReadingRealTables:
    def test_both_tables_are_found_workbook_wide(self, book: Workbook) -> None:
        assert book.table_names == ["SalesTable", "KeyTable"]

    def test_tables_are_found_per_sheet(self, book: Workbook) -> None:
        assert [t.name for t in book["Tabled"].tables] == ["SalesTable"]
        assert [t.name for t in book["Second"].tables] == ["KeyTable"]

    def test_the_extent_and_its_parts(self, book: Workbook) -> None:
        table = book.table("SalesTable")
        assert table.ref.a1 == "A1:C5", "the whole table, totals row included"
        assert table.header_row == 1
        assert table.data_range is not None
        assert table.data_range.a1 == "A2:C4"
        assert table.totals_row == 5
        assert table.has_totals_row is True

    def test_the_filter_excludes_the_totals_row(self, book: Workbook) -> None:
        """Giving the filter the table's full extent would put a dropdown on
        the totals row."""
        table = book.table("SalesTable")
        assert table.filter_ref is not None
        assert table.filter_ref.a1 == "A1:C4"
        assert table.filter_ref.bottom == table.ref.bottom - 1

    def test_the_columns(self, book: Workbook) -> None:
        table = book.table("SalesTable")
        assert table.column_names == ["Region", "Units", "Revenue"]
        assert [c.id for c in table.columns] == [1, 2, 3]

    def test_the_totals_row_details(self, book: Workbook) -> None:
        table = book.table("SalesTable")
        assert table.column("Region").totals_label == "Total"
        assert table.column("Units").totals_function == "sum"
        assert table.column("Region").totals_function is None

    def test_a_calculated_column_formula(self, book: Workbook) -> None:
        """Stored once for the column, the same idea as a shared formula."""
        assert book.table("SalesTable").column("Revenue").calculated_formula == "B2*10"
        assert book.table("SalesTable").column("Units").calculated_formula is None

    def test_the_style(self, book: Workbook) -> None:
        style = book.table("SalesTable").style
        assert style.name == "TableStyleMedium2"
        assert style.show_row_stripes is True
        assert style.show_first_column is False

    def test_a_table_with_no_totals_row(self, book: Workbook) -> None:
        table = book.table("KeyTable")
        assert table.has_totals_row is False
        assert table.totals_row is None
        assert table.ref.a1 == "A1:A2"
        assert table.data_range is not None
        # One data row and one column, so the block is a single cell and
        # renders as A2 rather than A2:A2.
        assert table.data_range.a1 == "A2"
        assert table.data_range.size == 1

    def test_column_lookup_is_case_insensitive(self, book: Workbook) -> None:
        assert book.table("SalesTable").column("units").name == "Units"

    def test_a_missing_column_names_what_is_there(self, book: Workbook) -> None:
        with pytest.raises(KeyError, match="Region, Units, Revenue"):
            book.table("SalesTable").column("Nope")

    def test_a_missing_table_names_what_is_there(self, book: Workbook) -> None:
        with pytest.raises(KeyError, match="SalesTable, KeyTable"):
            book.table("Nope")
        with pytest.raises(KeyError, match="SalesTable"):
            book["Tabled"].table("Nope")

    def test_a_sheet_with_no_tables(self, live_sample_xlsx: Path) -> None:
        assert Workbook.open(live_sample_xlsx)["Data"].tables == []
        assert Workbook.open(live_sample_xlsx).table_names == []

    def test_reading_tables_is_not_a_change(
        self, book: Workbook, live_structures_xlsx: Path
    ) -> None:
        for table in book.tables:
            _ = table.ref, table.columns, table.style, table.data_range, table.filter_ref
        assert book.is_modified is False
        assert book.to_bytes() == live_structures_xlsx.read_bytes()


class TestBuildingATablePart:
    def test_the_two_spellings_of_a_totals_row(self) -> None:
        """Excel writes the count when there is one and the shown flag when
        there is not; writing neither leaves it to guess."""
        with_totals = build_table_part(
            identifier=1,
            name="T",
            ref=RangeRef.parse("A1:B3"),
            column_names=["a", "b"],
            has_totals_row=True,
            style=TableStyle(),
        ).to_bytes()
        assert b'totalsRowCount="1"' in with_totals
        assert b"totalsRowShown" not in with_totals

        without = build_table_part(
            identifier=1,
            name="T",
            ref=RangeRef.parse("A1:B3"),
            column_names=["a", "b"],
            has_totals_row=False,
            style=TableStyle(),
        ).to_bytes()
        assert b'totalsRowShown="0"' in without
        assert b"totalsRowCount" not in without

    def test_the_filter_shrinks_for_a_totals_row(self) -> None:
        raw = build_table_part(
            identifier=1,
            name="T",
            ref=RangeRef.parse("A1:B4"),
            column_names=["a", "b"],
            has_totals_row=True,
            style=TableStyle(),
        ).to_bytes()
        assert b'ref="A1:B4"' in raw, "the table"
        assert b'<autoFilter ref="A1:B3"/>' in raw, "the filter, one row shorter"

    def test_children_are_in_schema_order(self) -> None:
        raw = build_table_part(
            identifier=1,
            name="T",
            ref=RangeRef.parse("A1:B3"),
            column_names=["a", "b"],
            has_totals_row=False,
            style=TableStyle(),
        ).to_text()
        assert raw.index("<autoFilter") < raw.index("<tableColumns")
        assert raw.index("<tableColumns") < raw.index("<tableStyleInfo")

    def test_the_name_and_display_name_match(self) -> None:
        raw = build_table_part(
            identifier=7,
            name="Mine",
            ref=RangeRef.parse("A1:A2"),
            column_names=["a"],
            has_totals_row=False,
            style=TableStyle(),
        ).to_bytes()
        assert b'name="Mine"' in raw
        assert b'displayName="Mine"' in raw
        assert b'id="7"' in raw

    def test_the_column_count_is_written(self) -> None:
        raw = build_table_part(
            identifier=1,
            name="T",
            ref=RangeRef.parse("A1:C2"),
            column_names=["a", "b", "c"],
            has_totals_row=False,
            style=TableStyle(),
        ).to_bytes()
        assert b'<tableColumns count="3">' in raw

    def test_it_parses_back(self) -> None:
        document = build_table_part(
            identifier=1,
            name="T",
            ref=RangeRef.parse("A1:B3"),
            column_names=["a", "b"],
            has_totals_row=False,
            style=TableStyle(name="TableStyleLight1"),
        )
        again = XmlDocument.parse(document.to_bytes())
        assert again.root.get("name") == "T"
        assert again.root.require("tableStyleInfo").get("name") == "TableStyleLight1"


class TestAddingATable:
    @pytest.fixture()
    def prepared(self, book: Workbook) -> Workbook:
        sheet = book["Tabled"]
        sheet["A10"].value = "Product"
        sheet["B10"].value = "Qty"
        sheet["A11"].value = "widget"
        sheet["B11"].value = 5
        sheet["A12"].value = "gadget"
        sheet["B12"].value = 9
        return book

    def test_all_four_pieces_are_written(self, prepared: Workbook) -> None:
        sheet = prepared["Tabled"]
        table = sheet.add_table("NewTable", "A10:B12")
        package = prepared.package

        assert package.has_part(table.part_name), "the part"
        assert package.content_types.of(table.part_name) == CT_TABLE, "the content type"
        targets = [
            r.target_part for r in package.relationships(sheet.part_name).by_type(RT_TABLE)
        ]
        assert table.part_name in targets, "the relationship"
        raw = package.read(sheet.part_name).decode()
        assert "<tableParts" in raw and "<tablePart " in raw, "the tableParts entry"

    def test_the_new_table_reads_back(self, prepared: Workbook) -> None:
        table = prepared["Tabled"].add_table("NewTable", "A10:B12")
        assert table.name == "NewTable"
        assert table.ref.a1 == "A10:B12"
        assert table.column_names == ["Product", "Qty"]
        assert table.data_range is not None
        assert table.data_range.a1 == "A11:B12"

    def test_it_survives_a_save(self, prepared: Workbook) -> None:
        prepared["Tabled"].add_table("NewTable", "A10:B12")
        reopened = Workbook.from_bytes(prepared.to_bytes())
        assert reopened.table_names == ["SalesTable", "NewTable", "KeyTable"]
        assert reopened.table("NewTable").column_names == ["Product", "Qty"]

    def test_the_part_name_and_id_avoid_collisions(self, prepared: Workbook) -> None:
        """Both are workbook-wide, so allocating from one sheet would clash."""
        table = prepared["Tabled"].add_table("NewTable", "A10:B12")
        assert table.part_name == "xl/tables/table3.xml"
        assert table.id == 3
        assert len({t.id for t in prepared.tables}) == 3

    def test_a_style_is_applied(self, prepared: Workbook) -> None:
        prepared["Tabled"].add_table(
            "NewTable", "A10:B12", style=TableStyle(name="TableStyleLight9", show_row_stripes=False)
        )
        reopened = Workbook.from_bytes(prepared.to_bytes())
        style = reopened.table("NewTable").style
        assert style.name == "TableStyleLight9"
        assert style.show_row_stripes is False

    def test_a_totals_row(self, prepared: Workbook) -> None:
        table = prepared["Tabled"].add_table("NewTable", "A10:B12", totals_row=True)
        assert table.has_totals_row is True
        assert table.totals_row == 12
        assert table.data_range is not None
        assert table.data_range.a1 == "A11:B11", "the last row became the totals row"
        assert table.filter_ref is not None
        assert table.filter_ref.a1 == "A10:B11"

    def test_a_totals_row_needs_a_data_row(self, prepared: Workbook) -> None:
        with pytest.raises(ValueError, match="no data between"):
            prepared["Tabled"].add_table("Small", "E1:E2", totals_row=True)

    def test_blank_headers_are_filled_in_the_cells_too(self, prepared: Workbook) -> None:
        """The column names have to equal the header cells' text, or Excel
        reconciles them by rewriting the part."""
        sheet = prepared["Second"]
        sheet["C2"].value = 1
        table = sheet.add_table("Filled", "C1:D2")
        assert table.column_names == ["Column1", "Column2"]
        assert sheet["C1"].value == "Column1"
        assert sheet["D1"].value == "Column2"

    def test_duplicate_headers_are_disambiguated_in_the_cells(self, prepared: Workbook) -> None:
        sheet = prepared["Second"]
        sheet["C1"].value = "dup"
        sheet["D1"].value = "dup"
        table = sheet.add_table("Dups", "C1:D2")
        assert table.column_names == ["dup", "dup2"]
        assert sheet["D1"].value == "dup2"

    def test_existing_headers_are_left_alone(self, prepared: Workbook) -> None:
        sheet = prepared["Tabled"]
        sheet.add_table("NewTable", "A10:B12")
        assert sheet["A10"].value == "Product", "not rewritten"

    def test_an_overlapping_table_is_refused(self, prepared: Workbook) -> None:
        with pytest.raises(ValueError, match="overlaps the table 'SalesTable'"):
            prepared["Tabled"].add_table("Clash", "B1:D3")

    def test_a_duplicate_name_is_refused_across_sheets(self, prepared: Workbook) -> None:
        """Table names are workbook-wide."""
        with pytest.raises(ValueError, match="already has a table"):
            prepared["Second"].add_table("SalesTable", "C1:C2")

    def test_an_invalid_name_is_refused_before_anything_is_written(self, prepared: Workbook) -> None:
        before = set(prepared.package.part_names())
        with pytest.raises(ValueError):
            prepared["Tabled"].add_table("Bad Name", "A10:B12")
        assert set(prepared.package.part_names()) == before
        assert prepared.table_names == ["SalesTable", "KeyTable"]

    def test_a_reversed_reference_normalizes(self, prepared: Workbook) -> None:
        assert prepared["Tabled"].add_table("NewTable", "B12:A10").ref.a1 == "A10:B12"

    def test_a_range_object_is_accepted(self, prepared: Workbook) -> None:
        table = prepared["Tabled"].add_table("NewTable", RangeRef.parse("A10:B12"))
        assert table.ref.a1 == "A10:B12"

    def test_two_tables_on_one_sheet(self, prepared: Workbook) -> None:
        sheet = prepared["Tabled"]
        sheet.add_table("First", "A10:B12")
        sheet["E10"].value = "Other"
        sheet["E11"].value = 1
        sheet.add_table("Second2", "E10:E11")
        reopened = Workbook.from_bytes(prepared.to_bytes())
        assert [t.name for t in reopened["Tabled"].tables] == ["SalesTable", "First", "Second2"]
        raw = reopened.package.read(reopened["Tabled"].part_name).decode()
        container = re.search(r"<tableParts[^>]*>", raw)
        assert container is not None
        assert 'count="3"' in container.group(0)

    def test_the_element_lands_in_schema_order(self, prepared: Workbook) -> None:
        from pyofficeeditor.excel._schema import WORKSHEET_CHILD_ORDER

        sheet = prepared["Tabled"]
        sheet.add_table("NewTable", "A10:B12")
        names = [
            child.name
            for child in sheet.document.root.elements()
            if child.name in WORKSHEET_CHILD_ORDER
        ]
        positions = [WORKSHEET_CHILD_ORDER.index(n) for n in names]
        assert positions == sorted(positions), names


class TestRemovingATable:
    def test_all_four_pieces_go(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        part = book.table("SalesTable").part_name
        sheet.remove_table("SalesTable")
        package = book.package
        assert not package.has_part(part)
        assert part not in package.content_types.overrides
        assert part not in [
            r.target_part for r in package.relationships(sheet.part_name).by_type(RT_TABLE)
        ]
        assert sheet.tables == []

    def test_the_values_stay(self, book: Workbook) -> None:
        """Which is what Excel's Convert to Range does."""
        sheet = book["Tabled"]
        sheet.remove_table("SalesTable")
        assert sheet["A1"].value == "Region"
        assert sheet["B2"].value == 120

    def test_it_survives_a_save(self, book: Workbook) -> None:
        book["Tabled"].remove_table("SalesTable")
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened.table_names == ["KeyTable"]
        assert not reopened.package.has_part("xl/tables/table1.xml")

    def test_the_element_goes_when_the_last_table_does(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        sheet.remove_table("SalesTable")
        assert b"tableParts" not in sheet.document.to_bytes()

    def test_removing_an_absent_table(self, book: Workbook) -> None:
        with pytest.raises(KeyError, match="has no table"):
            book["Tabled"].remove_table("Nope")

    def test_the_other_sheets_table_is_untouched(self, book: Workbook) -> None:
        book["Tabled"].remove_table("SalesTable")
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened.table("KeyTable").ref.a1 == "A1:A2"

    def test_the_name_frees_up(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        sheet.remove_table("SalesTable")
        sheet.add_table("SalesTable", "A1:C4")
        assert book.table("SalesTable").ref.a1 == "A1:C4"


class TestStyle:
    def test_setting_a_style(self, book: Workbook) -> None:
        table = book.table("SalesTable")
        table.style = TableStyle(name="TableStyleDark1", show_first_column=True)
        reopened = Workbook.from_bytes(book.to_bytes())
        style = reopened.table("SalesTable").style
        assert style.name == "TableStyleDark1"
        assert style.show_first_column is True

    def test_a_style_round_trips(self) -> None:
        for style in (
            TableStyle(),
            TableStyle(name=None, show_row_stripes=False),
            TableStyle(name="X", show_first_column=True, show_last_column=True),
            TableStyle(show_column_stripes=True),
        ):
            element = style.write()
            assert TableStyle.read(XmlDocument.parse(element.to_xml().encode()).root) == style

    def test_a_missing_style_element(self) -> None:
        assert TableStyle.read(None).name is None
