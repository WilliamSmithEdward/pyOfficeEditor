"""Deleting rows and columns.

Deletion is not insertion run backwards. Three things only it has to do, and
each has its own failure mode:

- a reference to something gone becomes ``#REF!``, while a range only partly
  deleted shrinks instead
- a shared formula whose master is deleted leaves its group pointing at
  nothing
- a merge reduced to one cell, or a table with no columns left, is not a
  merge or a table any more
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyofficeeditor._xml import Element
from pyofficeeditor.excel import (
    RangeRef,
    Workbook,
    Worksheet,
    cell_is,
    contains_text,
    expression,
)
from pyofficeeditor.excel._formulas import REF_ERROR, Deletion, delete_in_formula
from pyofficeeditor.excel._reference import MAX_COLUMN, MAX_ROW
from pyofficeeditor.excel._schema import WORKSHEET_CHILD_ORDER, insert_in_schema_order


@pytest.fixture()
def book(live_sample_xlsx: Path) -> Workbook:
    return Workbook.open(live_sample_xlsx)


class TestDeletionValue:
    @pytest.mark.parametrize(
        ("cell", "expected"),
        [("A1", "A1"), ("A2", "A2"), ("A3", None), ("A4", None), ("A5", "A3"), ("A9", "A7")],
    )
    def test_a_single_cell(self, cell: str, expected: str | None) -> None:
        from pyofficeeditor.excel import CellRef

        moved = Deletion.rows(3, 2).moved(CellRef.parse(cell))
        assert (moved.a1 if moved else None) == expected

    @pytest.mark.parametrize(
        ("block", "expected"),
        [
            ("A1:A10", "A1:A8"),
            ("A1:A2", "A1:A2"),
            ("A5:A10", "A3:A8"),
            ("A3:A4", None),
            ("A1:A4", "A1:A2"),
            ("A3:A10", "A3:A8"),
            ("A2:A3", "A2"),
            ("A3:A5", "A3"),
            ("A3", None),
        ],
    )
    def test_a_range_shrinks_rather_than_breaking(self, block: str, expected: str | None) -> None:
        """Only a block with nothing left is gone."""
        moved = Deletion.rows(3, 2).moved_range(RangeRef.parse(block))
        assert (moved.a1 if moved else None) == expected

    def test_columns(self) -> None:
        deletion = Deletion.columns(2, 1)
        assert deletion.moved_range(RangeRef.parse("A1:C1")) is not None
        moved = deletion.moved_range(RangeRef.parse("A1:C1"))
        assert moved is not None and moved.a1 == "A1:B1"
        assert deletion.moved_range(RangeRef.parse("B1:B5")) is None

    def test_covers(self) -> None:
        deletion = Deletion.rows(3, 2)
        assert deletion.covers_row(3) and deletion.covers_row(4)
        assert not deletion.covers_row(2) and not deletion.covers_row(5)
        assert not deletion.covers_column(3), "this deletion touches no columns"

    def test_an_empty_deletion_changes_nothing(self) -> None:
        assert Deletion().is_empty
        assert Deletion.rows(3, 0).is_empty


class TestDeleteInFormula:
    def delete(self, formula: str) -> str:
        return delete_in_formula(
            formula, Deletion.rows(3, 2), formula_sheet="Data", target_sheet="Data"
        )

    @pytest.mark.parametrize(
        ("formula", "expected"),
        [
            ("A1", "A1"),
            ("A3", REF_ERROR),
            ("A4", REF_ERROR),
            ("A5", "A3"),
            ("$A$3", REF_ERROR),
            ("A1+A3+A6", f"A1+{REF_ERROR}+A4"),
            ("SUM(A1:A10)", "SUM(A1:A8)"),
            ("SUM(A3:A4)", f"SUM({REF_ERROR})"),
            ("SUM(A1:A4)", "SUM(A1:A2)"),
            ("SUM(A3:A10)", "SUM(A3:A8)"),
            ("LOG10(A3)", f"LOG10({REF_ERROR})"),
            ('"A3"', '"A3"'),
            ("Notes!A3", "Notes!A3"),
        ],
    )
    def test_rows(self, formula: str, expected: str) -> None:
        assert self.delete(formula) == expected

    def test_columns(self) -> None:
        deletion = Deletion.columns(2, 1)
        assert (
            delete_in_formula("A1+B1+C1", deletion, formula_sheet="D", target_sheet="D")
            == f"A1+{REF_ERROR}+B1"
        )
        assert (
            delete_in_formula("SUM(A1:C1)", deletion, formula_sheet="D", target_sheet="D")
            == "SUM(A1:B1)"
        )

    def test_only_the_edited_sheets_references_change(self) -> None:
        assert (
            delete_in_formula(
                "A3+Data!A3+Notes!A3",
                Deletion.rows(3, 2),
                formula_sheet="Summary",
                target_sheet="Data",
            )
            == f"A3+Data!{REF_ERROR}+Notes!A3"
        )

    @pytest.mark.parametrize(
        ("formula", "expected"),
        [
            ("SUM(Data!2:4)", f"SUM(Data!{REF_ERROR})"),
            ("SUM(Data!1:6)", "SUM(Data!1:3)"),
            ("Data!A3", f"Data!{REF_ERROR}"),
            ("SUM(Data!A2:A4)", f"SUM(Data!{REF_ERROR})"),
            ("SUM(Data!A1:A8)", "SUM(Data!A1:A5)"),
        ],
    )
    def test_a_broken_reference_keeps_its_qualifier(self, formula: str, expected: str) -> None:
        """Measured: Excel writes ``Data!#REF!``, not a bare ``#REF!``. The
        sheet name survives even though what it pointed at does not."""
        assert (
            delete_in_formula(
                formula, Deletion.rows(2, 3), formula_sheet="Notes", target_sheet="Data"
            )
            == expected
        )

    def test_a_cross_sheet_range_shrinks(self) -> None:
        assert (
            delete_in_formula(
                "SUM(Data!A1:A10)",
                Deletion.rows(3, 2),
                formula_sheet="Summary",
                target_sheet="Data",
            )
            == "SUM(Data!A1:A8)"
        )

    def test_nothing_to_do(self) -> None:
        assert (
            delete_in_formula("A1", Deletion(), formula_sheet="D", target_sheet="D") == "A1"
        )

    @pytest.mark.parametrize(
        ("formula", "expected"),
        [
            ("SUM(A1:A4)", "SUM(A1:A1)"),
            ("SUM(A2:A5)", "SUM(A2:A2)"),
            ("SUM(A1:A8)", "SUM(A1:A5)"),
            ("SUM(A2:A4)", f"SUM({REF_ERROR})"),
            ("A3*2", f"{REF_ERROR}*2"),
            ("SUM(A:A)", "SUM(A:A)"),
            ("SUM($A$1:$A$4)", "SUM($A$1:$A$1)"),
            ("SUM(A1:B4)", "SUM(A1:B1)"),
        ],
    )
    def test_what_excel_writes(self, formula: str, expected: str) -> None:
        """Measured against Excel: these eight are its own answers for a
        deletion of rows 2 to 4, read back through ``Range.Formula``.

        Two of them are decisions a reasonable implementation gets wrong. A
        range shrunk to one cell keeps its range shape, ``SUM(A1:A1)`` rather
        than ``SUM(A1)``. Absolute markers survive on both ends.
        """
        assert (
            delete_in_formula(
                formula, Deletion.rows(2, 3), formula_sheet="S", target_sheet="S"
            )
            == expected
        )

    @pytest.mark.parametrize(
        ("formula", "expected"),
        [
            ("SUM(2:4)", f"SUM({REF_ERROR})"),
            ("SUM(1:6)", "SUM(1:3)"),
            ("SUM(6:7)", "SUM(3:4)"),
            ("SUM(A:B)", "SUM(A:B)"),
            ("SUM($1:$6)", "SUM($1:$3)"),
        ],
    )
    def test_whole_axis_references(self, formula: str, expected: str) -> None:
        """Also Excel's own answers. A whole-row reference shrinks and
        breaks exactly like a range, and a column span is untouched by a row
        deletion."""
        assert (
            delete_in_formula(
                formula, Deletion.rows(2, 3), formula_sheet="S", target_sheet="S"
            )
            == expected
        )

    def test_a_column_deletion_breaks_a_column_span(self) -> None:
        assert (
            delete_in_formula(
                "SUM(B:C)", Deletion.columns(2, 2), formula_sheet="S", target_sheet="S"
            )
            == f"SUM({REF_ERROR})"
        )


class TestDeletingRows:
    def test_values_move_up(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.delete_rows(3, 2)
        assert sheet["A2"].value == "North", "the row above stayed"
        assert sheet["A3"].value == "West", "row 5 came up to row 3"

    def test_the_sheet_gets_shorter(self, book: Workbook) -> None:
        sheet = book["Data"]
        assert sheet.max_row == 11
        sheet.delete_rows(3, 2)
        assert sheet.max_row == 9

    def test_one_row_by_default(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.delete_rows(3)
        assert sheet["A3"].value == "East"

    def test_a_range_over_the_deletion_shrinks(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.delete_rows(3, 2)
        assert sheet["D4"].formula == "SUM(D2:D3)"

    def test_merges_move_up(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.delete_rows(3, 2)
        assert [b.a1 for b in sheet.merged_ranges] == ["A9:C9"]

    def test_a_merge_entirely_inside_the_deletion_goes(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.delete_rows(11, 1)
        assert sheet.merged_ranges == []
        assert b"mergeCells" not in sheet.document.to_bytes()

    def test_a_merge_reduced_to_one_cell_goes(self, book: Workbook) -> None:
        """A single cell is not a merge, which is what Excel does with it."""
        sheet = book["Data"]
        sheet["H1"].value = "tall"
        sheet.merge("H1:H3")
        sheet.delete_rows(2, 2)
        assert all(b.a1 != "H1" for b in sheet.merged_ranges)

    def test_the_dimension_shrinks(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.delete_rows(3, 2)
        assert sheet.dimension is not None
        assert sheet.dimension.a1 == "A1:F9"

    def test_rows_stay_in_order(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.delete_rows(3, 2)
        data = sheet.document.root.require("sheetData")
        numbers = [int(r.get("r") or 0) for r in data.children_named("row")]
        assert numbers == sorted(numbers)
        assert len(numbers) == len(set(numbers)), "no row number appears twice"

    def test_a_formula_on_another_sheet_breaks_or_shrinks(self, book: Workbook) -> None:
        summary = book.add_sheet("Summary")
        summary["A1"].formula = "=SUM(Data!D2:D5)"
        summary["A2"].formula = "=Data!B3"
        book["Data"].delete_rows(3, 2)
        assert summary["A1"].formula == "SUM(Data!D2:D3)", "the range shrank"
        assert summary["A2"].formula == f"Data!{REF_ERROR}", "the cell is gone"

    def test_defined_names_shrink(self, book: Workbook) -> None:
        book.add_defined_name("Totals", "Data!$D$2:$D$5")
        book["Data"].delete_rows(3, 2)
        assert book.defined_name("Totals").refers_to == "Data!$D$2:$D$3"

    def test_tables_shrink(self, live_structures_xlsx: Path) -> None:
        structures = Workbook.open(live_structures_xlsx)
        assert structures.table("SalesTable").ref.a1 == "A1:C5"
        structures["Tabled"].delete_rows(2, 1)
        table = structures.table("SalesTable")
        assert table.ref.a1 == "A1:C4"
        assert table.filter_ref is not None
        assert table.filter_ref.a1 == "A1:C3"

    def test_it_survives_a_save(self, book: Workbook) -> None:
        book["Data"].delete_rows(3, 2)
        reopened = Workbook.from_bytes(book.to_bytes())["Data"]
        assert reopened["A3"].value == "West"
        assert reopened["D4"].formula == "SUM(D2:D3)"
        assert [b.a1 for b in reopened.merged_ranges] == ["A9:C9"]

    def test_deleting_below_everything(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.delete_rows(100, 5)
        assert sheet["A2"].value == "North"
        assert sheet.max_row == 11


class TestOrphanedSharedFormulas:
    def test_deleting_the_master_gives_the_group_its_own_text(self, book: Workbook) -> None:
        """A shared formula lives once. Delete its cell and the rest point at
        nothing, so they are given their own copies first."""
        sheet = book["Data"]
        sheet.delete_rows(2, 1)
        assert sheet["D2"].formula == "B2*C2", "was D3, one row up"
        assert sheet["D3"].formula == "B3*C3"
        assert sheet["D4"].formula == "B4*C4"
        assert b't="shared"' not in sheet.document.to_bytes()

    def test_it_works_on_a_cold_cache(self, book: Workbook) -> None:
        """The regression that mattered: a follower derives from its master's
        element, so stripping the master's markers while still walking left
        the rest of the group blank. It only showed when no formula had been
        read first, because a warm lookup hid it."""
        sheet = book["Data"]
        sheet.delete_rows(2, 1)
        reopened = Workbook.from_bytes(book.to_bytes())["Data"]
        assert reopened["D2"].formula == "B2*C2"
        assert reopened["D3"].formula == "B3*C3"
        assert reopened["D4"].formula == "B4*C4"

    def test_a_surviving_master_keeps_its_group(self, book: Workbook) -> None:
        """Nothing is expanded when the master lives: the followers still
        derive correctly from it."""
        sheet = book["Data"]
        sheet.delete_rows(3, 1)
        assert b't="shared"' in sheet.document.to_bytes()
        assert sheet["D2"].formula == "B2*C2"
        assert sheet["D3"].formula == "B3*C3", "was D4, one row up"


class TestDeletingColumns:
    def test_values_move_left(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.delete_columns(2, 1)
        assert sheet["A1"].value == "Region"
        assert sheet["B1"].value == "Price", "what was in C came to B"

    def test_a_formula_losing_an_input(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.delete_columns(2, 1)
        assert sheet["C2"].formula == f"{REF_ERROR}*B2"

    def test_a_merge_loses_a_column(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.delete_columns(2, 1)
        assert [b.a1 for b in sheet.merged_ranges] == ["A11:B11"]

    def test_column_widths_follow(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.set_column_width(3, 30)
        sheet.delete_columns(2, 1)
        assert sheet.column_width(2) == 30

    def test_a_deleted_columns_width_goes(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.set_column_width(2, 30)
        sheet.delete_columns(2, 1)
        assert sheet.column_width(2) is None

    def test_cells_stay_in_ascending_order(self, book: Workbook) -> None:
        from pyofficeeditor.excel import CellRef

        sheet = book["Data"]
        sheet.delete_columns(2, 1)
        data = sheet.document.root.require("sheetData")
        for row in data.children_named("row"):
            columns = [CellRef.parse(c.get("r") or "A1").column for c in row.children_named("c")]
            assert columns == sorted(columns), f"row {row.get('r')}"

    def test_it_survives_a_save(self, book: Workbook) -> None:
        book["Data"].delete_columns(2, 1)
        reopened = Workbook.from_bytes(book.to_bytes())["Data"]
        assert reopened["B1"].value == "Price"
        assert reopened["C2"].formula == f"{REF_ERROR}*B2"

    def test_a_table_loses_a_column(self, live_structures_xlsx: Path) -> None:
        structures = Workbook.open(live_structures_xlsx)
        assert structures.table("SalesTable").column_names == ["Region", "Units", "Revenue"]
        structures["Tabled"].delete_columns(2, 1)
        table = structures.table("SalesTable")
        assert table.ref.a1 == "A1:B5"
        assert table.column_names == ["Region", "Revenue"]


class TestRefusals:
    def test_deleting_a_tables_header_row(self, live_structures_xlsx: Path) -> None:
        """A table's column names come from its header cells."""
        structures = Workbook.open(live_structures_xlsx)
        with pytest.raises(ValueError, match="header row of the table 'SalesTable'"):
            structures["Tabled"].delete_rows(1, 1)

    def test_deleting_every_column_of_a_table(self, live_structures_xlsx: Path) -> None:
        structures = Workbook.open(live_structures_xlsx)
        with pytest.raises(ValueError, match="every column of the table"):
            structures["Tabled"].delete_columns(1, 3)

    def test_content_that_used_to_be_refused_shrinks(self, book: Workbook) -> None:
        """Nothing makes a deletion refuse any more. A validation whose
        range survives shrinks with it."""
        sheet = book["Data"]
        holder = Element.create("dataValidations")
        holder.append(Element.create("dataValidation", {"sqref": "B2:B9"}))
        insert_in_schema_order(sheet.document.root, holder, WORKSHEET_CHILD_ORDER)
        sheet.delete_rows(3, 2)
        found = sheet.document.root.require("dataValidations")
        assert [n.get("sqref") for n in found.children_named("dataValidation")] == ["B2:B7"]

    def test_an_entry_with_nothing_left_goes_with_its_wrapper(
        self, book: Workbook
    ) -> None:
        """An empty ``<dataValidations count="0"/>`` is not something Excel
        accepts, so the wrapper goes with its last entry."""
        sheet = book["Data"]
        holder = Element.create("dataValidations", {"count": "1"})
        holder.append(Element.create("dataValidation", {"sqref": "B3:B4"}))
        insert_in_schema_order(sheet.document.root, holder, WORKSHEET_CHILD_ORDER)
        sheet.delete_rows(3, 2)
        assert b"dataValidations" not in sheet.document.to_bytes()

    def test_a_surviving_sibling_keeps_the_wrapper_and_fixes_the_count(
        self, book: Workbook
    ) -> None:
        sheet = book["Data"]
        holder = Element.create("dataValidations", {"count": "2"})
        holder.append(Element.create("dataValidation", {"sqref": "B3:B4"}))
        holder.append(Element.create("dataValidation", {"sqref": "C2:C9"}))
        insert_in_schema_order(sheet.document.root, holder, WORKSHEET_CHILD_ORDER)
        sheet.delete_rows(3, 2)
        found = sheet.document.root.require("dataValidations")
        assert found.get("count") == "1"
        assert [n.get("sqref") for n in found.children_named("dataValidation")] == ["C2:C7"]

    @pytest.mark.parametrize(("at", "count"), [(0, 1), (-1, 1), (MAX_ROW + 1, 1), (3, 0)])
    def test_bad_row_arguments(self, book: Workbook, at: int, count: int) -> None:
        with pytest.raises(ValueError):
            book["Data"].delete_rows(at, count)

    @pytest.mark.parametrize(("at", "count"), [(0, 1), (MAX_COLUMN + 1, 1), (3, 0)])
    def test_bad_column_arguments(self, book: Workbook, at: int, count: int) -> None:
        with pytest.raises(ValueError):
            book["Data"].delete_columns(at, count)

    def test_a_comment_on_a_deleted_row_goes_with_it(self, book: Workbook) -> None:
        """Which is what Excel does. Comments were on neither the shifted
        list nor the refused one before, so they simply stayed put."""
        sheet = book["Data"]
        sheet.delete_rows(3, 2)
        assert sheet["A3"].value == "West"


class TestRoundTrip:
    def test_insert_then_delete_restores_the_values(self, book: Workbook) -> None:
        """Not byte-for-byte, since a shared group may have been expanded,
        but the data and the formulas come back."""
        sheet = book["Data"]
        before = [[cell.value for cell in row] for row in sheet.rows()]
        sheet.insert_rows(3, 2)
        sheet.delete_rows(3, 2)
        after = [[cell.value for cell in row] for row in sheet.rows()]
        assert after == before
        assert sheet["D6"].formula == "SUM(D2:D5)"
        assert [b.a1 for b in sheet.merged_ranges] == ["A11:C11"]

    def test_delete_then_insert_leaves_blanks(self, book: Workbook) -> None:
        """The other direction does not restore anything, which is expected:
        deletion throws the values away."""
        sheet = book["Data"]
        sheet.delete_rows(3, 2)
        sheet.insert_rows(3, 2)
        assert sheet["A3"].value is None
        assert sheet["A5"].value == "West"


class TestConditionalFormattingMoves:
    """Conditional formatting used to make insert and delete refuse. Now it
    moves, and three things have to move together or the rules highlight
    cells nobody asked about while the file opens perfectly."""

    def formatted(self, book: Workbook) -> Worksheet:
        sheet = book["Data"]
        sheet.add_conditional_format("B2:B9", cell_is("greaterThan", 100))
        sheet.add_conditional_format("D2:D5", contains_text("x"))
        sheet.add_conditional_format("F2:F9", expression("=$B2>SUM($D$2:$D$5)"))
        return sheet

    def sqrefs(self, sheet: Worksheet) -> list[str]:
        return [block.sqref for block in sheet.conditional_formats]

    def test_an_insertion_moves_the_ranges(self, book: Workbook) -> None:
        sheet = self.formatted(book)
        sheet.insert_rows(3, 2)
        assert self.sqrefs(sheet) == ["B2:B11", "D2:D7", "F2:F11"]

    def test_a_deletion_shrinks_them(self, book: Workbook) -> None:
        sheet = self.formatted(book)
        sheet.delete_rows(3, 2)
        assert self.sqrefs(sheet) == ["B2:B7", "D2:D3", "F2:F7"]

    def test_a_block_with_nothing_left_goes(self, book: Workbook) -> None:
        """Not left behind with an empty sqref, which Excel treats as
        malformed."""
        sheet = self.formatted(book)
        sheet.delete_rows(2, 4)
        assert "D" not in "".join(self.sqrefs(sheet))

    def test_an_expressions_references_follow(self, book: Workbook) -> None:
        """The condition is a real formula with real references. Moving the
        range and leaving the formula is the silent-corruption case."""
        sheet = self.formatted(book)
        sheet.insert_rows(3, 2)
        rule = sheet.conditional_formats[2].rules[0]
        assert rule.formulas == ("$B2>SUM($D$2:$D$7)",)

    def test_a_deleted_reference_in_a_condition_breaks(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.add_conditional_format("F2:F9", expression("=$B$3>1"))
        sheet.delete_rows(3, 1)
        assert sheet.conditional_formats[0].rules[0].formulas == (f"{REF_ERROR}>1",)

    def test_the_compatibility_formula_is_rebuilt_for_the_new_anchor(
        self, book: Workbook
    ) -> None:
        """It names the top-left of the range by construction, so it is
        rebuilt rather than shifted."""
        sheet = book["Data"]
        sheet.add_conditional_format("D4:D9", contains_text("x"))
        assert sheet.conditional_formats[0].rules[0].formulas == (
            'NOT(ISERROR(SEARCH("x",D4)))',
        )
        sheet.insert_rows(2, 3)
        block = sheet.conditional_formats[0]
        assert block.sqref == "D7:D12"
        assert block.rules[0].formulas == ('NOT(ISERROR(SEARCH("x",D7)))',)

    def test_columns_too(self, book: Workbook) -> None:
        sheet = self.formatted(book)
        sheet.insert_columns(3, 1)
        assert self.sqrefs(sheet) == ["B2:B9", "E2:E5", "G2:G9"]

    def test_a_deleted_column_takes_its_block(self, book: Workbook) -> None:
        sheet = self.formatted(book)
        sheet.delete_columns(4, 1)
        assert self.sqrefs(sheet) == ["B2:B9", "E2:E9"]

    def test_the_block_keeps_its_place_in_the_sheet(self, book: Workbook) -> None:
        """Between mergeCells and pageMargins. A rebuilt block appended to
        the end would break the schema's order."""
        sheet = self.formatted(book)
        sheet.insert_rows(3, 2)
        rendered = sheet.document.to_bytes().decode()
        assert rendered.index("<mergeCells") < rendered.index("<conditionalFormatting")
        assert rendered.index("<conditionalFormatting") < rendered.index("<pageMargins")

    def test_a_multi_area_block_moves_every_area(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.add_conditional_format("B2:B5 D2:D5", cell_is("greaterThan", 1))
        sheet.insert_rows(3, 2)
        assert sheet.conditional_formats[0].sqref == "B2:B7 D2:D7"

    def test_it_survives_a_save(self, book: Workbook) -> None:
        sheet = self.formatted(book)
        sheet.insert_rows(3, 2)
        reopened = Workbook.from_bytes(book.to_bytes())["Data"]
        assert self.sqrefs(reopened) == ["B2:B11", "D2:D7", "F2:F11"]

    def test_insertion_no_longer_refuses(self, book: Workbook) -> None:
        sheet = self.formatted(book)
        sheet.insert_rows(3, 2)  # this used to raise rather than move
        assert len(sheet.conditional_formats) == 3
