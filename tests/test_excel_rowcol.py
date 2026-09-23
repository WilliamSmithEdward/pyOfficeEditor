"""Inserting rows and columns, and the formula tokenizer it rests on.

A cell's address is written into the file in a dozen places. Missing one
gives a workbook that opens cleanly and points at the wrong cells, which no
byte comparison catches, so most of these tests are about the places that are
easy to forget.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyofficeeditor._xml import Element
from pyofficeeditor.excel import RangeRef, Workbook
from pyofficeeditor.excel._formulas import Shift, shift_formula, shift_range
from pyofficeeditor.excel._reference import MAX_COLUMN, MAX_ROW
from pyofficeeditor.excel._schema import WORKSHEET_CHILD_ORDER, insert_in_schema_order
from pyofficeeditor.excel._tokens import TokenKind, render, tokenize


@pytest.fixture()
def book(live_sample_xlsx: Path) -> Workbook:
    return Workbook.open(live_sample_xlsx)


class TestTokenizer:
    @pytest.mark.parametrize(
        "formula",
        [
            "B2*C2",
            "SUM(A1:A5)",
            "LOG10(A1)",
            '"A1" & B2',
            "'My Sheet'!B2",
            "Data!A1:A5",
            "IF(A1>0,LOG10(B1),Sheet2!C1)",
            "SUM(Table1[Amount])",
            'COUNTIF(A1:A10,">5")',
            "'It''s'!A1",
            "$B$2+B$3+$B4",
            "",
            "42",
            "TRUE",
            'CONCAT("unterminated',
        ],
    )
    def test_every_formula_round_trips_exactly(self, formula: str) -> None:
        """Which is what makes rewriting some tokens safe."""
        assert render(tokenize(formula)) == formula

    def test_a_function_name_is_not_a_reference(self) -> None:
        """``LOG10`` contains ``G10``."""
        references = [t.raw for t in tokenize("LOG10(A1)") if t.kind is TokenKind.REFERENCE]
        assert references == ["A1"]

    def test_a_string_is_not_searched(self) -> None:
        references = [t.raw for t in tokenize('"A1" & B2') if t.kind is TokenKind.REFERENCE]
        assert references == ["B2"]

    def test_a_reference_inside_a_quoted_sheet_name_is_not_one(self) -> None:
        tokens = tokenize("'My Sheet A1'!B2")
        references = [t.raw for t in tokens if t.kind is TokenKind.REFERENCE]
        assert references == ["B2"]

    def test_a_bare_sheet_qualifier(self) -> None:
        tokens = tokenize("Data!A1")
        assert [t.kind for t in tokens] == [TokenKind.SHEET, TokenKind.REFERENCE]
        assert tokens[0].value == "Data"
        assert tokens[1].sheet == "Data"

    def test_a_quoted_sheet_qualifier_unquotes(self) -> None:
        assert tokenize("'My Sheet'!A1")[0].value == "My Sheet"
        assert tokenize("'It''s'!A1")[0].value == "It's"

    def test_a_bare_reference_has_no_sheet(self) -> None:
        assert tokenize("A1")[0].sheet is None

    def test_both_ends_of_a_range_take_the_qualifier(self) -> None:
        """``Data!A1:A5`` is two references, both on Data."""
        references = [(t.raw, t.sheet) for t in tokenize("Data!A1:A5") if t.kind is TokenKind.REFERENCE]
        assert references == [("A1", "Data"), ("A5", "Data")]

    def test_a_qualifier_does_not_leak_past_its_range(self) -> None:
        references = [
            (t.raw, t.sheet) for t in tokenize("Data!A1:A5+B1") if t.kind is TokenKind.REFERENCE
        ]
        assert references == [("A1", "Data"), ("A5", "Data"), ("B1", None)]

    def test_two_qualifiers_in_one_formula(self) -> None:
        references = [
            (t.raw, t.sheet)
            for t in tokenize("SUM(Data!A1:A5,Notes!B1:B2)")
            if t.kind is TokenKind.REFERENCE
        ]
        assert references == [("A1", "Data"), ("A5", "Data"), ("B1", "Notes"), ("B2", "Notes")]

    def test_a_table_reference_is_not_a_cell_reference(self) -> None:
        assert [t for t in tokenize("SUM(Table1[Amount])") if t.kind is TokenKind.REFERENCE] == []


class TestWholeAxisTokens:
    """``A:A`` and ``2:4`` are references too, and they move.

    They get their own kind because they are not a pair of cells: ``A:A``
    is not ``A1:A1048576``, and Excel keeps the short form when it rewrites
    a formula. The risk is the reverse one, inventing an axis reference
    where the text only looks like one.
    """

    def axes(self, formula: str) -> list[str]:
        return [t.raw for t in tokenize(formula) if t.kind is TokenKind.AXIS]

    @pytest.mark.parametrize(
        ("formula", "expected"),
        [
            ("SUM(A:A)", ["A:A"]),
            ("SUM(2:4)", ["2:4"]),
            ("SUM($A:$C)", ["$A:$C"]),
            ("SUM($2:$4)", ["$2:$4"]),
            ("VLOOKUP(A1,A:C,2,0)", ["A:C"]),
            ("SUM(A1:A4)+SUM(2:4)", ["2:4"]),
        ],
    )
    def test_found(self, formula: str, expected: list[str]) -> None:
        assert self.axes(formula) == expected

    @pytest.mark.parametrize(
        "formula",
        [
            "SUM(A1:A4)",  # two cells, and a cell reference wins
            "LOG10(A1)",
            '"2:4"',  # inside a string
            "IF(A1=1,2,3)",
            "TIME(2,4,0)",  # commas, not a colon
            "SUM(Table1[Units])",
            "AA1:AA9",
        ],
    )
    def test_not_invented(self, formula: str) -> None:
        assert self.axes(formula) == []

    @pytest.mark.parametrize(
        ("formula", "sheet"),
        [("Data!2:4", "Data"), ("'My Sheet'!A:C", "My Sheet"), ("A:C", None)],
    )
    def test_the_qualifier_reaches_it(self, formula: str, sheet: str | None) -> None:
        found = [t for t in tokenize(formula) if t.kind is TokenKind.AXIS]
        assert len(found) == 1
        assert found[0].sheet == sheet

    @pytest.mark.parametrize(
        "formula",
        ["SUM(A:A)", "SUM(2:4)", "Data!2:4", '"2:4"', "SUM($A:$C)", "VLOOKUP(A1,A:C,2,0)"],
    )
    def test_round_trip(self, formula: str) -> None:
        assert "".join(t.raw for t in tokenize(formula)) == formula


class TestWholeAxisMoves:
    """Excel's own answers, measured through ``Range.Formula`` after
    inserting two rows at row 2 and after deleting rows 2 to 4."""

    @pytest.mark.parametrize(
        ("formula", "expected"),
        [
            ("SUM(2:4)", "SUM(4:6)"),
            ("SUM(1:6)", "SUM(1:8)"),
            ("SUM(6:7)", "SUM(8:9)"),
            ("SUM(A:A)", "SUM(A:A)"),
        ],
    )
    def test_insertion(self, formula: str, expected: str) -> None:
        assert (
            shift_formula(formula, Shift.rows(2, 2), formula_sheet="S", target_sheet="S")
            == expected
        )

    def test_a_column_insertion_moves_a_column_span(self) -> None:
        assert (
            shift_formula("SUM(A:C)", Shift.columns(2, 1), formula_sheet="S", target_sheet="S")
            == "SUM(A:D)"
        )

    def test_only_the_edited_sheet(self) -> None:
        assert (
            shift_formula(
                "SUM(2:4)+SUM(Data!2:4)",
                Shift.rows(2, 2),
                formula_sheet="Summary",
                target_sheet="Data",
            )
            == "SUM(2:4)+SUM(Data!4:6)"
        )


class TestShiftFormula:
    def shift(self, formula: str, shift: Shift) -> str:
        return shift_formula(formula, shift, formula_sheet="Data", target_sheet="Data")

    @pytest.mark.parametrize(
        ("formula", "expected"),
        [
            ("A5", "A7"),
            ("A1", "A1"),
            ("A1+A5+A10", "A1+A7+A12"),
            ("SUM(A1:A10)", "SUM(A1:A12)"),
            ("SUM(A6:A9)", "SUM(A8:A11)"),
            ("SUM(A1:A3)", "SUM(A1:A3)"),
            ("$A$5", "$A$7"),
            ("LOG10(A5)", "LOG10(A7)"),
            ('"A5"', '"A5"'),
        ],
    )
    def test_inserting_two_rows_at_five(self, formula: str, expected: str) -> None:
        assert self.shift(formula, Shift.rows(5, 2)) == expected

    def test_a_range_spanning_the_insertion_grows(self) -> None:
        """Rather than sliding: the top is above the insertion and stays."""
        assert self.shift("SUM(A1:A10)", Shift.rows(5, 2)) == "SUM(A1:A12)"

    def test_an_absolute_reference_still_moves(self) -> None:
        """Absoluteness governs copying, not insertion: the data moved, so
        the reference follows it."""
        assert self.shift("$A$5", Shift.rows(5, 2)) == "$A$7"

    @pytest.mark.parametrize(
        ("formula", "expected"),
        [("A1+B1+C1", "A1+C1+D1"), ("SUM(A1:C1)", "SUM(A1:D1)"), ("$B$1", "$C$1")],
    )
    def test_inserting_a_column(self, formula: str, expected: str) -> None:
        assert self.shift(formula, Shift.columns(2, 1)) == expected

    def test_only_the_edited_sheets_references_move(self) -> None:
        """A bare reference means the formula's own sheet, so a formula on
        Summary is untouched by an insertion into Data unless it says Data."""
        assert (
            shift_formula(
                "A5+Data!A5+Notes!A5",
                Shift.rows(5, 2),
                formula_sheet="Summary",
                target_sheet="Data",
            )
            == "A5+Data!A7+Notes!A5"
        )

    def test_a_bare_reference_on_the_edited_sheet_does_move(self) -> None:
        assert (
            shift_formula(
                "A5", Shift.rows(5, 2), formula_sheet="Data", target_sheet="Data"
            )
            == "A7"
        )

    def test_a_qualified_reference_to_another_sheet_stays(self) -> None:
        assert self.shift("Notes!A5", Shift.rows(5, 2)) == "Notes!A5"

    def test_nothing_to_do_returns_the_input(self) -> None:
        assert self.shift("A1", Shift.rows(5, 0)) == "A1"
        assert self.shift("", Shift.rows(5, 2)) == ""

    def test_a_reference_pushed_off_the_sheet_is_left_alone(self) -> None:
        """The caller refuses the insertion before this can matter; leaving
        it written is the conservative half."""
        assert self.shift(f"A{MAX_ROW}", Shift.rows(1, 5)) == f"A{MAX_ROW}"


class TestShiftRange:
    @pytest.mark.parametrize(
        ("block", "expected"),
        [
            ("A1:A10", "A1:A12"),
            ("A6:A9", "A8:A11"),
            ("A1:A3", "A1:A3"),
            ("A5", "A7"),
        ],
    )
    def test_rows(self, block: str, expected: str) -> None:
        assert shift_range(RangeRef.parse(block), Shift.rows(5, 2)).a1 == expected

    def test_columns(self) -> None:
        assert shift_range(RangeRef.parse("A1:C1"), Shift.columns(2, 1)).a1 == "A1:D1"

    def test_it_clamps_at_the_sheet_edge(self) -> None:
        moved = shift_range(RangeRef.parse(f"A{MAX_ROW}"), Shift.rows(1, 5))
        assert moved.start.row == MAX_ROW


class TestInsertingRows:
    def test_values_move_down(self, book: Workbook) -> None:
        sheet = book["Data"]
        assert sheet["A3"].value == "South"
        sheet.insert_rows(3, 2)
        assert sheet["A5"].value == "South"
        assert sheet["A2"].value == "North", "the row above did not move"

    def test_the_new_rows_are_blank(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.insert_rows(3, 2)
        assert sheet["A3"].value is None
        assert sheet["A4"].value is None

    def test_one_row_by_default(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.insert_rows(3)
        assert sheet["A4"].value == "South"

    def test_formulas_move_and_ranges_grow(self, book: Workbook) -> None:
        sheet = book["Data"]
        assert sheet["D6"].formula == "SUM(D2:D5)"
        sheet.insert_rows(3, 2)
        assert sheet["D8"].formula == "SUM(D2:D7)"

    def test_a_shared_formula_group_follows(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.insert_rows(3, 2)
        assert sheet["D2"].formula == "B2*C2", "the master stayed above the insertion"
        assert sheet["D5"].formula == "B5*C5", "and its follower reads its moved inputs"
        assert sheet["D7"].formula == "B7*C7"

    def test_merged_ranges_follow(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.insert_rows(3, 2)
        assert [b.a1 for b in sheet.merged_ranges] == ["A13:C13"]

    def test_the_dimension_grows(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.insert_rows(3, 2)
        assert sheet.dimension is not None
        assert sheet.dimension.a1 == "A1:F13"

    def test_rows_stay_in_order(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.insert_rows(3, 2)
        data = sheet.document.root.require("sheetData")
        numbers = [int(r.get("r") or 0) for r in data.children_named("row")]
        assert numbers == sorted(numbers)

    def test_cells_keep_their_column_and_gain_a_row(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.insert_rows(3, 2)
        raw = book.package.read(sheet.part_name).decode()
        assert 'r="A5"' in raw and 'r="F5"' in raw
        assert 'r="A3"' not in raw, "row 3 is empty, so it carries no cells"

    def test_a_formula_on_another_sheet_follows(self, book: Workbook) -> None:
        summary = book.add_sheet("Summary")
        summary["A1"].formula = "=SUM(Data!D2:D5)"
        summary["A2"].formula = "=Data!B2"
        book["Data"].insert_rows(3, 2)
        assert summary["A1"].formula == "SUM(Data!D2:D7)"
        assert summary["A2"].formula == "Data!B2", "it pointed above the insertion"

    def test_another_sheets_own_references_do_not_move(self, book: Workbook) -> None:
        summary = book.add_sheet("Summary")
        summary["A1"].formula = "=A5"
        book["Data"].insert_rows(3, 2)
        assert summary["A1"].formula == "A5"

    def test_defined_names_follow(self, book: Workbook) -> None:
        book.add_defined_name("Totals", "Data!$D$2:$D$5")
        book["Data"].insert_rows(3, 2)
        assert book.defined_name("Totals").refers_to == "Data!$D$2:$D$7"

    def test_tables_follow(self, live_structures_xlsx: Path) -> None:
        structures = Workbook.open(live_structures_xlsx)
        sheet = structures["Tabled"]
        assert structures.table("SalesTable").ref.a1 == "A1:C5"
        sheet.insert_rows(2, 1)
        table = structures.table("SalesTable")
        assert table.ref.a1 == "A1:C6"
        assert table.filter_ref is not None
        assert table.filter_ref.a1 == "A1:C5"

    def test_a_column_inserted_inside_a_table_joins_it(self, live_structures_xlsx: Path) -> None:
        """Measured: Excel names it ``ColumnN``, the smallest ``N`` free,
        numbers it after the highest id, and writes the name in the header.
        Left as it was, the table is wider than its columns and Excel
        refuses the workbook."""
        structures = Workbook.open(live_structures_xlsx)
        sheet = structures["Tabled"]
        sheet.insert_columns(2, 2)
        table = structures.table("SalesTable")
        assert table.ref.a1 == "A1:E5"
        assert table.column_names == ["Region", "Column1", "Column2", "Units", "Revenue"]
        assert [sheet["B1"].value, sheet["C1"].value] == ["Column1", "Column2"]
        container = table.document.root.require("tableColumns")
        assert container.get("count") == "5"
        ids = [entry.get("id") for entry in container.children_named("tableColumn")]
        assert ids == ["1", "4", "5", "2", "3"]

    def test_a_column_inserted_at_a_tables_edge_moves_it(self, live_structures_xlsx: Path) -> None:
        structures = Workbook.open(live_structures_xlsx)
        sheet = structures["Tabled"]
        sheet.insert_columns(1)
        assert structures.table("SalesTable").ref.a1 == "B1:D5"
        sheet.insert_columns(5)
        table = structures.table("SalesTable")
        assert table.ref.a1 == "B1:D5"
        assert table.column_names == ["Region", "Units", "Revenue"]

    def test_it_survives_a_save(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.insert_rows(3, 2)
        reopened = Workbook.from_bytes(book.to_bytes())["Data"]
        assert reopened["A5"].value == "South"
        assert reopened["D8"].formula == "SUM(D2:D7)"
        assert [b.a1 for b in reopened.merged_ranges] == ["A13:C13"]

    def test_inserting_at_row_one(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.insert_rows(1, 1)
        assert sheet["A2"].value == "Region", "the header moved down"
        assert sheet["A1"].value is None

    def test_inserting_below_everything(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.insert_rows(100, 1)
        assert sheet["A2"].value == "North", "nothing moved"


class TestInsertingColumns:
    def test_values_move_right(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.insert_columns(2, 1)
        assert sheet["A1"].value == "Region", "the column before stayed"
        assert sheet["C1"].value == "Units", "what was in B is now in C"
        assert sheet["B1"].value is None

    def test_formulas_move(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.insert_columns(2, 1)
        assert sheet["E2"].formula == "C2*D2"
        assert sheet["E6"].formula == "SUM(E2:E5)"

    def test_a_merge_spanning_the_insertion_grows(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.insert_columns(2, 1)
        assert [b.a1 for b in sheet.merged_ranges] == ["A11:D11"]

    def test_column_widths_follow(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.set_column_width(3, 30)
        sheet.insert_columns(2, 1)
        assert sheet.column_width(4) == 30
        assert sheet.column_width(3) is None

    def test_cells_stay_in_ascending_order(self, book: Workbook) -> None:
        from pyofficeeditor.excel import CellRef

        sheet = book["Data"]
        sheet.insert_columns(2, 1)
        data = sheet.document.root.require("sheetData")
        for row in data.children_named("row"):
            columns = [CellRef.parse(c.get("r") or "A1").column for c in row.children_named("c")]
            assert columns == sorted(columns), f"row {row.get('r')}"

    def test_it_survives_a_save(self, book: Workbook) -> None:
        book["Data"].insert_columns(2, 1)
        reopened = Workbook.from_bytes(book.to_bytes())["Data"]
        assert reopened["C1"].value == "Units"
        assert reopened["E2"].formula == "C2*D2"


class TestNothingIsRefused:
    """Every element that used to make an insertion raise now moves.

    The refusal list is empty and the machinery is gone, so these assert the
    replacement rather than the absence: each element's address has to come
    out somewhere new, not merely fail to raise.
    """

    @pytest.mark.parametrize(
        ("container", "entry", "attribute", "before", "after"),
        [
            ("dataValidations", "dataValidation", "sqref", "B2:B9", "B2:B11"),
            ("protectedRanges", "protectedRange", "sqref", "B2:B9", "B2:B11"),
            ("ignoredErrors", "ignoredError", "sqref", "B2:B9", "B2:B11"),
            ("dataConsolidate", "dataRef", "ref", "B2:B9", "B2:B11"),
        ],
    )
    def test_a_wrapped_address_moves(
        self, book: Workbook, container: str, entry: str, attribute: str, before: str, after: str
    ) -> None:
        sheet = book["Data"]
        holder = Element.create(container)
        holder.append(Element.create(entry, {attribute: before}))
        insert_in_schema_order(sheet.document.root, holder, WORKSHEET_CHILD_ORDER)
        sheet.insert_rows(3, 2)
        found = sheet.document.root.require(container).children_named(entry)
        assert [node.get(attribute) for node in found] == [after]

    def test_a_sort_state_and_its_condition(self, book: Workbook) -> None:
        sheet = book["Data"]
        state = Element.create("sortState", {"ref": "A2:C9"})
        state.append(Element.create("sortCondition", {"ref": "B2:B9"}))
        insert_in_schema_order(sheet.document.root, state, WORKSHEET_CHILD_ORDER)
        sheet.insert_rows(3, 2)
        moved = sheet.document.root.require("sortState")
        assert moved.get("ref") == "A2:C11"
        assert moved.require("sortCondition").get("ref") == "B2:B11"

    def test_a_scenarios_input_cells(self, book: Workbook) -> None:
        sheet = book["Data"]
        holder = Element.create("scenarios")
        scenario = Element.create("scenario", {"name": "High"})
        scenario.append(Element.create("inputCells", {"r": "A2", "val": "1"}))
        scenario.append(Element.create("inputCells", {"r": "A5", "val": "2"}))
        holder.append(scenario)
        insert_in_schema_order(sheet.document.root, holder, WORKSHEET_CHILD_ORDER)
        sheet.insert_rows(3, 2)
        cells = sheet.document.root.require("scenarios").require("scenario")
        assert [node.get("r") for node in cells.children_named("inputCells")] == ["A2", "A7"]

    def test_a_custom_sheet_views_selection(self, book: Workbook) -> None:
        sheet = book["Data"]
        holder = Element.create("customSheetViews")
        view = Element.create("customSheetView", {"guid": "{0}"})
        view.append(Element.create("selection", {"sqref": "B5:B9", "activeCell": "B5"}))
        holder.append(view)
        insert_in_schema_order(sheet.document.root, holder, WORKSHEET_CHILD_ORDER)
        sheet.insert_rows(3, 2)
        selection = sheet.document.root.require("customSheetViews").require(
            "customSheetView"
        ).require("selection")
        assert selection.get("sqref") == "B7:B11"
        assert selection.get("activeCell") == "B7"

    def test_an_inline_control_anchor(self, book: Workbook) -> None:
        """Zero-based, so row 9 is the tenth row and lands on the twelfth."""
        sheet = book["Data"]
        holder = Element.create("controls")
        anchor = Element.create("anchor")
        for end, row in (("from", "9"), ("to", "10")):
            node = Element.create(end)
            node.append(_indexed("xdr:col", "4"))
            node.append(_indexed("xdr:row", row))
            anchor.append(node)
        holder.append(anchor)
        insert_in_schema_order(sheet.document.root, holder, WORKSHEET_CHILD_ORDER)
        sheet.insert_rows(3, 2)
        ends = sheet.document.root.require("controls").require("anchor")
        rows = [
            ends.require(end).require("xdr:row").text for end in ("from", "to")
        ]
        assert rows == ["11", "12"]

    def test_an_extension_sqref(self, book: Workbook) -> None:
        """Everything newer than the 2006 schema addresses cells through an
        ``xm:sqref``, whatever the feature, so one rule covers them all."""
        sheet = book["Data"]
        extensions = Element.create("extLst")
        ext = Element.create("ext", {"uri": "{78C0D931-6437-407d-A8EE-F0AAD7539E65}"})
        block = Element.create("x14:conditionalFormatting")
        sqref = Element.create("xm:sqref")
        sqref.set_text("D2:D9")
        block.append(sqref)
        ext.append(block)
        extensions.append(ext)
        insert_in_schema_order(sheet.document.root, extensions, WORKSHEET_CHILD_ORDER)
        sheet.insert_rows(3, 2)
        rendered = sheet.document.to_bytes().decode()
        assert "<xm:sqref>D2:D11</xm:sqref>" in rendered

    def test_an_extension_formula(self, book: Workbook) -> None:
        sheet = book["Data"]
        extensions = Element.create("extLst")
        ext = Element.create("ext", {"uri": "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"})
        formula = Element.create("xm:f")
        formula.set_text("SUM(B2:B9)")
        ext.append(formula)
        extensions.append(ext)
        insert_in_schema_order(sheet.document.root, extensions, WORKSHEET_CHILD_ORDER)
        sheet.insert_rows(3, 2)
        assert "<xm:f>SUM(B2:B11)</xm:f>" in sheet.document.to_bytes().decode()


def _indexed(name: str, value: str) -> Element:
    node = Element.create(name)
    node.set_text(value)
    return node


class TestRefusals:
    @pytest.mark.parametrize(("at", "count"), [(0, 1), (-1, 1), (MAX_ROW + 1, 1), (3, 0), (3, -1)])
    def test_bad_arguments(self, book: Workbook, at: int, count: int) -> None:
        with pytest.raises(ValueError):
            book["Data"].insert_rows(at, count)

    @pytest.mark.parametrize(("at", "count"), [(0, 1), (MAX_COLUMN + 1, 1), (3, 0)])
    def test_bad_column_arguments(self, book: Workbook, at: int, count: int) -> None:
        with pytest.raises(ValueError):
            book["Data"].insert_columns(at, count)

    def test_pushing_content_off_the_sheet_is_refused(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet[f"A{MAX_ROW}"].value = "last row"
        with pytest.raises(ValueError, match=f"past {MAX_ROW}"):
            sheet.insert_rows(1, 1)

    def test_pushing_columns_off_the_sheet_is_refused(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet["XFD1"].value = "last column"
        with pytest.raises(ValueError, match=f"past {MAX_COLUMN}"):
            sheet.insert_columns(1, 1)


class TestNoOpDiscipline:
    def test_inserting_nowhere_useful_still_marks_the_workbook(self, book: Workbook) -> None:
        assert book.is_modified is False
        book["Data"].insert_rows(100, 1)
        assert book.is_modified is True

    def test_other_parts_keep_their_bytes(self, book: Workbook, live_sample_xlsx: Path) -> None:
        from pyofficeeditor._zip import ZipArchive

        book["Data"].insert_rows(3, 2)
        before = ZipArchive.from_bytes(live_sample_xlsx.read_bytes())
        after = ZipArchive.from_bytes(book.to_bytes())
        expected = {
            "xl/worksheets/sheet1.xml",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "[Content_Types].xml",
        }
        for name in after.names():
            if name in expected:
                continue
            assert after.member(name).stored == before.member(name).stored, name
