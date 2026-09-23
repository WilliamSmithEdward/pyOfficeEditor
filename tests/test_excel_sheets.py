"""Adding, removing, renaming and reordering sheets.

Four things have to line up for Excel to show a new sheet: the part, its
content-type override, a relationship from the workbook part, and an entry in
``<sheets>``. Miss one and the sheet is silently absent or the file does not
open, so each is asserted separately rather than trusting a round trip.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyofficeeditor._xml import Element, XmlDocument
from pyofficeeditor.excel import Workbook
from pyofficeeditor.excel._formulas import quote_sheet_name, rename_sheet_in_formula
from pyofficeeditor.excel._schema import (
    WORKBOOK_CHILD_ORDER,
    WORKSHEET_CHILD_ORDER,
    insert_in_schema_order,
)
from pyofficeeditor.excel.workbook import (
    CT_WORKSHEET,
    MAX_SHEET_NAME_LENGTH,
    RT_WORKSHEET,
    check_sheet_name,
)


@pytest.fixture()
def book(live_sample_xlsx: Path) -> Workbook:
    return Workbook.open(live_sample_xlsx)


class TestSheetNameRules:
    @pytest.mark.parametrize("name", ["Data", "A", "x" * MAX_SHEET_NAME_LENGTH, "Q1 2026", "a-b_c.d", "Ünïcode"])
    def test_names_excel_accepts(self, name: str) -> None:
        check_sheet_name(name)

    @pytest.mark.parametrize(
        ("name", "expected"),
        [
            ("", "cannot be empty"),
            ("x" * (MAX_SHEET_NAME_LENGTH + 1), "at most 31"),
            ("a:b", "formula reference"),
            ("a/b", "formula reference"),
            ("a\\b", "formula reference"),
            ("a?b", "formula reference"),
            ("a*b", "formula reference"),
            ("a[b]", "formula reference"),
            ("'quoted", "apostrophe"),
            ("quoted'", "apostrophe"),
            ("History", "reserved"),
            ("history", "reserved"),
        ],
    )
    def test_names_excel_refuses(self, name: str, expected: str) -> None:
        with pytest.raises(ValueError, match=expected):
            check_sheet_name(name)

    def test_uniqueness_ignores_case(self) -> None:
        with pytest.raises(ValueError, match="already has a sheet"):
            check_sheet_name("data", taken={"Data"})
        check_sheet_name("other", taken={"Data"})


class TestQuotingASheetName:
    @pytest.mark.parametrize(
        ("name", "expected"),
        [
            ("Data", "Data"),
            ("Notes", "Notes"),
            ("Sheet_1", "Sheet_1"),
            ("My Sheet", "'My Sheet'"),
            ("it's", "'it''s'"),
            ("has-dash", "'has-dash'"),
        ],
    )
    def test_quoting(self, name: str, expected: str) -> None:
        assert quote_sheet_name(name) == expected

    def test_a_name_that_looks_like_a_reference_is_quoted(self) -> None:
        """``=Q1!A1`` would read as a reference to column Q, so the name has
        to be quoted even though it is a bare identifier."""
        assert quote_sheet_name("Q1") == "'Q1'"
        assert quote_sheet_name("A1") == "'A1'"
        assert quote_sheet_name("XFD1") == "'XFD1'"


class TestRenamingInsideAFormula:
    @pytest.mark.parametrize(
        ("formula", "expected"),
        [
            ("Data!A1", "'Q1 Data'!A1"),
            ("'Data'!A1", "'Q1 Data'!A1"),
            ("SUM(Data!A1:A5)", "SUM('Q1 Data'!A1:A5)"),
            ("Data!A1+Notes!B2", "'Q1 Data'!A1+Notes!B2"),
            ("Data!A1+Data!A2", "'Q1 Data'!A1+'Q1 Data'!A2"),
        ],
    )
    def test_references_are_repointed(self, formula: str, expected: str) -> None:
        assert rename_sheet_in_formula(formula, "Data", "Q1 Data") == expected

    @pytest.mark.parametrize("formula", ["MyData!A1", "XData!A1", "DataX!A1", 'CONCAT("Data!A1")', '"Data"'])
    def test_what_must_not_be_touched(self, formula: str) -> None:
        """A longer name that merely contains the old one, and text that
        merely looks like a reference."""
        assert rename_sheet_in_formula(formula, "Data", "Q1 Data") == formula

    def test_quoted_becomes_bare_when_the_new_name_allows(self) -> None:
        assert rename_sheet_in_formula("'Q1 Data'!A1", "Q1 Data", "Data") == "Data!A1"

    def test_a_literal_beside_a_reference(self) -> None:
        assert (
            rename_sheet_in_formula('Data!A1&"Data"', "Data", "Q1")
            == "'Q1'!A1&\"Data\""
        )

    def test_renaming_to_itself_changes_nothing(self) -> None:
        assert rename_sheet_in_formula("Data!A1", "Data", "Data") == "Data!A1"


class TestSchemaOrder:
    def test_dimension_lands_after_sheet_pr(self) -> None:
        """A freshly authored sheet starts with ``sheetPr``. Putting
        ``dimension`` at the front would break the schema's sequence, and
        Excel refuses such a file rather than repairing it."""
        document = XmlDocument.parse(
            b'<worksheet><sheetPr codeName="Sheet1"/><sheetData/></worksheet>'
        )
        insert_in_schema_order(document.root, Element.create("dimension", {"ref": "A1"}), WORKSHEET_CHILD_ORDER)
        assert document.to_bytes() == (
            b'<worksheet><sheetPr codeName="Sheet1"/><dimension ref="A1"/><sheetData/></worksheet>'
        )

    def test_dimension_lands_first_when_there_is_no_sheet_pr(self) -> None:
        document = XmlDocument.parse(b"<worksheet><sheetData/></worksheet>")
        insert_in_schema_order(document.root, Element.create("dimension", {"ref": "A1"}), WORKSHEET_CHILD_ORDER)
        assert document.to_bytes() == b'<worksheet><dimension ref="A1"/><sheetData/></worksheet>'

    def test_an_unknown_name_is_appended(self) -> None:
        document = XmlDocument.parse(b"<worksheet><sheetData/></worksheet>")
        insert_in_schema_order(document.root, Element.create("somethingNew"), WORKSHEET_CHILD_ORDER)
        assert document.to_bytes() == b"<worksheet><sheetData/><somethingNew/></worksheet>"

    def test_the_order_is_a_sequence_with_no_duplicates(self) -> None:
        assert len(WORKSHEET_CHILD_ORDER) == len(set(WORKSHEET_CHILD_ORDER))
        assert WORKSHEET_CHILD_ORDER[0] == "sheetPr"
        assert WORKSHEET_CHILD_ORDER[-1] == "extLst"

    def test_writing_into_a_sheet_that_has_sheet_pr_keeps_the_order(
        self, live_sample_xlsx: Path
    ) -> None:
        """``empty.xlsx`` has ``sheetPr`` and a ``dimension``. A sheet this
        library adds has no ``sheetPr``; either way a write must not disturb
        the sequence."""
        book = Workbook.open(live_sample_xlsx)
        sheet = book.add_sheet("Fresh")
        sheet["B2"].value = 1
        names = [
            child.name
            for child in sheet.document.root.elements()
            if child.name in WORKSHEET_CHILD_ORDER
        ]
        positions = [WORKSHEET_CHILD_ORDER.index(n) for n in names]
        assert positions == sorted(positions), names


class TestAddingASheet:
    def test_all_four_pieces_are_written(self, book: Workbook) -> None:
        sheet = book.add_sheet("Summary")
        package = book.package

        assert package.has_part(sheet.part_name), "the part"
        assert package.content_types.of(sheet.part_name) == CT_WORKSHEET, "the content type"
        targets = [
            r.target_part for r in package.relationships("xl/workbook.xml").by_type(RT_WORKSHEET)
        ]
        assert sheet.part_name in targets, "the relationship"
        entries = [
            e.get("name")
            for e in book.package.xml("xl/workbook.xml").root.require("sheets").children_named("sheet")
        ]
        assert entries == ["Data", "Notes", "Summary"], "the sheets entry"

    def test_it_appends_by_default(self, book: Workbook) -> None:
        book.add_sheet("Summary")
        assert book.sheet_names == ["Data", "Notes", "Summary"]

    @pytest.mark.parametrize(
        ("index", "expected"),
        [
            (0, ["New", "Data", "Notes"]),
            (1, ["Data", "New", "Notes"]),
            (2, ["Data", "Notes", "New"]),
            (99, ["Data", "Notes", "New"]),
            (-5, ["New", "Data", "Notes"]),
        ],
    )
    def test_it_inserts_where_asked(self, book: Workbook, index: int, expected: list[str]) -> None:
        book.add_sheet("New", index=index)
        assert book.sheet_names == expected

    def test_the_new_sheet_is_usable(self, book: Workbook) -> None:
        sheet = book.add_sheet("Summary")
        sheet["A1"].value = "Total"
        sheet["B1"].formula = "=SUM(Data!D2:D5)"
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Summary"]["A1"].value == "Total"
        assert reopened["Summary"]["B1"].formula == "SUM(Data!D2:D5)"

    def test_the_part_name_avoids_collisions(self, book: Workbook) -> None:
        """The fixture already has sheet1 and sheet2."""
        assert book.add_sheet("Third").part_name == "xl/worksheets/sheet3.xml"
        assert book.add_sheet("Fourth").part_name == "xl/worksheets/sheet4.xml"

    def test_sheet_ids_stay_unique(self, book: Workbook) -> None:
        book.add_sheet("Third")
        entries = book.package.xml("xl/workbook.xml").root.require("sheets")
        ids = [e.get("sheetId") for e in entries.children_named("sheet")]
        assert len(ids) == len(set(ids))

    def test_a_duplicate_name_is_refused(self, book: Workbook) -> None:
        with pytest.raises(ValueError, match="already has a sheet"):
            book.add_sheet("Data")
        with pytest.raises(ValueError, match="already has a sheet"):
            book.add_sheet("data")

    def test_an_invalid_name_is_refused_before_anything_is_written(self, book: Workbook) -> None:
        before = set(book.package.part_names())
        with pytest.raises(ValueError):
            book.add_sheet("bad:name")
        assert set(book.package.part_names()) == before
        assert book.sheet_names == ["Data", "Notes"]


class TestRemovingASheet:
    def test_all_four_pieces_go(self, book: Workbook) -> None:
        part = book["Notes"].part_name
        book.remove_sheet("Notes")
        package = book.package
        assert not package.has_part(part)
        # The Override goes. The package-wide Default for .xml still matches
        # the name, which is harmless once no part is there to match.
        assert part not in package.content_types.overrides
        assert package.content_types.of(part) != CT_WORKSHEET
        assert part not in [
            r.target_part for r in package.relationships("xl/workbook.xml").by_type(RT_WORKSHEET)
        ]
        assert book.sheet_names == ["Data"]

    def test_it_survives_a_save(self, book: Workbook) -> None:
        book.remove_sheet("Notes")
        assert Workbook.from_bytes(book.to_bytes()).sheet_names == ["Data"]

    def test_removing_the_last_sheet_is_refused(self, book: Workbook) -> None:
        book.remove_sheet("Notes")
        with pytest.raises(ValueError, match="only sheet"):
            book.remove_sheet("Data")

    def test_removing_an_absent_sheet_raises(self, book: Workbook) -> None:
        with pytest.raises(KeyError, match="Data, Notes"):
            book.remove_sheet("Nope")

    def test_the_active_tab_is_clamped(self, book: Workbook) -> None:
        """``activeTab`` is an index. Pointing it past the end makes Excel
        repair the file."""
        book.add_sheet("Third")
        book.active = "Third"
        book.remove_sheet("Third")
        view = book.package.xml("xl/workbook.xml").root.require("bookViews").require("workbookView")
        assert int(view.get("activeTab") or "0") < len(book.sheet_names)


class TestRenamingASheet:
    def test_the_entry_changes(self, book: Workbook) -> None:
        book.rename_sheet("Data", "Q1 Data")
        assert book.sheet_names == ["Q1 Data", "Notes"]
        assert "Data" not in book
        assert book["Q1 Data"]["A2"].value == "North"

    def test_the_sheet_renames_itself_the_same_way(self, book: Workbook) -> None:
        """It only relabelled the object once, leaving the file's name as
        it was and the two disagreeing."""
        summary = book.add_sheet("Summary")
        summary["B1"].formula = "=SUM(Data!D2:D5)"
        sheet = book["Data"]
        sheet.rename("Q1 Data")
        assert (sheet.name, book.sheet_names) == ("Q1 Data", ["Q1 Data", "Notes", "Summary"])
        assert book["Q1 Data"] is sheet
        assert book["Summary"]["B1"].formula == "SUM('Q1 Data'!D2:D5)"

    def test_a_sheet_refuses_a_name_the_workbook_would(self, book: Workbook) -> None:
        with pytest.raises(ValueError):
            book["Data"].rename("Notes")
        assert book["Data"].name == "Data"

    def test_formulas_that_referenced_it_are_repointed(self, book: Workbook) -> None:
        summary = book.add_sheet("Summary")
        summary["B1"].formula = "=SUM(Data!D2:D5)"
        book.rename_sheet("Data", "Q1 Data")
        assert book["Summary"]["B1"].formula == "SUM('Q1 Data'!D2:D5)"

    def test_the_shared_formula_group_is_not_broken(self, book: Workbook) -> None:
        """Rewriting followers individually would give each its own text and
        destroy the group, so only masters carry text to rewrite."""
        book.rename_sheet("Data", "Q1 Data")
        sheet = book["Q1 Data"]
        assert sheet["D2"].formula == "B2*C2"
        assert sheet["D3"].formula == "B3*C3"
        raw = book.package.read(sheet.part_name)
        assert raw.count(b'<f t="shared" si="0"/>') == 3, "the followers stayed empty"

    def test_defined_names_are_repointed(self, book: Workbook) -> None:
        """A name scoped to a sheet points at it by name, so a rename that
        misses it leaves the name resolving to ``#REF!``."""
        root = book.package.xml("xl/workbook.xml").root
        names = Element.create("definedNames")
        entry = Element.create("definedName", {"name": "Totals"})
        entry.set_text("Data!$D$2:$D$5")
        names.append(entry)
        insert_in_schema_order(root, names, WORKBOOK_CHILD_ORDER)

        book.rename_sheet("Data", "Q1 Data")
        assert entry.text == "'Q1 Data'!$D$2:$D$5"

    def test_defined_names_land_in_schema_order(self, book: Workbook) -> None:
        """``definedNames`` sits between ``sheets`` and ``calcPr``."""
        root = book.package.xml("xl/workbook.xml").root
        insert_in_schema_order(root, Element.create("definedNames"), WORKBOOK_CHILD_ORDER)
        names = [
            child.name for child in root.elements() if child.name in WORKBOOK_CHILD_ORDER
        ]
        positions = [WORKBOOK_CHILD_ORDER.index(n) for n in names]
        assert positions == sorted(positions), names

    def test_a_duplicate_name_is_refused(self, book: Workbook) -> None:
        with pytest.raises(ValueError, match="already has a sheet"):
            book.rename_sheet("Data", "Notes")

    def test_renaming_to_the_same_name_is_a_no_op(self, book: Workbook, live_sample_xlsx: Path) -> None:
        book.rename_sheet("Data", "Data")
        assert book.to_bytes() == live_sample_xlsx.read_bytes()

    def test_an_invalid_name_is_refused(self, book: Workbook) -> None:
        with pytest.raises(ValueError, match="formula reference"):
            book.rename_sheet("Data", "bad[name]")
        assert book.sheet_names == ["Data", "Notes"]

    def test_it_survives_a_save(self, book: Workbook) -> None:
        book.rename_sheet("Data", "Q1 Data")
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened.sheet_names == ["Q1 Data", "Notes"]
        assert reopened["Q1 Data"]["A2"].value == "North"


class TestMovingASheet:
    def test_moving_to_the_front(self, book: Workbook) -> None:
        book.move_sheet("Notes", 0)
        assert book.sheet_names == ["Notes", "Data"]

    def test_moving_to_the_back(self, book: Workbook) -> None:
        book.move_sheet("Data", 1)
        assert book.sheet_names == ["Notes", "Data"]

    def test_an_index_out_of_range_clamps(self, book: Workbook) -> None:
        book.move_sheet("Data", 99)
        assert book.sheet_names == ["Notes", "Data"]
        book.move_sheet("Data", -5)
        assert book.sheet_names == ["Data", "Notes"]

    def test_moving_to_its_own_position_is_a_no_op(self, book: Workbook, live_sample_xlsx: Path) -> None:
        book.move_sheet("Data", 0)
        assert book.to_bytes() == live_sample_xlsx.read_bytes()

    def test_the_active_sheet_stays_active(self, book: Workbook) -> None:
        """``activeTab`` is an index, so a move has to follow the sheet."""
        assert book.active.name == "Data"
        book.move_sheet("Data", 1)
        assert book.active.name == "Data"

    def test_moving_an_absent_sheet_raises(self, book: Workbook) -> None:
        with pytest.raises(KeyError, match="Data, Notes"):
            book.move_sheet("Nope", 0)

    def test_it_survives_a_save(self, book: Workbook) -> None:
        book.move_sheet("Notes", 0)
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened.sheet_names == ["Notes", "Data"]
        assert reopened["Data"]["A2"].value == "North"


class TestCombined:
    def test_a_sequence_of_operations_holds_together(self, book: Workbook) -> None:
        summary = book.add_sheet("Summary")
        summary["A1"].value = "Total"
        summary["B1"].formula = "=SUM(Data!D2:D5)"
        book.add_sheet("Scratch", index=0)
        book.move_sheet("Summary", 1)
        book.rename_sheet("Data", "Q1 Data")
        book.remove_sheet("Scratch")

        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened.sheet_names == ["Summary", "Q1 Data", "Notes"]
        assert reopened["Summary"]["B1"].formula == "SUM('Q1 Data'!D2:D5)"
        assert reopened["Q1 Data"]["A2"].value == "North"
        assert reopened["Q1 Data"]["D3"].formula == "B3*C3"
        assert reopened["Notes"]["A1"].value == "Second sheet"

    def test_every_relationship_still_resolves(self, book: Workbook) -> None:
        book.add_sheet("Summary")
        book.rename_sheet("Data", "Q1 Data")
        book.remove_sheet("Notes")
        package = Workbook.from_bytes(book.to_bytes()).package
        for source in ["", *package.part_names()]:
            for relationship in package.relationships(source):
                if not relationship.is_external:
                    assert package.has_part(relationship.target_part), (
                        f"{source or 'root'} -> {relationship.target}"
                    )

    def test_every_part_still_has_a_content_type(self, book: Workbook) -> None:
        book.add_sheet("Summary")
        book.remove_sheet("Notes")
        package = Workbook.from_bytes(book.to_bytes()).package
        for name in package.part_names():
            if name == "[Content_Types].xml":
                continue
            assert package.content_types.of(name) is not None, name
