"""Defined names.

The dangerous part is scope. A sheet-scoped name records its sheet as a
*position* in the sheet order, so adding, removing or moving a sheet silently
rescopes every name after it unless the index moves too. Most of these tests
are about that.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from pyofficeeditor.excel import Workbook
from pyofficeeditor.excel._names import (
    BUILTIN_NAME_PREFIX,
    MAX_NAME_LENGTH,
    DefinedName,
    check_name,
)


@pytest.fixture()
def book(live_structures_xlsx: Path) -> Workbook:
    return Workbook.open(live_structures_xlsx)


class TestNameRules:
    @pytest.mark.parametrize(
        "name", ["Total", "_private", "a", "With.Dots", "x" * MAX_NAME_LENGTH, "\\odd"]
    )
    def test_names_excel_accepts(self, name: str) -> None:
        check_name(name)

    @pytest.mark.parametrize(
        ("name", "expected"),
        [
            ("", "cannot be empty"),
            ("x" * (MAX_NAME_LENGTH + 1), "at most 255"),
            ("has space", "no spaces"),
            ("1st", "must start with"),
            ("has-dash", "no spaces or operator"),
            ("A1", "reads as a cell reference"),
            ("C", "R1C1"),
            ("_xlnm.Print_Area", "reserves"),
        ],
    )
    def test_names_excel_refuses(self, name: str, expected: str) -> None:
        with pytest.raises(ValueError, match=expected):
            check_name(name)

    def test_the_message_says_what_kind_of_name(self) -> None:
        with pytest.raises(ValueError, match="table name"):
            check_name("has space", what="table name")

    def test_tables_and_defined_names_share_the_rules(self) -> None:
        """A table name is a defined name, so one check serves both."""
        from pyofficeeditor.excel._tables import check_table_name

        for bad in ("has space", "A1", "1st"):
            with pytest.raises(ValueError):
                check_table_name(bad)
            with pytest.raises(ValueError):
                check_name(bad)


class TestReadingRealNames:
    def test_both_scopes_are_read(self, book: Workbook) -> None:
        names = {entry.name: entry for entry in book.defined_names}
        assert names["TotalUnits"].scope is None, "workbook-wide"
        assert names["LocalRegion"].scope == "Tabled", "scoped to a sheet"

    def test_the_refers_to_text(self, book: Workbook) -> None:
        assert book.defined_name("TotalUnits").refers_to == "Tabled!$B$2:$B$4"
        assert book.defined_name("LocalRegion").refers_to == "Tabled!$A$2"

    def test_scope_is_reported_as_a_sheet_name(self, book: Workbook) -> None:
        """Not the index the file stores, which means a different sheet as
        soon as sheets are reordered."""
        assert book.defined_name("LocalRegion").scope == "Tabled"
        raw = book.package.read("xl/workbook.xml").decode()
        assert 'localSheetId="0"' in raw, "the file really does store a position"

    def test_lookup_by_scope(self, book: Workbook) -> None:
        assert book.defined_name("LocalRegion", scope="Tabled").name == "LocalRegion"
        with pytest.raises(KeyError, match="scoped to 'Second'"):
            book.defined_name("LocalRegion", scope="Second")

    def test_a_missing_name_says_what_is_there(self, book: Workbook) -> None:
        with pytest.raises(KeyError, match="LocalRegion, TotalUnits"):
            book.defined_name("Nope")

    def test_a_workbook_with_no_names(self, live_sample_xlsx: Path) -> None:
        assert Workbook.open(live_sample_xlsx).defined_names == []

    def test_reading_names_is_not_a_change(
        self, book: Workbook, live_structures_xlsx: Path
    ) -> None:
        _ = book.defined_names
        assert book.is_modified is False
        assert book.to_bytes() == live_structures_xlsx.read_bytes()

    def test_is_builtin(self) -> None:
        assert DefinedName(BUILTIN_NAME_PREFIX + "Print_Area", "x").is_builtin is True
        assert DefinedName("Mine", "x").is_builtin is False


class TestAddingAndRemoving:
    def test_adding_a_workbook_scoped_name(self, book: Workbook) -> None:
        book.add_defined_name("Alpha", "Tabled!$A$1")
        reopened = Workbook.from_bytes(book.to_bytes())
        entry = reopened.defined_name("Alpha")
        assert entry.refers_to == "Tabled!$A$1"
        assert entry.scope is None

    def test_adding_a_sheet_scoped_name(self, book: Workbook) -> None:
        book.add_defined_name("Local", "Second!$A$1", scope="Second")
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened.defined_name("Local").scope == "Second"

    def test_a_leading_equals_is_dropped(self, book: Workbook) -> None:
        book.add_defined_name("Alpha", "=Tabled!$A$1")
        assert book.defined_name("Alpha").refers_to == "Tabled!$A$1"

    def test_names_are_written_in_name_order(self, book: Workbook) -> None:
        book.add_defined_name("Alpha", "Tabled!$A$1")
        book.add_defined_name("Zulu", "Tabled!$A$1")
        raw = book.package.read("xl/workbook.xml").decode()
        found = re.findall(r'<definedName name="([^"]+)"', raw)
        assert found == sorted(found, key=str.casefold)

    def test_the_same_name_at_both_scopes(self, book: Workbook) -> None:
        """A workbook-wide name and a sheet-scoped one may share a name, and
        the sheet-scoped one wins on that sheet."""
        book.add_defined_name("Totals", "Tabled!$A$1")
        book.add_defined_name("Totals", "Second!$A$1", scope="Second")
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened.defined_name("Totals").scope is None, "the workbook one is preferred"
        assert reopened.defined_name("Totals", scope="Second").refers_to == "Second!$A$1"

    def test_a_duplicate_at_the_same_scope_is_refused(self, book: Workbook) -> None:
        with pytest.raises(ValueError, match="this workbook already has a defined name"):
            book.add_defined_name("TotalUnits", "Tabled!$A$1")

    def test_the_message_names_the_scope_the_collision_is_in(self, book: Workbook) -> None:
        book.add_defined_name("Local", "Second!$A$1", scope="Second")
        with pytest.raises(ValueError, match="the sheet 'Second' already has"):
            book.add_defined_name("Local", "Second!$B$1", scope="Second")

    def test_an_invalid_name_is_refused(self, book: Workbook) -> None:
        with pytest.raises(ValueError, match="no spaces"):
            book.add_defined_name("has space", "Tabled!$A$1")

    def test_scoping_to_a_sheet_that_does_not_exist(self, book: Workbook) -> None:
        with pytest.raises(KeyError, match="no sheet named"):
            book.add_defined_name("Alpha", "x", scope="Nope")

    def test_removing_a_name(self, book: Workbook) -> None:
        book.remove_defined_name("TotalUnits")
        assert "TotalUnits" not in [e.name for e in book.defined_names]
        reopened = Workbook.from_bytes(book.to_bytes())
        assert "TotalUnits" not in [e.name for e in reopened.defined_names]

    def test_removing_a_scoped_name_leaves_the_other(self, book: Workbook) -> None:
        book.add_defined_name("Totals", "Tabled!$A$1")
        book.add_defined_name("Totals", "Second!$A$1", scope="Second")
        book.remove_defined_name("Totals", scope="Second")
        assert book.defined_name("Totals").scope is None

    def test_the_element_goes_when_the_last_name_does(self, book: Workbook) -> None:
        book.remove_defined_name("TotalUnits")
        book.remove_defined_name("LocalRegion")
        assert b"definedNames" not in book.package.read("xl/workbook.xml")

    def test_removing_an_absent_name(self, book: Workbook) -> None:
        with pytest.raises(KeyError, match="no defined name"):
            book.remove_defined_name("Nope")

    def test_the_element_lands_in_schema_order(self, live_sample_xlsx: Path) -> None:
        """``definedNames`` sits between ``sheets`` and ``calcPr``."""
        from pyofficeeditor.excel._schema import WORKBOOK_CHILD_ORDER

        book = Workbook.open(live_sample_xlsx)
        book.add_defined_name("Alpha", "Data!$A$1")
        root = book.package.xml("xl/workbook.xml").root
        names = [c.name for c in root.elements() if c.name in WORKBOOK_CHILD_ORDER]
        positions = [WORKBOOK_CHILD_ORDER.index(n) for n in names]
        assert positions == sorted(positions), names


class TestScopeSurvivesSheetChanges:
    """``localSheetId`` is a position, so every sheet operation can silently
    rescope a name. These are the tests that matter."""

    def test_moving_a_sheet(self, book: Workbook) -> None:
        book.add_defined_name("OnSecond", "Second!$A$1", scope="Second")
        assert book.sheet_names == ["Tabled", "Second"]
        book.move_sheet("Tabled", 1)
        assert book.sheet_names == ["Second", "Tabled"]
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened.defined_name("LocalRegion").scope == "Tabled"
        assert reopened.defined_name("OnSecond").scope == "Second"

    def test_adding_a_sheet_in_front(self, book: Workbook) -> None:
        book.add_sheet("First", index=0)
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened.defined_name("LocalRegion").scope == "Tabled"

    def test_removing_an_earlier_sheet(self, book: Workbook) -> None:
        book.add_defined_name("OnSecond", "Second!$A$1", scope="Second")
        book.add_sheet("Throwaway", index=0)
        book.remove_sheet("Throwaway")
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened.defined_name("LocalRegion").scope == "Tabled"
        assert reopened.defined_name("OnSecond").scope == "Second"

    def test_removing_the_scoped_sheet_removes_the_name(self, book: Workbook) -> None:
        """Which is what Excel does: a name scoped to a sheet that is gone
        has nowhere to live."""
        book.add_defined_name("OnSecond", "Second!$A$1", scope="Second")
        book.remove_sheet("Second")
        assert "OnSecond" not in [e.name for e in book.defined_names]
        assert "LocalRegion" in [e.name for e in book.defined_names], "the other survives"

    def test_renaming_a_sheet_repoints_the_reference_and_keeps_the_scope(
        self, book: Workbook
    ) -> None:
        book.rename_sheet("Tabled", "Q1 Data")
        reopened = Workbook.from_bytes(book.to_bytes())
        entry = reopened.defined_name("LocalRegion")
        assert entry.scope == "Q1 Data", "the scope followed the rename"
        assert entry.refers_to == "'Q1 Data'!$A$2", "and so did the reference"

    def test_a_name_scoped_to_a_sheet_that_is_gone_is_dropped_on_remap(
        self, book: Workbook
    ) -> None:
        book.add_defined_name("OnSecond", "Second!$A$1", scope="Second")
        before = len(book.defined_names)
        book.remove_sheet("Second")
        assert len(book.defined_names) == before - 1
