"""Hyperlinks and outline grouping.

The XML quoted here is Excel's own. The thing worth measuring is that an
external link's address is not in the element at all.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyofficeeditor.excel import Workbook, Worksheet


class TestHyperlinks:
    @pytest.fixture()
    def sheet(self, live_sample_xlsx: Path) -> Worksheet:
        return Workbook.open(live_sample_xlsx)["Data"]

    def test_a_clean_sheet_has_none(self, sheet: Worksheet) -> None:
        assert sheet.hyperlinks == []

    def test_an_external_link_goes_in_a_relationship(self, sheet: Worksheet) -> None:
        """Not in the element: ``<hyperlink r:id="rId1"/>`` carries no
        address, and a reader that ignores the relationship finds a link
        with nowhere to go."""
        sheet.add_hyperlink("B2", "https://example.com/a")
        rendered = sheet.document.to_bytes().decode()
        assert "https://example.com" not in rendered, "the URL is not in the sheet"
        assert "r:id=" in rendered
        assert sheet.hyperlinks[0].target == "https://example.com/a"

    def test_the_relationship_is_marked_external(self, sheet: Worksheet) -> None:
        sheet.add_hyperlink("B2", "https://example.com/a")
        relationships = sheet.workbook.package.relationships(sheet.part_name)
        found = [r for r in relationships if r.target == "https://example.com/a"]
        assert found and found[0].is_external

    def test_an_internal_link_has_no_relationship(self, sheet: Worksheet) -> None:
        sheet.add_hyperlink("B2", location="Notes!A1")
        assert 'location="Notes!A1"' in sheet.document.to_bytes().decode()
        assert "r:id=" not in sheet.document.to_bytes().decode()
        link = sheet.hyperlinks[0]
        assert link.target is None
        assert link.is_external is False

    def test_both_together(self, sheet: Worksheet) -> None:
        """A URL with a fragment: the page in the relationship, the part
        after the hash in location."""
        sheet.add_hyperlink("B2", "https://example.com/b", location="frag")
        link = sheet.hyperlinks[0]
        assert link.target == "https://example.com/b"
        assert link.location == "frag"

    def test_a_link_with_nowhere_to_go(self, sheet: Worksheet) -> None:
        with pytest.raises(ValueError, match="needs somewhere to go"):
            sheet.add_hyperlink("B2")

    def test_a_tooltip_and_display_text(self, sheet: Worksheet) -> None:
        sheet.add_hyperlink("B2", "https://example.com", display="Example", tooltip="go")
        link = sheet.hyperlinks[0]
        assert link.display == "Example"
        assert link.tooltip == "go"

    def test_a_link_over_a_range(self, sheet: Worksheet) -> None:
        """``ref`` is a range, so one entry covers the block."""
        sheet.add_hyperlink("G2:G4", "https://example.com/c")
        assert sheet.hyperlinks[0].ref.a1 == "G2:G4"
        assert sheet.hyperlink_at("G3") is not None

    def test_finding_one_on_a_cell(self, sheet: Worksheet) -> None:
        sheet.add_hyperlink("B2", "https://example.com")
        assert sheet.hyperlink_at("B2") is not None
        assert sheet.hyperlink_at("Z99") is None

    def test_adding_over_an_existing_one_replaces_it(self, sheet: Worksheet) -> None:
        sheet.add_hyperlink("B2", "https://one.example")
        sheet.add_hyperlink("B2", "https://two.example")
        assert [link.target for link in sheet.hyperlinks] == ["https://two.example"]

    def test_removing_takes_the_relationship_too(self, sheet: Worksheet) -> None:
        """A relationship nothing points at is something Excel repairs."""
        sheet.add_hyperlink("B2", "https://example.com/a")
        assert sheet.remove_hyperlink("B2") == 1
        assert sheet.hyperlinks == []
        relationships = sheet.workbook.package.relationships(sheet.part_name)
        assert not [r for r in relationships if r.target == "https://example.com/a"]

    def test_the_container_goes_with_the_last_link(self, sheet: Worksheet) -> None:
        sheet.add_hyperlink("B2", "https://example.com")
        sheet.remove_hyperlink("B2")
        assert b"<hyperlinks" not in sheet.document.to_bytes()

    def test_removing_nothing(self, sheet: Worksheet) -> None:
        assert sheet.remove_hyperlink("B2") == 0

    def test_the_cells_value_is_untouched(self, sheet: Worksheet) -> None:
        """A hyperlink decorates whatever is already there."""
        before = sheet["A2"].value
        sheet.add_hyperlink("A2", "https://example.com")
        assert sheet["A2"].value == before

    def test_it_lands_in_schema_order(self, sheet: Worksheet) -> None:
        sheet.add_hyperlink("B2", "https://example.com")
        rendered = sheet.document.to_bytes().decode()
        assert rendered.index("<mergeCells") < rendered.index("<hyperlinks")
        assert rendered.index("<hyperlinks") < rendered.index("<pageMargins")

    def test_it_survives_a_save(self, sheet: Worksheet) -> None:
        sheet.add_hyperlink("B2", "https://example.com/a", tooltip="go")
        reopened = Workbook.from_bytes(sheet.workbook.to_bytes())["Data"]
        link = reopened.hyperlinks[0]
        assert link.target == "https://example.com/a"
        assert link.tooltip == "go"

    def test_it_moves_when_rows_are_inserted(self, sheet: Worksheet) -> None:
        sheet.add_hyperlink("B5", "https://example.com")
        sheet.insert_rows(3, 2)
        assert sheet.hyperlinks[0].ref.a1 == "B7"


class TestOutlineGrouping:
    @pytest.fixture()
    def sheet(self, live_sample_xlsx: Path) -> Worksheet:
        return Workbook.open(live_sample_xlsx)["Data"]

    def test_nothing_is_grouped_to_begin_with(self, sheet: Worksheet) -> None:
        assert sheet.row_outline_level(5) == 0
        assert sheet.column_outline_level(2) == 0

    def test_grouping_rows(self, sheet: Worksheet) -> None:
        sheet.group_rows(5, 8)
        assert [sheet.row_outline_level(n) for n in range(4, 10)] == [0, 1, 1, 1, 1, 0]

    def test_grouping_twice_nests(self, sheet: Worksheet) -> None:
        sheet.group_rows(5, 8)
        sheet.group_rows(6, 7)
        assert sheet.row_outline_level(6) == 2
        assert sheet.row_outline_level(5) == 1

    def test_a_collapsed_group_hides_its_rows(self, sheet: Worksheet) -> None:
        sheet.group_rows(5, 8, collapsed=True)
        assert sheet.row_hidden(5) is True
        assert sheet.row_hidden(9) is False, "the summary row stays"

    def test_ungrouping(self, sheet: Worksheet) -> None:
        sheet.group_rows(5, 8, collapsed=True)
        sheet.ungroup_rows(5, 8)
        assert sheet.row_outline_level(5) == 0
        assert sheet.row_hidden(5) is False

    def test_ungrouping_one_level_of_two(self, sheet: Worksheet) -> None:
        sheet.group_rows(5, 8)
        sheet.group_rows(6, 7)
        sheet.ungroup_rows(6, 7)
        assert sheet.row_outline_level(6) == 1

    def test_grouping_columns(self, sheet: Worksheet) -> None:
        sheet.group_columns(2, 3)
        assert sheet.column_outline_level(2) == 1
        assert sheet.column_outline_level(4) == 0

    def test_a_collapsed_column_group(self, sheet: Worksheet) -> None:
        sheet.group_columns(2, 3, collapsed=True)
        assert sheet.column_hidden(2) is True

    def test_ungrouping_columns(self, sheet: Worksheet) -> None:
        sheet.group_columns(2, 3)
        sheet.ungroup_columns(2, 3)
        assert sheet.column_outline_level(2) == 0

    def test_a_backwards_range(self, sheet: Worksheet) -> None:
        with pytest.raises(ValueError, match="not a range to group"):
            sheet.group_rows(8, 5)

    def test_summaries_are_below_and_right_by_default(self, sheet: Worksheet) -> None:
        assert sheet.summary_below is True
        assert sheet.summary_right is True

    def test_putting_summaries_above(self, sheet: Worksheet) -> None:
        sheet.summary_below = False
        assert sheet.summary_below is False
        assert 'summaryBelow="0"' in sheet.document.to_bytes().decode()

    def test_putting_them_back(self, sheet: Worksheet) -> None:
        """The default is below, so saying so explicitly is noise."""
        sheet.summary_below = False
        sheet.summary_below = True
        assert b"summaryBelow" not in sheet.document.to_bytes()

    def test_it_survives_a_save(self, sheet: Worksheet) -> None:
        sheet.group_rows(5, 8, collapsed=True)
        sheet.group_columns(2, 3)
        reopened = Workbook.from_bytes(sheet.workbook.to_bytes())["Data"]
        assert reopened.row_outline_level(5) == 1
        assert reopened.column_outline_level(2) == 1


class TestOutlineMarkup:
    """What Excel writes alongside a group, measured by grouping, folding
    and ungrouping in Excel and reading the file back."""

    @pytest.fixture()
    def sheet(self, live_sample_xlsx: Path) -> Worksheet:
        return Workbook.open(live_sample_xlsx)["Data"]

    def test_the_sheet_records_how_deep_its_outline_goes(self, sheet: Worksheet) -> None:
        sheet.group_rows(5, 8)
        sheet.group_rows(6, 7)
        sheet.group_columns(2, 3)
        head = sheet.document.root.require("sheetFormatPr")
        assert (head.get("outlineLevelRow"), head.get("outlineLevelCol")) == ("2", "1")
        sheet.ungroup_rows(6, 7)
        assert head.get("outlineLevelRow") == "1"
        sheet.ungroup_rows(5, 8)
        sheet.ungroup_columns(2, 3)
        assert (head.get("outlineLevelRow"), head.get("outlineLevelCol")) == (None, None)

    def test_a_folded_group_marks_its_summary_row(self, sheet: Worksheet) -> None:
        sheet.group_rows(5, 8, collapsed=True)
        rows = sheet.rows_by_number()
        assert rows[9].get("collapsed") == "1"
        assert all(rows[number].get("collapsed") is None for number in range(5, 9))

    def test_with_summaries_above_the_mark_goes_above(self, sheet: Worksheet) -> None:
        sheet.summary_below = False
        sheet.group_rows(5, 8, collapsed=True)
        rows = sheet.rows_by_number()
        assert rows[4].get("collapsed") == "1"
        assert rows[9].get("collapsed") is None

    def test_a_folded_inner_group(self, sheet: Worksheet) -> None:
        """As Excel wrote it: the inner rows hidden, and the inner summary,
        itself inside the outer group, marked."""
        sheet.group_rows(5, 9)
        sheet.group_rows(6, 7, collapsed=True)
        assert [sheet.row_hidden(number) for number in range(5, 10)] == [False, True, True, False, False]
        summary = sheet.rows_by_number()[8]
        assert (summary.get("collapsed"), summary.get("outlineLevel")) == ("1", "1")

    def test_ungrouping_drops_the_summary_mark(self, sheet: Worksheet) -> None:
        sheet.group_rows(5, 8, collapsed=True)
        sheet.ungroup_rows(5, 8)
        assert sheet.rows_by_number()[9].get("collapsed") is None

    def test_a_folded_column_group_marks_its_summary_column(self, sheet: Worksheet) -> None:
        sheet.group_columns(2, 3, collapsed=True)
        entries = list(sheet.document.root.require("cols").children_named("col"))
        assert [(entry.get("min"), entry.get("collapsed")) for entry in entries[:3]] == [
            ("2", None), ("3", None), ("4", "1")
        ]

    def test_grouped_columns_keep_the_standard_width(self, sheet: Worksheet) -> None:
        """Measured: a ``<col>`` with no width is a column of width 0, which
        Excel shows hidden whatever the outline says."""
        sheet.group_columns(2, 3, collapsed=True)
        for entry in sheet.document.root.require("cols").children_named("col"):
            assert entry.get("width"), entry.to_xml()
        assert sheet.column_width(2) == 9.140625, "Aptos Narrow 11's, the fixture's font"

    def test_ungrouping_columns_leaves_the_block_as_it_was(self, sheet: Worksheet) -> None:
        before = sheet.document.root.require("cols").to_xml()
        sheet.group_columns(2, 3, collapsed=True)
        sheet.ungroup_columns(2, 3)
        assert sheet.document.root.require("cols").to_xml() == before

    def test_restoring_a_width_keeps_a_width(self, sheet: Worksheet) -> None:
        """A column still hidden or grouped after its width is reset keeps
        an entry, and the entry keeps the standard width."""
        sheet.group_columns(2, 2)
        sheet.set_column_width(2, 20)
        sheet.set_column_width(2, None)
        assert sheet.column_width(2) == 9.140625
        assert sheet.column_outline_level(2) == 1

    def test_seven_levels_deep_and_no_further(self, sheet: Worksheet) -> None:
        """Measured: Excel's eighth Group fails."""
        for _ in range(7):
            sheet.group_rows(5, 8)
        with pytest.raises(ValueError, match="7 levels"):
            sheet.group_rows(4, 9)
        assert sheet.row_outline_level(5) == 7
        assert sheet.row_outline_level(4) == 0, "nothing changed"


class TestReadingWhatExcelWrote:
    def test_hyperlinks(self, live_links_xlsx: Path) -> None:
        sheet = Workbook.open(live_links_xlsx)["L"]
        by_ref = {link.ref.a1: link for link in sheet.hyperlinks}

        assert by_ref["E2"].target == "https://example.com/a"
        assert by_ref["E2"].tooltip == "go there"
        # An internal link: location and no relationship.
        assert by_ref["E3"].target is None
        assert by_ref["E3"].location == "L!A1"
        assert by_ref["E4"].target == "mailto:a@b.c"
        # Both, for a URL with a fragment.
        assert by_ref["E5"].target == "https://example.com/b"
        assert by_ref["E5"].location == "frag"
        # One entry over a whole range.
        assert by_ref["G2:G4"].target == "https://example.com/c"

    def test_grouping(self, live_links_xlsx: Path) -> None:
        sheet = Workbook.open(live_links_xlsx)["L"]
        assert sheet.row_outline_level(5) == 1
        assert sheet.row_outline_level(2) == 0
        assert sheet.column_outline_level(2) == 1
        assert sheet.summary_below is False, "the probe put summaries above"
        assert sheet.summary_right is False
