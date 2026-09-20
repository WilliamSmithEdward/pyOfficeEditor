"""Merged ranges.

A merge is two things at once: an entry in ``<mergeCells>``, and a block of
cells where only the top left holds a value. Excel writes the covered cells
anyway, empty and carrying the anchor's style, because that is how a border
renders across a merge.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from pyofficeeditor.excel import CellRef, RangeRef, Workbook


@pytest.fixture()
def book(live_sample_xlsx: Path) -> Workbook:
    return Workbook.open(live_sample_xlsx)


class TestRangeIntersection:
    @pytest.mark.parametrize(
        ("left", "right", "expected"),
        [
            ("A1:C3", "B2:D4", True),
            ("A1:C3", "C3:E5", True),
            ("A1:C3", "D1:F3", False),
            ("A1:C3", "A4:C6", False),
            ("A1:C3", "A1:C3", True),
            ("A1:C3", "B2", True),
            ("A1:C3", "D4", False),
            ("B2:B2", "B2", True),
        ],
    )
    def test_intersects(self, left: str, right: str, expected: bool) -> None:
        assert RangeRef.parse(left).intersects(RangeRef.parse(right)) is expected
        assert RangeRef.parse(right).intersects(RangeRef.parse(left)) is expected, "symmetric"

    @pytest.mark.parametrize(
        ("outer", "inner", "expected"),
        [
            ("A1:C3", "B2", True),
            ("A1:C3", "A1:C3", True),
            ("A1:C3", "B2:C3", True),
            ("A1:C3", "B2:D4", False),
            ("B2:C3", "A1:C3", False),
        ],
    )
    def test_contains(self, outer: str, inner: str, expected: bool) -> None:
        assert RangeRef.parse(outer).contains(RangeRef.parse(inner)) is expected


class TestReadingRealMerges:
    def test_the_fixtures_merge(self, book: Workbook) -> None:
        assert [block.a1 for block in book["Data"].merged_ranges] == ["A11:C11"]

    def test_the_anchor_holds_the_value(self, book: Workbook) -> None:
        cell = book["Data"]["A11"]
        assert cell.is_merged is True
        assert cell.is_merge_anchor is True
        assert cell.value == "Merged heading"

    def test_a_covered_cell_reads_as_empty(self, book: Workbook) -> None:
        """Which is why asking whether it is merged matters: the value is on
        the anchor, so an empty read is not an empty block."""
        cell = book["Data"]["B11"]
        assert cell.is_merged is True
        assert cell.is_merge_anchor is False
        assert cell.value is None
        assert cell.merged_range is not None
        assert cell.merged_range.a1 == "A11:C11"

    def test_a_cell_outside_the_merge(self, book: Workbook) -> None:
        assert book["Data"]["D11"].is_merged is False
        assert book["Data"]["D11"].merged_range is None

    def test_merged_range_at(self, book: Workbook) -> None:
        sheet = book["Data"]
        assert sheet.merged_range_at(CellRef.parse("C11")) is not None
        assert sheet.merged_range_at(CellRef.parse("A1")) is None

    def test_a_sheet_with_no_merges(self, book: Workbook) -> None:
        assert book["Notes"].merged_ranges == []
        assert book["Notes"]["A1"].is_merged is False


class TestMerging:
    def test_a_merge_is_recorded(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet["A20"].value = "wide heading"
        assert sheet.merge("A20:C20").a1 == "A20:C20"
        assert [b.a1 for b in sheet.merged_ranges] == ["A11:C11", "A20:C20"]

    def test_it_survives_a_save(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet["A20"].value = "wide heading"
        sheet.merge("A20:C20")
        reopened = Workbook.from_bytes(book.to_bytes())
        assert [b.a1 for b in reopened["Data"].merged_ranges] == ["A11:C11", "A20:C20"]
        assert reopened["Data"]["A20"].value == "wide heading"

    def test_the_covered_values_are_discarded(self, book: Workbook) -> None:
        """Excel does the same: a covered cell is not displayed, so data left
        in one would be invisible and misleading."""
        sheet = book["Data"]
        sheet["A20"].value = "kept"
        sheet["B20"].value = "discarded"
        sheet["C20"].value = 99
        sheet.merge("A20:C20")
        assert sheet["A20"].value == "kept"
        assert sheet["B20"].value is None
        assert sheet["C20"].value is None

    def test_a_covered_formula_goes_too(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet["B20"].formula = "=1+1"
        sheet.merge("A20:C20")
        assert sheet["B20"].formula is None

    def test_the_covered_cells_take_the_anchors_style(self, book: Workbook) -> None:
        """Which is how a border renders across a merge."""
        sheet = book["Data"]
        sheet["A20"].value = "heading"
        sheet["A20"].fill = "FFFF00"
        sheet.merge("A20:C20")
        anchor = sheet["A20"].style_index
        assert anchor is not None
        assert sheet["B20"].style_index == anchor
        assert sheet["C20"].style_index == anchor

    def test_the_covered_cells_are_still_written(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet["A20"].value = "heading"
        sheet.merge("A20:C20")
        raw = book.package.read(sheet.part_name).decode()
        row = re.search(r'<row r="20".*?</row>', raw)
        assert row is not None
        assert 'r="B20"' in row.group(0)
        assert 'r="C20"' in row.group(0)

    def test_a_one_cell_merge_is_refused(self, book: Workbook) -> None:
        with pytest.raises(ValueError, match="nothing to merge it with"):
            book["Data"].merge("A20")

    def test_an_overlapping_merge_is_refused(self, book: Workbook) -> None:
        """Excel repairs a worksheet whose merges overlap rather than
        rendering it."""
        with pytest.raises(ValueError, match="overlaps the merged range A11:C11"):
            book["Data"].merge("B11:D11")

    def test_a_merge_containing_an_existing_one_is_refused(self, book: Workbook) -> None:
        with pytest.raises(ValueError, match="overlaps"):
            book["Data"].merge("A10:D12")

    def test_a_refused_merge_changes_nothing(self, book: Workbook) -> None:
        sheet = book["Data"]
        with pytest.raises(ValueError):
            sheet.merge("B11:D11")
        assert [b.a1 for b in sheet.merged_ranges] == ["A11:C11"]

    def test_a_reversed_reference_normalizes(self, book: Workbook) -> None:
        assert book["Data"].merge("C20:A20").a1 == "A20:C20"

    def test_a_range_object_is_accepted(self, book: Workbook) -> None:
        assert book["Data"].merge(RangeRef.parse("A20:C20")).a1 == "A20:C20"

    def test_a_vertical_merge(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet["H1"].value = "tall"
        sheet.merge("H1:H4")
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"]["H1"].value == "tall"
        assert reopened["Data"]["H3"].merged_range is not None
        assert reopened["Data"]["H3"].merged_range.a1 == "H1:H4"

    def test_the_count_is_maintained(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.merge("A20:C20")
        sheet.merge("A21:C21")
        raw = book.package.read(sheet.part_name).decode()
        container = re.search(r"<mergeCells[^>]*>", raw)
        assert container is not None
        assert 'count="3"' in container.group(0)

    def test_the_element_lands_in_schema_order(self, book: Workbook) -> None:
        """``mergeCells`` sits between ``customSheetViews`` and
        ``phoneticPr``, and Excel refuses a part that breaks the sequence."""
        from pyofficeeditor.excel._schema import WORKSHEET_CHILD_ORDER

        sheet = book["Notes"]
        sheet["A1"].value = "x"
        sheet.merge("A1:B1")
        names = [
            child.name
            for child in sheet.document.root.elements()
            if child.name in WORKSHEET_CHILD_ORDER
        ]
        positions = [WORKSHEET_CHILD_ORDER.index(n) for n in names]
        assert positions == sorted(positions), names


class TestUnmerging:
    def test_by_its_own_reference(self, book: Workbook) -> None:
        sheet = book["Data"]
        assert sheet.unmerge("A11:C11").a1 == "A11:C11"
        assert sheet.merged_ranges == []

    def test_by_a_cell_inside_it(self, book: Workbook) -> None:
        """The way Excel's own command works on a selection."""
        sheet = book["Data"]
        assert sheet.unmerge("B11").a1 == "A11:C11"
        assert sheet.merged_ranges == []

    def test_by_the_anchor(self, book: Workbook) -> None:
        assert book["Data"].unmerge("A11").a1 == "A11:C11"

    def test_it_survives_a_save(self, book: Workbook) -> None:
        book["Data"].unmerge("A11:C11")
        assert Workbook.from_bytes(book.to_bytes())["Data"].merged_ranges == []

    def test_the_anchors_value_stays(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.unmerge("A11:C11")
        assert sheet["A11"].value == "Merged heading"

    def test_the_element_goes_when_the_last_merge_does(self, book: Workbook) -> None:
        """Excel omits ``mergeCells`` rather than writing ``count="0"``."""
        sheet = book["Data"]
        sheet.unmerge("A11:C11")
        assert b"mergeCells" not in sheet.document.to_bytes()
        assert b"mergeCells" not in book.package.read(sheet.part_name)

    def test_unmerging_something_that_is_not_merged(self, book: Workbook) -> None:
        with pytest.raises(ValueError, match="not a merged range"):
            book["Data"].unmerge("E5")

    def test_unmerging_a_partial_reference(self, book: Workbook) -> None:
        """A block that overlaps a merge without being it, and is not a single
        cell, matches nothing."""
        with pytest.raises(ValueError, match="not a merged range"):
            book["Data"].unmerge("A11:B11")

    def test_on_a_sheet_with_no_merges(self, book: Workbook) -> None:
        with pytest.raises(ValueError, match="not a merged range"):
            book["Notes"].unmerge("A1:B1")

    def test_merging_again_after_unmerging(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet.unmerge("A11:C11")
        sheet.merge("A11:D11")
        reopened = Workbook.from_bytes(book.to_bytes())
        assert [b.a1 for b in reopened["Data"].merged_ranges] == ["A11:D11"]


class TestNoOpDiscipline:
    def test_reading_merges_is_not_a_change(self, book: Workbook, live_sample_xlsx: Path) -> None:
        for sheet in book:
            _ = sheet.merged_ranges
            for row in sheet.rows():
                for cell in row:
                    _ = cell.is_merged, cell.merged_range, cell.is_merge_anchor
        assert book.is_modified is False
        assert book.to_bytes() == live_sample_xlsx.read_bytes()
