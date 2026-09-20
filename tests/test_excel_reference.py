"""A1 notation.

Pure arithmetic, so it is tested exhaustively at the boundaries: bijective
base-26 carries, Excel's real limits, and the absolute-marker behavior that
shared-formula translation depends on.
"""

from __future__ import annotations

import pytest

from pyofficeeditor.excel._reference import (
    MAX_COLUMN,
    MAX_ROW,
    CellRef,
    RangeRef,
    column_index,
    column_letter,
)


class TestColumnLetters:
    @pytest.mark.parametrize(
        ("index", "letters"),
        [
            (1, "A"),
            (2, "B"),
            (25, "Y"),
            (26, "Z"),
            (27, "AA"),
            (28, "AB"),
            (51, "AY"),
            (52, "AZ"),
            (53, "BA"),
            (78, "BZ"),
            (79, "CA"),
            (702, "ZZ"),
            (703, "AAA"),
            (704, "AAB"),
            (16384, "XFD"),
        ],
    )
    def test_the_carries_that_bijective_base_26_gets_wrong(self, index: int, letters: str) -> None:
        assert column_letter(index) == letters
        assert column_index(letters) == index

    def test_every_column_round_trips(self) -> None:
        for index in range(1, MAX_COLUMN + 1):
            assert column_index(column_letter(index)) == index

    def test_case_does_not_matter_on_the_way_in(self) -> None:
        assert column_index("aa") == column_index("AA") == column_index("aA")

    @pytest.mark.parametrize("index", [0, -1, MAX_COLUMN + 1, 999999])
    def test_an_index_outside_the_sheet_is_refused(self, index: int) -> None:
        with pytest.raises(ValueError, match="column"):
            column_letter(index)

    @pytest.mark.parametrize("letters", ["", "1", "A1", "-", "ABCD", "A B", "ZZZ", "XFE"])
    def test_a_name_that_is_not_a_column_is_refused(self, letters: str) -> None:
        with pytest.raises(ValueError):
            column_index(letters)

    def test_the_last_column_and_the_one_past_it(self) -> None:
        assert column_index("XFD") == MAX_COLUMN
        with pytest.raises(ValueError, match="past XFD"):
            column_index("XFE")


class TestCellRef:
    @pytest.mark.parametrize(
        ("text", "row", "column", "absolute_row", "absolute_column"),
        [
            ("A1", 1, 1, False, False),
            ("$A1", 1, 1, False, True),
            ("A$1", 1, 1, True, False),
            ("$A$1", 1, 1, True, True),
            ("B2", 2, 2, False, False),
            ("XFD1048576", MAX_ROW, MAX_COLUMN, False, False),
            ("aa10", 10, 27, False, False),
            ("  C3  ", 3, 3, False, False),
        ],
    )
    def test_parse(
        self, text: str, row: int, column: int, absolute_row: bool, absolute_column: bool
    ) -> None:
        reference = CellRef.parse(text)
        assert (reference.row, reference.column) == (row, column)
        assert reference.absolute_row is absolute_row
        assert reference.absolute_column is absolute_column

    @pytest.mark.parametrize("text", ["A1", "$A1", "A$1", "$A$1", "XFD1048576", "B12"])
    def test_a1_round_trips_with_its_markers(self, text: str) -> None:
        assert CellRef.parse(text).a1 == text

    @pytest.mark.parametrize(
        "text", ["", "A", "1", "A0", "1A", "$", "$1", "A1:B2", "A 1", "ZZZ1", "A1048577", "XFE1", "-A1"]
    )
    def test_malformed_references_are_refused(self, text: str) -> None:
        with pytest.raises(ValueError):
            CellRef.parse(text)

    def test_the_row_and_column_limits_bind_on_construction(self) -> None:
        CellRef(MAX_ROW, MAX_COLUMN)
        with pytest.raises(ValueError, match="row"):
            CellRef(MAX_ROW + 1, 1)
        with pytest.raises(ValueError, match="column"):
            CellRef(1, MAX_COLUMN + 1)
        with pytest.raises(ValueError, match="row"):
            CellRef(0, 1)

    def test_markers_do_not_change_position_but_do_change_identity(self) -> None:
        plain = CellRef.parse("A1")
        pinned = CellRef.parse("$A$1")
        assert plain.sort_key == pinned.sort_key
        assert plain != pinned, "the two are written differently and must stay so"
        assert plain == pinned.relative

    def test_offset_moves_everything_and_keeps_the_markers(self) -> None:
        moved = CellRef.parse("$B$2").offset(rows=3, columns=1)
        assert moved.a1 == "$C$5"

    def test_reading_order(self) -> None:
        unsorted = [CellRef.parse(t) for t in ("B1", "A2", "A1", "C1", "B2")]
        assert [c.a1 for c in sorted(unsorted)] == ["A1", "B1", "C1", "A2", "B2"]

    def test_str_is_the_a1_form(self) -> None:
        assert str(CellRef.parse("$D$9")) == "$D$9"


class TestTranslation:
    """What a shared formula's followers need."""

    def test_a_relative_reference_moves(self) -> None:
        assert CellRef.parse("B2").translated(rows=1, columns=0).a1 == "B3"
        assert CellRef.parse("B2").translated(rows=0, columns=2).a1 == "D2"
        assert CellRef.parse("B2").translated(rows=3, columns=3).a1 == "E5"

    def test_an_absolute_reference_stays(self) -> None:
        assert CellRef.parse("$B$2").translated(rows=5, columns=5).a1 == "$B$2"

    def test_a_mixed_reference_moves_only_its_relative_half(self) -> None:
        assert CellRef.parse("$B2").translated(rows=1, columns=9).a1 == "$B3"
        assert CellRef.parse("B$2").translated(rows=9, columns=1).a1 == "C$2"

    def test_the_fixture_case(self) -> None:
        """Excel stored ``B2*C2`` once for D2:D5. D3's formula is the
        master's translated down one row."""
        assert CellRef.parse("B2").translated(rows=1, columns=0).a1 == "B3"
        assert CellRef.parse("C2").translated(rows=3, columns=0).a1 == "C5"

    def test_translating_off_the_sheet_is_refused(self) -> None:
        with pytest.raises(ValueError, match="row"):
            CellRef.parse("A1").translated(rows=-1, columns=0)


class TestRangeRef:
    def test_parse_a_block(self) -> None:
        block = RangeRef.parse("A1:C3")
        assert (block.top, block.left, block.bottom, block.right) == (1, 1, 3, 3)
        assert (block.height, block.width, block.size) == (3, 3, 9)

    def test_a_bare_cell_is_a_range_of_one(self) -> None:
        block = RangeRef.parse("A1")
        assert block.is_single_cell
        assert block.size == 1
        assert block.a1 == "A1", "Excel writes a one-cell dimension as A1, not A1:A1"

    def test_a1_round_trips(self) -> None:
        for text in ("A1", "A1:C3", "$A$1:$C$3", "B2:B2"):
            expected = "B2" if text == "B2:B2" else text
            assert RangeRef.parse(text).a1 == expected

    def test_a_reversed_block_normalizes(self) -> None:
        assert RangeRef.parse("C3:A1").normalized.a1 == "A1:C3"
        assert RangeRef.parse("C3:A1").size == 9, "the size is right before normalizing too"

    def test_containment(self) -> None:
        block = RangeRef.parse("B2:D4")
        assert CellRef.parse("C3") in block
        assert CellRef.parse("B2") in block
        assert CellRef.parse("D4") in block
        assert CellRef.parse("A1") not in block
        assert CellRef.parse("E4") not in block
        assert CellRef.parse("C5") not in block

    def test_containment_ignores_markers(self) -> None:
        assert CellRef.parse("$C$3") in RangeRef.parse("B2:D4")

    def test_cells_are_in_reading_order(self) -> None:
        assert [c.a1 for c in RangeRef.parse("A1:B2").cells()] == ["A1", "B1", "A2", "B2"]

    def test_rows_group_by_row(self) -> None:
        assert [[c.a1 for c in row] for row in RangeRef.parse("A1:B2").rows()] == [
            ["A1", "B1"],
            ["A2", "B2"],
        ]

    def test_bounding(self) -> None:
        cells = [CellRef.parse(t) for t in ("C1", "A3", "B2")]
        assert RangeRef.bounding(cells).a1 == "A1:C3"

    def test_bounding_one_cell(self) -> None:
        assert RangeRef.bounding([CellRef.parse("D4")]).a1 == "D4"

    def test_bounding_nothing_is_refused(self) -> None:
        with pytest.raises(ValueError, match="no bounding range"):
            RangeRef.bounding([])

    def test_expanded(self) -> None:
        assert RangeRef.parse("B2:C3").expanded(CellRef.parse("A5")).a1 == "A2:C5"
        assert RangeRef.parse("B2:C3").expanded(CellRef.parse("C3")).a1 == "B2:C3"

    def test_a_whole_sheet_range_does_not_overflow(self) -> None:
        whole = RangeRef.parse(f"A1:XFD{MAX_ROW}")
        assert whole.size == MAX_COLUMN * MAX_ROW
        assert CellRef(MAX_ROW, MAX_COLUMN) in whole

    @pytest.mark.parametrize("text", ["", ":", "A1:", ":B2", "A1:B2:C3", "A0:B1", "junk"])
    def test_malformed_ranges_are_refused(self, text: str) -> None:
        with pytest.raises(ValueError):
            RangeRef.parse(text)

    def test_str_is_the_a1_form(self) -> None:
        assert str(RangeRef.parse("A1:B9")) == "A1:B9"


class TestAgainstRealExcelOutput:
    def test_the_fixture_dimension_parses(self, live_sample_xlsx: object) -> None:
        """Whatever Excel put in ``dimension`` must be a range this module
        reads, since every sheet has one and the cell layer trusts it."""
        from pyofficeeditor.opc import OpcPackage

        package = OpcPackage.open(str(live_sample_xlsx))
        for name in package.part_names():
            if not name.startswith("xl/worksheets/sheet"):
                continue
            dimension = package.xml(name).root.child("dimension")
            assert dimension is not None, name
            reference = dimension.get("ref")
            assert reference is not None
            block = RangeRef.parse(reference)
            assert block.size >= 1, f"{name}: {reference}"

    def test_every_cell_reference_in_the_fixture_parses(self, live_sample_xlsx: object) -> None:
        from pyofficeeditor.opc import OpcPackage

        package = OpcPackage.open(str(live_sample_xlsx))
        seen = 0
        for name in package.part_names():
            if not name.startswith("xl/worksheets/sheet"):
                continue
            for cell in package.xml(name).root.require("sheetData").descendants("c"):
                reference = cell.get("r")
                assert reference is not None, f"{name}: a cell with no r attribute"
                CellRef.parse(reference)
                seen += 1
        assert seen > 30, "the sample fixture should have a useful number of cells"
