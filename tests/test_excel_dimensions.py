"""Column widths, row heights, hiding and frozen panes.

The two are stored nothing alike. A row's settings sit on its own element; a
column's sit in a separate block whose entries cover *ranges*, so setting one
column inside a span means splitting the span rather than editing it. Editing
it would resize every column it covers.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from pyofficeeditor._xml import XmlDocument
from pyofficeeditor.excel import CellRef, Workbook
from pyofficeeditor.excel._dimensions import Freeze, column_entry, isolate_column


@pytest.fixture()
def book(live_structures_xlsx: Path) -> Workbook:
    return Workbook.open(live_structures_xlsx)


def cols(text: bytes) -> XmlDocument:
    return XmlDocument.parse(text)


class TestSplittingAColumnSpan:
    def test_isolating_the_middle_makes_three_entries(self) -> None:
        document = cols(b'<cols><col min="1" max="5" width="10"/></cols>')
        isolate_column(document.root, 3).set("width", "99")
        assert document.to_bytes() == (
            b'<cols><col min="1" max="2" width="10"/>'
            b'<col min="3" max="3" width="99"/>'
            b'<col min="4" max="5" width="10"/></cols>'
        )

    def test_isolating_the_head_makes_two(self) -> None:
        document = cols(b'<cols><col min="1" max="5" width="10"/></cols>')
        isolate_column(document.root, 1).set("width", "99")
        assert document.to_bytes() == (
            b'<cols><col min="1" max="1" width="99"/><col min="2" max="5" width="10"/></cols>'
        )

    def test_isolating_the_tail_makes_two(self) -> None:
        document = cols(b'<cols><col min="1" max="5" width="10"/></cols>')
        isolate_column(document.root, 5).set("width", "99")
        assert document.to_bytes() == (
            b'<cols><col min="1" max="4" width="10"/><col min="5" max="5" width="99"/></cols>'
        )

    def test_every_attribute_is_carried_to_each_piece(self) -> None:
        """A column that was hidden and styled stays hidden and styled."""
        document = cols(
            b'<cols><col min="1" max="5" width="10" hidden="1" customWidth="1" '
            b'style="3" outlineLevel="2" bestFit="1" collapsed="1"/></cols>'
        )
        isolate_column(document.root, 3)
        raw = document.to_bytes().decode()
        assert raw.count('hidden="1"') == 3
        assert raw.count('style="3"') == 3
        assert raw.count('outlineLevel="2"') == 3
        assert raw.count('collapsed="1"') == 3

    def test_an_already_single_entry_is_returned_as_is(self) -> None:
        document = cols(b'<cols><col min="3" max="3" width="10"/></cols>')
        entry = isolate_column(document.root, 3)
        assert entry.get("width") == "10"
        assert len(list(document.root.children_named("col"))) == 1

    def test_a_column_with_no_entry_gets_one(self) -> None:
        document = cols(b'<cols><col min="5" max="5" width="10"/></cols>')
        isolate_column(document.root, 2).set("width", "7")
        assert document.to_bytes() == (
            b'<cols><col min="2" max="2" width="7"/><col min="5" max="5" width="10"/></cols>'
        )

    def test_entries_stay_in_ascending_order(self) -> None:
        document = cols(b"<cols></cols>")
        for column in (5, 1, 9, 3):
            isolate_column(document.root, column).set("width", "1")
        mins = [int(e.get("min") or 0) for e in document.root.children_named("col")]
        assert mins == sorted(mins) == [1, 3, 5, 9]

    def test_column_entry_finds_a_span(self) -> None:
        document = cols(b'<cols><col min="2" max="6" width="10"/></cols>')
        assert column_entry(document.root, 4) is not None
        assert column_entry(document.root, 1) is None
        assert column_entry(document.root, 7) is None

    @pytest.mark.parametrize("column", [0, -1, 16385])
    def test_a_column_outside_the_sheet_is_refused(self, column: int) -> None:
        document = cols(b"<cols></cols>")
        with pytest.raises(ValueError, match="column"):
            isolate_column(document.root, column)


class TestFreezeValue:
    @pytest.mark.parametrize(
        ("cell", "columns", "rows"),
        [("A1", 0, 0), ("B2", 1, 1), ("A2", 0, 1), ("B1", 1, 0), ("C3", 2, 2), ("D10", 3, 9)],
    )
    def test_at(self, cell: str, columns: int, rows: int) -> None:
        freeze = Freeze.at(CellRef.parse(cell))
        assert (freeze.columns, freeze.rows) == (columns, rows)
        assert freeze.top_left.a1 == cell

    def test_a1_is_not_frozen(self) -> None:
        assert Freeze.at(CellRef.parse("A1")).is_frozen is False
        assert Freeze().is_frozen is False
        assert Freeze(rows=1).is_frozen is True

    @pytest.mark.parametrize(
        ("cell", "expected"),
        [
            (
                "B2",
                b'<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>',
            ),
            ("A2", b'<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'),
            ("B1", b'<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>'),
            (
                "C3",
                b'<pane xSplit="2" ySplit="2" topLeftCell="C3" activePane="bottomRight" state="frozen"/>',
            ),
        ],
    )
    def test_it_writes_what_excel_writes(self, cell: str, expected: bytes) -> None:
        """Byte for byte, including activePane, which decides where the
        cursor lands and is the easiest part to get wrong."""
        assert Freeze.at(CellRef.parse(cell)).write().to_xml().encode() == expected

    def test_read_round_trips(self) -> None:
        for cell in ("B2", "A2", "B1", "C3"):
            freeze = Freeze.at(CellRef.parse(cell))
            assert Freeze.read(XmlDocument.parse(freeze.write().to_xml().encode()).root) == freeze

    def test_a_split_that_is_not_frozen_reads_as_unfrozen(self) -> None:
        """A split pane is a different feature; only a frozen one counts."""
        element = XmlDocument.parse(b'<pane xSplit="1000" ySplit="1000" state="split"/>').root
        assert Freeze.read(element).is_frozen is False

    def test_no_pane_reads_as_unfrozen(self) -> None:
        assert Freeze.read(None) == Freeze()


class TestReadingRealDimensions:
    def test_the_widths_excel_wrote(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        assert sheet.column_width(1) == 18.6328125
        assert sheet.column_width(3) == 10.08984375
        assert sheet.column_width(2) is None, "no entry means the default"

    def test_the_width_is_not_the_number_a_person_types(self, book: Workbook) -> None:
        """The fixture set ColumnWidth = 18 through VBA and Excel stored
        18.6328125. The offset depends on the font's maximum digit width,
        which is not in the file, so no conversion is offered."""
        assert book["Tabled"].column_width(1) != 18

    def test_the_row_heights_excel_wrote(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        assert sheet.row_height(1) == 24, "points, and exact"
        assert sheet.row_height(2) is None, "no ht means the default"
        assert sheet.default_row_height == 14.5

    def test_nothing_is_hidden_or_frozen_to_begin_with(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        assert sheet.column_hidden(1) is False
        assert sheet.row_hidden(1) is False
        assert sheet.freeze.is_frozen is False

    def test_reading_dimensions_is_not_a_change(
        self, book: Workbook, live_structures_xlsx: Path
    ) -> None:
        for sheet in book:
            for index in range(1, 8):
                _ = sheet.column_width(index), sheet.column_hidden(index)
                _ = sheet.row_height(index), sheet.row_hidden(index)
            _ = sheet.freeze, sheet.default_row_height
        assert book.is_modified is False
        assert book.to_bytes() == live_structures_xlsx.read_bytes()


class TestWritingDimensions:
    def test_a_width_survives_a_save(self, book: Workbook) -> None:
        book["Tabled"].set_column_width(2, 25.5)
        assert Workbook.from_bytes(book.to_bytes())["Tabled"].column_width(2) == 25.5

    def test_custom_width_accompanies_the_value(self, book: Workbook) -> None:
        """Without the flag Excel ignores the width, so the change would look
        like it never happened."""
        sheet = book["Tabled"]
        sheet.set_column_width(2, 25.5)
        raw = book.package.read(sheet.part_name).decode()
        entry = re.search(r'<col min="2"[^>]*/>', raw)
        assert entry is not None
        assert 'customWidth="1"' in entry.group(0)

    def test_clearing_a_width_restores_the_default(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        sheet.set_column_width(1, None)
        assert sheet.column_width(1) is None
        assert Workbook.from_bytes(book.to_bytes())["Tabled"].column_width(1) is None

    def test_setting_one_column_leaves_its_neighbours(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        sheet.set_column_width(2, 25.5)
        assert sheet.column_width(1) == 18.6328125
        assert sheet.column_width(3) == 10.08984375

    def test_a_negative_width_is_refused(self, book: Workbook) -> None:
        with pytest.raises(ValueError, match="cannot be negative"):
            book["Tabled"].set_column_width(2, -1)

    def test_hiding_and_unhiding_a_column(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        sheet.set_column_hidden(4, True)
        assert sheet.column_hidden(4) is True
        assert Workbook.from_bytes(book.to_bytes())["Tabled"].column_hidden(4) is True
        sheet.set_column_hidden(4, False)
        assert sheet.column_hidden(4) is False

    def test_hiding_a_column_keeps_its_width(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        sheet.set_column_hidden(1, True)
        assert sheet.column_width(1) == 18.6328125

    def test_a_row_height_survives_a_save(self, book: Workbook) -> None:
        book["Tabled"].set_row_height(3, 40)
        assert Workbook.from_bytes(book.to_bytes())["Tabled"].row_height(3) == 40

    def test_custom_height_accompanies_the_value(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        sheet.set_row_height(3, 40)
        raw = book.package.read(sheet.part_name).decode()
        row = re.search(r'<row r="3"[^>]*>', raw)
        assert row is not None
        assert 'customHeight="1"' in row.group(0)

    def test_a_fractional_height(self, book: Workbook) -> None:
        book["Tabled"].set_row_height(3, 30.5)
        assert Workbook.from_bytes(book.to_bytes())["Tabled"].row_height(3) == 30.5

    def test_clearing_a_height(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        sheet.set_row_height(1, None)
        assert sheet.row_height(1) is None

    def test_a_negative_height_is_refused(self, book: Workbook) -> None:
        with pytest.raises(ValueError, match="cannot be negative"):
            book["Tabled"].set_row_height(3, -1)

    def test_hiding_a_row(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        sheet.set_row_hidden(5, True)
        assert sheet.row_hidden(5) is True
        assert Workbook.from_bytes(book.to_bytes())["Tabled"].row_hidden(5) is True

    def test_setting_a_dimension_on_an_absent_row_creates_it(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        assert sheet.row_height(50) is None
        sheet.set_row_height(50, 33)
        assert Workbook.from_bytes(book.to_bytes())["Tabled"].row_height(50) == 33

    def test_rows_stay_in_order_after_creating_one(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        sheet.set_row_height(50, 33)
        sheet.set_row_height(20, 22)
        data = sheet.document.root.require("sheetData")
        numbers = [int(r.get("r") or 0) for r in data.children_named("row")]
        assert numbers == sorted(numbers)

    def test_an_empty_cols_block_is_removed(self, book: Workbook) -> None:
        """Excel omits it rather than writing an empty one."""
        sheet = book["Notes"] if "Notes" in book else book["Second"]
        sheet.set_column_width(2, 10)
        sheet.set_column_width(2, None)
        assert b"<cols" not in sheet.document.to_bytes()

    def test_the_cols_block_lands_in_schema_order(self, book: Workbook) -> None:
        from pyofficeeditor.excel._schema import WORKSHEET_CHILD_ORDER

        sheet = book["Second"]
        sheet.set_column_width(3, 12)
        names = [
            child.name
            for child in sheet.document.root.elements()
            if child.name in WORKSHEET_CHILD_ORDER
        ]
        positions = [WORKSHEET_CHILD_ORDER.index(n) for n in names]
        assert positions == sorted(positions), names


class TestFreezingPanes:
    def test_freezing_and_reading_back(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        assert sheet.freeze_panes("B2") == Freeze(columns=1, rows=1)
        assert sheet.freeze == Freeze(columns=1, rows=1)
        assert Workbook.from_bytes(book.to_bytes())["Tabled"].freeze == Freeze(columns=1, rows=1)

    @pytest.mark.parametrize(
        ("cell", "expected"), [("A2", Freeze(rows=1)), ("B1", Freeze(columns=1)), ("C3", Freeze(2, 2))]
    )
    def test_the_three_shapes(self, book: Workbook, cell: str, expected: Freeze) -> None:
        book["Tabled"].freeze_panes(cell)
        assert Workbook.from_bytes(book.to_bytes())["Tabled"].freeze == expected

    def test_freezing_at_a1_is_unfreezing(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        sheet.freeze_panes("B2")
        sheet.freeze_panes("A1")
        assert sheet.freeze.is_frozen is False
        assert b"<pane" not in sheet.document.to_bytes()

    def test_unfreezing_with_none(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        sheet.freeze_panes("B2")
        sheet.freeze_panes(None)
        assert sheet.freeze.is_frozen is False

    def test_refreezing_replaces_rather_than_stacking(self, book: Workbook) -> None:
        sheet = book["Tabled"]
        sheet.freeze_panes("B2")
        sheet.freeze_panes("D5")
        assert sheet.document.to_bytes().count(b"<pane") == 1
        assert sheet.freeze == Freeze(columns=3, rows=4)

    def test_the_pane_is_the_first_child_of_the_view(self, book: Workbook) -> None:
        """CT_SheetView is a sequence and pane leads it."""
        sheet = book["Tabled"]
        sheet.freeze_panes("B2")
        view = sheet.document.root.require("sheetViews").require("sheetView")
        assert next(iter(view.elements())).name == "pane"

    def test_a_cell_reference_object_is_accepted(self, book: Workbook) -> None:
        assert book["Tabled"].freeze_panes(CellRef.parse("B2")) == Freeze(columns=1, rows=1)

    def test_freezing_a_sheet_with_no_view(self, book: Workbook) -> None:
        sheet = book.add_sheet("Fresh")
        sheet.freeze_panes("B2")
        assert Workbook.from_bytes(book.to_bytes())["Fresh"].freeze == Freeze(columns=1, rows=1)
