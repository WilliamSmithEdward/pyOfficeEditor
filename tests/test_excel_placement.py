"""Drawn objects under inserted and deleted rows and columns.

Excel gives every shape, picture, chart, control and embedded object one of
three placements, and each takes an edit differently; a note's box follows
its cell. :mod:`pyofficeeditor.excel._placement` holds the rules, each
measured against Excel's own edits. These tests pin them without Office;
the live gate is what compares them with Excel.
"""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

import pytest

from pyofficeeditor._xml import Element
from pyofficeeditor.excel import Workbook, Worksheet
from pyofficeeditor.excel._formulas import Deletion, Shift
from pyofficeeditor.excel._placement import (
    AxisEdit,
    Corner,
    anchor_placement,
    move_box,
    move_vml,
    record_placement,
    swallows,
    vml_placement,
    walk,
    with_cells,
)
from pyofficeeditor.excel._shapes import EMU_PER_POINT, SheetGrid, anchors_in

#: The relationship Excel names an ActiveX control's part by.
RT_ACTIVEX = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control"


def copy_of(source: Path, tmp_path: Path) -> Workbook:
    target = tmp_path / source.name
    shutil.copy(source, target)
    return Workbook.open(target)


@pytest.fixture
def shapes(tmp_path: Path, live_shapes_xlsm: Path) -> Workbook:
    """Excel's shapes: rectangles, a group and a button, which move and size
    with their cells, over rows 3 and 4 that Excel made 30 points tall."""
    return copy_of(live_shapes_xlsm, tmp_path)


@pytest.fixture
def pictures(tmp_path: Path, live_pictures_xlsx: Path) -> Workbook:
    """Excel's pictures, which it places to move without sizing."""
    return copy_of(live_pictures_xlsx, tmp_path)


@pytest.fixture
def charts(tmp_path: Path, live_charts_xlsx: Path) -> Workbook:
    return copy_of(live_charts_xlsx, tmp_path)


def anchor_of(sheet: Worksheet, name: str) -> Element:
    """The drawing anchor a listed shape sits in: the first copy, for one
    Excel wraps in ``mc:AlternateContent``."""
    located = getattr(sheet, "_located")(name)
    assert located is not None, name
    return anchors_in(located[1])[0]


def chart_part_of(sheet: Worksheet, name: str) -> str:
    """The part a listed chart's frame points at."""
    located = getattr(sheet, "_located")(name)
    assert located is not None, name
    part, node = located[0], located[1]
    chart = next(node.descendants("chart"))
    return sheet.workbook.package.relationships(part).by_id(chart.get("r:id") or "").target_part


def corner(anchor: Element, end: str) -> tuple[int, int, int, int]:
    """An anchor end as column, column offset, row, row offset."""
    node = anchor.require(end)
    return tuple(int(node.require(name).text) for name in ("col", "colOff", "row", "rowOff"))  # type: ignore[return-value]


def vml_of(sheet: Worksheet) -> str:
    parts: list[str] = getattr(sheet, "_legacy_vml_parts")()
    return sheet.workbook.package.read(parts[0]).decode("utf-8")


def button_anchor(sheet: Worksheet) -> list[int]:
    """The eight numbers of the sheet's one button's VML anchor."""
    found = re.search(
        r'<x:ClientData ObjectType="Button">.*?<x:Anchor>(.*?)</x:Anchor>', vml_of(sheet), re.DOTALL
    )
    assert found is not None
    return [int(number) for number in re.findall(r"-?\d+", found.group(1))]


def note_anchor(sheet: Worksheet) -> list[int]:
    """The eight numbers of the sheet's one note's box."""
    found = re.search(
        r'<x:ClientData ObjectType="Note">.*?<x:Anchor>(.*?)</x:Anchor>', vml_of(sheet), re.DOTALL
    )
    assert found is not None
    return [int(number) for number in re.findall(r"-?\d+", found.group(1))]


def uniform(rows: float = 15.0) -> SheetGrid:
    return SheetGrid(default_row=rows)


def rows_edit(at: int, count: int, *, deleting: bool = False) -> AxisEdit:
    """An edit over rows all 15 points tall, before and after."""
    shift = None if deleting else Shift.rows(at, count)
    deletion = Deletion.rows(at, count) if deleting else None
    edit = AxisEdit.of(shift, deletion, uniform(), uniform())
    assert edit is not None
    return edit


ROW = 15 * EMU_PER_POINT


class TestTheRules:
    def test_an_edge_moves_with_its_cell(self) -> None:
        edit = rows_edit(3, 2)
        assert with_cells(Corner(1, 100), edit).corner == Corner(1, 100)
        moved = with_cells(Corner(2, 100), edit)
        assert (moved.corner, moved.distance) == (Corner(4, 100), 2 * ROW)

    def test_a_far_edge_on_the_line_stays(self) -> None:
        """Measured: a shape whose bottom sits exactly on the top of the
        row inserted does not grow; its top would move down."""
        edit = rows_edit(3, 2)
        assert with_cells(Corner(2, 0), edit, far=True).corner == Corner(2, 0)
        assert with_cells(Corner(2, 0), edit).corner == Corner(4, 0)
        assert with_cells(Corner(2, 1), edit, far=True).corner == Corner(4, 1)

    def test_an_edge_whose_row_goes_lands_on_the_boundary(self) -> None:
        edit = rows_edit(3, 2, deleting=True)
        moved = with_cells(Corner(3, 500), edit)
        assert (moved.corner, moved.distance, moved.collapsed) == (Corner(2, 0), -(ROW + 500), True)
        assert with_cells(Corner(5, 7), edit).corner == Corner(3, 7)

    def test_a_move_only_box_keeps_its_size(self) -> None:
        box = move_box(Corner(1, 100), Corner(4, 200), "moveOnly", rows_edit(3, 1))
        assert box.first.corner == Corner(1, 100)
        assert box.last is not None and box.last.corner == Corner(4, 200)
        assert box.last.distance == box.first.distance == 0

    def test_a_free_box_stays(self) -> None:
        taller = AxisEdit.of(Shift.rows(1, 1), None, uniform(), SheetGrid(default_row=15.0, row_heights={0: 30.0}))
        assert taller is not None
        box = move_box(Corner(2, 0), Corner(3, 0), "free", taller)
        # 30 points down is the whole of the new 30-point row.
        assert box.first.corner == Corner(1, 0)
        assert box.last is not None and box.last.corner == Corner(2, 0)
        assert box.first.distance == box.last.distance == 0

    def test_a_move_and_size_box_its_rows_all_go_is_swallowed(self) -> None:
        edit = rows_edit(3, 3, deleting=True)
        assert move_box(Corner(2, 0), Corner(4, 10), "moveAndSize", edit).swallowed
        assert move_box(Corner(2, 0), Corner(5, 0), "moveAndSize", edit).swallowed
        assert not move_box(Corner(2, 0), Corner(5, 1), "moveAndSize", edit).swallowed
        assert not move_box(Corner(2, 0), Corner(4, 10), "moveOnly", edit).swallowed
        assert swallows(Corner(2, 0), Corner(5, 0), 2, 5)
        assert not swallows(Corner(1, 9), Corner(3, 0), 2, 5)

    def test_walking_skips_a_row_of_no_height_and_lands_on_a_line(self) -> None:
        edit = AxisEdit.of(
            Shift.rows(1, 1), None, uniform(), SheetGrid(default_row=15.0, row_heights={1: 0.0})
        )
        assert edit is not None
        assert walk(edit, Corner(0, 0), ROW) == Corner(2, 0)
        assert walk(edit, Corner(2, 0), -1) == Corner(0, ROW - 1)

    def test_placements_as_each_part_records_them(self) -> None:
        anchor = Element.create("xdr:twoCellAnchor")
        assert anchor_placement(anchor) == "moveAndSize"
        anchor.set("editAs", "oneCell")
        assert anchor_placement(anchor) == "moveOnly"
        anchor.set("editAs", "absolute")
        assert anchor_placement(anchor) == "free"
        assert anchor_placement(Element.create("xdr:oneCellAnchor")) == "moveOnly"
        assert record_placement(Element.create("anchor", {"moveWithCells": "1", "sizeWithCells": "1"})) == "moveAndSize"
        assert record_placement(Element.create("anchor", {"moveWithCells": "1"})) == "moveOnly"
        assert record_placement(Element.create("anchor")) == "free"

    def test_the_vml_flags_read_backwards(self) -> None:
        assert vml_placement('<x:ClientData ObjectType="Button"></x:ClientData>') == "moveAndSize"
        assert vml_placement("<x:ClientData><x:SizeWithCells/></x:ClientData>") == "moveOnly"
        assert vml_placement("<x:ClientData><x:MoveWithCells/><x:SizeWithCells/></x:ClientData>") == "free"

    def test_a_note_box_follows_its_cell_in_the_vml(self) -> None:
        """Measured: a row inserted at a note's cell moves the whole box,
        though its top sits in the row above."""
        vml = (
            '<x:ClientData ObjectType="Note">\r\n  <x:Anchor>\r\n    6, 15, 4, 10, 8, 31, 8, 9</x:Anchor>'
            "\r\n  <x:Row>5</x:Row>\r\n  <x:Column>5</x:Column>\r\n </x:ClientData>"
        )
        moved = move_vml(vml, rows_edit(6, 1))
        assert "6, 15, 5, 10, 8, 31, 9, 9</x:Anchor>" in moved
        # The note's own cell is moved by the caller, after the box.
        assert "<x:Row>5</x:Row>" in moved


class TestInsertingRows:
    def test_the_transform_moves_with_the_anchor(self, shapes: Workbook) -> None:
        sheet = shapes["Shapes"]
        before = sheet.shape("Box")
        sheet.insert_rows(1, 3)
        after = sheet.shape("Box")
        assert (before.top, before.cells) == (50.0, "B3:D6")
        assert (after.top, after.cells) == (50.0 + 3 * 14.5, "B6:D9")
        assert (after.left, after.width, after.height) == (before.left, before.width, before.height)

    def test_a_row_inside_stretches_a_shape_that_sizes_with_its_cells(self, shapes: Workbook) -> None:
        sheet = shapes["Shapes"]
        sheet.insert_rows(4)
        box = sheet.shape("Box")
        assert (box.top, box.height) == (50.0, 60.0 + 14.5)

    def test_columns_move_the_transform_too(self, shapes: Workbook) -> None:
        sheet = shapes["Shapes"]
        sheet.insert_columns(1, 2)
        box = sheet.shape("Box")
        assert (box.left, box.cells) == (100.0 + 2 * 48.0, "D3:F6")

    def test_a_bottom_edge_on_the_line_stays(self, shapes: Workbook) -> None:
        sheet = shapes["Shapes"]
        anchor = anchor_of(sheet, "Box")
        end = anchor.require("to")
        end.require("row").set_text("5")
        end.require("rowOff").set_text("0")
        sheet.insert_rows(6)
        assert corner(anchor, "to")[2:] == (5, 0)
        assert sheet.shape("Box").height == 60.0

    def test_a_group_moves_and_its_members_stay_as_excel_leaves_them(self, shapes: Workbook) -> None:
        sheet = shapes["Shapes"]
        sheet.insert_rows(1, 2)
        rendered = anchor_of(sheet, "Pair").to_xml()
        assert f'<a:off x="5080000" y="{1905000 + 2 * 184150}"/>' in rendered
        assert '<a:chOff x="5080000" y="1905000"/>' in rendered
        assert rendered.count('y="1905000"') == 3, "the child offset and both members"
        assert sheet.shape("Pair").top == 150.0 + 29.0

    def test_a_picture_that_moves_without_sizing_keeps_its_size(self, pictures: Workbook) -> None:
        """Its far corner stays on the sheet, which over rows of one height
        is the same row and offset, as Excel writes it."""
        sheet = pictures[0]
        before = sheet.shape("Natural")
        sheet.insert_rows(6)
        after = sheet.shape("Natural")
        assert (after.top, after.height) == (before.top, before.height)
        assert corner(anchor_of(sheet, "Natural"), "to")[2:] == (7, 63405)

    def test_a_free_floating_shape_stays_where_it_is(self, shapes: Workbook) -> None:
        sheet = shapes["Shapes"]
        anchor = anchor_of(sheet, "Box")
        anchor.set("editAs", "absolute")
        sheet.insert_rows(1, 3)
        assert sheet.shape("Box").top == 50.0
        # 50 points down, past three new rows of 14.5, is 6.5 into the fourth.
        assert corner(anchor, "from")[2:] == (3, round(6.5 * EMU_PER_POINT))

    def test_a_buttons_three_anchors_move_together(self, shapes: Workbook) -> None:
        """A row inserted inside the button stretches it in its drawing
        twin, its record on the sheet, which is the one Excel places it by,
        measured, and its VML. Each keeps its own offset: the VML's are in
        the pixels of the screen that wrote it, here one at 150 percent."""
        sheet = shapes["Shapes"]
        vml_before = button_anchor(sheet)
        sheet.insert_rows(20)
        twin = corner(anchor_of(sheet, "Go"), "to")
        record = next(
            node for node in sheet.document.root.descendants("anchor")
            if node.parent is not None and node.parent.name.endswith("controlPr")
        )
        assert twin == corner(record, "to") == (2, 647700, 21, 114300)
        assert button_anchor(sheet)[6:] == [21, vml_before[7]]

    def test_a_note_box_follows_its_cell(self, shapes: Workbook) -> None:
        sheet = shapes["Shapes"]
        sheet.set_comment("F6", "moves with its cell")
        before = note_anchor(sheet)
        sheet.insert_rows(6)
        after = note_anchor(sheet)
        assert after[2] == before[2] + 1 and after[6] == before[6] + 1
        assert (after[3], after[7]) == (before[3], before[7])


class TestDeletingRows:
    def test_an_edge_whose_row_goes_lands_on_the_boundary(self, shapes: Workbook) -> None:
        """Row 3 is 30 points tall. The top was 21 points into it, and
        lands on its top line; the bottom rises by the whole row."""
        sheet = shapes["Shapes"]
        sheet.delete_rows(3)
        box = sheet.shape("Box")
        assert corner(anchor_of(sheet, "Box"), "from")[2:] == (2, 0)
        assert (box.top, box.height) == (50.0 - 21.0, 60.0 + 21.0 - 30.0)

    def test_a_shape_whose_rows_all_go_goes(self, shapes: Workbook, tmp_path: Path) -> None:
        sheet = shapes["Shapes"]
        sheet.delete_rows(3, 4)
        assert [shape.name for shape in sheet.shapes] == ["Note", "Edge", "Pair", "Go"]
        shapes.save()
        assert dangling(tmp_path / "shapes.xlsm") == []

    def test_so_does_a_shape_whose_columns_all_go(self, shapes: Workbook) -> None:
        sheet = shapes["Shapes"]
        sheet.delete_columns(2, 3)
        assert "Box" not in [shape.name for shape in sheet.shapes]
        assert "Rounded" in [shape.name for shape in sheet.shapes]

    def test_a_control_goes_with_its_record_part_and_vml(self, shapes: Workbook, tmp_path: Path) -> None:
        sheet = shapes["Shapes"]
        go = sheet.shape("Go")
        assert go.control is not None
        sheet.delete_rows(19, 3)
        assert "Go" not in [shape.name for shape in sheet.shapes]
        shapes.save()
        with zipfile.ZipFile(tmp_path / "shapes.xlsm") as archive:
            names = set(archive.namelist())
            text = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
            vml = "".join(archive.read(name).decode("utf-8") for name in names if name.endswith(".vml"))
        assert go.control.part_name not in names
        assert f'shapeId="{go.shape_id}"' not in text and "AlternateContent" not in text
        assert f"_x0000_s{go.shape_id}" not in vml
        assert dangling(tmp_path / "shapes.xlsm") == []

    def test_a_chart_goes_with_its_part(self, charts: Workbook, tmp_path: Path) -> None:
        sheet = charts["Data"]
        line = chart_part_of(sheet, "Line")
        sheet.delete_rows(16, 14)
        assert [shape.name for shape in sheet.shapes] == ["Columns"]
        assert not charts.package.has_part(line)
        charts.save()
        assert dangling(tmp_path / "charts.xlsx") == []

    def test_a_picture_lands_on_the_boundary_at_its_size(self, pictures: Workbook) -> None:
        sheet = pictures[0]
        before = sheet.shape("Natural")
        sheet.delete_rows(4, 6)
        after = sheet.shape("Natural")
        assert corner(anchor_of(sheet, "Natural"), "from")[2:] == (3, 0)
        assert (after.top, after.height) == (before.top - 5.0, before.height)

    def test_a_free_floating_shape_stays(self, shapes: Workbook) -> None:
        sheet = shapes["Shapes"]
        anchor_of(sheet, "Box").set("editAs", "absolute")
        sheet.delete_rows(3, 4)
        box = sheet.shape("Box")
        assert (box.top, box.height) == (50.0, 60.0)

    def test_an_activex_control_in_the_way_refuses_it(self, shapes: Workbook) -> None:
        """Pointed at an ActiveX part, as Excel's record of one is."""
        sheet = shapes["Shapes"]
        control = next(sheet.document.root.descendants("control"))
        folder, _, leaf = sheet.part_name.rpartition("/")
        rels = shapes.package.xml(f"{folder}/_rels/{leaf}.rels")
        entry = next(node for node in rels.root.children_named("Relationship") if node.get("Id") == control.get("r:id"))
        entry.set("Type", RT_ACTIVEX)
        with pytest.raises(ValueError, match="an ActiveX control"):
            sheet.delete_rows(19, 3)
        assert "Go" in [shape.name for shape in sheet.shapes]
        assert sheet.shape("Box").cells == "B3:D6", "nothing moved"

    def test_a_notes_box_stays_when_rows_inside_it_go(self, shapes: Workbook) -> None:
        """Measured: the box follows the cell, and the cell has not moved."""
        sheet = shapes["Shapes"]
        sheet.set_comment("F6", "stays")
        before = note_anchor(sheet)
        sheet.delete_rows(8, 2)
        assert note_anchor(sheet) == before


def dangling(path: Path) -> list[str]:
    """Relationships pointing at a part the package does not hold."""
    missing: list[str] = []
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        for entry in names:
            if not entry.endswith(".rels"):
                continue
            base = entry.rsplit("_rels/", 1)[0]
            for relationship in re.findall(r"<Relationship\b[^>]*>", archive.read(entry).decode("utf-8")):
                target = re.search(r'Target="([^"]+)"', relationship)
                if target is None or 'TargetMode="External"' in relationship:
                    continue
                resolved = f"{base}{target.group(1)}"
                while "/../" in resolved:
                    head, _, tail = resolved.partition("/../")
                    resolved = f"{head.rsplit('/', 1)[0]}/{tail}"
                if resolved not in names:
                    missing.append(f"{entry} -> {target.group(1)}")
    return missing
