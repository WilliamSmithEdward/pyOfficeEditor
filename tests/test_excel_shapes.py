"""Shapes on a sheet.

``shapes.xlsm`` was built by Excel with one of every kind a sheet can hold,
and ``shapes_answers.json`` beside it is what Excel then said about each one
through its own object model. The reader is held to those answers rather
than to a reading of the markup.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyofficeeditor.excel import Workbook, Worksheet
from pyofficeeditor.excel._shapes import (
    AUTO_SHAPE_TYPES,
    DEFAULT_COLUMN_POINTS,
    EMU_PER_POINT,
    MSO_TYPE,
    PRESET_GEOMETRY,
    SheetGrid,
    characters_to_points,
    emu,
    points,
)


@pytest.fixture(scope="module")
def sheet(live_shapes_xlsm: Path) -> Worksheet:
    return Workbook.open(live_shapes_xlsm)["Shapes"]


@pytest.fixture(scope="module")
def answers(live_shapes_answers: Path) -> dict[str, dict[str, float]]:
    return json.loads(live_shapes_answers.read_text(encoding="utf-8"))


class TestUnits:
    def test_a_point_is_12700_emu(self) -> None:
        assert EMU_PER_POINT == 12700
        assert points(1270000) == 100.0
        assert emu(100.0) == 1270000

    def test_they_round_trip(self) -> None:
        assert points(emu(37.5)) == 37.5

    def test_the_default_column(self) -> None:
        """8.43 characters is 64 pixels, and a point is three quarters of
        one, so exactly 48."""
        assert characters_to_points(8.43) == DEFAULT_COLUMN_POINTS == 48.0

    def test_rounding_to_whole_pixels_first(self) -> None:
        """Without it the default comes out 47.9 rather than 48."""
        assert characters_to_points(8.43) != (8.43 * 7 + 5) * 0.75


class TestTheGrid:
    def test_an_empty_grid_uses_the_defaults(self) -> None:
        grid = SheetGrid()
        assert grid.x(0) == 0.0
        assert grid.x(2) == 96.0, "two default columns"
        assert grid.y(2) == 29.0, "two default rows"

    def test_an_offset_is_emu(self) -> None:
        assert SheetGrid().x(0, 12700) == 1.0

    def test_column_widths_come_from_the_sheet(self, sheet: Worksheet) -> None:
        grid = SheetGrid.of(sheet.document.root)
        assert grid.column_widths[0] == characters_to_points(12.6328125)
        assert 3 not in grid.column_widths, "only A to C were widened"

    def test_row_heights_come_from_the_sheet(self, sheet: Worksheet) -> None:
        grid = SheetGrid.of(sheet.document.root)
        # The probe set rows 3 and 4, which are 2 and 3 zero-based.
        assert grid.row_heights[2] == 30.0
        assert grid.default_row == 14.5


class TestReadingWhatExcelWrote:
    def test_every_shape_is_found(
        self, sheet: Worksheet, answers: dict[str, dict[str, float]]
    ) -> None:
        assert {shape.name for shape in sheet.shapes} == set(answers)

    def test_each_kind_matches_excels_own_type(
        self, sheet: Worksheet, answers: dict[str, dict[str, float]]
    ) -> None:
        for shape in sheet.shapes:
            assert shape.mso_type == answers[shape.name]["type"], shape.name

    def test_the_kinds(self, sheet: Worksheet) -> None:
        kinds = {shape.name: shape.kind for shape in sheet.shapes}
        assert kinds["Box"] == "shape"
        assert kinds["Note"] == "textBox", "an sp with txBox, not its own element"
        assert kinds["Edge"] == "line"
        assert kinds["Pair"] == "group"
        assert kinds["Go"] == "formControl"

    def test_the_geometry(self, sheet: Worksheet) -> None:
        geometry = {shape.name: shape.geometry for shape in sheet.shapes}
        assert geometry["Box"] == "rect"
        assert geometry["Rounded"] == "roundRect"
        assert geometry["Oval"] == "ellipse"

    @pytest.mark.parametrize("name", ["Box", "Rounded", "Oval", "Note", "Edge", "Pair"])
    def test_the_box_is_exact_for_an_ordinary_shape(
        self, sheet: Worksheet, answers: dict[str, dict[str, float]], name: str
    ) -> None:
        """Because the transform is used where it says anything, and Excel's
        own Left, Top, Width and Height agree with it to the point."""
        shape = sheet.shape(name)
        said = answers[name]
        assert (shape.left, shape.top) == (said["left"], said["top"])
        assert (shape.width, shape.height) == (said["width"], said["height"])

    def test_a_form_controls_box_comes_from_the_anchor(
        self, sheet: Worksheet, answers: dict[str, dict[str, float]]
    ) -> None:
        """Its transform is all zeros, so the grid has to answer, and the
        grid is not exact: a column width is stored in characters of a font
        whose metrics the file does not carry. A quarter of a point out is
        what to expect here, and nowhere else."""
        shape = sheet.shape("Go")
        said = answers["Go"]
        assert abs(shape.left - said["left"]) <= 0.5
        assert abs(shape.top - said["top"]) <= 0.5
        assert shape.top == said["top"], "rows are exact; only columns are not"

    def test_a_drawing_shapes_macro_is_on_the_shape(self, sheet: Worksheet) -> None:
        assert sheet.shape("Box").macro == "Clicked"

    def test_a_form_controls_macro_is_not(self, sheet: Worksheet) -> None:
        """It comes from the sheet's own ``<control>``, where the bracketed
        number names the workbook."""
        assert sheet.shape("Go").macro == "[1]!Clicked"

    def test_a_shape_with_no_macro(self, sheet: Worksheet) -> None:
        assert sheet.shape("Oval").macro == ""

    def test_text(self, sheet: Worksheet) -> None:
        assert sheet.shape("Box").text == "hello"

    def test_a_line_break_is_normalised(self, sheet: Worksheet) -> None:
        """Excel stores it as CRLF inside the run and reports it back as one
        character."""
        assert sheet.shape("Note").text == "first\nsecond"
        assert "\r" not in sheet.shape("Note").text

    def test_a_group_holds_its_members(self, sheet: Worksheet) -> None:
        pair = sheet.shape("Pair")
        assert {child.name for child in pair.children} == {"GroupA", "GroupB"}

    def test_the_members_are_not_listed_separately(self, sheet: Worksheet) -> None:
        """Excel reports the group, not the two shapes inside it."""
        assert "GroupA" not in {shape.name for shape in sheet.shapes}

    def test_shape_ids_are_the_drawings_own(self, sheet: Worksheet) -> None:
        """A control is numbered from 1025 while drawing shapes start at 2,
        and the number is what links a control to its record."""
        assert sheet.shape("Box").shape_id == 2
        assert sheet.shape("Go").shape_id == 1025

    def test_looking_one_up_by_name(self, sheet: Worksheet) -> None:
        assert sheet.shape("Box").name == "Box"

    def test_an_unknown_name(self, sheet: Worksheet) -> None:
        with pytest.raises(KeyError, match="no shape named"):
            sheet.shape("Nope")


class TestASheetWithNoShapes:
    def test_it_reports_none(self, live_sample_xlsx: Path) -> None:
        assert Workbook.open(live_sample_xlsx)["Data"].shapes == []

    def test_mso_type_covers_every_kind(self) -> None:
        """A kind with no number would report 1 and look like an AutoShape."""
        from pyofficeeditor.excel._shapes import ShapeKind

        named = set(MSO_TYPE)
        assert named >= {"shape", "textBox", "line", "group", "formControl"}
        assert "other" not in named, "the catch-all deliberately has no number"
        del ShapeKind


class TestPresetGeometry:
    """The table is held to a workbook in which Excel made one shape of
    each type, because every plausible wrong answer here is a real preset
    name belonging to some other shape: the file stays valid and the wrong
    shape appears."""

    def measured(self, path: Path) -> dict[int, str]:
        import re

        from pyofficeeditor.opc import OpcPackage

        raw = OpcPackage.open(path).read("xl/drawings/drawing1.xml").decode()
        return {
            int(m.group(1)): m.group(2)
            for m in re.finditer(r'name="T(\d+)".*?<a:prstGeom prst="([^"]+)"', raw, re.S)
        }

    def test_every_entry_is_what_excel_wrote(self, live_geometry_xlsx: Path) -> None:
        for number, preset in self.measured(live_geometry_xlsx).items():
            assert PRESET_GEOMETRY.get(number) == preset, number

    def test_the_table_claims_nothing_unmeasured(self, live_geometry_xlsx: Path) -> None:
        assert set(PRESET_GEOMETRY) == set(self.measured(live_geometry_xlsx))

    @pytest.mark.parametrize(
        ("number", "preset"),
        [(11, "plus"), (12, "pentagon"), (16, "foldedCorner"), (17, "smileyFace")],
    )
    def test_the_four_that_are_easy_to_get_wrong(self, number: int, preset: str) -> None:
        """cross, star5, can and cube are the plausible answers, and all
        four are real presets belonging to other shapes."""
        assert PRESET_GEOMETRY[number] == preset

    def test_the_five_pointed_star_is_92_not_12(self) -> None:
        assert PRESET_GEOMETRY[92] == "star5"
        assert PRESET_GEOMETRY[12] != "star5"

    def test_the_inverse_agrees(self) -> None:
        for number, preset in PRESET_GEOMETRY.items():
            assert AUTO_SHAPE_TYPES[preset] == number

    def test_a_shape_reports_its_auto_shape_type(self, sheet: Worksheet) -> None:
        assert sheet.shape("Box").auto_shape_type == 1
        assert sheet.shape("Rounded").auto_shape_type == 5
        assert sheet.shape("Oval").auto_shape_type == 9

    def test_a_text_box_answers_one_too(self, sheet: Worksheet) -> None:
        """Which is what Excel reports: both are drawn as a rectangle."""
        assert sheet.shape("Note").auto_shape_type == 1

    def test_a_geometry_with_no_number(self, sheet: Worksheet) -> None:
        assert sheet.shape("Edge").auto_shape_type is None
