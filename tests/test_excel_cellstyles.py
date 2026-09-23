"""Named cell styles.

``styles.xlsx`` was built by Excel with every built-in style it lists
applied to a cell of its own, a style of the workbook's own, and styles put
over formatting a cell already had. ``styles_answers.json`` records what
Excel showed for each cell and what each style includes. Reading is held
to those, and writing to the formats Excel's own cells ended up with.
"""

from __future__ import annotations

import json
import re
import shutil
import zipfile
from dataclasses import replace
from pathlib import Path

import pytest

from pyofficeeditor.excel import (
    Alignment,
    CellFormat,
    CellRef,
    CellStyle,
    Color,
    Fill,
    Font,
    Workbook,
    Worksheet,
)
from pyofficeeditor.excel._cellstyles import ASPECTS, BUILTIN_STYLE_NAMES

Answers = dict[str, dict[str, object]]


@pytest.fixture(scope="module")
def answers(live_styles_answers: Path) -> Answers:
    return json.loads(live_styles_answers.read_text(encoding="utf-8"))


@pytest.fixture
def excels(tmp_path: Path, live_styles_xlsx: Path) -> Worksheet:
    target = tmp_path / "styles.xlsx"
    shutil.copy(live_styles_xlsx, target)
    return Workbook.open(target)["Styles"]


@pytest.fixture
def empty(tmp_path: Path, live_empty_xlsx: Path) -> Worksheet:
    target = tmp_path / "empty.xlsx"
    shutil.copy(live_empty_xlsx, target)
    return Workbook.open(target)[0]


def rows(answers: Answers) -> list[tuple[int, str]]:
    """Each row of column B, and the style Excel gave it."""
    found = [(int(key[1:]), str(value["style"])) for key, value in answers.items() if re.fullmatch(r"B\d+", key)]
    return sorted(found)


def unanchored(cell_format: CellFormat) -> CellFormat:
    """A format without the index of its style, which differs between two
    workbooks holding the same styles."""
    return replace(cell_format, style_id=0)


def saved(sheet: Worksheet) -> Path:
    sheet.workbook.save()
    path = sheet.workbook.path
    assert path is not None
    return path


class TestReadingExcelsOwn:
    def test_every_style_the_workbook_defines(self, excels: Worksheet, answers: Answers) -> None:
        names = [style.name for style in excels.workbook.cell_styles]
        expected = sorted([name for _, name in rows(answers)] + ["Mine", "Normal"], key=str.casefold)
        assert names == expected, "every one, in name order"

    def test_what_each_style_sets(self, excels: Worksheet, answers: Answers) -> None:
        for style in excels.workbook.cell_styles:
            said = answers.get(f"includes {style.name}")
            if said is None:
                continue
            assert style.aspects == tuple(aspect for aspect in ASPECTS if said[aspect]), style.name

    def test_excels_numbers_for_its_own(self, excels: Worksheet) -> None:
        for style in excels.workbook.cell_styles:
            if style.name == "Mine":
                assert style.builtin_id is None
            else:
                assert style.builtin_id is not None
                assert BUILTIN_STYLE_NAMES[style.builtin_id] == style.name

    def test_each_cells_style(self, excels: Worksheet, answers: Answers) -> None:
        for row, name in rows(answers):
            assert excels.get_cell_style(CellRef(row, 2)) == name
        assert excels["D1"].style == "Mine"
        assert excels["A1"].style == "Normal", "a cell with no style of its own"

    def test_a_style_over_formatting_keeps_what_it_does_not_set(self, excels: Worksheet) -> None:
        """Good over a bold, centred cell showing two decimals."""
        good = excels.workbook.cell_style("Good")
        assert good is not None
        found = excels["D2"].format
        assert (found.number_format, found.alignment.horizontal) == ("0.00", "center")
        assert (found.font, found.fill) == (good.format.font, good.format.fill)


class TestWritingAsExcelDoes:
    def test_each_of_excels_styles_on_a_cell(self, empty: Worksheet, excels: Worksheet, answers: Answers) -> None:
        for row, name in rows(answers):
            reference = CellRef(row, 2)
            empty.set_cell_style(reference, name)
            assert empty.get_cell_style(reference) == name
            assert unanchored(empty.get_format(reference)) == unanchored(excels.get_format(reference)), name

    def test_each_style_is_defined_as_excel_defines_it(self, empty: Worksheet, excels: Worksheet) -> None:
        theirs = {style.name: style for style in excels.workbook.cell_styles}
        for name in sorted(BUILTIN_STYLE_NAMES.values()):
            if name == "Normal":
                continue
            empty.set_cell_style(CellRef.parse("A1"), name)
            ours = empty.workbook.cell_style(name)
            assert ours is not None
            expected = theirs[name]
            assert (unanchored(ours.format), ours.aspects, ours.builtin_id) == (
                unanchored(expected.format),
                expected.aspects,
                expected.builtin_id,
            ), name

    def test_a_style_over_formatting(self, empty: Worksheet, excels: Worksheet) -> None:
        cell = empty["D2"]
        cell.format = CellFormat(
            number_format="0.00", font=replace(cell.format.font, bold=True), alignment=Alignment(horizontal="center")
        )
        cell.style = "Good"
        assert unanchored(cell.format) == unanchored(excels["D2"].format)

    def test_a_style_then_formatting(self, empty: Worksheet, excels: Worksheet) -> None:
        cell = empty["D3"]
        cell.style = "Currency"
        cell.font = replace(cell.font, italic=True)
        assert cell.style == "Currency"
        assert unanchored(cell.format) == unanchored(excels["D3"].format)

    def test_a_style_of_the_workbooks_own(self, empty: Worksheet, excels: Worksheet) -> None:
        normal = empty.get_format(CellRef.parse("A1")).font
        mine = empty.workbook.add_cell_style(
            "Mine",
            CellFormat(
                font=replace(normal, bold=True, color=Color(rgb="FF0000FF")),
                fill=Fill(pattern="solid", foreground=Color(rgb="FFFFFF00"), background=Color(indexed=64)),
            ),
            aspects=("font", "fill"),
        )
        assert (mine.name, mine.aspects, mine.builtin_id) == ("Mine", ("font", "fill"), None)
        empty["D1"].style = "Mine"
        assert unanchored(empty["D1"].format) == unanchored(excels["D1"].format)

    def test_styles_are_kept_in_name_order(self, empty: Worksheet) -> None:
        for name in ("Title", "Bad", "20% - Accent3", "Normal", "Comma"):
            empty["A1"].style = name
        names = [style.name for style in empty.workbook.cell_styles]
        assert names == ["20% - Accent3", "Bad", "Comma", "Normal", "Title"]

    def test_a_style_is_defined_once(self, empty: Worksheet) -> None:
        empty["A1"].style = "Good"
        empty["A2"].style = "Good"
        assert [style.name for style in empty.workbook.cell_styles].count("Good") == 1
        assert empty["A1"].style_index == empty["A2"].style_index

    def test_a_name_is_matched_whatever_its_case(self, empty: Worksheet) -> None:
        empty["A1"].style = "heading 1"
        assert empty["A1"].style == "Heading 1"

    def test_normal_takes_everything_back(self, empty: Worksheet) -> None:
        cell = empty["A1"]
        cell.style = "Check Cell"
        cell.style = "Normal"
        assert unanchored(cell.format) == unanchored(empty["Z99"].format)

    def test_comma_writes_its_format_as_excel_does(self, empty: Worksheet, live_styles_xlsx: Path) -> None:
        """Under the built-in id, with the code written out."""
        empty["A1"].style = "Comma"
        pattern = r'<numFmt numFmtId="43"[^>]*/>'
        with zipfile.ZipFile(saved(empty)) as ours, zipfile.ZipFile(live_styles_xlsx) as theirs:
            written = re.findall(pattern, ours.read("xl/styles.xml").decode("utf-8"))
            assert written == re.findall(pattern, theirs.read("xl/styles.xml").decode("utf-8"))

    def test_a_range_takes_a_style(self, empty: Worksheet) -> None:
        empty.range("A1:B2").apply_style("Note")
        assert {empty.get_cell_style(reference) for reference in empty.range("A1:B2").reference.cells()} == {"Note"}

    def test_it_survives_a_save(self, empty: Worksheet) -> None:
        empty["A1"].style = "Title"
        empty.workbook.add_cell_style("Mine", CellFormat(font=Font(name="Arial", size=12)))
        again = Workbook.open(saved(empty))
        assert again[0]["A1"].style == "Title"
        mine = again.cell_style("mine")
        assert isinstance(mine, CellStyle) and mine.format.font == Font(name="Arial", size=12)


class TestRefusals:
    def test_a_style_no_one_has(self, empty: Worksheet) -> None:
        with pytest.raises(ValueError, match="no cell style named 'Fancy'"):
            empty["A1"].style = "Fancy"

    @pytest.mark.parametrize("name", ["Good", "good", "Normal", "", "  "])
    def test_a_name_that_cannot_be_the_workbooks_own(self, empty: Worksheet, name: str) -> None:
        with pytest.raises(ValueError):
            empty.workbook.add_cell_style(name, CellFormat())

    def test_a_name_the_workbook_has(self, empty: Worksheet) -> None:
        empty.workbook.add_cell_style("Mine", CellFormat())
        with pytest.raises(ValueError, match="already has"):
            empty.workbook.add_cell_style("MINE", CellFormat())

    def test_an_aspect_that_is_not_one(self, empty: Worksheet) -> None:
        with pytest.raises(ValueError, match="not aspects"):
            empty.workbook.add_cell_style("Mine", CellFormat(), aspects=("colour",))  # type: ignore[list-item]


class TestTheTheme:
    def test_the_typefaces_excels_styles_name(self, empty: Worksheet) -> None:
        assert empty.workbook.theme_fonts == ("Aptos Display", "Aptos Narrow")
