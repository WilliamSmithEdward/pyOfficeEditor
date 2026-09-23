"""Text in more than one font in one cell.

``richtext.xlsx`` was built through Excel's ``Characters``, and
``richtext_answers.json`` beside it records the font Excel reported at every
change. Reading is held to those answers, and writing to the very entries
Excel wrote, byte for byte. Rich text other writers leave, which Excel never
writes, is held to the fonts Excel showed for it when measured; the live
gate asks Excel again.
"""

from __future__ import annotations

import json
import re
import shutil
import zipfile
from collections.abc import Sequence
from dataclasses import replace
from pathlib import Path

import pytest

from pyofficeeditor.excel import Alignment, CellFormat, CellRef, Color, Font, TextRun, Workbook, Worksheet

RED = Color(rgb="FFFF0000")

#: The seven cells of ``richtext.xlsx``, asked of this library.
SEVEN: dict[str, list[TextRun | str]] = {
    "A1": [TextRun("bold", Font(bold=True)), " plain"],
    "A2": ["plain ", TextRun("red", Font(color=RED)), " end"],
    "A3": [TextRun("big", Font(size=16)), " and ", TextRun("italic", Font(italic=True))],
    "A4": [TextRun("under", Font(underline="single")), " ", TextRun("over", Font(script="superscript"))],
    "A5": ["font ", TextRun("change", Font(name="Courier New", family=3))],
    "A6": ["two\nlines ", TextRun("bold", Font(bold=True))],
    "A7": [TextRun("all", Font(bold=True, italic=True)), TextRun(" bold", Font(bold=True))],
}

#: Excel's numbers for an underline.
UNDERLINE = {None: -4142, "single": 2}

Answers = dict[str, dict[str, list[dict[str, object]]]]


def excel_colour(color: Color | None) -> int:
    """A colour as Excel's object model gives it, blue in the high byte.
    Automatic and ``Text 1``, theme colour 1, both show as black in the
    Office theme these workbooks carry."""
    if color is None or color == Color(theme=1):
        return 0
    assert color.rgb is not None, color
    red, green, blue = (int(color.rgb[at : at + 2], 16) for at in (2, 4, 6))
    return red + (green << 8) + (blue << 16)


def changes(runs: Sequence[TextRun]) -> list[dict[str, object]]:
    """Each place the font changes, described the way the builder's VBA
    reports it: character by character, so two runs in one font are one."""
    found: list[dict[str, object]] = []
    last: dict[str, object] | None = None
    position = 0
    for run in runs:
        font = run.font
        assert font is not None
        described: dict[str, object] = {
            "name": font.name,
            "size": font.size,
            "bold": font.bold,
            "italic": font.italic,
            "color": excel_colour(font.color),
            "underline": UNDERLINE[font.underline],
            "superscript": font.script == "superscript",
        }
        for _ in run.text:
            position += 1
            if described != last:
                found.append({"start": position, **described})
            last = described
    return found


def entries(path: Path) -> list[str]:
    """The ``<si>`` entries of a saved workbook, exactly as written."""
    with zipfile.ZipFile(path) as package:
        text = package.read("xl/sharedStrings.xml").decode("utf-8")
    return re.findall(r"<si>.*?</si>", text, flags=re.S)


def saved(sheet: Worksheet) -> Path:
    """Save the sheet's workbook where it was opened from, and say where."""
    sheet.workbook.save()
    path = sheet.workbook.path
    assert path is not None
    return path


def runs_of(sheet: Worksheet, address: str) -> tuple[TextRun, ...]:
    runs = sheet.get_rich_text(CellRef.parse(address))
    assert runs is not None, address
    return runs


@pytest.fixture(scope="module")
def answers(live_richtext_answers: Path) -> Answers:
    return json.loads(live_richtext_answers.read_text(encoding="utf-8"))


@pytest.fixture
def excels(tmp_path: Path, live_richtext_xlsx: Path) -> Worksheet:
    target = tmp_path / "richtext.xlsx"
    shutil.copy(live_richtext_xlsx, target)
    return Workbook.open(target)["Rich"]


@pytest.fixture
def empty(tmp_path: Path, live_empty_xlsx: Path) -> Worksheet:
    """A sheet Excel saved with nothing on it, in Aptos Narrow 11 as
    ``richtext.xlsx`` is."""
    target = tmp_path / "empty.xlsx"
    shutil.copy(live_empty_xlsx, target)
    return Workbook.open(target)[0]


def write_seven(sheet: Worksheet) -> None:
    for address, runs in SEVEN.items():
        sheet.set_rich_text(CellRef.parse(address), runs)


class TestReadingExcelsOwn:
    @pytest.mark.parametrize("address", list(SEVEN))
    def test_every_font_change_where_excel_reported_it(
        self, excels: Worksheet, answers: Answers, address: str
    ) -> None:
        assert changes(runs_of(excels, address)) == answers[address]["runs"]

    def test_the_first_run_is_in_the_cells_font(self, excels: Worksheet) -> None:
        """Excel wrote no font for it. The cell's is bold, so it is too."""
        first = runs_of(excels, "A1")[0]
        assert first == TextRun("bold", excels.get_format(CellRef.parse("A1")).font)
        assert first.font is not None and first.font.bold

    def test_the_runs_make_up_the_value(self, excels: Worksheet) -> None:
        for address in SEVEN:
            text = "".join(run.text for run in runs_of(excels, address))
            assert text == excels.get_value(CellRef.parse(address)), address

    def test_a_line_break_reads_as_one_character(self, excels: Worksheet) -> None:
        """Excel stored it as CRLF, and counts it as one."""
        assert runs_of(excels, "A6")[0].text == "two\nlines "

    def test_a_cell_with_no_text_has_no_runs(self, empty: Worksheet) -> None:
        empty.set_value(CellRef.parse("B1"), 12)
        assert empty.get_rich_text(CellRef.parse("B1")) is None
        assert empty.get_rich_text(CellRef.parse("B2")) is None

    def test_plain_text_is_one_run_in_the_cells_font(self, empty: Worksheet) -> None:
        reference = CellRef.parse("B1")
        empty.set_value(reference, "plain")
        assert empty.get_rich_text(reference) == (TextRun("plain", empty.get_format(reference).font),)


#: The cells' font in the foreign workbook, and the workbook's default.
COURIER = Font(name="Courier New", size=16, color=RED, family=3)
ARIAL = Font(name="Arial", size=10, color=Color(rgb="FF00B050"), family=2)
#: A run's own font where it sets nothing but what is added to it: the
#: default font's typeface and size, and an automatic colour.
BARE = Font(name="Arial", size=10, family=2)

#: What Excel showed for each run of ``FOREIGN_RICH_TEXT``, measured with
#: ``Characters``: typeface, size, bold, italic and colour agreed for every
#: character, and an automatic colour reported as automatic.
SHOWN: dict[str, list[tuple[str, Font]]] = {
    "A1": [("a", COURIER), ("b", COURIER), ("c", replace(BARE, bold=True)), ("d", ARIAL)],
    "A2": [("x", replace(BARE, bold=True)), ("y", ARIAL)],
    "A3": [("p", COURIER), ("q", replace(BARE, size=20))],
    "A4": [("m", ARIAL), ("n", replace(BARE, italic=True))],
    "A5": [("e", COURIER), ("f", Font(name="Times New Roman", size=10))],
    "A6": [("j", COURIER), ("k", COURIER)],
    "A7": [("u", replace(BARE, bold=True))],
    "A8": [("x", replace(BARE, bold=True)), ("y", ARIAL)],
    "A9": [("a", COURIER), ("b", replace(COURIER, bold=True)), ("c", ARIAL)],
    "A10": [("n", COURIER), ("o", replace(BARE, italic=True, color=Color(theme=1)))],
    "A11": [("x", replace(BARE, bold=True)), ("y", ARIAL)],
    "A12": [("a", COURIER), ("b", replace(BARE, italic=True))],
}


class TestReadingOtherWriters:
    """A run with no font is in the cell's until a run has one, and in the
    workbook's default font after that. A run's font takes what it leaves
    unsaid from the default font too, not from the cell's."""

    @pytest.fixture
    def foreign(self, foreign_rich_text: Path, tmp_path: Path) -> Worksheet:
        target = tmp_path / "foreign.xlsx"
        shutil.copy(foreign_rich_text, target)
        return Workbook.open(target)[0]

    @pytest.mark.parametrize("address", list(SHOWN))
    def test_each_run_in_the_font_excel_showed(self, foreign: Worksheet, address: str) -> None:
        assert [(run.text, run.font) for run in runs_of(foreign, address)] == SHOWN[address]

    def test_the_cells_are_in_the_fonts_the_runs_are_told_apart_by(self, foreign: Worksheet) -> None:
        assert foreign.get_format(CellRef.parse("A1")).font == COURIER
        assert foreign.get_format(CellRef.parse("A4")).font == ARIAL


class TestWriting:
    def test_the_entries_are_the_ones_excel_wrote(self, empty: Worksheet, live_richtext_xlsx: Path) -> None:
        write_seven(empty)
        assert entries(saved(empty)) == entries(live_richtext_xlsx)

    def test_each_cell_takes_the_font_excel_gave_it(self, empty: Worksheet, excels: Worksheet) -> None:
        """The first run's: Excel writes that run with no font of its own."""
        write_seven(empty)
        for address in SEVEN:
            reference = CellRef.parse(address)
            assert empty.get_format(reference).font == excels.get_format(reference).font, address

    def test_what_is_written_reads_as_excel_read_its_own(self, empty: Worksheet, answers: Answers) -> None:
        write_seven(empty)
        for address in SEVEN:
            assert changes(runs_of(empty, address)) == answers[address]["runs"], address

    def test_it_reads_the_same_after_a_save(self, empty: Worksheet) -> None:
        write_seven(empty)
        before = {address: runs_of(empty, address) for address in SEVEN}
        again = Workbook.open(saved(empty))[0]
        assert {address: runs_of(again, address) for address in SEVEN} == before

    def test_one_run_is_plain_text_in_its_font(self, empty: Worksheet) -> None:
        reference = CellRef.parse("B1")
        empty.set_rich_text(reference, [TextRun("loud", Font(bold=True))])
        font = empty.get_format(reference).font
        assert font.bold
        assert empty.get_rich_text(reference) == (TextRun("loud", font),)
        assert entries(saved(empty)) == ["<si><t>loud</t></si>"]

    def test_a_plain_first_run_leaves_the_cell_as_it_was(self, empty: Worksheet) -> None:
        reference = CellRef.parse("B1")
        empty.set_rich_text(reference, ["plain ", TextRun("red", Font(color=RED))])
        assert empty.style_index(reference) is None

    def test_the_first_runs_font_keeps_the_rest_of_the_format(self, empty: Worksheet) -> None:
        reference = CellRef.parse("B1")
        empty.set_format(reference, CellFormat(number_format="@", alignment=Alignment(wrap_text=True)))
        empty.set_rich_text(reference, [TextRun("bold", Font(bold=True)), " plain"])
        wanted = empty.get_format(reference)
        assert (wanted.number_format, wanted.alignment.wrap_text, wanted.font.bold) == ("@", True, True)

    def test_a_run_takes_what_it_leaves_unsaid_from_the_cell(self, empty: Worksheet) -> None:
        """Typeface, size and colour; bold, italic and the like are its own."""
        reference = CellRef.parse("B1")
        empty.set_format(reference, CellFormat(font=Font(name="Courier New", size=14, italic=True, color=RED, family=3)))
        empty.set_rich_text(reference, ["plain ", TextRun("bold", Font(bold=True))])
        second = runs_of(empty, "B1")[1]
        assert second.font == Font(name="Courier New", size=14, bold=True, color=RED, family=3)

    def test_a_new_typeface_leaves_the_theme_behind(self, empty: Worksheet) -> None:
        """The cell's font follows the theme; a run in Courier must not,
        or the theme's typeface would show in its place."""
        empty.set_rich_text(CellRef.parse("B1"), ["font ", TextRun("change", Font(name="Courier New", family=3))])
        (entry,) = entries(saved(empty))
        assert '<rFont val="Courier New"/><family val="3"/></rPr>' in entry
        assert "scheme" not in entry.split("<t>change</t>")[0].rsplit("<rPr>", 1)[1]

    def test_a_line_break_is_written_as_excel_writes_it(self, empty: Worksheet) -> None:
        empty.set_rich_text(CellRef.parse("B1"), ["two\nlines ", TextRun("bold", Font(bold=True))])
        path = saved(empty)
        (entry,) = entries(path)
        assert "two\r\nlines " in entry
        assert runs_of(Workbook.open(path)[0], "B1")[0].text == "two\nlines "

    def test_the_same_runs_are_stored_once(self, empty: Worksheet) -> None:
        runs: list[TextRun | str] = [TextRun("a", Font(bold=True)), "b"]
        empty.set_rich_text(CellRef.parse("B1"), runs)
        empty.set_rich_text(CellRef.parse("B2"), runs)
        assert len(entries(saved(empty))) == 1

    def test_a_formula_gives_way(self, empty: Worksheet) -> None:
        reference = CellRef.parse("B1")
        empty.set_formula(reference, "1+1")
        empty.set_rich_text(reference, ["one ", TextRun("two", Font(bold=True))])
        assert empty.get_formula(reference) is None
        assert empty.get_value(reference) == "one two"

    @pytest.mark.parametrize("runs", [[], [""], ["", TextRun("")]])
    def test_some_text_is_needed(self, empty: Worksheet, runs: list[TextRun | str]) -> None:
        with pytest.raises(ValueError, match="needs some text"):
            empty.set_rich_text(CellRef.parse("B1"), runs)

    def test_a_cell_reads_and_writes_its_runs(self, empty: Worksheet) -> None:
        cell = empty["B1"]
        cell.rich_text = ["plain ", TextRun("bold", Font(bold=True))]
        assert cell.rich_text == empty.get_rich_text(CellRef.parse("B1"))
        assert empty["B9"].rich_text is None
