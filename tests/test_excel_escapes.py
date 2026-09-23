"""Text XML cannot carry as it is, read and written as Excel does.

``escapes.xlsx`` was built by Excel with control characters, a lone
carriage return and a literal ``_x0041_`` everywhere a workbook keeps text,
and ``escapes_answers.json`` records the characters Excel's object model
reported for each piece. Reading is held to those answers, and writing to
the markup Excel wrote; the live gate has Excel open what is written.
"""

from __future__ import annotations

import json
import re
import shutil
import zipfile
from pathlib import Path

import pytest

from pyofficeeditor.excel import (
    CellRef,
    DataValidation,
    HeaderFooter,
    HeaderFooterText,
    Workbook,
    Worksheet,
    expression,
)

#: The cells Excel was given text in, in order.
TEXT_CELLS = [f"Text!A{row}" for row in range(1, 15)]
#: The table's headers, as Excel was given them before it made the table.
HEADERS = ["head\ntwo", "_x0041_", "plain", "ctl\x01x", "tab\tx", "cr\rx"]

Answers = dict[str, dict[str, str]]


@pytest.fixture(scope="module")
def answers(live_escapes_answers: Path) -> Answers:
    return json.loads(live_escapes_answers.read_text(encoding="utf-8"))


@pytest.fixture
def excels(tmp_path: Path, live_escapes_xlsx: Path) -> Workbook:
    target = tmp_path / "escapes.xlsx"
    shutil.copy(live_escapes_xlsx, target)
    return Workbook.open(target)


@pytest.fixture
def empty(tmp_path: Path, live_empty_xlsx: Path) -> Worksheet:
    target = tmp_path / "empty.xlsx"
    shutil.copy(live_empty_xlsx, target)
    return Workbook.open(target)[0]


def said(answers: Answers, key: str) -> str:
    return answers[key]["text"]


def part(path: Path, name: str) -> str:
    with zipfile.ZipFile(path) as package:
        return package.read(name).decode("utf-8")


def saved(sheet: Worksheet) -> Path:
    sheet.workbook.save()
    path = sheet.workbook.path
    assert path is not None
    return path


class TestReadingExcelsOwn:
    @pytest.mark.parametrize("key", TEXT_CELLS)
    def test_the_text_in_each_cell(self, excels: Workbook, answers: Answers, key: str) -> None:
        sheet, _, address = key.partition("!")
        assert excels[sheet].get_value(CellRef.parse(address)) == said(answers, key)

    @pytest.mark.parametrize("address", ["B1", "B2"])
    def test_a_formulas_text_result(self, excels: Workbook, answers: Answers, address: str) -> None:
        assert excels["Text"].get_value(CellRef.parse(address)) == said(answers, f"Text!{address}")

    @pytest.mark.parametrize("address", ["B1", "B2", "B3"])
    def test_a_formula(self, excels: Workbook, answers: Answers, address: str) -> None:
        """B3 names the sheet called ``a_x0041_b``, which its formula spells
        ``a_x005F_x0041_b``."""
        formula = excels["Text"].get_formula(CellRef.parse(address))
        assert f"={formula}" == said(answers, f"Text!{address} formula")

    def test_a_note(self, excels: Workbook, answers: Answers) -> None:
        note = excels["Text"].comment("C1")
        assert note is not None and note.text == said(answers, "note")

    def test_a_validations_messages(self, excels: Workbook, answers: Answers) -> None:
        (rule,) = excels["Text"].data_validations
        assert (rule.prompt_title, rule.prompt_message, rule.error_message) == (
            said(answers, "validation title"),
            said(answers, "validation message"),
            said(answers, "validation error"),
        )

    def test_a_hyperlinks_tip(self, excels: Workbook, answers: Answers) -> None:
        (link,) = excels["Text"].hyperlinks
        assert link.tooltip == said(answers, "hyperlink tip")

    def test_a_tables_columns_and_headers(self, excels: Workbook, answers: Answers) -> None:
        sheet = excels["Text"]
        (table,) = sheet.tables
        count = len(HEADERS)
        assert table.column_names == [said(answers, f"table column {index}") for index in range(1, count + 1)]
        headers = [sheet.get_value(CellRef(1, 7 + offset)) for offset in range(count)]
        assert headers == [said(answers, f"table header {index}") for index in range(1, count + 1)]

    def test_a_page_header(self, excels: Workbook, answers: Answers) -> None:
        assert excels["Text"].header_footer.odd_header.center == said(answers, "page header")

    def test_a_sheets_name(self, excels: Workbook, answers: Answers) -> None:
        assert excels.sheet_names == ["Text", said(answers, "sheet name")]
        assert excels["a_x0041_b"].get_value(CellRef.parse("A1")) == 7

    def test_a_defined_names_comment(self, excels: Workbook, answers: Answers) -> None:
        (name,) = excels.defined_names
        assert name.comment == said(answers, "name comment")


class TestWritingAsExcelWrote:
    def test_cell_text(self, empty: Worksheet, answers: Answers, live_escapes_xlsx: Path) -> None:
        for row, key in enumerate(TEXT_CELLS, start=1):
            empty.set_value(CellRef(row, 1), said(answers, key))
        ours = re.findall(r"<si>.*?</si>", part(saved(empty), "xl/sharedStrings.xml"), flags=re.S)
        theirs = re.findall(r"<si>.*?</si>", part(live_escapes_xlsx, "xl/sharedStrings.xml"), flags=re.S)
        assert ours == theirs[: len(TEXT_CELLS)]

    def test_a_formula(self, empty: Worksheet) -> None:
        empty.set_formula(CellRef.parse("B2"), '="_x0041_"')
        assert '<f>"_x005F_x0041_"</f>' in part(saved(empty), "xl/worksheets/sheet1.xml")

    def test_a_note(self, empty: Worksheet, answers: Answers, live_escapes_xlsx: Path) -> None:
        empty.set_comment("C1", said(answers, "note"))
        ours = re.search(r"<t>.*?</t>", part(saved(empty), "xl/comments1.xml"), flags=re.S)
        theirs = re.search(r"<t>.*?</t>", part(live_escapes_xlsx, "xl/comments1.xml"), flags=re.S)
        assert ours is not None and theirs is not None
        assert ours.group(0) == theirs.group(0)

    def test_a_validations_messages(self, empty: Worksheet, answers: Answers) -> None:
        rule = DataValidation.whole_number(
            1,
            9,
            prompt_title=said(answers, "validation title"),
            prompt_message=said(answers, "validation message"),
            error_message=said(answers, "validation error"),
        )
        empty.add_data_validation("D1", rule)
        sheet = part(saved(empty), "xl/worksheets/sheet1.xml")
        for written in ('error="er_x000d_ror"', 'promptTitle="ti_x0001_t"', 'prompt="in_x000a_put_x0001_ _x005f_x0041_"'):
            assert written in sheet

    def test_a_hyperlinks_tip(self, empty: Worksheet, answers: Answers) -> None:
        empty.add_hyperlink("E1", "https://example.com/", tooltip=said(answers, "hyperlink tip"))
        assert 'tooltip="tip_x0001__x000a_x"' in part(saved(empty), "xl/worksheets/sheet1.xml")

    def test_a_tables_columns(self, empty: Worksheet, live_escapes_xlsx: Path) -> None:
        """A header XML cannot hold is cleaned as Excel cleans it, in the
        column's name and in its cell."""
        for offset, header in enumerate(HEADERS):
            empty.set_value(CellRef(1, 7 + offset), header)
            empty.set_value(CellRef(2, 7 + offset), offset)
        empty.add_table("Escapes", "G1:L2")
        pattern = r'<tableColumn [^>]*?name="([^"]*)"'
        theirs = re.findall(pattern, part(live_escapes_xlsx, "xl/tables/table1.xml"))
        assert re.findall(pattern, part(saved(empty), "xl/tables/table1.xml")) == theirs
        assert empty.get_value(CellRef.parse("J1")) == "ctl\ufffdx"

    def test_a_page_header(self, empty: Worksheet, answers: Answers) -> None:
        empty.header_footer = HeaderFooter(odd_header=HeaderFooterText(center=said(answers, "page header")))
        assert "<oddHeader>&amp;Chead\r\ner _x005F_x0041_</oddHeader>" in part(saved(empty), "xl/worksheets/sheet1.xml")

    def test_a_sheets_name(self, empty: Worksheet, answers: Answers) -> None:
        empty.workbook.add_sheet(said(answers, "sheet name"))
        assert 'name="a_x005f_x0041_b"' in part(saved(empty), "xl/workbook.xml")

    def test_a_defined_names_comment(self, empty: Worksheet, answers: Answers) -> None:
        empty.workbook.add_defined_name("nm", "1", comment=said(answers, "name comment"))
        assert 'comment="c_x000a__x005f_x0041_"' in part(saved(empty), "xl/workbook.xml")


class TestRoundTrips:
    def test_every_character_below_a_space(self, empty: Worksheet) -> None:
        text = "".join(chr(code) for code in range(1, 32)) + "_x0041_\ufffe\uffff"
        empty.set_value(CellRef.parse("A1"), text)
        empty.set_value(CellRef.parse("A2"), "\r\n\r")
        again = Workbook.open(saved(empty))[0]
        assert again.get_value(CellRef.parse("A1")) == text
        assert again.get_value(CellRef.parse("A2")) == "\r\n\r"

    def test_a_sheet_named_like_an_escape_keeps_its_formulas(self, empty: Worksheet) -> None:
        book = empty.workbook
        other = book.add_sheet("a_x0041_b")
        other.set_value(CellRef.parse("A1"), 7)
        empty.set_formula(CellRef.parse("A1"), "a_x0041_b!A1")
        book.rename_sheet("a_x0041_b", "b_x0042_c")
        assert empty.get_formula(CellRef.parse("A1")) == "b_x0042_c!A1"
        assert book.sheet_names[-1] == "b_x0042_c"

    def test_a_rules_formula_moves_on_a_sheet_named_like_an_escape(self, empty: Worksheet) -> None:
        """A rule holds its formula as it reads and the formula code works
        on it as stored, so the name has to be decoded once, not twice."""
        empty.rename("a_x0041_b")
        empty.set_formula(CellRef.parse("C1"), "a_x0041_b!A1*2")
        empty.add_conditional_format("B1:B3", expression("a_x0041_b!$A$1>0"))
        empty.insert_rows(1, 2)
        (block,) = empty.conditional_formats
        assert block.rules[0].formulas == ("a_x0041_b!$A$3>0",)
        assert empty.get_formula(CellRef.parse("C3")) == "a_x0041_b!A3*2"
