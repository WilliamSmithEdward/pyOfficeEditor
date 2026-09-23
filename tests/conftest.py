"""Shared fixtures.

The committed packages were authored by Excel; the generated ones are
authored by openpyxl, which writes the same parts a different legal way.
Testing against both is the point: a reader that only ever sees one
producer's output encodes that producer's habits as rules.
"""

from __future__ import annotations

import io
import re
import shutil
import zipfile
from pathlib import Path

import pytest

FIXTURES = Path(__file__).parent / "fixtures"
EXCEL_FIXTURES = FIXTURES / "excel"

#: Authored by Excel: growth hints, mc:Ignorable, a CRLF after the
#: declaration.  The byte-fidelity gate.
MINIMAL_XLSM = EXCEL_FIXTURES / "excel_authored_minimal.xlsm"
#: Authored by Excel: 32 parts and a wider relationship graph.
POWERQUERY_XLSX = EXCEL_FIXTURES / "excel_authored_powerquery.xlsx"
#: Authored by Excel: the binary workbook format, whose worksheets are
#: .bin parts rather than XML.  Proves the container layer does not assume
#: its members are text.
BINARY_XLSB = EXCEL_FIXTURES / "excel_authored_binary.xlsb"
#: Authored by real Excel through pyvbaharness, on demand.  Tests that need
#: these skip when scripts/build_excel_fixtures.py has not been run.
LIVE_EMPTY_XLSX = EXCEL_FIXTURES / "empty.xlsx"
LIVE_SAMPLE_XLSX = EXCEL_FIXTURES / "sample.xlsx"
#: Authored by Excel: two ListObjects, one with a totals row and a calculated
#: column, plus defined names at both scopes, column widths and a hyperlink.
LIVE_STRUCTURES_XLSX = EXCEL_FIXTURES / "structures.xlsx"
#: Authored by Excel: every element that used to make an insertion refuse.
#: Data validation with formulas, a protected range, a saved sort,
#: scenarios, a shape in a drawing part, a form control and a comment,
#: the last two anchored through VML.
LIVE_REFUSED_XLSX = EXCEL_FIXTURES / "refused.xlsx"
#: Authored by Excel: protection with and without a password, a tab colour,
#: view settings, grouping, page setup and three visibility states.
LIVE_SETTINGS_XLSX = EXCEL_FIXTURES / "settings.xlsx"
#: Authored by Excel: hyperlinks of every kind, an autofilter with two
#: sorts of criteria, and outline grouping on rows and columns.
LIVE_LINKS_XLSX = EXCEL_FIXTURES / "links.xlsx"
#: Authored by Excel: one of every shape a sheet can hold, and beside it
#: what Excel then said about each through its own object model.
LIVE_SHAPES_XLSM = EXCEL_FIXTURES / "shapes.xlsm"
LIVE_SHAPES_ANSWERS = EXCEL_FIXTURES / "shapes_answers.json"
#: Authored by Excel: one shape of each MsoAutoShapeType, so the preset
#: geometry table is held to what Excel actually wrote for each number.
LIVE_GEOMETRY_XLSX = EXCEL_FIXTURES / "geometry.xlsx"
#: Authored by Excel: one of every Forms control that has something to
#: say, wired to cells and ranges, and beside it what Excel's own object
#: model answered for each. ``shapes.xlsm`` carries a Button and nothing
#: else, so it can prove nothing about a linked cell or a value.
LIVE_CONTROLS_XLSM = EXCEL_FIXTURES / "controls.xlsm"
LIVE_CONTROLS_ANSWERS = EXCEL_FIXTURES / "controls_answers.json"
#: Authored by Excel: one autofilter criterion per sheet, with what Excel
#: answered for each and how many rows it hid. The hidden counts are the
#: point: Excel stores a filter twice and does not recompute it on open,
#: so they are what an evaluator here has to reproduce.
LIVE_FILTERS_XLSX = EXCEL_FIXTURES / "filters.xlsx"
LIVE_FILTERS_ANSWERS = EXCEL_FIXTURES / "filters_answers.json"
#: Authored by Excel: notes of every shape, one sharing its VML part with a
#: form control, and what Excel's object model said of each.
LIVE_COMMENTS_XLSX = EXCEL_FIXTURES / "comments.xlsx"
LIVE_COMMENTS_ANSWERS = EXCEL_FIXTURES / "comments_answers.json"
#: Authored by Excel: pictures put in at their own size, at 144 dots to the
#: inch, stretched, and a GIF, the same image twice among them, and what
#: Excel's object model said of each.
LIVE_PICTURES_XLSX = EXCEL_FIXTURES / "pictures.xlsx"
LIVE_PICTURES_ANSWERS = EXCEL_FIXTURES / "pictures_answers.json"
#: Authored by Excel: text in more than one font in a cell, made through
#: ``Characters``, and the font Excel reported at each change.
LIVE_RICHTEXT_XLSX = EXCEL_FIXTURES / "richtext.xlsx"
LIVE_RICHTEXT_ANSWERS = EXCEL_FIXTURES / "richtext_answers.json"
#: Authored by Excel: charts on a data sheet, on another sheet and on a chart
#: sheet, with each chart's references before any edit and after each of
#: nine, read back from the file Excel saved after it.
LIVE_CHARTS_XLSX = EXCEL_FIXTURES / "charts.xlsx"
LIVE_CHARTS_ANSWERS = EXCEL_FIXTURES / "charts_answers.json"
#: Authored by Excel: every built-in cell style applied to a cell of its own,
#: a style of the workbook's own, and styles put over existing formatting,
#: with what Excel showed for each cell and what each style includes.
LIVE_STYLES_XLSX = EXCEL_FIXTURES / "styles.xlsx"
LIVE_STYLES_ANSWERS = EXCEL_FIXTURES / "styles_answers.json"
#: Authored by Excel: text XML cannot carry as it is, control characters,
#: a lone carriage return and a literal "_x0041_", everywhere a workbook
#: keeps text, and the characters Excel reported for each piece.
LIVE_ESCAPES_XLSX = EXCEL_FIXTURES / "escapes.xlsx"
LIVE_ESCAPES_ANSWERS = EXCEL_FIXTURES / "escapes_answers.json"
#: Rich text as other writers leave it, which Excel never writes itself:
#: runs with no font after one that has a font, fonts written in part, an
#: empty ``<rPr/>``, and runs held inline in the cell. The cells are in red
#: Courier New 16 and the workbook's default font is green Arial 10, so
#: which font a run shows in cannot be mistaken. A4 alone is in the default.
FOREIGN_RICH_TEXT = {
    "A1": "<si><r><t>a</t></r><r><t>b</t></r><r><rPr><b/></rPr><t>c</t></r><r><t>d</t></r></si>",
    "A2": "<si><r><rPr><b/></rPr><t>x</t></r><r><t>y</t></r></si>",
    "A3": '<si><r><t>p</t></r><r><rPr><sz val="20"/></rPr><t>q</t></r></si>',
    "A4": "<si><r><t>m</t></r><r><rPr><i/></rPr><t>n</t></r></si>",
    "A5": '<si><r><t>e</t></r><r><rPr><rFont val="Times New Roman"/></rPr><t>f</t></r></si>',
    "A6": "<si><r><t>j</t></r><r><rPr/><t>k</t></r></si>",
    "A7": "<si><r><rPr><b/></rPr><t>u</t></r></si>",
    "A8": "<si><r><rPr><b/></rPr><t>x</t></r><r><rPr/><t>y</t></r></si>",
    "A9": (
        '<si><r><t>a</t></r><r><rPr><b/><sz val="16"/><color rgb="FFFF0000"/><rFont val="Courier New"/>'
        '<family val="3"/></rPr><t>b</t></r><r><t>c</t></r></si>'
    ),
    "A10": '<si><r><t>n</t></r><r><rPr><i/><color theme="1"/></rPr><t>o</t></r></si>',
    "A11": "<is><r><rPr><b/></rPr><t>x</t></r><r><t>y</t></r></is>",
    "A12": "<is><r><t>a</t></r><r><rPr><i/></rPr><t>b</t></r></is>",
}
#: The default font those cells are measured against.
FOREIGN_DEFAULT_FONT = '<font><sz val="10"/><color rgb="FF00B050"/><name val="Arial"/><family val="2"/></font>'
#: Measured by Excel: the text it shows for values under some five hundred
#: format codes in both date systems, and what each builtin format id means.
NUMBER_FORMATS_JSON = EXCEL_FIXTURES / "number_formats.json"


@pytest.fixture(scope="session")
def minimal_xlsm_bytes() -> bytes:
    return MINIMAL_XLSM.read_bytes()


@pytest.fixture(scope="session")
def powerquery_xlsx_bytes() -> bytes:
    return POWERQUERY_XLSX.read_bytes()


@pytest.fixture(scope="session")
def binary_xlsb_bytes() -> bytes:
    return BINARY_XLSB.read_bytes()


@pytest.fixture(scope="session")
def excel_authored_packages() -> list[tuple[str, bytes]]:
    """Every committed Excel-authored package, for the fidelity gates."""
    return [(p.name, p.read_bytes()) for p in (MINIMAL_XLSM, POWERQUERY_XLSX, BINARY_XLSB)]


@pytest.fixture(scope="session")
def openpyxl_xlsx_bytes() -> bytes:
    """A workbook openpyxl wrote: a sharedStrings part, formula cells, and
    none of Excel's growth hints."""
    openpyxl = pytest.importorskip("openpyxl", reason="the dev extra provides openpyxl")
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Data"
    sheet["A1"] = "Region"
    sheet["B1"] = "Units"
    sheet["A2"] = "North"
    sheet["B2"] = 120
    sheet["A3"] = "South"
    sheet["B3"] = 340
    sheet["C2"] = "=B2*2"
    sheet["A5"] = "  padded  "
    sheet["A6"] = "amp & lt < gt >"
    workbook.create_sheet("Notes")["A1"] = "Second sheet"
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


@pytest.fixture(scope="session")
def data_descriptor_zip_bytes() -> bytes:
    """An archive whose members carry data descriptors.

    CPython's zipfile writes them when the output stream cannot seek, which
    is the shape a streaming producer emits: the local header zeroes the
    CRC and both sizes and only the central directory states the truth.
    """

    class Unseekable(io.RawIOBase):
        def __init__(self) -> None:
            self.buffer = bytearray()

        def writable(self) -> bool:
            return True

        def write(self, data: object) -> int:  # pyright: ignore[reportIncompatibleMethodOverride]
            payload = bytes(data)  # type: ignore[arg-type]
            self.buffer += payload
            return len(payload)

        def seekable(self) -> bool:
            return False

    sink = Unseekable()
    with zipfile.ZipFile(sink, "w", zipfile.ZIP_DEFLATED) as archive:  # type: ignore[arg-type]
        archive.writestr("[Content_Types].xml", b'<?xml version="1.0"?><Types/>')
        archive.writestr("a.xml", b"<a/>" * 100)
        archive.writestr("b.xml", b"<b/>" * 100)
    return bytes(sink.buffer)


@pytest.fixture(scope="session")
def live_empty_xlsx() -> Path:
    """The path to the workbook Excel saves when nothing is put in it."""
    if not LIVE_EMPTY_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author empty.xlsx with real Excel")
    return LIVE_EMPTY_XLSX


@pytest.fixture(scope="session")
def live_sample_xlsx() -> Path:
    """The path to the Excel-authored sample, or a skip if it is absent.

    Session-scoped because it is a constant, and because the live gate's
    module-scoped fixtures request it: a narrower scope cannot be consumed by
    a wider one.
    """
    if not LIVE_SAMPLE_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author sample.xlsx with real Excel")
    return LIVE_SAMPLE_XLSX


@pytest.fixture(scope="session")
def live_structures_xlsx() -> Path:
    """The Excel-authored workbook carrying tables and defined names."""
    if not LIVE_STRUCTURES_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author structures.xlsx with real Excel")
    return LIVE_STRUCTURES_XLSX


@pytest.fixture(scope="session")
def live_refused_xlsx() -> Path:
    """The Excel-authored workbook carrying every once-refused element."""
    if not LIVE_REFUSED_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author refused.xlsx with real Excel")
    return LIVE_REFUSED_XLSX


@pytest.fixture(scope="session")
def live_settings_xlsx() -> Path:
    """The Excel-authored workbook carrying the sheet settings."""
    if not LIVE_SETTINGS_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author settings.xlsx with real Excel")
    return LIVE_SETTINGS_XLSX


@pytest.fixture(scope="session")
def live_links_xlsx() -> Path:
    """The Excel-authored workbook carrying links and grouping."""
    if not LIVE_LINKS_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author links.xlsx with real Excel")
    return LIVE_LINKS_XLSX


@pytest.fixture(scope="session")
def live_shapes_xlsm() -> Path:
    """The Excel-authored workbook carrying one of every shape."""
    if not LIVE_SHAPES_XLSM.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author shapes.xlsm with real Excel")
    return LIVE_SHAPES_XLSM


@pytest.fixture(scope="session")
def live_shapes_answers() -> Path:
    """What Excel said about those shapes, measured beside the workbook."""
    if not LIVE_SHAPES_ANSWERS.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to measure the shapes with real Excel")
    return LIVE_SHAPES_ANSWERS


@pytest.fixture(scope="session")
def live_geometry_xlsx() -> Path:
    """One shape of each MsoAutoShapeType, authored by Excel."""
    if not LIVE_GEOMETRY_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author geometry.xlsx with real Excel")
    return LIVE_GEOMETRY_XLSX


@pytest.fixture(scope="session")
def live_controls_xlsm() -> Path:
    """The Excel-authored workbook carrying wired-up Forms controls."""
    if not LIVE_CONTROLS_XLSM.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author controls.xlsm with real Excel")
    return LIVE_CONTROLS_XLSM


@pytest.fixture(scope="session")
def live_controls_answers() -> Path:
    """What Excel said about those controls, measured beside the workbook."""
    if not LIVE_CONTROLS_ANSWERS.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to measure the controls with real Excel")
    return LIVE_CONTROLS_ANSWERS


@pytest.fixture(scope="session")
def live_filters_xlsx() -> Path:
    """One autofilter criterion per sheet, authored by Excel."""
    if not LIVE_FILTERS_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author filters.xlsx with real Excel")
    return LIVE_FILTERS_XLSX


@pytest.fixture(scope="session")
def live_filters_answers() -> Path:
    """What Excel said about those filters, and how many rows it hid."""
    if not LIVE_FILTERS_ANSWERS.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to measure the filters with real Excel")
    return LIVE_FILTERS_ANSWERS


@pytest.fixture(scope="session")
def live_comments_xlsx() -> Path:
    """Notes authored by Excel."""
    if not LIVE_COMMENTS_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author comments.xlsx with real Excel")
    return LIVE_COMMENTS_XLSX


@pytest.fixture(scope="session")
def live_comments_answers() -> Path:
    """What Excel's object model said of those notes."""
    if not LIVE_COMMENTS_ANSWERS.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to measure the notes with real Excel")
    return LIVE_COMMENTS_ANSWERS


@pytest.fixture(scope="session")
def live_pictures_xlsx() -> Path:
    """Pictures authored by Excel."""
    if not LIVE_PICTURES_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author pictures.xlsx with real Excel")
    return LIVE_PICTURES_XLSX


@pytest.fixture(scope="session")
def live_pictures_answers() -> Path:
    """What Excel's object model said of those pictures."""
    if not LIVE_PICTURES_ANSWERS.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to measure the pictures with real Excel")
    return LIVE_PICTURES_ANSWERS


@pytest.fixture(scope="session")
def live_richtext_xlsx() -> Path:
    """Text in several fonts, authored by Excel."""
    if not LIVE_RICHTEXT_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author richtext.xlsx with real Excel")
    return LIVE_RICHTEXT_XLSX


@pytest.fixture(scope="session")
def live_richtext_answers() -> Path:
    """The font Excel reported at each change in that text."""
    if not LIVE_RICHTEXT_ANSWERS.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to measure the text with real Excel")
    return LIVE_RICHTEXT_ANSWERS


@pytest.fixture(scope="session")
def live_charts_xlsx() -> Path:
    """Charts, authored by Excel."""
    if not LIVE_CHARTS_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author charts.xlsx with real Excel")
    return LIVE_CHARTS_XLSX


@pytest.fixture(scope="session")
def live_charts_answers() -> Path:
    """Each chart's references as Excel read them, before and after each edit."""
    if not LIVE_CHARTS_ANSWERS.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to measure the charts with real Excel")
    return LIVE_CHARTS_ANSWERS


@pytest.fixture(scope="session")
def live_styles_xlsx() -> Path:
    """Cell styles, authored by Excel."""
    if not LIVE_STYLES_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author styles.xlsx with real Excel")
    return LIVE_STYLES_XLSX


@pytest.fixture(scope="session")
def live_styles_answers() -> Path:
    """What Excel showed for each styled cell, and what each style includes."""
    if not LIVE_STYLES_ANSWERS.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to measure the styles with real Excel")
    return LIVE_STYLES_ANSWERS


@pytest.fixture(scope="session")
def live_escapes_xlsx() -> Path:
    """Text XML cannot carry as it is, authored by Excel."""
    if not LIVE_ESCAPES_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author escapes.xlsx with real Excel")
    return LIVE_ESCAPES_XLSX


@pytest.fixture(scope="session")
def live_escapes_answers() -> Path:
    """The characters Excel reported for each piece of that text."""
    if not LIVE_ESCAPES_ANSWERS.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to measure the text with real Excel")
    return LIVE_ESCAPES_ANSWERS


@pytest.fixture(scope="session")
def foreign_rich_text(live_empty_xlsx: Path, tmp_path_factory: pytest.TempPathFactory) -> Path:
    """A workbook holding :data:`FOREIGN_RICH_TEXT`.

    Built from ``empty.xlsx``: this library formats the cells and fills
    them with markers, then the markers' entries, the default font, and
    the two inline cells are put in place by hand, as another writer would
    have written them.
    """
    from pyofficeeditor.excel import CellFormat, CellRef, Color, Font, Workbook

    base = tmp_path_factory.mktemp("foreign") / "base.xlsx"
    shutil.copy(live_empty_xlsx, base)
    courier = CellFormat(font=Font(name="Courier New", size=16, color=Color(rgb="FFFF0000"), family=3))
    with Workbook.open(base) as book:
        sheet = book[0]
        for address in FOREIGN_RICH_TEXT:
            if address != "A4":
                sheet.set_format(CellRef.parse(address), courier)
            sheet.set_value(CellRef.parse(address), f"@{address}@")
        book.save()

    def by_hand(name: str, text: str) -> str:
        if name == "xl/styles.xml":
            text, count = re.subn(r"<font>.*?</font>", FOREIGN_DEFAULT_FONT, text, count=1, flags=re.S)
            assert count == 1, "the default font was not replaced"
            return text
        for address, markup in FOREIGN_RICH_TEXT.items():
            if name == "xl/sharedStrings.xml" and markup.startswith("<si>"):
                text, count = re.subn(re.escape(f"<si><t>@{address}@</t></si>"), markup, text)
            elif name == "xl/worksheets/sheet1.xml" and markup.startswith("<is>"):
                pattern = rf'<c r="{address}"( s="\d+")? t="s"><v>\d+</v></c>'
                text, count = re.subn(pattern, rf'<c r="{address}"\1 t="inlineStr">{markup}</c>', text)
            else:
                continue
            assert count == 1, f"{address} was not put in place in {name}"
        return text

    target = base.with_name("foreign.xlsx")
    rewritten = {"xl/sharedStrings.xml", "xl/worksheets/sheet1.xml", "xl/styles.xml"}
    with zipfile.ZipFile(base) as source, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as built:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename in rewritten:
                data = by_hand(item.filename, data.decode("utf-8")).encode("utf-8")
            built.writestr(item, data)
    return target


@pytest.fixture(scope="session")
def number_formats_json() -> Path:
    """What Excel showed for values under number formats, measured."""
    if not NUMBER_FORMATS_JSON.is_file():
        pytest.skip("run scripts/measure_number_formats.py to measure the formats with real Excel")
    return NUMBER_FORMATS_JSON
