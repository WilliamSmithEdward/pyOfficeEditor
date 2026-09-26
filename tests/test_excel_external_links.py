"""Formulas that read another workbook, calculated from what the link caches.

Measured in Excel: Dest.xlsx read Source.xlsx through the formulas below,
was saved with both open, and was then opened alone, links not updated,
B1 changed to 20 and everything recalculated. The answers below are what
Excel showed then. The package here is built from the parts Excel wrote,
the link's own path left out, so no machine's folders are recorded.
"""

from __future__ import annotations

import io
import zipfile

from pyofficeeditor.excel import CellError, CellValue, Workbook

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

#: The link as Excel wrote it, less the alternate absolute path newer
#: Excel records beside the relative one.
LINK = (
    f'<externalLink xmlns="{MAIN}"><externalBook xmlns:r="{RELS}" r:id="rId1">'
    '<sheetNames><sheetName val="Sheet1"/><sheetName val="Other"/></sheetNames>'
    "<definedNames><definedName name=\"Rate\" refersTo=\"='Sheet1'!$C$1\"/></definedNames>"
    '<sheetDataSet><sheetData sheetId="0">'
    '<row r="1"><cell r="A1"><v>1</v></cell><cell r="B1" t="str"><v>a</v></cell><cell r="C1"><v>10</v></cell>'
    '<cell r="D1" t="str"><v>x</v></cell></row>'
    '<row r="2"><cell r="A2"><v>2</v></cell><cell r="B2" t="str"><v>b</v></cell><cell r="C2"><v>20</v></cell></row>'
    '<row r="3"><cell r="A3"><v>3</v></cell><cell r="B3" t="str"><v>c</v></cell><cell r="C3"><v>30</v></cell></row>'
    '<row r="4"><cell r="A4"><v>4</v></cell><cell r="B4" t="str"><v>d</v></cell><cell r="C4"><v>40</v></cell></row>'
    '<row r="5"><cell r="A5"><v>5</v></cell><cell r="B5" t="str"><v>e</v></cell><cell r="C5"><v>50</v></cell></row>'
    '</sheetData><sheetData sheetId="1"><row r="1"><cell r="A1"><v>7</v></cell></row></sheetData>'
    "</sheetDataSet></externalBook></externalLink>"
)

#: Each formula, in the file's own spelling, and what Excel showed.
FORMULAS: list[tuple[str, CellValue]] = [
    ("[1]Sheet1!A1", 1),
    ("[1]Sheet1!A2*B1", 40),
    ("SUM([1]Sheet1!A1:A5)", 15),
    ('VLOOKUP("c",[1]Sheet1!B1:C5,2,FALSE)', 30),
    ("INDEX([1]Sheet1!A1:A5,3)", 3),
    ("[1]!Rate*100", 1000),
    ("[1]Other!A1*2", 14),
    ("SUM([1]Sheet1:Other!A1)", 8),
    ("[1]Sheet1!A6", 0),
    ("[1]Sheet1!D1", "x"),
    ("ISREF([1]Sheet1!A1)", True),
    ("SUMPRODUCT([1]Sheet1!A1:A5*(B1+1))", 315),
    ("SUM([1]Sheet1!A:A)", 15),
    ("COUNTBLANK([1]Sheet1!A1:A6)", 1),
    ("RANK(3,[1]Sheet1!A1:A5)", 3),
    ("SUBTOTAL(9,[1]Sheet1!A1:A5)", 15),
    ("ROW([1]Sheet1!A3)", 3),
    # A closed workbook's range, which these want as a range of their own.
    ('COUNTIF([1]Sheet1!A1:A5,">2")', CellError("#VALUE!")),
    ('SUMIFS([1]Sheet1!A1:A5,[1]Sheet1!B1:B5,"c")', CellError("#VALUE!")),
    ("DSUM([1]Sheet1!A1:C5,3,[1]Sheet1!A1:A2)", CellError("#VALUE!")),
    ("OFFSET([1]Sheet1!A1,1,0)", CellError("#VALUE!")),
    ('INDIRECT("[Source.xlsx]Sheet1!A1")', CellError("#REF!")),
    ('CELL("address",[1]Sheet1!A2)', CellError("#N/A")),
    ("SHEET([1]Sheet1!A1)", CellError("#N/A")),
    ("ISFORMULA([1]Sheet1!A1)", CellError("#N/A")),
]


def _escaped(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _package() -> bytes:
    cells = "".join(
        f'<row r="{row}"><c r="A{row}"><f>{_escaped(formula)}</f></c>'
        + ('<c r="B1"><v>10</v></c>' if row == 1 else "")
        + "</row>"
        for row, (formula, _) in enumerate(FORMULAS, start=1)
    )
    parts = {
        "[Content_Types].xml": (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/externalLinks/externalLink1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>'
            "</Types>"
        ),
        "_rels/.rels": (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{RELS}/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>"
        ),
        "xl/workbook.xml": (
            f'<workbook xmlns="{MAIN}" xmlns:r="{RELS}"><sheets>'
            '<sheet name="Dest" sheetId="1" r:id="rId1"/></sheets>'
            '<externalReferences><externalReference r:id="rId2"/></externalReferences></workbook>'
        ),
        "xl/_rels/workbook.xml.rels": (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{RELS}/worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{RELS}/externalLink" Target="externalLinks/externalLink1.xml"/>'
            "</Relationships>"
        ),
        "xl/worksheets/sheet1.xml": f'<worksheet xmlns="{MAIN}"><sheetData>{cells}</sheetData></worksheet>',
        "xl/externalLinks/externalLink1.xml": LINK,
        "xl/externalLinks/_rels/externalLink1.xml.rels": (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{RELS}/externalLinkPath" Target="Source.xlsx" TargetMode="External"/>'
            "</Relationships>"
        ),
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as package:
        for name, text in parts.items():
            package.writestr(name, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + text)
    return buffer.getvalue()


def test_a_closed_workbook_is_read_from_its_links_cache() -> None:
    with Workbook.from_bytes(_package()) as book:
        sheet = book["Dest"]
        sheet["B1"] = 20
        report = book.calculate()
        assert report.complete
        got = [sheet[f"A{row}"].value for row in range(1, len(FORMULAS) + 1)]
    assert got == [want for _, want in FORMULAS]
