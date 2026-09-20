# pyvbaharness ships no type stubs and is an optional extra, so strict
# inference cannot see through it here. The suppressions are scoped to this
# file, which is the only one that imports it.
# pyright: reportMissingTypeStubs=false, reportUnknownMemberType=false
# pyright: reportAttributeAccessIssue=false, reportUnknownArgumentType=false
# pyright: reportUnknownVariableType=false, reportMissingParameterType=false
"""The live gate: does Excel accept what this library writes?

Every other test compares bytes against bytes. This one asks the only
authority that matters. It opens a workbook this library wrote in real
Excel, checks that Excel did not have to repair it, and reads the cells back
through Excel's own object model.

It is the strongest available check on things nothing else can prove:

- the package is valid, since Excel repairs or refuses one that is not
- ``fullCalcOnLoad`` works, since a changed input must produce a
  recalculated formula result rather than the stale cache still in the file
- a date written as a serial plus a number format really shows as a date
- a style index paints the cell it was meant to, since a wrong one paints
  the wrong cell rather than erroring

Opt in with ``RUN_LIVE_EXCEL=1``. It needs Windows, Excel, and the ``live``
extra.

**One session, shared.** The harness holds a machine-wide mutex and acquires
it with a zero timeout, so creating a session per test races against the
previous one's teardown and fails intermittently. Office automation is
sequential by contract, so these tests share a single Excel instance: it is
also three times faster. If the lock is genuinely held by other work on the
machine, the fixture skips rather than fails, because that is not a defect in
the code under test.
"""

from __future__ import annotations

import datetime as dt
import os
from collections.abc import Iterator
from pathlib import Path

import pytest

from pyofficeeditor.excel import Alignment, Border, CellError, Workbook

pytestmark = [
    pytest.mark.live,
    pytest.mark.skipif(
        os.environ.get("RUN_LIVE_EXCEL") != "1",
        reason="set RUN_LIVE_EXCEL=1 to drive real Excel",
    ),
]


@pytest.fixture(scope="module")
def excel() -> Iterator[object]:
    """One Excel instance for every test in this module."""
    try:
        from pyvbaharness import ExcelSession, SessionLockHeld
    except ImportError:  # pragma: no cover - depends on the environment
        pytest.skip('pyvbaharness is not installed; pip install -e ".[dev,live]"')

    try:
        session = ExcelSession()
    except SessionLockHeld as exc:  # pragma: no cover - depends on the machine
        pytest.skip(f"another pyvbaharness session holds the Excel lock: {exc}")

    try:
        session.new_workbook()
        yield session
    finally:
        session.close()


def probe(excel: object, source: str, path: Path, name: str) -> dict[str, str]:
    """Run a VBA probe against a file and split its reply into fields.

    Each probe gets its own module name: the session is shared, so injecting
    several sources that all define ``Probe`` under one name would leave VBA
    with an ambiguous procedure.
    """
    result = excel.run_vba(  # type: ignore[attr-defined]
        source, proc="Probe", args=(str(path),), timeout=180, module_name=f"Probe_{name}"
    )
    assert result.outcome == "passed", f"Excel refused the workbook: {result!r}"
    return dict(part.split("=", 1) for part in str(result.value).split("|"))


# --------------------------------------------------------------------------
# Values, dates and recalculation
# --------------------------------------------------------------------------

_VALUE_PROBE = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Data")

    ' Excel recalculates on load when fullCalcOnLoad is set. D2 is =B2*C2
    ' with B2 changed and C2 cleared, so a correct file gives 0 and a stale
    ' cache gives the 510 still written in the bytes.
    parts = "D2=" & CStr(ws.Range("D2").Value)
    parts = parts & "|D6=" & CStr(ws.Range("D6").Value)
    parts = parts & "|A2=" & ws.Range("A2").Value
    parts = parts & "|H1=" & ws.Range("H1").Value
    parts = parts & "|H2=" & CStr(ws.Range("H2").Value)
    parts = parts & "|H3=" & CStr(ws.Range("H3").Value)
    parts = parts & "|H4=" & Format(ws.Range("H4").Value, "yyyy-mm-dd")
    parts = parts & "|H4isdate=" & CStr(IsDate(ws.Range("H4").Value))
    parts = parts & "|H5=" & Format(ws.Range("H5").Value, "yyyy-mm-dd hh:nn")
    parts = parts & "|H7=" & CStr(ws.Range("H7").Text)
    parts = parts & "|A9=[" & ws.Range("A9").Value & "]"
    parts = parts & "|Z50=" & ws.Range("Z50").Value
    parts = parts & "|C2empty=" & CStr(IsEmpty(ws.Range("C2").Value))
    parts = parts & "|Notes=" & wb.Worksheets("Notes").Range("B5").Value
    parts = parts & "|sheets=" & CStr(wb.Worksheets.Count)

    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""


@pytest.fixture(scope="module")
def edited(live_sample_xlsx: Path, tmp_path_factory: pytest.TempPathFactory) -> Path:
    """A workbook this library edited, on disk, ready for Excel."""
    book = Workbook.open(live_sample_xlsx)
    sheet = book["Data"]
    sheet["B2"].value = 200
    sheet["H1"].value = "brand new string"
    sheet["H2"].value = 3.5
    sheet["H3"].value = True
    sheet["H4"].value = dt.date(2026, 7, 4)
    sheet["H5"].value = dt.datetime(2026, 7, 4, 13, 45)
    sheet["H7"].value = CellError("#N/A")
    sheet["A9"].value = "  still padded  "
    sheet["Z50"].value = "far away"
    sheet["C2"].clear()
    book["Notes"]["B5"].value = "on the second sheet"
    target = tmp_path_factory.mktemp("live") / "edited.xlsx"
    book.save(target)
    return target


def test_excel_opens_the_edited_workbook_and_recalculates(excel: object, edited: Path) -> None:
    """The gate. If Excel had to repair the file, opening it raises here."""
    seen = probe(excel, _VALUE_PROBE, edited, "values")

    # C2 was cleared, so =B2*C2 is 200 * empty = 0. The point is that Excel
    # recalculated at all: the file still carries the cached 510.
    assert seen["D2"] == "0", f"D2 was not recalculated; Excel reported {seen['D2']}"
    assert seen["D6"] != "3265", "D6's cached total survived, so nothing recalculated"
    # 0 + 1445 + 522.5 + 787.5. That total only comes out if the shared
    # formula's followers were translated the way this library translates them.
    assert seen["D6"] == "2755", "the shared formula group recalculated as expected"

    assert seen["A2"] == "North", "an existing shared string"
    assert seen["H1"] == "brand new string", "a shared string this library added"
    assert seen["H2"] == "3.5"
    assert seen["H3"] == "True"
    assert seen["H4"] == "2026-07-04"
    assert seen["H4isdate"] == "True", "the number format has to make it a date"
    assert seen["H5"] == "2026-07-04 13:45"
    assert seen["H7"] == "#N/A", "an error value, not the text of one"
    assert seen["A9"] == "[  still padded  ]", "xml:space held the spaces"
    assert seen["Z50"] == "far away"
    assert seen["C2empty"] == "True"
    assert seen["Notes"] == "on the second sheet"
    assert seen["sheets"] == "2"


# --------------------------------------------------------------------------
# Sheets
# --------------------------------------------------------------------------

_SHEET_PROBE = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String
    Dim i As Long

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)

    For i = 1 To wb.Worksheets.Count
        parts = parts & wb.Worksheets(i).Name & ","
    Next i
    parts = "order=" & parts

    Set ws = wb.Worksheets("Summary")
    parts = parts & "|A1=" & ws.Range("A1").Value
    parts = parts & "|B1formula=" & ws.Range("B1").Formula
    parts = parts & "|B1=" & CStr(ws.Range("B1").Value)
    parts = parts & "|renamed=" & wb.Worksheets("Q1 Data").Range("A2").Value
    parts = parts & "|D3=" & wb.Worksheets("Q1 Data").Range("D3").Formula

    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""


def test_excel_accepts_added_renamed_and_reordered_sheets(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """A part this library created from nothing is where Excel is strictest:
    the content type, the relationship and the schema's child order all have
    to be right or it refuses the file."""
    book = Workbook.open(live_sample_xlsx)
    summary = book.add_sheet("Summary")
    summary["A1"].value = "Total"
    summary["B1"].formula = "=SUM(Data!D2:D5)"
    book.add_sheet("Scratch", index=0)
    book.move_sheet("Summary", 1)
    book.rename_sheet("Data", "Q1 Data")
    book.remove_sheet("Scratch")
    target = tmp_path / "sheets.xlsx"
    book.save(target)

    seen = probe(excel, _SHEET_PROBE, target, "sheets")

    assert seen["order"] == "Summary,Q1 Data,Notes,"
    assert seen["A1"] == "Total"
    # Excel spells the rename back with the quoting the new name needs.
    assert seen["B1formula"] == "=SUM('Q1 Data'!D2:D5)"
    assert seen["B1"] == "3265", "the formula computed against the renamed sheet"
    assert seen["renamed"] == "North", "the renamed sheet kept its contents"
    assert seen["D3"] == "=B3*C3", "the shared formula group survived the rename"


# --------------------------------------------------------------------------
# Formatting
# --------------------------------------------------------------------------

_FORMAT_PROBE = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Data")

    parts = "B2bold=" & CStr(ws.Range("B2").Font.Bold)
    parts = parts & "|B2name=" & ws.Range("B2").Font.Name
    parts = parts & "|B2size=" & CStr(ws.Range("B2").Font.Size)
    parts = parts & "|B3fill=" & CStr(ws.Range("B3").Interior.Color)
    parts = parts & "|B3pattern=" & CStr(ws.Range("B3").Interior.Pattern)
    parts = parts & "|B4edge=" & CStr(ws.Range("B4").Borders(xlEdgeLeft).LineStyle)
    parts = parts & "|B4weight=" & CStr(ws.Range("B4").Borders(xlEdgeLeft).Weight)
    parts = parts & "|B4colour=" & CStr(ws.Range("B4").Borders(xlEdgeLeft).Color)
    parts = parts & "|B5halign=" & CStr(ws.Range("B5").HorizontalAlignment)
    parts = parts & "|B5wrap=" & CStr(ws.Range("B5").WrapText)
    parts = parts & "|C3italic=" & CStr(ws.Range("C3").Font.Italic)
    parts = parts & "|C3value=" & CStr(ws.Range("C3").Value)
    parts = parts & "|A1bold=" & CStr(ws.Range("A1").Font.Bold)
    parts = parts & "|A1italic=" & CStr(ws.Range("A1").Font.Italic)
    parts = parts & "|B1italic=" & CStr(ws.Range("B1").Font.Italic)
    parts = parts & "|F2isdate=" & CStr(IsDate(ws.Range("F2").Value))
    parts = parts & "|D6=" & CStr(ws.Range("D6").Value)

    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""


def test_excel_renders_the_formatting_this_library_writes(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """Excel is strict about the style tables, and a wrong index paints the
    wrong cell rather than erroring, so Excel has to be asked what it sees."""
    book = Workbook.open(live_sample_xlsx)
    sheet = book["Data"]
    sheet["B2"].format = sheet["B2"].format.with_font(bold=True)
    sheet["B3"].fill = "FFFF00"
    sheet["B4"].border = Border.all_sides("thin", "FF0000")
    sheet["B5"].alignment = Alignment(horizontal="center", wrap_text=True)
    sheet.range("C2:C5").apply_font(italic=True)
    sheet["A1"].format = sheet["A1"].format.with_font(italic=True)
    target = tmp_path / "formatted.xlsx"
    book.save(target)

    seen = probe(excel, _FORMAT_PROBE, target, "formats")

    assert seen["B2bold"] == "True"
    # The cell had no style of its own, so its format resolved through
    # cellXfs[0]. Had it resolved to an empty format instead, the new font
    # would carry no name or size and Excel would render the default
    # typeface rather than the workbook's.
    assert seen["B2name"] == "Aptos Narrow", "bolding must not change the typeface"
    assert seen["B2size"] == "11"

    # Excel reports Interior.Color as BGR, so FFFF00 comes back as 65535.
    assert seen["B3fill"] == "65535", "a yellow solid fill"
    assert seen["B3pattern"] == "1", "xlSolid"

    assert seen["B4edge"] == "1", "xlContinuous"
    assert seen["B4weight"] == "2", "xlThin"
    assert seen["B4colour"] == "255", "red, as BGR"

    assert seen["B5halign"] == "-4108", "xlCenter"
    assert seen["B5wrap"] == "True"

    assert seen["C3italic"] == "True"
    assert seen["C3value"] == "4.25", "formatting a cell left its value alone"

    # A1 and B1 shared one cellXfs entry. Italicising A1 must not italicise
    # B1, which is what editing the shared entry in place would have done.
    assert seen["A1italic"] == "True"
    assert seen["B1italic"] == "False", "its neighbour shared the entry and is untouched"
    assert seen["A1bold"] == "True", "and A1 kept the bold it already had"

    assert seen["F2isdate"] == "True", "the existing date format survived"
    assert seen["D6"] == "3265", "and the formulas still compute"


# --------------------------------------------------------------------------
# Merged ranges
# --------------------------------------------------------------------------

_MERGE_PROBE = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Data")

    parts = "A20merged=" & CStr(ws.Range("A20").MergeCells)
    parts = parts & "|A20area=" & ws.Range("A20").MergeArea.Address(False, False)
    parts = parts & "|A20value=" & ws.Range("A20").Value
    parts = parts & "|B20empty=" & CStr(IsEmpty(ws.Range("B20").Value))
    parts = parts & "|H1area=" & ws.Range("H1").MergeArea.Address(False, False)
    parts = parts & "|A11merged=" & CStr(ws.Range("A11").MergeCells)
    parts = parts & "|D11merged=" & CStr(ws.Range("D11").MergeCells)
    ' Range.MergeCells over a mixed block returns Null, which CStr refuses,
    ' so ask cell by cell rather than over UsedRange.
    parts = parts & "|A11area=" & ws.Range("A11").MergeArea.Address(False, False)

    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""


def test_excel_renders_the_merges_this_library_writes(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """Overlapping merges make Excel repair a worksheet, so the only way to
    know a merge is well formed is to let Excel open it."""
    book = Workbook.open(live_sample_xlsx)
    sheet = book["Data"]
    sheet["A20"].value = "wide heading"
    sheet["B20"].value = "this is discarded"
    sheet.merge("A20:C20")
    sheet["H1"].value = "tall"
    sheet.merge("H1:H4")
    target = tmp_path / "merged.xlsx"
    book.save(target)

    seen = probe(excel, _MERGE_PROBE, target, "merges")

    assert seen["A20merged"] == "True"
    assert seen["A20area"] == "A20:C20", "Excel agrees on the extent"
    assert seen["A20value"] == "wide heading", "the anchor kept its value"
    assert seen["B20empty"] == "True", "the covered value was discarded"
    assert seen["H1area"] == "H1:H4", "a vertical merge too"
    assert seen["A11merged"] == "True", "the fixture's own merge survived"
    assert seen["A11area"] == "A11:C11", "at its original extent"
    assert seen["D11merged"] == "False", "and a cell outside it is not merged"


# --------------------------------------------------------------------------
# Tables
# --------------------------------------------------------------------------

_TABLE_PROBE = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim parts As String
    Dim i As Long

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Tabled")

    parts = "count=" & CStr(ws.ListObjects.Count)
    For i = 1 To ws.ListObjects.Count
        parts = parts & "," & ws.ListObjects(i).Name
    Next i

    Set lo = ws.ListObjects("NewTable")
    parts = parts & "|newRange=" & lo.Range.Address(False, False)
    parts = parts & "|newHeaders=" & lo.HeaderRowRange.Address(False, False)
    parts = parts & "|newData=" & lo.DataBodyRange.Address(False, False)
    parts = parts & "|newStyle=" & lo.TableStyle
    parts = parts & "|newCols="
    For i = 1 To lo.ListColumns.Count
        parts = parts & lo.ListColumns(i).Name & ";"
    Next i
    parts = parts & "|newRows=" & CStr(lo.ListRows.Count)
    ' A table name is a defined name, so a formula must resolve through it.
    parts = parts & "|structured=" & CStr(Application.Evaluate("SUM(NewTable[Qty])"))

    ' The fixture's own table, with its totals row, must be untouched.
    Set lo = ws.ListObjects("SalesTable")
    parts = parts & "|salesRange=" & lo.Range.Address(False, False)
    parts = parts & "|salesTotals=" & CStr(lo.ShowTotals)
    parts = parts & "|salesTotal=" & CStr(lo.TotalsRowRange.Cells(1, 2).Value)

    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""


def test_excel_accepts_a_table_this_library_created(
    excel: object, live_structures_xlsx: Path, tmp_path: Path
) -> None:
    """A table is four wired-together pieces, and Excel ignores or refuses a
    table whose wiring is wrong, so only Excel can confirm it."""
    from pyofficeeditor.excel._tables import TableStyle

    book = Workbook.open(live_structures_xlsx)
    sheet = book["Tabled"]
    sheet["A10"].value = "Product"
    sheet["B10"].value = "Qty"
    sheet["A11"].value = "widget"
    sheet["B11"].value = 5
    sheet["A12"].value = "gadget"
    sheet["B12"].value = 9
    sheet.add_table("NewTable", "A10:B12", style=TableStyle(name="TableStyleLight9"))
    target = tmp_path / "tabled.xlsx"
    book.save(target)

    seen = probe(excel, _TABLE_PROBE, target, "tables")

    assert seen["count"].startswith("2"), f"Excel sees {seen['count']}"
    assert "NewTable" in seen["count"] and "SalesTable" in seen["count"]

    assert seen["newRange"] == "A10:B12"
    assert seen["newHeaders"] == "A10:B10"
    assert seen["newData"] == "A11:B12"
    assert seen["newStyle"] == "TableStyleLight9"
    assert seen["newCols"] == "Product;Qty;"
    assert seen["newRows"] == "2"
    # A structured reference only resolves if Excel registered the name.
    assert seen["structured"] == "14", "SUM(NewTable[Qty]) is 5 + 9"

    assert seen["salesRange"] == "A1:C5", "the fixture's table is untouched"
    assert seen["salesTotals"] == "True"
    assert seen["salesTotal"] == "555", "its totals row still sums 120+340+95"


# --------------------------------------------------------------------------
# The no-op case
# --------------------------------------------------------------------------

_OPEN_PROBE = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Probe = "D2=" & CStr(wb.Worksheets("Data").Range("D2").Value) & _
            "|A2=" & wb.Worksheets("Data").Range("A2").Value
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Function
"""


def test_excel_opens_an_untouched_workbook(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """A no-op save must produce a file Excel is equally happy with."""
    book = Workbook.open(live_sample_xlsx)
    target = tmp_path / "untouched.xlsx"
    target.write_bytes(book.to_bytes())
    assert target.read_bytes() == live_sample_xlsx.read_bytes(), "a no-op save changed bytes"

    seen = probe(excel, _OPEN_PROBE, target, "untouched")
    assert seen["D2"] == "510", "nothing changed, so the cached value stands"
    assert seen["A2"] == "North"
