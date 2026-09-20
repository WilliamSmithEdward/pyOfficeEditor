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
# Dimensions and frozen panes
# --------------------------------------------------------------------------

_DIMENSION_PROBE = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Tabled")

    ' The width Excel reports back for what we stored, which is the number
    ' a person would type rather than the one in the file.
    parts = "Bwidth=" & CStr(ws.Columns("B").ColumnWidth)
    parts = parts & "|Awidth=" & CStr(ws.Columns("A").ColumnWidth)
    parts = parts & "|Dhidden=" & CStr(ws.Columns("D").Hidden)
    parts = parts & "|Chidden=" & CStr(ws.Columns("C").Hidden)
    parts = parts & "|row3=" & CStr(ws.Rows(3).RowHeight)
    parts = parts & "|row5hidden=" & CStr(ws.Rows(5).Hidden)
    parts = parts & "|row4hidden=" & CStr(ws.Rows(4).Hidden)

    ' Freezing is the fiddly one: the wrong activePane leaves the cursor in
    ' a pane the user cannot see.
    '
    ' Not SplitRow and SplitColumn. A workbook Excel froze itself reports 0
    ' for both through this route, so they discriminate nothing. Panes.Count
    ' and where the scrolling area starts do.
    ws.Activate
    parts = parts & "|frozen=" & CStr(ActiveWindow.FreezePanes)
    parts = parts & "|panes=" & CStr(ActiveWindow.Panes.Count)
    parts = parts & "|visibleTop=" & ActiveWindow.VisibleRange.Cells(1, 1).Address(False, False)

    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""


def test_excel_applies_the_dimensions_this_library_writes(
    excel: object, live_structures_xlsx: Path, tmp_path: Path
) -> None:
    """A width without customWidth is ignored, and a freeze with the wrong
    activePane leaves the cursor somewhere invisible. Only Excel can say."""
    book = Workbook.open(live_structures_xlsx)
    sheet = book["Tabled"]
    # Copy the width Excel itself stored for column A onto column B, which
    # is the one operation the units make exact.
    sheet.set_column_width(2, sheet.column_width(1) or 20.0)
    sheet.set_column_hidden(4, True)
    sheet.set_row_height(3, 40)
    sheet.set_row_hidden(5, True)
    sheet.freeze_panes("B2")
    target = tmp_path / "dimensions.xlsx"
    book.save(target)

    seen = probe(excel, _DIMENSION_PROBE, target, "dimensions")

    assert seen["Bwidth"] == seen["Awidth"], "a copied width reproduces exactly"
    assert seen["Dhidden"] == "True"
    assert seen["Chidden"] == "False", "isolating D did not hide its neighbour"
    assert seen["row3"] == "40", "row heights are points, and exact"
    assert seen["row5hidden"] == "True"
    assert seen["row4hidden"] == "False"

    assert seen["frozen"] == "True"
    assert seen["panes"] == "4", "frozen on both axes makes four panes"
    assert seen["visibleTop"] == "B2", "row 1 and column A are pinned above and left"


# --------------------------------------------------------------------------
# Defined names
# --------------------------------------------------------------------------

_NAME_PROBE = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim parts As String

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)

    ' A name Excel accepts resolves in a formula. One whose scope index is
    ' wrong resolves to the wrong sheet, or not at all.
    parts = "added=" & CStr(Application.Evaluate(wb.Names("Alpha").RefersTo))
    parts = parts & "|workbookScope=" & CStr(wb.Names("TotalUnits").RefersTo)
    parts = parts & "|sum=" & CStr(Application.Evaluate("SUM(TotalUnits)"))

    ' A sheet-scoped name is reached through the sheet, and its index has to
    ' still point at that sheet after the move.
    parts = parts & "|scoped=" & CStr(wb.Worksheets("Tabled").Names("LocalRegion").RefersTo)
    parts = parts & "|order="
    Dim i As Long
    For i = 1 To wb.Worksheets.Count
        parts = parts & wb.Worksheets(i).Name & ","
    Next i

    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""


def test_excel_resolves_the_defined_names_this_library_writes(
    excel: object, live_structures_xlsx: Path, tmp_path: Path
) -> None:
    """A sheet-scoped name stores its sheet as a position, so moving a sheet
    silently rescopes it unless the index moves too. Excel is the only thing
    that can say which sheet a name actually reached."""
    book = Workbook.open(live_structures_xlsx)
    book.add_defined_name("Alpha", "Tabled!$B$2")
    book.move_sheet("Tabled", 1)
    target = tmp_path / "named.xlsx"
    book.save(target)

    seen = probe(excel, _NAME_PROBE, target, "names")

    assert seen["order"] == "Second,Tabled,", "the move took"
    assert seen["added"] == "120", "Tabled!B2 holds 120"
    assert seen["workbookScope"] == "=Tabled!$B$2:$B$4"
    assert seen["sum"] == "555", "120 + 340 + 95"
    # The scope index had to follow the sheet from position 0 to position 1.
    assert seen["scoped"] == "=Tabled!$A$2", "still scoped to the sheet it named"


# --------------------------------------------------------------------------
# Inserting rows and columns
# --------------------------------------------------------------------------

_INSERT_PROBE = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Data")

    ' The values moved down two rows, and the rows opened up are blank.
    parts = "B2=" & CStr(ws.Range("B2").Value)
    parts = parts & "|A3blank=" & CStr(IsEmpty(ws.Range("A3").Value))
    parts = parts & "|A4blank=" & CStr(IsEmpty(ws.Range("A4").Value))
    parts = parts & "|A5=" & ws.Range("A5").Value
    parts = parts & "|B5=" & CStr(ws.Range("B5").Value)

    ' The formulas moved and their ranges grew across the insertion. These
    ' are computed values, so Excel had to accept and evaluate them.
    parts = parts & "|D8formula=" & ws.Range("D8").Formula
    parts = parts & "|D8=" & CStr(ws.Range("D8").Value)
    parts = parts & "|D5formula=" & ws.Range("D5").Formula
    parts = parts & "|D5=" & CStr(ws.Range("D5").Value)

    ' The merge moved with its row.
    parts = parts & "|merge=" & ws.Range("A13").MergeArea.Address(False, False)

    ' A formula on another sheet that reads from this one, and a defined
    ' name pointing into it, both had to follow.
    parts = parts & "|summary=" & wb.Worksheets("Summary").Range("A1").Formula
    parts = parts & "|summaryValue=" & CStr(wb.Worksheets("Summary").Range("A1").Value)
    parts = parts & "|above=" & wb.Worksheets("Summary").Range("A2").Formula
    parts = parts & "|name=" & wb.Names("Totals").RefersTo
    parts = parts & "|nameSum=" & CStr(Application.Evaluate("SUM(Totals)"))

    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""


def test_excel_accepts_inserted_rows_and_follows_every_reference(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """The riskiest operation here. A cell's address is written into the file
    in a dozen places; missing one gives a workbook that opens cleanly and is
    wrong, which no byte comparison can catch."""
    book = Workbook.open(live_sample_xlsx)
    summary = book.add_sheet("Summary")
    summary["A1"].formula = "=SUM(Data!D2:D5)"
    summary["A2"].formula = "=Data!B2"
    book.add_defined_name("Totals", "Data!$D$2:$D$5")

    book["Data"].insert_rows(3, 2)
    target = tmp_path / "inserted.xlsx"
    book.save(target)

    seen = probe(excel, _INSERT_PROBE, target, "insert")

    assert seen["B2"] == "120", "the row above the insertion did not move"
    assert seen["A3blank"] == "True", "two blank rows opened up"
    assert seen["A4blank"] == "True"
    assert seen["A5"] == "South", "what was in row 3 is now in row 5"
    assert seen["B5"] == "340"

    # SUM(D2:D5) spanned the insertion, so it grew rather than sliding.
    assert seen["D8formula"] == "=SUM(D2:D7)"
    assert seen["D8"] == "3265", "and still totals the same four rows"
    # D3's shared formula became D5's, reading the row its inputs moved to.
    assert seen["D5formula"] == "=B5*C5"
    assert seen["D5"] == "1445", "340 * 4.25"

    assert seen["merge"] == "A13:C13", "the merge moved down two rows"

    assert seen["summary"] == "=SUM(Data!D2:D7)", "a formula on another sheet followed"
    assert seen["summaryValue"] == "3265"
    assert seen["above"] == "=Data!B2", "and one pointing above the insertion did not"

    assert seen["name"] == "=Data!$D$2:$D$7", "the defined name followed too"
    assert seen["nameSum"] == "3265"


def test_excel_accepts_inserted_columns(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    book = Workbook.open(live_sample_xlsx)
    book["Data"].insert_columns(2, 1)
    target = tmp_path / "inserted_columns.xlsx"
    book.save(target)

    source = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Data")
    parts = "A1=" & ws.Range("A1").Value
    parts = parts & "|B1blank=" & CStr(IsEmpty(ws.Range("B1").Value))
    parts = parts & "|C1=" & ws.Range("C1").Value
    parts = parts & "|E2formula=" & ws.Range("E2").Formula
    parts = parts & "|E2=" & CStr(ws.Range("E2").Value)
    parts = parts & "|E6=" & CStr(ws.Range("E6").Value)
    parts = parts & "|merge=" & ws.Range("A11").MergeArea.Address(False, False)
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""
    seen = probe(excel, source, target, "insertcols")

    assert seen["A1"] == "Region", "the column before the insertion stayed"
    assert seen["B1blank"] == "True", "a blank column opened up"
    assert seen["C1"] == "Units", "what was in B is now in C"
    assert seen["E2formula"] == "=C2*D2", "the formula's inputs moved with it"
    assert seen["E2"] == "510", "and it still computes the same answer"
    assert seen["E6"] == "3265"
    assert seen["merge"] == "A11:D11", "the merge grew across the insertion"


# --------------------------------------------------------------------------
# Deleting rows and columns
# --------------------------------------------------------------------------

_DELETE_PROBE = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Data")

    parts = "A2=" & ws.Range("A2").Value
    parts = parts & "|A3=" & ws.Range("A3").Value
    parts = parts & "|used=" & ws.UsedRange.Address(False, False)

    ' The total's range shrank by the two deleted rows and still computes.
    parts = parts & "|D4formula=" & ws.Range("D4").Formula
    parts = parts & "|D4=" & CStr(ws.Range("D4").Value)
    parts = parts & "|D3formula=" & ws.Range("D3").Formula
    parts = parts & "|D3=" & CStr(ws.Range("D3").Value)

    parts = parts & "|merge=" & ws.Range("A9").MergeArea.Address(False, False)

    ' A formula that pointed into the deleted rows. Excel writes #REF! for
    ' one of its own, so what matters is that it reads back the same way.
    parts = parts & "|broken=" & wb.Worksheets("Summary").Range("A2").Formula
    parts = parts & "|brokenText=" & wb.Worksheets("Summary").Range("A2").Text
    parts = parts & "|shrunk=" & wb.Worksheets("Summary").Range("A1").Formula
    parts = parts & "|shrunkValue=" & CStr(wb.Worksheets("Summary").Range("A1").Value)
    parts = parts & "|name=" & wb.Names("Totals").RefersTo

    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""


def test_excel_accepts_deleted_rows_and_the_references_that_broke(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """Deletion has a failure mode insertion does not: a reference to
    something gone becomes ``#REF!``, and only Excel can confirm the spelling
    it reads back is one it accepts."""
    book = Workbook.open(live_sample_xlsx)
    summary = book.add_sheet("Summary")
    summary["A1"].formula = "=SUM(Data!D2:D5)"
    summary["A2"].formula = "=Data!B3"
    book.add_defined_name("Totals", "Data!$D$2:$D$5")

    book["Data"].delete_rows(3, 2)
    target = tmp_path / "deleted.xlsx"
    book.save(target)

    seen = probe(excel, _DELETE_PROBE, target, "delete")

    assert seen["A2"] == "North", "the row above the deletion stayed"
    assert seen["A3"] == "West", "row 5 came up to row 3"
    assert seen["used"] == "A1:F9", "the sheet is two rows shorter"

    # SUM(D2:D5) lost two of its rows, so it shrank rather than breaking.
    assert seen["D4formula"] == "=SUM(D2:D3)"
    assert seen["D4"] == "1297.5", "510 + 787.5, the two surviving rows"
    # The shared formula that was at D5 is now at D3 and reads its new row.
    assert seen["D3formula"] == "=B3*C3"
    assert seen["D3"] == "787.5", "210 * 3.75"

    assert seen["merge"] == "A9:C9", "the merge came up two rows"

    assert "#REF!" in seen["broken"], f"Excel reports {seen['broken']}"
    assert seen["brokenText"] == "#REF!"
    assert seen["shrunk"] == "=SUM(Data!D2:D3)", "the cross-sheet range shrank"
    assert seen["shrunkValue"] == "1297.5"
    assert seen["name"] == "=Data!$D$2:$D$3", "and so did the defined name"


def test_excel_accepts_a_deleted_shared_formula_master(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """A shared formula lives once, on its group's first cell. Deleting that
    cell leaves the rest pointing at nothing unless the group is given its own
    text first."""
    book = Workbook.open(live_sample_xlsx)
    book["Data"].delete_rows(2, 1)
    target = tmp_path / "orphaned.xlsx"
    book.save(target)

    source = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Data")
    parts = "D2=" & ws.Range("D2").Formula & "|D2v=" & CStr(ws.Range("D2").Value)
    parts = parts & "|D3=" & ws.Range("D3").Formula & "|D3v=" & CStr(ws.Range("D3").Value)
    parts = parts & "|D4=" & ws.Range("D4").Formula & "|D4v=" & CStr(ws.Range("D4").Value)
    parts = parts & "|D5=" & ws.Range("D5").Formula & "|D5v=" & CStr(ws.Range("D5").Value)
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""
    seen = probe(excel, source, target, "orphaned")

    # The three survivors each kept a correct formula for their new row.
    assert seen["D2"] == "=B2*C2"
    assert seen["D2v"] == "1445", "340 * 4.25, the row that was 3"
    assert seen["D3"] == "=B3*C3"
    assert seen["D3v"] == "522.5", "95 * 5.5"
    assert seen["D4"] == "=B4*C4"
    assert seen["D4v"] == "787.5", "210 * 3.75"
    assert seen["D5"] == "=SUM(D2:D4)", "the total lost one row"
    assert seen["D5v"] == "2755"


def test_excel_accepts_deleted_columns(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    book = Workbook.open(live_sample_xlsx)
    book["Data"].delete_columns(2, 1)
    target = tmp_path / "deleted_columns.xlsx"
    book.save(target)

    source = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Data")
    parts = "A1=" & ws.Range("A1").Value
    parts = parts & "|B1=" & ws.Range("B1").Value
    parts = parts & "|C2=" & ws.Range("C2").Formula
    parts = parts & "|C2text=" & ws.Range("C2").Text
    parts = parts & "|merge=" & ws.Range("A11").MergeArea.Address(False, False)
    parts = parts & "|used=" & ws.UsedRange.Address(False, False)
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""
    seen = probe(excel, source, target, "deletecols")

    assert seen["A1"] == "Region"
    assert seen["B1"] == "Price", "what was in C came to B"
    # =B2*C2 lost B, so half the formula is gone.
    assert "#REF!" in seen["C2"], f"Excel reports {seen['C2']}"
    assert seen["C2text"] == "#REF!"
    assert seen["merge"] == "A11:B11", "the merge lost a column"
    assert seen["used"] == "A1:E11"


def test_excel_agrees_about_whole_axis_references(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """``SUM(2:4)`` and ``SUM(A:A)`` are references too, and they move.

    The offline tests pin this library's output against answers measured
    from Excel. This one closes the loop: Excel opens a file where this
    library did the moving, and reports the same formulas and the same
    values it would have produced itself.
    """
    def written(where: Path, edit: str) -> Path:
        # The formulas live on Notes, not Data. On Data they would sit in
        # the very rows being moved, and a whole-row SUM containing its own
        # cell is circular besides. This also exercises the qualified path,
        # where the sheet name decides whether a reference moves at all.
        book = Workbook.open(live_sample_xlsx)
        notes = book["Notes"]
        notes["D1"].formula = "=SUM(Data!2:4)"
        notes["D2"].formula = "=SUM(Data!1:6)"
        notes["D3"].formula = "=SUM(Data!6:7)"
        notes["D4"].formula = "=SUM(Data!B:B)"
        # Bare, so it addresses Notes and must not move at all.
        notes["D5"].formula = "=SUM(8:9)"
        if edit == "insert":
            book["Data"].insert_rows(2, 2)
        else:
            book["Data"].delete_rows(2, 3)
        book.save(where)
        return where

    source = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Notes")
    parts = "D1=" & ws.Range("D1").Formula
    parts = parts & "|D2=" & ws.Range("D2").Formula
    parts = parts & "|D3=" & ws.Range("D3").Formula
    parts = parts & "|D4=" & ws.Range("D4").Formula
    parts = parts & "|D5=" & ws.Range("D5").Formula
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""

    seen = probe(excel, source, written(tmp_path / "axis_inserted.xlsx", "insert"), "axisinsert")
    assert seen["D1"] == "=SUM(Data!4:6)", "both ends moved past the insertion"
    assert seen["D2"] == "=SUM(Data!1:8)", "a span across the insertion grew"
    assert seen["D3"] == "=SUM(Data!8:9)"
    assert seen["D4"] == "=SUM(Data!B:B)", "a row insertion cannot touch a column span"
    assert seen["D5"] == "=SUM(8:9)", "bare, so it addresses Notes and stays"

    seen = probe(excel, source, written(tmp_path / "axis_deleted.xlsx", "delete"), "axisdelete")
    assert seen["D1"] == "=SUM(Data!#REF!)", "gone, but still qualified"
    assert seen["D2"] == "=SUM(Data!1:3)", "a span across the deletion shrank"
    assert seen["D3"] == "=SUM(Data!3:4)"
    assert seen["D4"] == "=SUM(Data!B:B)"
    assert seen["D5"] == "=SUM(8:9)"


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
