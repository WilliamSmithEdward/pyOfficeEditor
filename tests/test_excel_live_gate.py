# pyvbaharness ships no type stubs and is installed only on demand, so strict
# inference cannot see through it here. The suppressions are scoped to this
# file, which is the only one that imports it. reportMissingImports is among
# them because the live group is Windows-and-Office only: CI installs [dev]
# alone, where the import genuinely is not resolvable and the try/except
# around it is what makes that fine.
# pyright: reportMissingTypeStubs=false, reportUnknownMemberType=false
# pyright: reportMissingImports=false
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
dependency group, which is not a published extra: pyvbaharness is test
equipment and must not be installable from the released package.

**One session, shared.** The harness holds a machine-wide mutex and acquires
it with a zero timeout, so creating a session per test races against the
previous one's teardown and fails intermittently. Office automation is
sequential by contract, so these tests share a single Excel instance: it is
also three times faster. If the lock is genuinely held by other work on the
machine, the fixture skips rather than fails, because that is not a defect in
the code under test. ``LIVE_EXCEL_LOCK_WAIT`` gives it that many seconds to
wait for the other work to finish first.
"""

from __future__ import annotations

import datetime as dt
import os
import shutil
from collections.abc import Iterator
from pathlib import Path

import pytest

from pyofficeeditor.excel import (
    Alignment,
    Border,
    CellError,
    DataValidation,
    DateGroup,
    Dxf,
    DynamicFilter,
    FilterColumn,
    HeaderFooter,
    HeaderFooterText,
    PageMargins,
    PageSetup,
    PrintOptions,
    SheetProtection,
    Top10Filter,
    ValueFilter,
    Workbook,
    cell_is,
    contains_text,
    criteria,
    duplicates,
    expression,
)

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
        from pyvbaharness.session import HarnessConfig
    except ImportError:  # pragma: no cover - depends on the environment
        pytest.skip('pyvbaharness is not installed; pip install -e ".[dev]" --group live')

    # How long to wait for another session to finish with Excel, in seconds;
    # by default the gate does not wait.
    wait = float(os.environ.get("LIVE_EXCEL_LOCK_WAIT", "0"))
    try:
        session = ExcelSession(HarnessConfig(lock_wait_s=wait))
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
    """A height without customHeight is ignored, and a freeze with the wrong
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


_OUTLINE_PROBE = r"""
Private Function Detail(ByVal r As Range) As String
    On Error Resume Next
    Detail = CStr(r.ShowDetail)
    If Err.Number <> 0 Then Detail = "none"
    Err.Clear
End Function

Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String
    Dim i As Long
    Dim hidden As String

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Data")
    For i = 1 To 8
        If ws.Rows(i).Hidden Then hidden = hidden & "r" & i & ","
    Next i
    For i = 1 To 8
        If ws.Columns(i).Hidden Then hidden = hidden & "c" & i & ","
    Next i
    parts = "hidden=" & hidden
    parts = parts & "|rowSummary=" & Detail(ws.Rows(6))
    parts = parts & "|columnSummary=" & Detail(ws.Columns(4))
    parts = parts & "|summaryWidth=" & CStr(ws.Columns(4).Width) & "|plainWidth=" & CStr(ws.Columns(1).Width)
    ws.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
    hidden = ""
    For i = 1 To 8
        If ws.Rows(i).Hidden Or ws.Columns(i).Hidden Then hidden = hidden & i & ","
    Next i
    parts = parts & "|expanded=" & hidden & "|groupedWidth=" & CStr(ws.Columns(2).Width)
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""


def test_excel_folds_and_unfolds_the_groups_this_library_writes(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """A folded group is three things in the file: the rows or columns
    hidden, the summary marked, and the sheet's outline depth. And every
    column entry carries a width, since one without is a column of width 0
    that no unfolding brings back."""
    target = tmp_path / "outline.xlsx"
    shutil.copy(live_sample_xlsx, target)
    with Workbook.open(target) as book:
        sheet = book["Data"]
        sheet.group_rows(3, 5, collapsed=True)
        sheet.group_columns(2, 3, collapsed=True)
        book.save()

    seen = probe(excel, _OUTLINE_PROBE, target, "outline")

    assert seen["hidden"] == "r3,r4,r5,c2,c3,"
    assert seen["rowSummary"] == "False", "Excel reads the row group as folded"
    assert seen["columnSummary"] == "False", "and the column group"
    assert seen["expanded"] == "", "unfolding shows everything"
    # The standard width is measured at 100% display scaling; at another
    # scaling the pixels round differently, by a few percent.
    plain = float(seen["plainWidth"])
    for key in ("summaryWidth", "groupedWidth"):
        assert abs(float(seen[key]) - plain) <= plain * 0.1, f"{key} {seen[key]} against {plain}"


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


def test_excel_paints_the_conditional_formats_this_library_writes(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """The only check that catches an inert rule.

    A conditional formatting rule can be present, schema-valid, counted by
    ``FormatConditions.Count`` and still never fire: that is what happens
    when the compatibility ``<formula>`` beside an attribute-driven rule is
    left out. Nothing offline detects it, because the file looks right.
    ``DisplayFormat`` is what Excel actually renders, so a rule that does
    nothing shows up here as an unpainted cell.
    """
    book = Workbook.open(live_sample_xlsx)
    sheet = book["Data"]
    pink = Dxf.of(fill="FFC7CE", color="9C0006", bold=True)

    for row in range(1, 8):
        sheet.cell(row, 12).value = row * 10  # L: 10..70
        sheet.cell(row, 13).value = f"item {row}"  # M
        sheet.cell(row, 14).value = row % 3  # N: repeats

    sheet.add_conditional_format("L1:L7", cell_is("greaterThan", 50), dxf=pink)
    sheet.add_conditional_format("M1:M7", contains_text("item 3"), dxf=pink)
    sheet.add_conditional_format("N1:N7", duplicates(), dxf=pink)
    sheet.add_conditional_format("O1:O7", expression("=MOD($L1,20)=0"), dxf=pink)

    target = tmp_path / "conditional.xlsx"
    book.save(target)

    source = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Data")
    parts = "count=" & CStr(ws.Cells.FormatConditions.Count)
    parts = parts & "|cellIs_hit=" & CStr(ws.Range("L6").DisplayFormat.Interior.Color)
    parts = parts & "|cellIs_miss=" & CStr(ws.Range("L2").DisplayFormat.Interior.Color)
    parts = parts & "|text_hit=" & CStr(ws.Range("M3").DisplayFormat.Interior.Color)
    parts = parts & "|text_miss=" & CStr(ws.Range("M1").DisplayFormat.Interior.Color)
    parts = parts & "|dupe_hit=" & CStr(ws.Range("N1").DisplayFormat.Interior.Color)
    parts = parts & "|expr_hit=" & CStr(ws.Range("O2").DisplayFormat.Interior.Color)
    parts = parts & "|expr_miss=" & CStr(ws.Range("O1").DisplayFormat.Interior.Color)
    parts = parts & "|bold=" & CStr(ws.Range("L6").DisplayFormat.Font.Bold)
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""
    seen = probe(excel, source, target, "conditional")

    painted = "13551615"  # FFC7CE, as VBA reports a colour: blue-green-red
    unpainted = "16777215"  # plain white

    assert seen["count"] == "4", "every rule reached the file"
    assert seen["cellIs_hit"] == painted, "60 is greater than 50"
    assert seen["cellIs_miss"] == unpainted, "20 is not"
    # The one that needs the compatibility formula. Without it this is white
    # while the rule still counts above, which is the whole point of the gate.
    assert seen["text_hit"] == painted, "'item 3' contains 'item 3'"
    assert seen["text_miss"] == unpainted, "'item 1' does not"
    assert seen["dupe_hit"] == painted, "1 appears more than once in N"
    assert seen["expr_hit"] == painted, "MOD(20,20) is 0"
    assert seen["expr_miss"] == unpainted, "MOD(10,20) is not"
    assert seen["bold"] == "True", "the dxf's font reached the cell too"


def test_excel_agrees_where_the_once_refused_content_landed(
    excel: object, live_refused_xlsx: Path, tmp_path: Path
) -> None:
    """The gate for an empty refusal list.

    Every element that used to make an insertion raise is on this sheet:
    data validation with its formulas, a protected range, a saved sort,
    scenarios, a shape, a form control and a comment. Each of them records a
    cell address somewhere the worksheet XML does not reach, and Excel is
    the only thing that can say whether they all landed together.

    It caught two real defects while it was being written. The drawing part
    was never touched, because the anchor search looked for a wrapper called
    ``<anchor>`` and a drawing calls it ``<xdr:twoCellAnchor>``. And a
    comment's VML shape names its own cell in ``<x:Row>``/``<x:Column>``,
    separate from its ``<x:Anchor>``; moving one and not the other left the
    VML and the comments part disagreeing, and Excel refused to open the
    workbook at all rather than repairing it.
    """
    book = Workbook.open(live_refused_xlsx)
    book["R"].insert_rows(3, 2)
    target = tmp_path / "once_refused.xlsx"
    book.save(target)

    source = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("R")
    parts = "dvList=" & ws.Range("G4").Validation.Formula1
    parts = parts & "|dvBound=" & ws.Range("H4").Validation.Formula2
    parts = parts & "|dvCustom=" & ws.Range("I4").Validation.Formula1
    parts = parts & "|protected=" & ws.Protection.AllowEditRanges(1).Range.Address(False, False)
    parts = parts & "|sortRange=" & ws.Sort.Rng.Address(False, False)
    parts = parts & "|sortKey=" & ws.Sort.SortFields(1).Key.Address(False, False)
    parts = parts & "|scenario=" & ws.Scenarios(1).ChangingCells.Address(False, False)
    parts = parts & "|shape=" & ws.Shapes(1).TopLeftCell.Address(False, False)
    parts = parts & "|button=" & ws.Shapes(2).TopLeftCell.Address(False, False)
    parts = parts & "|comment=" & ws.Comments(1).Parent.Address(False, False)
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""
    seen = probe(excel, source, target, "oncerefused")

    # A range starting at row 2 keeps its start, which is above the
    # insertion, and only its far end moves.
    assert seen["protected"] == "B2:B11"
    assert seen["sortRange"] == "A2:C11"
    assert seen["sortKey"] == "B2:B11"
    # A2 stays and A3 becomes A5, so the two are no longer contiguous.
    assert seen["scenario"] == "A2,A5"

    # A validation's bounds are formulas and carry references of their own.
    assert seen["dvList"] == "=$E$1:$E$2", "above the insertion, so unmoved"
    assert seen["dvBound"] == "=$B$7", "was $B$5"
    assert seen["dvCustom"] == "=ISNUMBER($A4)", "relative, was $A2"

    # Zero-based anchors, in a drawing part and in VML respectively.
    assert seen["shape"] == "E8", "anchored at zero-based row 5, so row 6"
    assert seen["button"] == "E12", "zero-based row 9, so row 10"
    assert seen["comment"] == "C5", "was C3, in two places that must agree"


def test_excel_enforces_the_validations_this_library_writes(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """Excel's own object model, on rules this library authored.

    ``InCellDropdown`` is the reason this gate exists. The attribute behind
    it, ``showDropDown``, is inverted: Excel writes it for a list whose
    dropdown is turned *off* and omits it for the ordinary list that has
    one. Nothing offline can tell which way round that is, and getting it
    backwards gives every dropdown in the workbook the wrong behaviour while
    the file stays valid.
    """
    book = Workbook.open(live_sample_xlsx)
    sheet = book["Data"]
    sheet["K1"].value = "red"
    sheet["K2"].value = "green"

    sheet.add_data_validation("L2:L9", DataValidation.any_of(["red", "green", "blue"]))
    sheet.add_data_validation(
        "M2:M9", DataValidation.any_of(["red", "green"], hide_dropdown=True)
    )
    sheet.add_data_validation("N2:N9", DataValidation.any_of("=$K$1:$K$2"))
    sheet.add_data_validation(
        "O2:O9",
        DataValidation.whole_number(1, 10)
        .with_error("That is not one to ten.", title="No")
        .with_prompt("One to ten.", title="Heads up"),
    )
    sheet.add_data_validation(
        "Q2:Q9", DataValidation.date(dt.date(2026, 1, 1), operator="greaterThan")
    )
    sheet.add_data_validation("T2:T9", DataValidation.custom("=ISNUMBER(T2)"))
    sheet.add_data_validation(
        "U2:U9",
        DataValidation.whole_number(5, operator="equal").with_error("nope", style="warning"),
    )

    target = tmp_path / "validated.xlsx"
    book.save(target)

    source = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Data")
    parts = "listType=" & CStr(ws.Range("L3").Validation.Type)
    parts = parts & "|listF1=" & ws.Range("L3").Validation.Formula1
    parts = parts & "|listDrop=" & CStr(ws.Range("L3").Validation.InCellDropdown)
    parts = parts & "|hiddenDrop=" & CStr(ws.Range("M3").Validation.InCellDropdown)
    parts = parts & "|rangeF1=" & ws.Range("N3").Validation.Formula1
    parts = parts & "|wholeOp=" & CStr(ws.Range("O3").Validation.Operator)
    parts = parts & "|wholeF1=" & ws.Range("O3").Validation.Formula1
    parts = parts & "|wholeF2=" & ws.Range("O3").Validation.Formula2
    parts = parts & "|errTitle=" & ws.Range("O3").Validation.ErrorTitle
    parts = parts & "|errMsg=" & ws.Range("O3").Validation.ErrorMessage
    parts = parts & "|inTitle=" & ws.Range("O3").Validation.InputTitle
    parts = parts & "|dateF1=" & ws.Range("Q3").Validation.Formula1
    parts = parts & "|customF1=" & ws.Range("T3").Validation.Formula1
    parts = parts & "|warnStyle=" & CStr(ws.Range("U3").Validation.AlertStyle)
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""
    seen = probe(excel, source, target, "validation")

    assert seen["listType"] == "3", "xlValidateList"
    assert seen["listF1"] == "red,green,blue", "the quotes are the file's, not the value's"
    # The inverted pair, which is the whole point of the gate.
    assert seen["listDrop"] == "True", "no showDropDown attribute means a dropdown"
    assert seen["hiddenDrop"] == "False", 'showDropDown="1" means no dropdown'

    assert seen["rangeF1"] == "=$K$1:$K$2"
    assert seen["wholeOp"] == "1", "xlBetween, which the file leaves unwritten"
    assert seen["wholeF1"] == "1"
    assert seen["wholeF2"] == "10"
    assert seen["errTitle"] == "No"
    assert seen["errMsg"] == "That is not one to ten."
    assert seen["inTitle"] == "Heads up"
    # 46023 in the file. Excel shows the serial back as the date it denotes.
    assert seen["dateF1"] == "1/1/2026"
    # Relative to the range's top-left, so it reads as T3 on row 3.
    assert seen["customF1"] == "=ISNUMBER(T3)"
    assert seen["warnStyle"] == "2", "xlValidAlertWarning"


def test_excel_agrees_about_protection_and_appearance(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """The two things only Excel can settle.

    Every protection flag names a lock rather than a permission, so
    ``formatCells="0"`` means formatting is allowed. And a password is a
    SHA-512 hash of the salt plus the password in UTF-16LE, spun a hundred
    thousand times; if this library computed it differently the sheet would
    still look protected and the password would simply not work.

    Excel is asked to unprotect with the password, which fails loudly if the
    hash is wrong, and to report each allowance through ``Protection``.
    """
    book = Workbook.open(live_sample_xlsx)
    sheet = book["Data"]
    sheet.tab_color = "FF0000"
    sheet.show_gridlines = False
    sheet.show_headings = False
    sheet.zoom = 85
    sheet.protect(
        SheetProtection(allow_sort=True, allow_autofilter=True, allow_format_cells=True),
        password="hunter2",
    )
    book["Notes"].visible = "veryHidden"

    target = tmp_path / "protected.xlsx"
    book.save(target)

    source = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Data")

    parts = "protected=" & CStr(ws.ProtectContents)
    parts = parts & "|sort=" & CStr(ws.Protection.AllowSorting)
    parts = parts & "|filter=" & CStr(ws.Protection.AllowFiltering)
    parts = parts & "|format=" & CStr(ws.Protection.AllowFormattingCells)
    parts = parts & "|insertRows=" & CStr(ws.Protection.AllowInsertingRows)
    parts = parts & "|tab=" & CStr(ws.Tab.Color)
    parts = parts & "|grid=" & CStr(wb.Windows(1).DisplayGridlines)
    parts = parts & "|headings=" & CStr(wb.Windows(1).DisplayHeadings)
    parts = parts & "|zoom=" & CStr(wb.Windows(1).Zoom)
    parts = parts & "|notes=" & CStr(wb.Worksheets("Notes").Visible)

    ' The password check. A wrong hash raises here.
    On Error Resume Next
    ws.Unprotect Password:="hunter2"
    parts = parts & "|unprotectErr=" & CStr(Err.Number)
    Err.Clear
    On Error GoTo 0
    parts = parts & "|afterUnprotect=" & CStr(ws.ProtectContents)

    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""
    seen = probe(excel, source, target, "protection")

    assert seen["protected"] == "True"
    # Each of these was written as name="0", which is what lifts the lock.
    assert seen["sort"] == "True", 'sort="0" means sorting is allowed'
    assert seen["filter"] == "True"
    assert seen["format"] == "True"
    # And this one was never written, which is what leaves the lock on.
    assert seen["insertRows"] == "False", "absent means blocked"

    assert seen["tab"] == "255", "BGR, so pure red is 255"
    assert seen["grid"] == "False"
    assert seen["headings"] == "False"
    assert seen["zoom"] == "85"
    assert seen["notes"] == "2", "xlSheetVeryHidden"

    # The hash is right, so Excel accepts the password and the sheet opens.
    assert seen["unprotectErr"] == "0", "Excel rejected the password"
    assert seen["afterUnprotect"] == "False"


def test_excel_agrees_about_how_a_sheet_prints(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """Page setup, read back through Excel's own PageSetup.

    Two of these cannot be checked any other way. Margins are inches in the
    file and points in the object model, so a unit slip is invisible offline
    and gives an inch-and-a-half margin where a quarter inch was asked for.
    And fitToWidth and fitToHeight do nothing until the separate
    ``sheetPr/pageSetUpPr/@fitToPage`` flag turns them on, so a sheet can
    carry the right numbers and still print at full size.
    """
    book = Workbook.open(live_sample_xlsx)
    sheet = book["Data"]
    sheet.page_margins = PageMargins(0.25, 0.25, 1.0, 1.0, 0.5, 0.5)
    sheet.page_setup = PageSetup.on(
        "A4",
        orientation="landscape",
        fit_to_width=1,
        fit_to_height=2,
        first_page_number=3,
        use_first_page_number=True,
    )
    sheet.fit_to_page = True
    sheet.print_options = PrintOptions(
        horizontal_centered=True, headings=True, gridlines=True
    )
    sheet.header_footer = HeaderFooter(
        odd_header=HeaderFooterText(left="left head", center="centre"),
        odd_footer=HeaderFooterText(right="Page &P of &N"),
    )
    sheet.print_area = "A1:D9"
    sheet.print_titles = "$1:$1"

    target = tmp_path / "printed.xlsx"
    book.save(target)

    source = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Data")
    With ws.PageSetup
        parts = "left=" & CStr(.LeftMargin)
        parts = parts & "|top=" & CStr(.TopMargin)
        parts = parts & "|orient=" & CStr(.Orientation)
        parts = parts & "|paper=" & CStr(.PaperSize)
        parts = parts & "|zoomOff=" & CStr(.Zoom)
        parts = parts & "|wide=" & CStr(.FitToPagesWide)
        parts = parts & "|tall=" & CStr(.FitToPagesTall)
        parts = parts & "|firstPage=" & CStr(.FirstPageNumber)
        parts = parts & "|centred=" & CStr(.CenterHorizontally)
        parts = parts & "|grid=" & CStr(.PrintGridlines)
        parts = parts & "|headings=" & CStr(.PrintHeadings)
        parts = parts & "|lhead=" & .LeftHeader
        parts = parts & "|chead=" & .CenterHeader
        parts = parts & "|rfoot=" & .RightFooter
        parts = parts & "|area=" & .PrintArea
        parts = parts & "|titles=" & .PrintTitleRows
    End With
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""
    seen = probe(excel, source, target, "printing")

    # Inches in the file, points here: a quarter inch is eighteen points.
    assert seen["left"] == "18", "0.25 inch"
    assert seen["top"] == "72", "1 inch"

    assert seen["orient"] == "2", "xlLandscape"
    assert seen["paper"] == "9", "xlPaperA4"

    # Zoom reports False exactly when fit-to-page is on, which is the flag
    # that lives on sheetPr rather than pageSetup.
    assert seen["zoomOff"] == "False", "fitToPage did not take"
    assert seen["wide"] == "1"
    assert seen["tall"] == "2"

    assert seen["firstPage"] == "3"
    assert seen["centred"] == "True"
    assert seen["grid"] == "True"
    assert seen["headings"] == "True"

    # One string in the file, three boxes here.
    assert seen["lhead"] == "left head"
    assert seen["chead"] == "centre"
    assert seen["rfoot"] == "Page &P of &N"

    # Both of these are defined names rather than attributes.
    assert seen["area"] == "$A$1:$D$9"
    assert seen["titles"] == "$1:$1"


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


# --------------------------------------------------------------------------
# Shapes this library adds
# --------------------------------------------------------------------------

_SHAPE_PROBE = r"""
Private Function Safe(ByVal s As Shape, ByVal which As String) As String
    On Error Resume Next
    If which = "link" Then Safe = s.ControlFormat.LinkedCell
    If which = "list" Then Safe = s.ControlFormat.ListFillRange
    If which = "value" Then Safe = CStr(s.ControlFormat.Value)
    If Err.Number <> 0 Then Safe = ""
    Err.Clear
End Function

Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim s As Shape
    Dim parts As String

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets(1)

    parts = "count=" & CStr(ws.Shapes.Count)
    parts = parts & "|standardHeight=" & CStr(ws.StandardHeight)

    Set s = ws.Shapes("Box")
    parts = parts & "|boxType=" & CStr(s.Type)
    parts = parts & "|boxTop=" & CStr(s.Top)
    parts = parts & "|boxWidth=" & CStr(s.Width)
    parts = parts & "|boxHeight=" & CStr(s.Height)
    parts = parts & "|boxText=" & s.TextFrame.Characters.Text
    parts = parts & "|boxMacro=" & s.OnAction

    Set s = ws.Shapes("Edge")
    parts = parts & "|edgeType=" & CStr(s.Type)

    Set s = ws.Shapes("Note")
    parts = parts & "|noteType=" & CStr(s.Type)

    Set s = ws.Shapes("Press")
    parts = parts & "|btnType=" & CStr(s.Type)
    parts = parts & "|btnMacro=" & s.OnAction
    parts = parts & "|btnText=" & s.TextFrame.Characters.Text

    Set s = ws.Shapes("Tick")
    parts = parts & "|tickType=" & CStr(s.Type)
    parts = parts & "|tickValue=" & Safe(s, "value")

    Set s = ws.Shapes("Pick")
    parts = parts & "|pickList=" & Safe(s, "list")
    parts = parts & "|pickValue=" & Safe(s, "value")

    Set s = ws.Shapes("Step")
    parts = parts & "|stepValue=" & Safe(s, "value")

    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""


def test_excel_accepts_shapes_this_library_adds(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """The gate for writing shapes, on a sheet that has none of the parts.

    It has to make the drawing, the VML and a control part, declare the
    content types and prefixes, and wire four relationships. Every one of
    those was got wrong first time and Excel refused the workbook rather
    than repairing it, which is why this test opens the file rather than
    comparing bytes.

    The values are deliberately set without a linked cell. Measured: a
    control with one takes its state from that cell on load, so a tick box
    stored ticked and linked to an empty cell opens unticked, and the
    stored value would prove nothing.
    """
    target = tmp_path / "added.xlsx"
    shutil.copy(live_sample_xlsx, target)

    with Workbook.open(target) as book:
        sheet = book["Data"]
        sheet.add_shape(
            "Box", left=300, top=20, width=120, height=50,
            text="Hello", macro="Run", geometry="roundRect",
        )
        sheet.add_shape("Edge", left=300, top=90, width=120, height=10, kind="line")
        sheet.add_shape(
            "Note", left=300, top=110, width=120, height=40, kind="textBox", text="Two"
        )
        sheet.add_form_control(
            "Press", left=300, top=160, width=100, height=30, text="Click", macro="Run"
        )
        sheet.add_form_control(
            "Tick", kind="CheckBox", left=300, top=200, width=110, height=20, value=1
        )
        sheet.add_form_control(
            "Pick", kind="Drop", left=300, top=230, width=110, height=20,
            list_range="$A$1:$A$3", value=2,
        )
        sheet.add_form_control(
            "Step", kind="Spin", left=300, top=260, width=20, height=30,
            value=7, maximum=50,
        )
        book.save()

    seen = probe(excel, _SHAPE_PROBE, target, "added")

    assert seen["count"] == "7"

    # A drawing shape: msoAutoShape, at the box it was given, with its
    # text and the macro a click runs.
    assert seen["boxType"] == "1"
    # A row with no height of its own is as tall as Excel's default row at
    # the display scaling it runs at, which the file only hints at: 14.5
    # points where sample.xlsx was written, 15 at 100%. The box is anchored
    # from the second row to the fifth, so its top moves by the difference
    # once and its height by it three times.
    drift = float(seen["standardHeight"]) - 14.5
    assert float(seen["boxTop"]) == 20 + drift
    assert seen["boxWidth"] == "120"
    assert float(seen["boxHeight"]) == 50 + 3 * drift
    assert seen["boxText"] == "Hello"
    # A drawing shape reports the bare name it stores. Only a control is
    # reported qualified by the workbook, because only a control stores
    # it that way: see btnMacro below.
    assert seen["boxMacro"] == "Run"

    # msoLine, which needs prst="line": a connector carrying the rect
    # every other shape gets makes Excel refuse the package.
    assert seen["edgeType"] == "9"
    assert seen["noteType"] == "17", "msoTextBox, which is an sp with a flag"

    # msoFormControl for all four, which is the whole four-part assembly
    # agreeing: drawing, sheet record, control part and VML.
    assert seen["btnType"] == "8"
    assert seen["btnMacro"].endswith("!Run")
    assert seen["btnText"] == "Click"

    assert seen["tickType"] == "8"
    assert seen["tickValue"] == "1", "xlOn"
    assert seen["pickList"] == "$A$1:$A$3"
    assert seen["pickValue"] == "2", "the second item, which is sel and not val"
    assert seen["stepValue"] == "7", "a spinner whose maximum was left at 0 reads 0"


_NOTES_PROBE = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim c As Comment
    Dim out As String

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    For Each ws In wb.Worksheets
        For Each c In ws.Comments
            out = out & ws.Name & "!" & c.Parent.Address(False, False) & "=" & _
                  Replace(c.Text, vbLf, "\n") & ";" & c.Author & ";" & CStr(c.Visible) & "|"
        Next c
        out = out & ws.Name & "!shapes=" & CStr(ws.Shapes.Count) & "|"
    Next ws
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = Left(out, Len(out) - 1)
End Function
"""


def test_excel_reads_the_notes_this_library_writes(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """A note is three things that have to agree: its text in the comments
    part, its box in the VML part, and ``<legacyDrawing>`` on the sheet.
    Beside a form control the two share one VML part and one sequence of
    shape ids, and Excel refused the first version of that."""
    target = tmp_path / "notes.xlsx"
    shutil.copy(live_sample_xlsx, target)
    with Workbook.open(target) as book:
        data = book["Data"]
        data.set_comment("C3", "plain note", author="Ada")
        data.set_comment("E5", "two\nlines", visible=True)
        data.set_comment("A1", "corner")
        data.set_comment("B8", "removed again")
        data.remove_comment("B8")
        notes = book["Notes"]
        notes.set_comment("B2", "second sheet")
        notes.add_form_control("Press", left=150, top=40, width=60, height=20, text="Click")
        notes.set_comment("D6", "after the button", author="Bo")
        book.save()

    result = excel.run_vba(  # type: ignore[attr-defined]
        _NOTES_PROBE, proc="Probe", args=(str(target),), timeout=180, module_name="Probe_notes"
    )
    assert result.outcome == "passed", f"Excel refused the workbook: {result!r}"
    seen = dict(part.split("=", 1) for part in str(result.value).split("|"))
    assert seen == {
        "Data!A1": "corner;;False",
        "Data!C3": "plain note;Ada;False",
        "Data!E5": "two\\nlines;;True",
        "Data!shapes": "3",
        "Notes!B2": "second sheet;;False",
        "Notes!D6": "after the button;Bo;False",
        "Notes!shapes": "3",
    }


_THREADS_PROBE = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim t As CommentThreaded
    Dim r As CommentThreaded
    Dim out As String

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    For Each ws In wb.Worksheets
        For Each t In ws.CommentsThreaded
            out = out & ws.Name & "!" & t.Parent.Address(False, False) & "=" & _
                  Replace(t.Text, vbLf, "\n") & ";" & t.Author.Name & ";" & CStr(t.Resolved)
            For Each r In t.Replies
                out = out & ";" & r.Text & "/" & r.Author.Name
            Next r
            out = out & "|"
        Next t
        out = out & ws.Name & "!notes=" & CStr(ws.Comments.Count) & "|"
    Next ws
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = Left(out, Len(out) - 1)
End Function
"""


def test_excel_reads_the_threads_this_library_writes(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """A thread is four things: its entries in the sheet's threads part,
    their authors in the workbook's person list, the placeholder note Excel
    keeps beside it, and that note's box. Excel lists replies written at
    one moment by their ids, so their order is part of what it checks."""
    target = tmp_path / "threads.xlsx"
    shutil.copy(live_sample_xlsx, target)
    when = dt.datetime(2026, 9, 22, 18, 30, tzinfo=dt.timezone.utc)
    with Workbook.open(target) as book:
        data = book["Data"]
        data.add_threaded_comment("C3", "first line\nsecond line", author="Ada", when=when)
        data.add_threaded_reply("C3", "a reply", author="Bo", when=when)
        data.add_threaded_reply("C3", "and another", author="Ada", when=when)
        data.add_threaded_comment("B2", "done with", author="Bo", when=when)
        data.resolve_threaded_comment("B2")
        data.set_comment("E5", "a plain note beside them")
        book["Notes"].add_threaded_comment("A1", "second sheet", author="Cy", when=when)
        book.save()

    result = excel.run_vba(  # type: ignore[attr-defined]
        _THREADS_PROBE, proc="Probe", args=(str(target),), timeout=180, module_name="Probe_threads"
    )
    assert result.outcome == "passed", f"Excel refused the workbook: {result!r}"
    seen = dict(part.split("=", 1) for part in str(result.value).split("|"))
    assert seen == {
        "Data!B2": "done with;Bo;True",
        "Data!C3": "first line\\nsecond line;Ada;False;a reply/Bo;and another/Ada",
        "Data!notes": "1",
        "Notes!A1": "second sheet;Cy;False",
        "Notes!notes": "0",
    }


_REMOVED_PROBE = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim s As Shape
    Dim names As String

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets(1)
    For Each s In ws.Shapes
        names = names & s.Name & ","
    Next s
    Probe = "count=" & CStr(ws.Shapes.Count) & "|names=" & names
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Function
"""


def test_excel_accepts_a_workbook_with_shapes_removed(
    excel: object, live_controls_xlsm: Path, tmp_path: Path
) -> None:
    """Removing a control has to take four things with it.

    A relationship left pointing at a control part that is gone is the
    failure that stays invisible until Excel opens the file, and then it
    is the whole workbook that gets repaired rather than the one control.
    """
    target = tmp_path / "removed.xlsm"
    shutil.copy(live_controls_xlsm, target)

    with Workbook.open(target) as book:
        sheet = book["Controls"]
        sheet.remove_shape("Plain")
        sheet.remove_shape("Tick")
        sheet.remove_shape("Step")
        book.save()

    seen = probe(excel, _REMOVED_PROBE, target, "removed")
    assert seen["count"] == "10"
    gone = {"Plain", "Tick", "Step"}
    left = {name for name in seen["names"].split(",") if name}
    assert not (gone & left)
    assert "Go" in left, "the others are untouched"


_MACRO_PROBE = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets(1)
    Probe = "control=" & ws.Shapes("Go").OnAction & _
            "|drawing=" & ws.Shapes("Plain").OnAction
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Function
"""


def test_excel_runs_the_macro_a_shape_was_pointed_at(
    excel: object, live_controls_xlsm: Path, tmp_path: Path
) -> None:
    """Measured: the sheet's ``<controlPr macro=...>`` is what Excel
    reads, and the VML's ``<x:FmlaMacro>`` is ignored. Both are written,
    and this checks the one that counts."""
    target = tmp_path / "macros.xlsm"
    shutil.copy(live_controls_xlsm, target)

    with Workbook.open(target) as book:
        sheet = book["Controls"]
        sheet.set_shape_macro("Go", "Renamed")
        sheet.set_shape_macro("Plain", "OnPlain")
        book.save()

    seen = probe(excel, _MACRO_PROBE, target, "macros")
    assert seen["control"].endswith("!Renamed")
    assert seen["drawing"] == "OnPlain", "a drawing shape carries a bare name"


# --------------------------------------------------------------------------
# Autofilters
# --------------------------------------------------------------------------

_FILTER_PROBE = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long
    Dim hidden As String
    Dim shown As Long

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Data")

    For i = 2 To 9
        If ws.Rows(i).Hidden Then
            hidden = hidden & i & ","
        Else
            shown = shown + 1
        End If
    Next i

    Probe = "mode=" & CStr(ws.AutoFilterMode) & _
            "|range=" & ws.AutoFilter.Range.Address(True, True, 1, False) & _
            "|hidden=" & hidden & "|shown=" & CStr(shown) & _
            "|header=" & CStr(ws.Rows(1).Hidden)
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Function
"""


def test_excel_shows_a_filter_this_library_applied(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """The gate for autofilters, and it is about the rows rather than the
    criteria.

    Excel does not recompute a filter when a workbook opens: it reads the
    criteria to light the dropdowns and takes which rows are out of view
    from a flag on each row. So a file carrying criteria alone opens
    claiming to be filtered with everything showing, and only Excel can
    confirm this library wrote both halves and that they agree.
    """
    target = tmp_path / "filtered.xlsx"
    shutil.copy(live_sample_xlsx, target)

    with Workbook.open(target) as book:
        sheet = book["Data"]
        wanted = [
            row
            for row in range(2, 10)
            if str(sheet[f"A{row}"].value or "") == "North"
        ]
        assert wanted, "the fixture should carry some North rows to keep"
        sheet.set_auto_filter(
            "A1:D9", [FilterColumn(0, ValueFilter(values=("North",)))]
        )
        book.save()

    seen = probe(excel, _FILTER_PROBE, target, "filtered")

    assert seen["mode"] == "True", "the filter is on"
    assert seen["range"] == "$A$1:$D$9"
    assert seen["header"] == "False", "the header row stays in view"
    assert int(seen["shown"]) == len(wanted), (
        "Excel shows a different number of rows than the criteria keep, which "
        "is the criteria and the hidden flags disagreeing"
    )

    hidden = {int(one) for one in seen["hidden"].split(",") if one}
    assert hidden == set(range(2, 10)) - set(wanted)


def test_excel_shows_every_row_after_the_filter_is_cleared(
    excel: object, live_sample_xlsx: Path, tmp_path: Path
) -> None:
    """Clearing has the same split to get right in reverse: dropping the
    criteria alone would leave the excluded rows hidden with nothing left
    to say why."""
    target = tmp_path / "unfiltered.xlsx"
    shutil.copy(live_sample_xlsx, target)

    with Workbook.open(target) as book:
        sheet = book["Data"]
        sheet.set_auto_filter(
            "A1:D9", [FilterColumn(0, ValueFilter(values=("North",)))]
        )
        sheet.clear_auto_filter()
        book.save()

    source = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long
    Dim hidden As Long

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Data")
    For i = 1 To 9
        If ws.Rows(i).Hidden Then hidden = hidden + 1
    Next i
    Probe = "mode=" & CStr(ws.AutoFilterMode) & "|hidden=" & CStr(hidden)
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Function
"""
    seen = probe(excel, source, target, "unfiltered")
    assert seen["mode"] == "False", "the filter is gone"
    assert seen["hidden"] == "0", "and nothing is left hidden by it"


_HIDDEN_ROWS_VBA = r"""
Private Function Hidden(ByVal ws As Worksheet) As String
    Dim i As Long
    Dim out As String
    For i = 2 To 30
        If ws.Rows(i).Hidden Then out = out & i & ","
    Next i
    Hidden = out
End Function
"""

_REAPPLY_PROBE = _HIDDEN_ROWS_VBA + r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim out As String
    Dim before As String

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    For Each ws In wb.Worksheets
        If ws.AutoFilterMode Then
            before = Hidden(ws)
            ws.AutoFilter.ApplyFilter
            out = out & ws.Name & "=" & before & ";" & Hidden(ws) & ";" & CStr(ws.FilterMode) & "|"
        End If
    Next ws
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = Left(out, Len(out) - 1)
End Function
"""


def test_excel_applies_each_filter_as_this_library_did(
    excel: object, live_filters_xlsx: Path, tmp_path: Path
) -> None:
    """A different criterion on every sheet, applied here and saved; then
    Excel opens the file and applies each filter again itself. The rows it
    hides have to be the rows already hidden, sheet by sheet. That is the
    evaluator held to Excel on criteria Excel's own dropdowns would write,
    and the markup held to Excel, since a filter it could not read would
    not re-apply."""
    target = tmp_path / "reapplied.xlsx"
    shutil.copy(live_filters_xlsx, target)
    with Workbook.open(target) as book:
        book["Values"].set_auto_filter("A1:C21", [FilterColumn(0, criteria("=r1*"))])
        book["Compare"].set_auto_filter(
            "A1:C21", [FilterColumn(0, criteria("<>r5")), FilterColumn(1, criteria(">=10", "<40"))]
        )
        book["Between"].set_auto_filter("A1:C21", [FilterColumn(1, Top10Filter(5))])
        book["TopTen"].set_auto_filter("A1:C21", [FilterColumn(1, DynamicFilter("belowAverage"))])
        book["Blanks"].set_auto_filter("A1:C21", [FilterColumn(0, criteria("="))])
        # Measured from today, which is the day Excel measures it from too.
        book["Dynamic"].set_auto_filter("A1:C21", [FilterColumn(2, DynamicFilter("thisMonth"))])
        book["Numbers"].set_auto_filter("A1:C21", [FilterColumn(2, criteria(">=6/1/2026"))])
        book["Dates"].set_auto_filter(
            "A1:C21", [FilterColumn(2, ValueFilter(("1/5/2026",), date_groups=(DateGroup("month", 2026, 3),)))]
        )
        # And a column inserted into a filtered range, which renumbers it.
        book["Compare"].insert_columns(2)
        book.save()

    result = excel.run_vba(  # type: ignore[attr-defined]
        _REAPPLY_PROBE, proc="Probe", args=(str(target),), timeout=180, module_name="Probe_reapply"
    )
    assert result.outcome == "passed", f"Excel refused the workbook: {result!r}"
    seen = dict(part.split("=", 1) for part in str(result.value).split("|"))
    assert len(seen) == 8
    for name, reply in seen.items():
        before, after, mode = reply.split(";")
        assert before == after, f"{name}: this library hid {before}, Excel hides {after}"
        assert before, f"{name}: the criterion should hide something"
        assert mode == "True", name


_TABLE_REAPPLY_PROBE = _HIDDEN_ROWS_VBA + r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim out As String
    Dim before As String

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            before = Hidden(ws)
            lo.AutoFilter.ApplyFilter
            out = out & ws.Name & "=" & before & ";" & Hidden(ws) & ";" & CStr(ws.AutoFilterMode) & "|"
        Next lo
    Next ws
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = Left(out, Len(out) - 1)
End Function
"""


def test_excel_applies_each_table_filter_as_this_library_did(
    excel: object, live_filters_xlsx: Path, tmp_path: Path
) -> None:
    """The same criteria as the gate above, each on a table made over the
    sheet's data in place of the sheet's own filter. Excel applies each
    table's filter again and has to hide the rows already hidden."""
    target = tmp_path / "tables_reapplied.xlsx"
    shutil.copy(live_filters_xlsx, target)
    wanted = {
        "Values": [FilterColumn(0, criteria("=r1*"))],
        "Compare": [FilterColumn(0, criteria("<>r5")), FilterColumn(1, criteria(">=10", "<40"))],
        "Between": [FilterColumn(1, Top10Filter(5))],
        "TopTen": [FilterColumn(1, DynamicFilter("belowAverage"))],
        "Blanks": [FilterColumn(0, criteria("="))],
        "Dynamic": [FilterColumn(2, DynamicFilter("thisMonth"))],
        "Numbers": [FilterColumn(2, criteria(">=6/1/2026"))],
        "Dates": [FilterColumn(2, ValueFilter(("1/5/2026",), date_groups=(DateGroup("month", 2026, 3),)))],
    }
    with Workbook.open(target) as book:
        for name, columns in wanted.items():
            sheet = book[name]
            sheet.clear_auto_filter()
            sheet.add_table(f"Table{name}", "A1:C21")
            sheet.set_table_filter(f"Table{name}", columns)
        # A column inserted into a filtered table, which renumbers it.
        book["Compare"].insert_columns(2)
        book.save()

    result = excel.run_vba(  # type: ignore[attr-defined]
        _TABLE_REAPPLY_PROBE, proc="Probe", args=(str(target),), timeout=180, module_name="Probe_tables"
    )
    assert result.outcome == "passed", f"Excel refused the workbook: {result!r}"
    seen = dict(part.split("=", 1) for part in str(result.value).split("|"))
    assert len(seen) == 8
    for name, reply in seen.items():
        before, after, sheet_filter = reply.split(";")
        assert before == after, f"{name}: this library hid {before}, Excel hides {after}"
        assert before, f"{name}: the criterion should hide something"
        assert sheet_filter == "False", f"{name}: the sheet's own filter is off"
