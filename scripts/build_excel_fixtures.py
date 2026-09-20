"""Author the Excel test fixtures with real Excel.

The library's fidelity gates only mean something against bytes Excel
actually wrote.  openpyxl writes the same parts a different legal way: no
growth-hint extra fields, different ZIP version words, a different styles
part.  Both are worth testing against, and only one of them can be
generated without Office.

Run this on a Windows machine with Excel installed:

    python -m pip install -e ".[dev]" --group live
    python scripts/build_excel_fixtures.py          # only what is missing
    python scripts/build_excel_fixtures.py --force  # rebuild everything

A fixture that already exists is left alone. Excel does not produce the same
bytes twice, since it stamps every part with a fresh revision GUID, so
rebuilding a committed fixture churns it for no benefit and buries the real
change in a diff. ``--force`` is there for when a fixture's *content* needs to
change, which is a deliberate act.

Excel runs hidden and the harness terminates the instance it owns by recorded
process id even if a step wedges.
"""

from __future__ import annotations

import sys
from pathlib import Path

FIXTURES = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "excel"

#: XlFileFormat values.  Excel will not save .xlsx through the harness's
#: save_as, which only knows the macro-enabled formats, so the format is
#: named explicitly in VBA.
XL_OPEN_XML_WORKBOOK = 51

_BUILD_SAMPLE = r"""
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Data"

    ' Text, so the workbook gets a real sharedStrings part.
    ws.Range("A1").Value = "Region"
    ws.Range("B1").Value = "Units"
    ws.Range("C1").Value = "Price"
    ws.Range("D1").Value = "Total"
    ws.Range("E1").Value = "Shipped"
    ws.Range("F1").Value = "Ordered"
    ws.Range("A1:F1").Font.Bold = True

    ws.Range("A2").Value = "North"
    ws.Range("A3").Value = "South"
    ws.Range("A4").Value = "East"
    ws.Range("A5").Value = "West"

    ' Numbers, including one that is deliberately not an integer.
    ws.Range("B2").Value = 120
    ws.Range("B3").Value = 340
    ws.Range("B4").Value = 95
    ws.Range("B5").Value = 210
    ws.Range("C2").Value = 4.25
    ws.Range("C3").Value = 4.25
    ws.Range("C4").Value = 5.5
    ws.Range("C5").Value = 3.75

    ' Formulas, so there is something with a cached value to read.
    ws.Range("D2:D5").Formula = "=B2*C2"
    ws.Range("D6").Formula = "=SUM(D2:D5)"

    ' Booleans and dates, the two types that do not look like numbers.
    ws.Range("E2").Value = True
    ws.Range("E3").Value = False
    ws.Range("E4").Value = True
    ws.Range("E5").Value = False
    ws.Range("F2").Value = DateSerial(2026, 1, 15)
    ws.Range("F3").Value = DateSerial(2026, 2, 1)
    ws.Range("F4").Value = DateSerial(2026, 2, 28)
    ws.Range("F5").Value = DateSerial(2026, 3, 9)
    ws.Range("F2:F5").NumberFormat = "yyyy-mm-dd"

    ' A string with characters that have to be escaped, and one with
    ' leading and trailing spaces that needs xml:space="preserve".
    ws.Range("A8").Value = "ampersand & angle < bracket > quote "" done"
    ws.Range("A9").Value = "  padded  "

    ' An inline formula error, so the error cell type appears.
    ws.Range("B8").Formula = "=1/0"

    ' A merged range and a second sheet, so structure is exercised too.
    ws.Range("A11:C11").Merge
    ws.Range("A11").Value = "Merged heading"

    wb.Worksheets.Add After:=ws
    wb.Worksheets(2).Name = "Notes"
    wb.Worksheets(2).Range("A1").Value = "Second sheet"

    ws.Activate
    ws.Range("A1").Select

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True

    Build = wb.FullName
End Function
"""

_BUILD_EMPTY = r"""
Public Function Build(ByVal Target As String) As String
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = ActiveWorkbook.FullName
End Function
"""

_BUILD_STRUCTURES = r"""
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Tabled"

    ws.Range("A1").Value = "Region"
    ws.Range("B1").Value = "Units"
    ws.Range("C1").Value = "Revenue"
    ws.Range("A2").Value = "North"
    ws.Range("B2").Value = 120
    ws.Range("C2").Formula = "=B2*10"
    ws.Range("A3").Value = "South"
    ws.Range("B3").Value = 340
    ws.Range("C3").Formula = "=B3*10"
    ws.Range("A4").Value = "East"
    ws.Range("B4").Value = 95
    ws.Range("C4").Formula = "=B4*10"

    ' A real ListObject, with a totals row so the shape is complete.
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C4"), , xlYes)
    lo.Name = "SalesTable"
    lo.TableStyle = "TableStyleMedium2"
    lo.ShowTotals = True
    lo.ListColumns("Units").TotalsCalculation = xlTotalsCalculationSum

    ' A second, plainer table on another sheet, so two table parts exist.
    wb.Worksheets.Add After:=ws
    wb.Worksheets(2).Name = "Second"
    wb.Worksheets(2).Range("A1").Value = "Key"
    wb.Worksheets(2).Range("A2").Value = "k1"
    wb.Worksheets(2).ListObjects.Add(xlSrcRange, _
        wb.Worksheets(2).Range("A1:A2"), , xlYes).Name = "KeyTable"

    ' A workbook-scoped defined name and a sheet-scoped one, since renaming a
    ' sheet has to repoint both.
    wb.Names.Add Name:="TotalUnits", RefersTo:="=Tabled!$B$2:$B$4"
    ws.Names.Add Name:="LocalRegion", RefersTo:="=Tabled!$A$2"

    ' Column widths and a row height, which live outside sheetData.
    ws.Columns("A").ColumnWidth = 18
    ws.Rows(1).RowHeight = 24

    ' A hyperlink, which needs its own relationship from the sheet part.
    ws.Hyperlinks.Add Anchor:=ws.Range("E1"), Address:="https://example.invalid/", _
        TextToDisplay:="a link"

    ws.Activate
    ws.Range("A1").Select

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = wb.FullName
End Function
"""


def build(session: object, source: str, target: Path, label: str, *, force: bool) -> None:
    from pyvbaharness import ExcelSession

    assert isinstance(session, ExcelSession)
    if target.exists():
        if not force:
            print(f"  kept  {target.name} (already there; pass --force to rebuild)")
            return
        target.unlink()
    session.reset_sheets()
    result = session.run_vba(source, proc="Build", args=(str(target),), timeout=180)
    if result.outcome != "passed":
        raise SystemExit(f"{label}: Excel refused the build ({result.outcome}): {result!r}")
    if not target.is_file():
        raise SystemExit(f"{label}: Excel reported success but {target} is not there.")
    print(f"  wrote {target.relative_to(FIXTURES.parent.parent.parent)} ({target.stat().st_size} bytes)")



#: Every element that used to make an insertion refuse, on one sheet.
#: Between them they cover all five address notations, including the two
#: zero-based ones and the two cases where an address is stored twice and
#: Excel refuses the workbook if the copies disagree.
_BUILD_REFUSED = r"""
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "R"

    For i = 1 To 12
        ws.Cells(i, 1).Value = i
        ws.Cells(i, 2).Value = i * 3
        ws.Cells(i, 3).Value = "n" & i
    Next i
    ws.Range("E1").Value = "alpha"
    ws.Range("E2").Value = "beta"

    ' Data validation: a list from a range, a bound that is a reference,
    ' and a custom formula with a relative reference.
    With ws.Range("G2:G9").Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=$E$1:$E$2"
    End With
    With ws.Range("H2:H9").Validation
        .Delete
        .Add Type:=1, AlertStyle:=1, Operator:=1, Formula1:="1", Formula2:="=$B$5"
    End With
    With ws.Range("I2:I9,K2:K4").Validation
        .Delete
        .Add Type:=7, AlertStyle:=1, Operator:=1, Formula1:="=ISNUMBER($A2)"
    End With

    ws.Protection.AllowEditRanges.Add Title:="Editable", Range:=ws.Range("B2:B9")

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("B2:B9"), Order:=2
        .SetRange ws.Range("A1:C9")
        .Header = 1
        .Apply
    End With

    ws.Scenarios.Add Name:="High", ChangingCells:=ws.Range("A2:A3"), Values:=Array(99, 98)

    ' A shape, whose anchor lives in a drawing part.
    ws.Shapes.AddShape 1, ws.Range("E6").Left, ws.Range("E6").Top, 80, 40
    ' A form control, anchored inline and again in VML.
    ws.Buttons.Add ws.Range("E10").Left, ws.Range("E10").Top, 70, 24
    ' A comment, whose cell is in the comments part and again in VML.
    ws.Range("C3").AddComment "note here"

    Build = "ok"
End Function
"""


#: Sheet-level settings: protection with and without a password, a tab
#: colour, view flags, outline grouping, page setup and three visibility
#: states across four sheets.
_BUILD_SETTINGS = r"""
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim other As Worksheet

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "S"
    ws.Range("A1:D20").Value = 1

    wb.Windows(1).Zoom = 85
    ws.Tab.Color = RGB(255, 0, 0)
    wb.Windows(1).DisplayGridlines = False
    wb.Windows(1).DisplayHeadings = False

    With ws.PageSetup
        .Orientation = 2
        .PaperSize = 9
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(1)
        .BottomMargin = Application.InchesToPoints(1)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.5)
        .CenterHorizontally = True
        .PrintArea = "$A$1:$D$20"
        .PrintTitleRows = "$1:$1"
        .PrintTitleColumns = "$A:$A"
        .LeftHeader = "left head"
        .CenterHeader = "&Bbold centre&B"
        .RightFooter = "Page &P of &N"
        .PrintGridlines = True
        .PrintHeadings = True
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 2
        .FirstPageNumber = 3
    End With

    ws.Rows("5:8").Group
    ws.Columns("B:C").Group

    Set other = wb.Worksheets.Add
    other.Name = "Hidden"
    other.Visible = 0
    Set other = wb.Worksheets.Add
    other.Name = "VeryHidden"
    other.Visible = 2

    ws.Protect Password:="secret", DrawingObjects:=True, Contents:=True, Scenarios:=True

    Set other = wb.Worksheets.Add
    other.Name = "Plain"
    other.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowFormattingCells:=True, AllowSorting:=True

    Build = "ok"
End Function
"""


#: Hyperlinks of every kind, an autofilter, and outline grouping.
_BUILD_LINKS = r"""
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "L"
    ws.Range("A1").Value = "Region"
    ws.Range("B1").Value = "Units"
    ws.Range("C1").Value = "When"
    For i = 2 To 12
        ws.Cells(i, 1).Value = "r" & i
        ws.Cells(i, 2).Value = i * 2
        ws.Cells(i, 3).Value = DateSerial(2026, 1, i)
    Next i

    ws.Hyperlinks.Add Anchor:=ws.Range("E2"), Address:="https://example.com/a", _
        ScreenTip:="go there", TextToDisplay:="Example"
    ws.Hyperlinks.Add Anchor:=ws.Range("E3"), Address:="", SubAddress:="L!A1", _
        TextToDisplay:="top"
    ws.Hyperlinks.Add Anchor:=ws.Range("E4"), Address:="mailto:a@b.c", TextToDisplay:="mail"
    ws.Hyperlinks.Add Anchor:=ws.Range("E5"), Address:="https://example.com/b", _
        SubAddress:="frag", TextToDisplay:="both"
    ws.Hyperlinks.Add Anchor:=ws.Range("G2:G4"), Address:="https://example.com/c"

    ws.Range("A1:C12").AutoFilter Field:=1, Criteria1:=Array("r2", "r3"), Operator:=7
    ws.Range("A1:C12").AutoFilter Field:=2, Criteria1:=">=10"

    ws.Rows("5:8").Group
    ws.Rows("9:11").Group
    ws.Rows("5:8").EntireRow.Hidden = True
    ws.Columns("B:C").Group
    ws.Outline.SummaryRow = 0
    ws.Outline.SummaryColumn = 0

    Build = "ok"
End Function
"""


#: One shape of each MsoAutoShapeType the table names, so the preset
#: geometry mapping is held to what Excel writes rather than transcribed.
_BUILD_GEOMETRY = r"""
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim s As Shape
    Dim kinds As Variant
    Dim i As Long

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "G"

    kinds = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, _
                  19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 92, 93, 94, 95, 96)
    For i = LBound(kinds) To UBound(kinds)
        On Error Resume Next
        Set s = ws.Shapes.AddShape(kinds(i), 10 + i * 5, 10, 40, 30)
        If Err.Number = 0 Then s.Name = "T" & kinds(i)
        Err.Clear
        On Error GoTo 0
    Next i

    Build = "ok"
End Function
"""

def main() -> int:
    try:
        from pyvbaharness import ExcelSession
    except ImportError:
        print(
            'pyvbaharness is not installed. Install the live extra:\n'
            '    python -m pip install -e ".[dev]" --group live',
            file=sys.stderr,
        )
        return 2

    force = "--force" in sys.argv
    wanted = [
        ("empty.xlsx", _BUILD_EMPTY),
        ("sample.xlsx", _BUILD_SAMPLE),
        ("structures.xlsx", _BUILD_STRUCTURES),
        ("refused.xlsx", _BUILD_REFUSED),
        ("settings.xlsx", _BUILD_SETTINGS),
        ("links.xlsx", _BUILD_LINKS),
        ("geometry.xlsx", _BUILD_GEOMETRY),
    ]
    if not force and all((FIXTURES / name).exists() for name, _ in wanted):
        print("every fixture is already there; nothing to do (pass --force to rebuild)")
        return 0

    FIXTURES.mkdir(parents=True, exist_ok=True)
    print(f"authoring fixtures in {FIXTURES}")
    with ExcelSession() as excel:
        for name, source in wanted:
            excel.new_workbook()
            build(excel, source, FIXTURES / name, name, force=force)
    print("done")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
