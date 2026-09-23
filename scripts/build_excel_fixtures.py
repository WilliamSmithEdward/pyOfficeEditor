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

import json
import os
import struct
import sys
import zlib
from collections.abc import Callable
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
    print(f"  wrote {target.name} ({target.stat().st_size} bytes)")



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

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = wb.FullName
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

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = wb.FullName
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

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = wb.FullName
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

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = wb.FullName
End Function
"""

#: One of every Forms control that has something to say, wired up, so the
#: control reader is held to what Excel writes rather than to the schema.
#:
#: The committed shapes fixture carries a Button and nothing else, whose
#: whole part is ``objectType="Button" lockText="1"``. It cannot show that
#: a linked cell is read correctly because it has no linked cell, and it
#: cannot show a value at all.
#:
#: The measuring matters because the current value is spelled three
#: different ways and none of them is a single ``val`` attribute:
#:
#: - a check box and an option button write ``checked="Checked"`` or
#:   ``checked="Mixed"``, and write nothing when they are off
#: - a drop down and a list box write ``sel``, the 1-based selection,
#:   and leave ``val`` at 0
#: - a spinner and a scroll bar write ``val``
#:
#: So a reader that reports ``val`` answers 0 for a ticked box and 0 for a
#: drop down with the third item chosen. Both look plausible and both are
#: wrong. Excel's own object model answers -4146 for an unticked box,
#: which is ``xlOff`` and not 0, and that is the number recorded here.
_BUILD_CONTROLS = r"""
Private Function SafeLinked(ByVal s As Shape) As String
    On Error Resume Next
    SafeLinked = s.ControlFormat.LinkedCell
    If Err.Number <> 0 Then SafeLinked = ""
    Err.Clear
End Function

Private Function SafeList(ByVal s As Shape) As String
    On Error Resume Next
    SafeList = s.ControlFormat.ListFillRange
    If Err.Number <> 0 Then SafeList = ""
    Err.Clear
End Function

Private Function SafeValue(ByVal s As Shape) As String
    On Error Resume Next
    SafeValue = CStr(s.ControlFormat.Value)
    If Err.Number <> 0 Then SafeValue = ""
    Err.Clear
End Function

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim s As Shape
    Dim out As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Controls"

    ws.Range("A1").Value = "Alpha"
    ws.Range("A2").Value = "Beta"
    ws.Range("A3").Value = "Gamma"

    ' A button: a macro, and nothing wired to a cell.
    Set s = ws.Shapes.AddFormControl(xlButtonControl, 200, 20, 90, 30)
    s.Name = "Go"
    s.TextFrame.Characters.Text = "Press"
    s.OnAction = "Clicked"

    ' A ticked check box with a linked cell: the case the consumer pins.
    Set s = ws.Shapes.AddFormControl(xlCheckBox, 200, 60, 110, 20)
    s.Name = "Tick"
    s.TextFrame.Characters.Text = "Enabled"
    s.ControlFormat.LinkedCell = "$D$6"
    s.ControlFormat.Value = xlOn

    ' An unticked one, which writes no checked attribute at all.
    Set s = ws.Shapes.AddFormControl(xlCheckBox, 200, 90, 110, 20)
    s.Name = "Untick"
    s.ControlFormat.LinkedCell = "$D$7"
    s.ControlFormat.Value = xlOff

    ' And the third state, which is neither.
    Set s = ws.Shapes.AddFormControl(xlCheckBox, 200, 120, 110, 20)
    s.Name = "Part"
    s.ControlFormat.Value = xlMixed

    ' A drop down with the second item chosen: sel moves, val stays 0.
    Set s = ws.Shapes.AddFormControl(xlDropDown, 200, 150, 110, 20)
    s.Name = "Pick"
    s.ControlFormat.ListFillRange = "$A$1:$A$3"
    s.ControlFormat.LinkedCell = "$D$8"
    s.ControlFormat.Value = 2

    ' A list box, wired the same way.
    Set s = ws.Shapes.AddFormControl(xlListBox, 200, 180, 110, 50)
    s.Name = "Many"
    s.ControlFormat.ListFillRange = "$A$1:$A$3"
    s.ControlFormat.LinkedCell = "$D$9"

    ' Two option buttons, so firstButton shows up on the first of them.
    Set s = ws.Shapes.AddFormControl(xlOptionButton, 200, 240, 110, 20)
    s.Name = "First"
    s.TextFrame.Characters.Text = "Choose me"
    s.ControlFormat.LinkedCell = "$D$10"
    s.ControlFormat.Value = xlOn

    Set s = ws.Shapes.AddFormControl(xlOptionButton, 200, 270, 110, 20)
    s.Name = "Second"
    s.TextFrame.Characters.Text = "Or me"

    ' A spinner and a scroll bar, which carry val and a range.
    Set s = ws.Shapes.AddFormControl(xlSpinner, 200, 300, 20, 30)
    s.Name = "Step"
    s.ControlFormat.LinkedCell = "$D$11"
    s.ControlFormat.Min = 0
    s.ControlFormat.Max = 50
    s.ControlFormat.Value = 7

    Set s = ws.Shapes.AddFormControl(xlScrollBar, 200, 340, 110, 20)
    s.Name = "Slide"
    s.ControlFormat.LinkedCell = "$D$12"
    s.ControlFormat.Min = 0
    s.ControlFormat.Max = 100
    s.ControlFormat.Value = 25

    ' A group box and a label, wired to nothing, to prove absence reads
    ' as absence rather than as zero.
    Set s = ws.Shapes.AddFormControl(xlGroupBox, 340, 240, 130, 50)
    s.Name = "Set"
    s.TextFrame.Characters.Text = "A group"

    Set s = ws.Shapes.AddFormControl(xlLabel, 340, 300, 110, 20)
    s.Name = "Caption"
    s.TextFrame.Characters.Text = "A label"

    ' An ordinary AutoShape too, so the reader is made to tell a drawing
    ' shape from a control on one sheet.
    Set s = ws.Shapes.AddShape(1, 20, 20, 80, 40)
    s.Name = "Plain"

    ' Saved before the measuring, deliberately. OnAction reports the macro
    ' qualified by the workbook holding it, so reading it from an unsaved
    ' book records "Book8!Clicked" and the number changes every rebuild.
    Application.DisplayAlerts = False
    wb.SaveAs Target, 52
    Application.DisplayAlerts = True

    For Each s In ws.Shapes
        out = out & s.Name & "|" & s.Type & "|" & s.OnAction & "|" _
            & SafeLinked(s) & "|" & SafeList(s) & "|" & SafeValue(s) & vbLf
    Next s

    Build = out
End Function
"""

#: One of every autofilter criterion, each on its own sheet.
#:
#: Each kind is a different child element rather than a variation on one,
#: so a model built from the two kinds a typical file carries would be
#: wrong about the other four. ``links.xlsx`` has a value list and one
#: comparison; this adds the rest.
#:
#: One kind per sheet, because stacking them fails: after three criteria
#: almost no rows are left showing and Excel refuses a top-ten call that
#: has nothing to rank. That is the probe's problem rather than a limit
#: worth recording, and separate sheets avoid it.
#:
#: What each sheet is for is the row count it hides as much as the markup.
#: Excel stores a filter twice, the criteria and a ``hidden`` flag on each
#: row, and does not recompute the first when the workbook opens.
_BUILD_FILTERS = r"""
Private Function Safe(ByVal f As Object, ByVal which As Long) As String
    On Error Resume Next
    If which = 1 Then Safe = CStr(f.Criteria1)
    If which = 2 Then Safe = CStr(f.Criteria2)
    If which = 3 Then Safe = CStr(f.Operator)
    If Err.Number <> 0 Then Safe = "n/a"
    Err.Clear
End Function

Private Sub Fill(ByVal ws As Worksheet)
    Dim i As Long
    ws.Range("A1").Value = "Name"
    ws.Range("B1").Value = "Units"
    ws.Range("C1").Value = "When"
    For i = 2 To 21
        ws.Cells(i, 1).Value = "r" & i
        ws.Cells(i, 2).Value = (i * 7) Mod 50
        ws.Cells(i, 3).Value = DateSerial(2026, ((i - 2) Mod 12) + 1, 5)
    Next i
    ' One blank, so a blanks filter has something to keep.
    ws.Range("A15").ClearContents
End Sub

Private Function Report(ByVal ws As Worksheet) As String
    Dim i As Long
    Dim f As Object
    Dim hidden As Long
    Dim out As String

    out = ws.Name & ";" & CStr(ws.AutoFilterMode)
    If ws.AutoFilterMode Then
        For i = 1 To ws.AutoFilter.Filters.Count
            Set f = ws.AutoFilter.Filters(i)
            If f.On Then
                out = out & ";" & CStr(i) & ":" & Safe(f, 1) & _
                      ":" & Safe(f, 2) & ":" & Safe(f, 3)
            End If
        Next i
    End If
    For i = 2 To 21
        If ws.Rows(i).Hidden Then hidden = hidden + 1
    Next i
    Report = out & ";hidden=" & CStr(hidden)
End Function

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim out As String
    Dim names As Variant
    Dim i As Long

    Set wb = ActiveWorkbook
    names = Array("Values", "Compare", "Between", "TopTen", "Blanks", "Dynamic", _
                  "Numbers", "Dates")

    For i = LBound(names) To UBound(names)
        If wb.Worksheets.Count <= i Then wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)
        Set ws = wb.Worksheets(i + 1)
        ws.Name = names(i)
        Fill ws
    Next i

    ' A value list.
    wb.Worksheets("Values").Range("A1:C21").AutoFilter _
        Field:=1, Criteria1:=Array("r2", "r3", "r4"), Operator:=7

    ' One comparison.
    wb.Worksheets("Compare").Range("A1:C21").AutoFilter Field:=2, Criteria1:=">=10"

    ' Two comparisons joined by And.
    wb.Worksheets("Between").Range("A1:C21").AutoFilter _
        Field:=2, Criteria1:=">=5", Operator:=1, Criteria2:="<=40"

    ' Top ten items, on a sheet nothing else has narrowed.
    wb.Worksheets("TopTen").Range("A1:C21").AutoFilter _
        Field:=2, Criteria1:="3", Operator:=3

    ' Blanks, which is its own spelling rather than a value.
    wb.Worksheets("Blanks").Range("A1:C21").AutoFilter Field:=1, Criteria1:="="

    ' Above average, which stores the average Excel worked out.
    wb.Worksheets("Dynamic").Range("A1:C21").AutoFilter Field:=2, Operator:=11, Criteria1:=33

    ' A value list over numbers, which writes the plain number as text.
    wb.Worksheets("Numbers").Range("A1:C21").AutoFilter _
        Field:=2, Criteria1:=Array("14", "21", "28"), Operator:=7

    ' And over dates, which through the object model keeps the text it was
    ' given rather than writing a dateGroupItem.
    wb.Worksheets("Dates").Range("A1:C21").AutoFilter _
        Field:=3, Criteria1:=Array("1/5/2026", "3/5/2026"), Operator:=7

    For i = 1 To wb.Worksheets.Count
        out = out & Report(wb.Worksheets(i)) & "|"
    Next i

    Application.DisplayAlerts = False
    wb.SaveAs Target, 51
    Application.DisplayAlerts = True

    Build = out
End Function
"""

#: Notes, which Excel's object model still calls comments: plain, on two
#: lines, showing, resized, on the first row and column, with a word in
#: bold, and one beside a form control, since the two share a VML part.
#:
#: The reply is one line per note, with what Excel's object model says of
#: it. The text comes back with its line breaks spelled ``\n``, because the
#: harness turns a control character into a space.
_BUILD_COMMENTS = r"""
Private Function Report(ByVal c As Comment) As String
    Report = c.Parent.Parent.Name & "!" & c.Parent.Address(False, False) & "|" & _
             Replace(c.Text, vbLf, "\n") & "|" & c.Author & "|" & CStr(c.Visible) & "|" & _
             CStr(c.Shape.Left) & "|" & CStr(c.Shape.Top) & "|" & _
             CStr(c.Shape.Width) & "|" & CStr(c.Shape.Height)
End Function

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim both As Worksheet
    Dim c As Comment
    Dim out As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Notes"
    ws.Range("A1:D6").Value = 1

    ws.Range("C3").AddComment "plain note"
    Set c = ws.Range("E5").AddComment("two" & vbLf & "lines")
    c.Visible = True
    Set c = ws.Range("B8").AddComment("sized")
    c.Shape.Width = 200
    c.Shape.Height = 80
    ws.Range("A1").AddComment "corner"
    Set c = ws.Range("D10").AddComment("bold start")
    c.Shape.TextFrame.Characters(1, 4).Font.Bold = True

    Set both = wb.Worksheets.Add(After:=ws)
    both.Name = "Both"
    both.Buttons.Add 100, 30, 60, 20
    both.Range("B2").AddComment "beside a button"

    For Each ws In wb.Worksheets
        For Each c In ws.Comments
            out = out & Report(c) & vbLf
        Next c
    Next ws

    wb.Worksheets(1).Activate
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = out
End Function
"""

def _png(width: int, height: int, dpi: int | None) -> bytes:
    """A plain red PNG, with a ``pHYs`` chunk when ``dpi`` is given."""

    def chunk(kind: bytes, body: bytes) -> bytes:
        crc = zlib.crc32(kind + body) & 0xFFFFFFFF
        return struct.pack(">I", len(body)) + kind + body + struct.pack(">I", crc)

    rows = b"".join(b"\x00" + b"\xff\x00\x00" * width for _ in range(height))
    body = chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
    if dpi is not None:
        per_metre = round(dpi / 0.0254)
        body += chunk(b"pHYs", struct.pack(">IIB", per_metre, per_metre, 1))
    body += chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b"")
    return b"\x89PNG\r\n\x1a\n" + body


#: The one-pixel GIF every web page used to carry.
_GIF = (
    b"GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff!\xf9\x04\x01\x00\x00\x00\x00"
    b",\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;"
)

#: Pictures as ``Shapes.AddPicture`` puts them in: at their own size, one
#: at 144 dots to the inch, one stretched, a GIF, and one with alternative
#: text, the same image twice among them. VBA writes the images itself from
#: hex, so the recipe needs nothing on disk.
_BUILD_PICTURES = r"""
Private Function Written(ByVal Name As String, ByVal Hex As String) As String
    Dim f As Integer
    Dim i As Long
    Dim data() As Byte
    ReDim data(Len(Hex) \ 2 - 1)
    For i = 0 To UBound(data)
        data(i) = CByte("&H" & Mid(Hex, i * 2 + 1, 2))
    Next i
    Written = Environ("TEMP") & "\" & Name
    f = FreeFile
    Open Written For Binary Access Write As #f
    Put #f, , data
    Close #f
End Function

Private Function Report(ByVal s As Shape) As String
    Report = s.Name & "|" & CStr(s.Type) & "|" & CStr(s.Left) & "|" & CStr(s.Top) & "|" & _
             CStr(s.Width) & "|" & CStr(s.Height) & "|" & s.AlternativeText & "|" & CStr(s.LockAspectRatio)
End Function

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim s As Shape
    Dim plain As String
    Dim dense As String
    Dim dot As String
    Dim out As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Pictures"
    plain = Written("pyofficeeditor_plain.png", "{PLAIN}")
    dense = Written("pyofficeeditor_dense.png", "{DENSE}")
    dot = Written("pyofficeeditor_dot.gif", "{DOT}")

    Set s = ws.Shapes.AddPicture(plain, msoFalse, msoTrue, 100, 50, -1, -1)
    s.Name = "Natural"
    Set s = ws.Shapes.AddPicture(dense, msoFalse, msoTrue, 300, 50, -1, -1)
    s.Name = "Dense"
    Set s = ws.Shapes.AddPicture(plain, msoFalse, msoTrue, 300, 200, 60, 30)
    s.Name = "Stretched"
    Set s = ws.Shapes.AddPicture(dot, msoFalse, msoTrue, 100, 300, -1, -1)
    s.Name = "Dot"
    s.AlternativeText = "a single pixel"

    For Each s In ws.Shapes
        out = out & Report(s) & vbLf
    Next s
    Kill plain
    Kill dense
    Kill dot

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = out
End Function
""".replace("{PLAIN}", _png(120, 80, 96).hex().upper()).replace(
    "{DENSE}", _png(120, 80, 144).hex().upper()
).replace("{DOT}", _GIF.hex().upper())

#: Text in more than one font in one cell, made through ``Characters`` the
#: way a person makes it: a bold word, a red one, a bigger one, underline
#: and superscript, another typeface, a line break, and a cell whose own
#: font is bold. The reply is one line per cell with each font change Excel
#: reports: where it starts, then the typeface, size, bold, italic, colour,
#: underline and superscript.
_BUILD_RICHTEXT = r"""
Private Function Describe(ByVal c As Range) As String
    Dim i As Long
    Dim f As Font
    Dim out As String
    Dim last As String
    Dim now As String
    For i = 1 To Len(c.Value)
        Set f = c.Characters(i, 1).Font
        now = f.Name & "," & CStr(f.Size) & "," & CStr(f.Bold) & "," & CStr(f.Italic) & "," & _
              CStr(f.Color) & "," & CStr(f.Underline) & "," & CStr(f.Superscript)
        If now <> last Then out = out & CStr(i) & ":" & now & ";"
        last = now
    Next i
    Describe = out
End Function

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long
    Dim out As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Rich"

    ws.Range("A1").Value = "bold plain"
    ws.Range("A1").Characters(1, 4).Font.Bold = True
    ws.Range("A2").Value = "plain red end"
    ws.Range("A2").Characters(7, 3).Font.Color = RGB(255, 0, 0)
    ws.Range("A3").Value = "big and italic"
    ws.Range("A3").Characters(1, 3).Font.Size = 16
    ws.Range("A3").Characters(9, 6).Font.Italic = True
    ws.Range("A4").Value = "under over"
    ws.Range("A4").Characters(1, 5).Font.Underline = xlUnderlineStyleSingle
    ws.Range("A4").Characters(7, 4).Font.Superscript = True
    ws.Range("A5").Value = "font change"
    ws.Range("A5").Characters(6, 6).Font.Name = "Courier New"
    ws.Range("A6").Value = "two" & vbLf & "lines bold"
    ws.Range("A6").Characters(11, 4).Font.Bold = True
    ws.Range("A6").WrapText = True
    ws.Range("A7").Value = "all bold"
    ws.Range("A7").Font.Bold = True
    ws.Range("A7").Characters(1, 3).Font.Italic = True

    For i = 1 To 7
        out = out & "A" & CStr(i) & "|" & Describe(ws.Cells(i, 1)) & vbLf
    Next i

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = out
End Function
"""

#: Which fields of the reported line mean what.
_CONTROL_FIELDS = ("type", "macro", "linked_cell", "list_range", "value")

#: The same for a picture, after its name.
_PICTURE_FIELDS = ("type", "left", "top", "width", "height", "description", "locks_aspect")

#: The same for a note, after its sheet-qualified address.
_COMMENT_FIELDS = ("text", "author", "visible", "left", "top", "width", "height")


#: What a measured fixture's reply turns into: an entry per thing measured.
Answers = dict[str, dict[str, object]]


def build_answers(
    session: object,
    source: str,
    target: Path,
    label: str,
    *,
    force: bool,
    parse: Callable[[str], Answers],
) -> None:
    """Author a fixture and record what Excel's object model said about it.

    The measurements are the point of this one, so they are written beside
    the workbook instead of being read once and remembered. A reader that
    disagrees with them is wrong about Excel, which is the only authority
    that settles it. ``parse`` reads the fixture's own reply.
    """
    from pyvbaharness import ExcelSession

    assert isinstance(session, ExcelSession)
    answers_path = target.with_name(f"{target.stem}_answers.json")
    if target.exists() and answers_path.exists():
        if not force:
            print(f"  kept  {target.name} (already there; pass --force to rebuild)")
            return
        target.unlink()
        answers_path.unlink(missing_ok=True)

    session.reset_sheets()
    result = session.run_vba(source, proc="Build", args=(str(target),), timeout=180)
    if result.outcome != "passed":
        raise SystemExit(f"{label}: Excel refused the build ({result.outcome}): {result!r}")
    if not target.is_file():
        raise SystemExit(f"{label}: Excel reported success but {target} is not there.")

    answers = parse(str(result.value or ""))
    answers_path.write_text(
        json.dumps(answers, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    print(f"  wrote {target.name} ({target.stat().st_size} bytes) and {answers_path.name}")


def control_answers(reply: str) -> Answers:
    """One line per control: its name, then the fields of ``_CONTROL_FIELDS``."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        parts = line.split("|")
        name = parts[0]
        fields = (parts[1:] + [""] * len(_CONTROL_FIELDS))[: len(_CONTROL_FIELDS)]
        entry: dict[str, object] = {}
        for key, raw in zip(_CONTROL_FIELDS, fields, strict=True):
            if key in ("type", "value"):
                entry[key] = int(raw) if raw.strip("-").isdigit() else raw
            else:
                entry[key] = raw
        answers[name] = entry
    return answers


def richtext_answers(reply: str) -> Answers:
    """One line per cell: its address, then each font change as
    ``start:name,size,bold,italic,colour,underline,superscript;``."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        address, _, body = line.partition("|")
        runs: list[dict[str, object]] = []
        for piece in body.split(";"):
            if not piece:
                continue
            start, _, fields = piece.partition(":")
            name, size, bold, italic, color, underline, superscript = fields.split(",")
            runs.append({
                "start": int(start),
                "name": name,
                "size": float(size),
                "bold": bold == "True",
                "italic": italic == "True",
                "color": int(color),
                "underline": int(underline),
                "superscript": superscript == "True",
            })
        answers[address] = {"runs": runs}
    return answers


def picture_answers(reply: str) -> Answers:
    """One line per picture: its name, then the fields of ``_PICTURE_FIELDS``."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        parts = line.split("|")
        fields = (parts[1:] + [""] * len(_PICTURE_FIELDS))[: len(_PICTURE_FIELDS)]
        entry: dict[str, object] = {}
        for key, raw in zip(_PICTURE_FIELDS, fields, strict=True):
            if key == "type":
                entry[key] = int(raw)
            elif key in ("left", "top", "width", "height"):
                entry[key] = float(raw)
            elif key == "locks_aspect":
                entry[key] = raw == "-1"
            else:
                entry[key] = raw
        answers[parts[0]] = entry
    return answers


def comment_answers(reply: str) -> Answers:
    """One line per note: ``Sheet!A1``, then the fields of ``_COMMENT_FIELDS``."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        parts = line.split("|")
        fields = (parts[1:] + [""] * len(_COMMENT_FIELDS))[: len(_COMMENT_FIELDS)]
        entry: dict[str, object] = {}
        for key, raw in zip(_COMMENT_FIELDS, fields, strict=True):
            if key == "text":
                entry[key] = raw.replace("\\n", "\n")
            elif key == "visible":
                entry[key] = raw == "True"
            elif key in ("left", "top", "width", "height"):
                entry[key] = float(raw)
            else:
                entry[key] = raw
        answers[parts[0]] = entry
    return answers


def filter_answers(reply: str) -> Answers:
    """One entry per sheet: the name, whether the autofilter is on, each
    live filter's criteria and operator, and how many rows Excel hid. The
    last is the measurement that matters, because Excel does not recompute
    a filter when the workbook opens."""
    answers: Answers = {}
    for row in reply.split("|"):
        if not row.strip():
            continue
        parts = row.split(";")
        name, mode = parts[0], parts[1] == "True"
        hidden = 0
        filters: dict[str, dict[str, str]] = {}
        for piece in parts[2:]:
            if piece.startswith("hidden="):
                hidden = int(piece.partition("=")[2] or 0)
                continue
            index, _, rest = piece.partition(":")
            first, second, operator = [*rest.split(":"), "", "", ""][:3]
            filters[index] = {
                "criteria1": first,
                "criteria2": second,
                "operator": operator,
            }
        answers[name] = {"on": mode, "hidden": hidden, "filters": filters}
    return answers


def main() -> int:
    try:
        from pyvbaharness import ExcelSession
        from pyvbaharness.session import HarnessConfig
    except ImportError:
        print(
            'pyvbaharness is not installed. Install the live extra:\n'
            '    python -m pip install -e ".[dev]" --group live',
            file=sys.stderr,
        )
        return 2
    # Seconds to wait for another session to finish with Excel, as the live
    # gate waits; by default the build does not wait.
    wait = float(os.environ.get("LIVE_EXCEL_LOCK_WAIT", "0"))

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
    #: Fixtures whose measurements are recorded beside them, each with what
    #: reads its reply back.
    measured = [
        ("controls.xlsm", _BUILD_CONTROLS, control_answers),
        ("filters.xlsx", _BUILD_FILTERS, filter_answers),
        ("comments.xlsx", _BUILD_COMMENTS, comment_answers),
        ("pictures.xlsx", _BUILD_PICTURES, picture_answers),
        ("richtext.xlsx", _BUILD_RICHTEXT, richtext_answers),
    ]

    everything = [name for name, _ in wanted] + [name for name, _, _ in measured]
    everything += [f"{Path(name).stem}_answers.json" for name, _, _ in measured]
    if not force and all((FIXTURES / name).exists() for name in everything):
        print("every fixture is already there; nothing to do (pass --force to rebuild)")
        return 0

    FIXTURES.mkdir(parents=True, exist_ok=True)
    print(f"authoring fixtures in {FIXTURES}")
    with ExcelSession(HarnessConfig(lock_wait_s=wait)) as excel:
        for name, source in wanted:
            excel.new_workbook()
            build(excel, source, FIXTURES / name, name, force=force)
        for name, source, parse in measured:
            excel.new_workbook()
            build_answers(excel, source, FIXTURES / name, name, force=force, parse=parse)
    print("done")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
