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

import html
import json
import os
import posixpath
import re
import struct
import sys
import zipfile
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

#: Text XML cannot carry as it is, everywhere a workbook keeps text: cells
#: with control characters, a lone carriage return, a CRLF and a literal
#: "_x0041_", formula results and a formula naming a sheet called
#: "a_x0041_b", a note, a validation's messages, a hyperlink's tip, a
#: table's headers, a page header and a defined name's comment. The reply
#: is one line per piece of text: what it is, then the UTF-16 code of each
#: character Excel reports for it, in hex, so no character is lost on the
#: way. Triple single quotes, because the formulas hold three double ones.
_BUILD_ESCAPES = r'''
Private Function Codes(ByVal s As String) As String
    Dim i As Long
    Dim out As String
    For i = 1 To Len(s)
        out = out & Hex(AscW(Mid(s, i, 1))) & " "
    Next i
    Codes = Trim(out)
End Function

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim other As Worksheet
    Dim lo As ListObject
    Dim nm As Name
    Dim texts As Variant
    Dim headers As Variant
    Dim i As Long
    Dim out As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Text"
    Set other = wb.Worksheets.Add(After:=ws)
    other.Name = "a_x0041_b"
    other.Range("A1").Value = 7

    texts = Array("a" & Chr(1) & "b", "tab" & vbTab & "x", "cr" & vbCr & "x", "crlf" & vbCrLf & "x", _
                  "lf" & vbLf & "x", "_x0041_", "_X0041_", "_x004_", "a_x005F_b", _
                  "bell" & Chr(7) & "end" & Chr(31), "high" & ChrW(&HFFFE) & "x", "_x0001_" & Chr(1), _
                  "__x0041_", "_x00e9_")
    For i = 0 To UBound(texts)
        ws.Cells(i + 1, 1).Value = texts(i)
    Next i

    ws.Range("B1").Formula = "=""a""&CHAR(1)&""b""&CHAR(10)&""c""&CHAR(13)&""d"""
    ws.Range("B2").Formula = "=""_x0041_"""
    ws.Range("B3").Formula = "='a_x0041_b'!A1"

    ws.Range("C1").AddComment "note" & Chr(1) & "x" & vbLf & "y" & vbCr & "z _x0041_"

    With ws.Range("D1").Validation
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="1", Formula2:="9"
        .InputTitle = "ti" & Chr(1) & "t"
        .InputMessage = "in" & vbLf & "put" & Chr(1) & " _x0041_"
        .ErrorMessage = "er" & vbCr & "ror"
    End With

    ws.Hyperlinks.Add Anchor:=ws.Range("E1"), Address:="https://example.com/", _
                      ScreenTip:="tip" & Chr(1) & vbLf & "x", TextToDisplay:="link"

    headers = Array("head" & vbLf & "two", "_x0041_", "plain", "ctl" & Chr(1) & "x", _
                    "tab" & vbTab & "x", "cr" & vbCr & "x")
    For i = 0 To UBound(headers)
        ws.Cells(1, 7 + i).Value = headers(i)
        ws.Cells(2, 7 + i).Value = i
    Next i
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 7), ws.Cells(2, 7 + UBound(headers))), , xlYes)

    ws.PageSetup.CenterHeader = "head" & vbLf & "er _x0041_"

    Set nm = wb.Names.Add(Name:="nm", RefersTo:="=1")
    nm.Comment = "c" & vbLf & "_x0041_"

    For i = 0 To UBound(texts)
        out = out & "Text!A" & CStr(i + 1) & "|" & Codes(CStr(ws.Cells(i + 1, 1).Value)) & vbLf
    Next i
    For i = 1 To 3
        out = out & "Text!B" & CStr(i) & "|" & Codes(CStr(ws.Cells(i, 2).Value)) & vbLf
        out = out & "Text!B" & CStr(i) & " formula|" & Codes(ws.Cells(i, 2).Formula) & vbLf
    Next i
    out = out & "note|" & Codes(ws.Range("C1").Comment.Text) & vbLf
    out = out & "validation title|" & Codes(ws.Range("D1").Validation.InputTitle) & vbLf
    out = out & "validation message|" & Codes(ws.Range("D1").Validation.InputMessage) & vbLf
    out = out & "validation error|" & Codes(ws.Range("D1").Validation.ErrorMessage) & vbLf
    out = out & "hyperlink tip|" & Codes(ws.Hyperlinks(1).ScreenTip) & vbLf
    For i = 1 To lo.ListColumns.Count
        out = out & "table column " & CStr(i) & "|" & Codes(lo.ListColumns(i).Name) & vbLf
        out = out & "table header " & CStr(i) & "|" & Codes(CStr(lo.HeaderRowRange.Cells(1, i).Value)) & vbLf
    Next i
    out = out & "page header|" & Codes(ws.PageSetup.CenterHeader) & vbLf
    out = out & "sheet name|" & Codes(other.Name) & vbLf
    out = out & "name comment|" & Codes(nm.Comment) & vbLf

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = out
End Function
'''

#: Every built-in cell style Excel lists, each applied to a cell of its own
#: in column B beside its name, a style of the workbook's own in D1, and
#: two styles put over formatting a cell already had: "Good" over a bold,
#: centred cell showing two decimals, and "Currency" then italic. The reply
#: is a line per cell describing what Excel shows, and a line per built-in
#: style saying which aspects it sets.
_BUILD_STYLES = r'''
Private Function Describe(ByVal c As Range) As String
    Dim s As String
    Dim edges As Variant
    Dim e As Variant
    s = c.Style.Name & "|" & c.NumberFormat & "|" & c.Font.Name & "|" & CStr(c.Font.Size) & "|" & _
        CStr(c.Font.Bold) & "|" & CStr(c.Font.Italic) & "|" & CStr(c.Font.Color) & "|" & _
        CStr(c.Interior.Pattern) & "|" & CStr(c.Interior.Color) & "|" & CStr(c.HorizontalAlignment)
    edges = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
    For Each e In edges
        s = s & "|" & CStr(c.Borders(e).LineStyle) & "," & CStr(c.Borders(e).Weight) & "," & CStr(c.Borders(e).Color)
    Next e
    Describe = s
End Function

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim st As Style
    Dim mine As Style
    Dim i As Long
    Dim r As Long
    Dim out As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Styles"
    For Each st In wb.Styles
        If st.BuiltIn And st.Name <> "Normal" Then
            i = i + 1
            ws.Cells(i, 1).Value = st.Name
            ws.Cells(i, 2).Style = st.Name
            ws.Cells(i, 2).Value = 1234.5
            out = out & "includes|" & st.Name & "|" & CStr(st.IncludeNumber) & "," & CStr(st.IncludeFont) & "," & _
                  CStr(st.IncludePatterns) & "," & CStr(st.IncludeBorder) & "," & CStr(st.IncludeAlignment) & "," & _
                  CStr(st.IncludeProtection) & vbLf
        End If
    Next st

    Set mine = wb.Styles.Add("Mine")
    mine.IncludeNumber = False
    mine.IncludeAlignment = False
    mine.IncludeBorder = False
    mine.IncludeProtection = False
    mine.Font.Bold = True
    mine.Font.Color = RGB(0, 0, 255)
    mine.Interior.Color = RGB(255, 255, 0)
    ws.Range("D1").Style = "Mine"
    ws.Range("D1").Value = "mine"

    ws.Range("D2").Value = 1.5
    ws.Range("D2").NumberFormat = "0.00"
    ws.Range("D2").Font.Bold = True
    ws.Range("D2").HorizontalAlignment = xlCenter
    ws.Range("D2").Style = "Good"
    ws.Range("D3").Value = 2.5
    ws.Range("D3").Style = "Currency"
    ws.Range("D3").Font.Italic = True

    For r = 1 To i
        out = out & "B" & CStr(r) & "|" & Describe(ws.Cells(r, 2)) & vbLf
    Next r
    For r = 1 To 3
        out = out & "D" & CStr(r) & "|" & Describe(ws.Cells(r, 4)) & vbLf
    Next r

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = out
End Function
'''

#: Charts on a data sheet, on another sheet and on a chart sheet: columns
#: with a typed title, a line over two areas, a pie whose title is linked
#: to a cell, a scatter, and bars. After the workbook is saved, each edit a
#: chart's references have to follow is made on a fresh copy by Excel and
#: saved to the temporary folder. The references are read from those files
#: rather than from Excel's object model, which reports a deleted reference
#: by its old address until the file is reopened, and then refuses to
#: report the series at all. The reply is each chart's series formulas and
#: title before any edit, and where each edited file is.
_BUILD_CHARTS = r'''
Private Function Report(ByVal wb As Workbook, ByVal label As String) As String
    Dim out As String
    Dim co As ChartObject
    Dim ws As Worksheet
    Dim ch As Chart
    For Each ws In wb.Worksheets
        For Each co In ws.ChartObjects
            out = out & label & "|" & ws.Name & "/" & co.Name & "|" & Series(co.Chart) & vbLf
        Next co
    Next ws
    For Each ch In wb.Charts
        out = out & label & "|" & ch.Name & "|" & Series(ch) & vbLf
    Next ch
    Report = out
End Function

Private Function Series(ByVal ch As Chart) As String
    Dim s As Series
    Dim out As String
    For Each s In ch.SeriesCollection
        out = out & s.Formula & ";"
    Next s
    out = out & "|"
    If ch.HasTitle Then out = out & ch.ChartTitle.Formula
    Series = out
End Function

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rep As Worksheet
    Dim co As ChartObject
    Dim cs As Chart
    Dim r As Long
    Dim out As String
    Dim ops As Variant
    Dim i As Long
    Dim again As Workbook
    Dim edited As String
    Dim base As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Data"
    ws.Range("A1:D1").Value = Array("Month", "Sales", "Costs", "Profit")
    For r = 2 To 6
        ws.Cells(r, 1).Value = "M" & CStr(r - 1)
        ws.Cells(r, 2).Value = r * 10
        ws.Cells(r, 3).Value = r * 4
        ws.Cells(r, 4).Formula = "=B" & CStr(r) & "-C" & CStr(r)
    Next r

    Set co = ws.ChartObjects.Add(Left:=300, Top:=10, Width:=300, Height:=200)
    co.Name = "Columns"
    co.Chart.ChartType = xlColumnClustered
    co.Chart.SetSourceData Source:=ws.Range("A1:C6")
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = "Plain title"

    Set co = ws.ChartObjects.Add(Left:=300, Top:=220, Width:=300, Height:=200)
    co.Name = "Line"
    co.Chart.ChartType = xlLine
    co.Chart.SetSourceData Source:=ws.Range("A1:A6,D1:D6")

    Set rep = wb.Worksheets.Add(After:=ws)
    rep.Name = "Report"
    Set co = rep.ChartObjects.Add(Left:=10, Top:=10, Width:=300, Height:=200)
    co.Name = "Pie"
    co.Chart.ChartType = xlPie
    co.Chart.SetSourceData Source:=ws.Range("A1:B6")
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Formula = "=Data!$B$1"

    Set co = rep.ChartObjects.Add(Left:=10, Top:=220, Width:=300, Height:=200)
    co.Name = "Scatter"
    co.Chart.ChartType = xlXYScatter
    co.Chart.SetSourceData Source:=ws.Range("B1:C6")

    Set cs = wb.Charts.Add(After:=wb.Sheets(wb.Sheets.Count))
    cs.Name = "Bars"
    cs.ChartType = xlBarClustered
    cs.SetSourceData Source:=ws.Range("A1:B6")

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51

    ' The fixture stays open as the active workbook, so everything is read
    ' from a copy of it opened afresh, as each edit's result is. The copies
    ' go in the temporary folder, so a build that fails leaves none beside
    ' the fixtures.
    base = Environ("TEMP") & "\pyofficeeditor_charts_base.xlsx"
    wb.SaveCopyAs base
    Set again = Workbooks.Open(base)
    out = Report(again, "base")
    again.Close SaveChanges:=False
    out = out & "file|base|" & base & vbLf
    ops = Array("insert rows 3:4", "insert row 1", "delete row 4", "delete rows 2:6", "delete row 1", _
                "insert column B", "delete column C", "delete column A", "rename Data")
    For i = 0 To UBound(ops)
        Set again = Workbooks.Open(base)
        Select Case ops(i)
            Case "insert rows 3:4": again.Worksheets("Data").Rows("3:4").Insert
            Case "insert row 1": again.Worksheets("Data").Rows(1).Insert
            Case "delete row 4": again.Worksheets("Data").Rows(4).Delete
            Case "delete rows 2:6": again.Worksheets("Data").Rows("2:6").Delete
            Case "delete row 1": again.Worksheets("Data").Rows(1).Delete
            Case "insert column B": again.Worksheets("Data").Columns("B").Insert
            Case "delete column C": again.Worksheets("Data").Columns("C").Delete
            Case "delete column A": again.Worksheets("Data").Columns("A").Delete
            Case "rename Data": again.Worksheets("Data").Name = "Q1 Data"
        End Select
        edited = Environ("TEMP") & "\pyofficeeditor_charts_" & CStr(i) & ".xlsx"
        again.SaveAs Filename:=edited, FileFormat:=51
        again.Close SaveChanges:=False
        out = out & "file|" & ops(i) & "|" & edited & vbLf
    Next i
    Application.DisplayAlerts = True
    Build = out
End Function
'''

#: Data!A1:C11 feeding two pivot tables, each from a cache of its own:
#: Summary on a sheet of its own at A3, and Beside on the data sheet at F3,
#: where the edits reach it. Each edit is made by Excel on a fresh copy and
#: saved to the temporary folder, or reported refused; the pivot tables are
#: read from those files. The reply is where each file is.
_BUILD_PIVOTS = r'''
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pv As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim r As Long
    Dim ops As Variant
    Dim i As Long
    Dim again As Workbook
    Dim out As String
    Dim base As String
    Dim edited As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Data"
    ws.Range("A1:C1").Value = Array("Region", "Product", "Sales")
    For r = 2 To 11
        ws.Cells(r, 1).Value = Choose((r Mod 3) + 1, "North", "South", "West")
        ws.Cells(r, 2).Value = Choose((r Mod 2) + 1, "Tea", "Coffee")
        ws.Cells(r, 3).Value = r * 10
    Next r
    Set pv = wb.Worksheets.Add(After:=ws)
    pv.Name = "Pivot"
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data!R1C1:R11C3")
    Set pt = pc.CreatePivotTable(TableDestination:=pv.Range("A3"), TableName:="Summary")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Total", xlSum
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data!R1C1:R11C3")
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("F3"), TableName:="Beside")
    pt.PivotFields("Product").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Sum", xlSum

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    base = Environ("TEMP") & "\pyofficeeditor_pivots_base.xlsx"
    wb.SaveCopyAs base
    out = "file|base|" & base & vbLf

    ops = Array("insert row 1", "insert row 3", "insert row 7", "insert rows 5:6", "delete row 3", _
                "delete row 4", "delete rows 3:6", "delete rows 2:7", "delete rows 1:11", "delete rows 1:2", _
                "delete row 7", "insert column A", "insert column F", "insert column H", "delete column B", _
                "delete column F", "delete columns E:H", "insert row 1 on Pivot", "rename Data")
    For i = 0 To UBound(ops)
        Set again = Workbooks.Open(base)
        Set ws = again.Worksheets("Data")
        On Error Resume Next
        Select Case ops(i)
            Case "insert row 1": ws.Rows(1).Insert
            Case "insert row 3": ws.Rows(3).Insert
            Case "insert row 7": ws.Rows(7).Insert
            Case "insert rows 5:6": ws.Rows("5:6").Insert
            Case "delete row 3": ws.Rows(3).Delete
            Case "delete row 4": ws.Rows(4).Delete
            Case "delete rows 3:6": ws.Rows("3:6").Delete
            Case "delete rows 2:7": ws.Rows("2:7").Delete
            Case "delete rows 1:11": ws.Rows("1:11").Delete
            Case "delete rows 1:2": ws.Rows("1:2").Delete
            Case "delete row 7": ws.Rows(7).Delete
            Case "insert column A": ws.Columns("A").Insert
            Case "insert column F": ws.Columns("F").Insert
            Case "insert column H": ws.Columns("H").Insert
            Case "delete column B": ws.Columns("B").Delete
            Case "delete column F": ws.Columns("F").Delete
            Case "delete columns E:H": ws.Columns("E:H").Delete
            Case "insert row 1 on Pivot": again.Worksheets("Pivot").Rows(1).Insert
            Case "rename Data": ws.Name = "Q1 Data"
        End Select
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            again.Close SaveChanges:=False
            out = out & "refused|" & ops(i) & vbLf
        Else
            On Error GoTo 0
            edited = Environ("TEMP") & "\pyofficeeditor_pivots_" & CStr(i) & ".xlsx"
            again.SaveAs Filename:=edited, FileFormat:=51
            again.Close SaveChanges:=False
            out = out & "file|" & ops(i) & "|" & edited & vbLf
        End If
    Next i
    Application.DisplayAlerts = True
    Build = out
End Function
'''

#: GETPIVOTDATA's measurements, on sheet Q: each formula below with Excel's
#: answer cached beside it, which is what the engine is held to. They read
#: pivot tables in each layout Excel offers, over Data!A1:F13 and the small
#: table in Data!H1:J5: nested, across, with two values, tabular, outline
#: with subtotals at the bottom or none, filtered, without totals, with
#: values down the side, with renamed items and fields, hidden items, dates
#: grouped by years and months, numbers in bins, a calculated field, and
#: values shown as a share of the total. Grouping and a calculated field
#: change a cache for every pivot table sharing it, so those tables have
#: caches of their own. Excel removes personal information as it saves.
_PIVOT_DATA_FORMULAS = [
    'GETPIVOTDATA("Sales",PT1!$A$3)',
    'GETPIVOTDATA("Sum of Sales",PT1!$A$3,"Region","East")',
    'GETPIVOTDATA("sales",PT1!$A$3,"region","east")',
    'GETPIVOTDATA("Sales",PT1!$A$3,"Region","Nope")',
    'GETPIVOTDATA("Qty",PT1!$A$3)',
    'GETPIVOTDATA("Sales",PT1!$B$5)',
    'GETPIVOTDATA("Sales",PT1!$A$1)',
    'GETPIVOTDATA("Sales",PT1!$A$3:$B$8,"Region","West")',
    'GETPIVOTDATA("Sales",PT1!$E$3,"Region","East","Product","Pen")',
    'GETPIVOTDATA("Sales",PT1!$E$3,"Product","Pen")',
    'GETPIVOTDATA("Sales",PT1!$E$3,"Region","East")',
    'GETPIVOTDATA("Sales",PT1!$E$3,"Product","Pen","Region","East")',
    'GETPIVOTDATA("Sales",PT1!$J$3,"Year",2023)',
    'GETPIVOTDATA("Sales",PT1!$J$3,"Year","2023")',
    'GETPIVOTDATA("Sales",PT1!$J$3,"Region","East","Year",2024)',
    'GETPIVOTDATA("Sales",PT1!$J$3,"Region","East")',
    'GETPIVOTDATA("Sales",PT1!$J$3)',
    'GETPIVOTDATA("Sales",PT1!$A$30,"Region","East","Year",2024)',
    'GETPIVOTDATA("Qty",PT1!$A$30,"Region","East","Year",2024)',
    'GETPIVOTDATA("Sum of Qty",PT1!$A$30,"Year",2023)',
    'GETPIVOTDATA("Qty",PT1!$A$30)',
    'GETPIVOTDATA("Sales",PT1!$J$30,"Region","East","Product","Ink")',
    'GETPIVOTDATA("Sales",PT1!$J$30,"Region","West")',
    'GETPIVOTDATA("Sales",PT1!$J$30)',
    'GETPIVOTDATA("Sales",PT1!$A$62,"Product","Pen")',
    'GETPIVOTDATA("Sales",PT1!$A$62,"Year",2023)',
    'GETPIVOTDATA("Sales",PT1!$A$62,"Year",2024)',
    'GETPIVOTDATA("Sales",PT1!$A$62,"Year",2023,"Product","Ink")',
    'GETPIVOTDATA("Sales",PT1!$A$62)',
    'GETPIVOTDATA("Sales",PT1!$J$62,"Region","East","Year",2023)',
    'GETPIVOTDATA("Sales",PT1!$J$62,"Region","East")',
    'GETPIVOTDATA("Sales",PT1!$J$62)',
    'GETPIVOTDATA("Sales",PT1!$A$80,"Region","North")',
    'GETPIVOTDATA("Count of Product",PT1!$A$80,"Region","North")',
    'GETPIVOTDATA("Product",PT1!$A$80)',
    'GETPIVOTDATA("Revenue",PT1!$J$80,"Region","Orient")',
    'GETPIVOTDATA("Revenue",PT1!$J$80,"Region","East")',
    'GETPIVOTDATA("Sales",PT1!$J$80)',
    'GETPIVOTDATA("Sum of Sales",PT1!$J$80)',
    'GETPIVOTDATA("Qty",PT1!$A$105,"Year",2024)',
    'GETPIVOTDATA("Qty",PT1!$A$105,"Year","2024")',
    'GETPIVOTDATA("Qty",PT1!$A$105,"Year",2024.0)',
    'GETPIVOTDATA("Qty",PT1!$A$105,"Year",TRUE)',
    'GETPIVOTDATA("Sales",PT2!$A$3,"Product","Ink","Region","South")',
    'GETPIVOTDATA("Sales",PT2!$A$3,"Product","Ink","Region","East")',
    'GETPIVOTDATA("Sales",PT2!$A$3,"Region","South")',
    'GETPIVOTDATA("Sales",PT2!$J$3,"Region","East")',
    'GETPIVOTDATA("Sales",PT2!$J$3,"Region","East","Product","Pen")',
    'GETPIVOTDATA("Sales",PT2!$A$30,"Region","East")',
    'GETPIVOTDATA("Sales",PT2!$A$30,"Region","East","Product","Ink")',
    'GETPIVOTDATA("Sales",PT2!$A$30)',
    'GETPIVOTDATA("Sales",PT2!$J$32,"Year",2023)',
    'GETPIVOTDATA("Sales",PT2!$J$32,"Product","Pen")',
    'GETPIVOTDATA("Sales",PT2!$J$30)',
    'GETPIVOTDATA("Sales",PT2!$K$30)',
    'GETPIVOTDATA("Sales",PT2!$J$31)',
    'GETPIVOTDATA("Qty",PT2!$A$60,"Day",DATE(2023,1,15))',
    'GETPIVOTDATA("Qty",PT2!$A$60,"Day","1/15/2023")',
    'GETPIVOTDATA("Qty",PT2!$A$60,"Day",44941)',
    'GETPIVOTDATA("Qty",PT2!$A$60,"Day","2023-01-15")',
    'GETPIVOTDATA("Qty",PT2!$A$60,"Day",DATE(2024,12,13))',
    'GETPIVOTDATA("Qty",PT2!$A$60)',
    'GETPIVOTDATA("Sales",PT2!$J$60,"Region","East")',
    'GETPIVOTDATA("Sum of Sales",PT2!$J$60,"Region","East")',
    'GETPIVOTDATA("Average of Sales",PT2!$J$60,"Region","East")',
    'GETPIVOTDATA("Sales",PT2!$A$95,"Area","East")',
    'GETPIVOTDATA("Sales",PT2!$A$95,"Region","East")',
    'GETPIVOTDATA("Sales",PT2!$J$95,"Region","West")',
    'GETPIVOTDATA("Sales",PT2!$J$95,"Region","East")',
    'GETPIVOTDATA("Sales",PT2!$J$95)',
    'GETPIVOTDATA("Sales",PT2!$A$110,"Region","East","Product","Ink")',
    'GETPIVOTDATA("Qty",PT2!$A$110,"Region","East","Product","Ink")',
    'GETPIVOTDATA("Sales",PT2!$A$110,"Region","East")',
    'GETPIVOTDATA("Qty",PT2!$A$110,"Region","East")',
    'GETPIVOTDATA("Qty",PT2!$A$110)',
    'GETPIVOTDATA("Sales",PT2!$A$110,"Product","Ink")',
    'GETPIVOTDATA("Sales",PT2!$A$3,"Product","Pen","Nope","x")',
    'GETPIVOTDATA("Sales",PT2!$A$3,"Product")',
    'GETPIVOTDATA("Sales",PT2!$A$3,"Day",DATE(2023,1,15))',
    'GETPIVOTDATA("Sales",PT2!$A$3:$Z$120)',
    'GETPIVOTDATA("Sales","PT2!A3")',
    'GETPIVOTDATA(1,PT2!$A$3)',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",2023)',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)","2023")',
    'GETPIVOTDATA("Qty",PT3!$A$3,"years (day)",2024)',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",DATE(2023,6,1))',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",2023.5)',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",2024,"Months (Day)","Mar")',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",2024,"Months (Day)","mar")',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",2024,"Months (Day)",3)',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",2024,"Months (Day)","March")',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",2024,"Months (Day)",DATE(2024,3,10))',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Months (Day)","Jan")',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Day",DATE(2023,1,15))',
    'GETPIVOTDATA("Qty",PT3!$A$3)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales","0-99")',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",0)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",50)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",99)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",99.5)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",100)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",150)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",350)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",399)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",400)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",-5)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales","0")',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",TRUE)',
    'GETPIVOTDATA("Profit",PT3!$A$40,"Region","East")',
    'GETPIVOTDATA("Share",PT3!$J$40,"Region","East")',
    'GETPIVOTDATA("Score",PT3!$A$60,"Person","Bob")',
    'GETPIVOTDATA("Score",PT3!$A$60,"Person","Ann")',
    'GETPIVOTDATA("Score",PT3!$A$60,"Person","Cat")',
    'GETPIVOTDATA("Score",PT3!$A$60,"Team","Red")',
    'GETPIVOTDATA("Score",PT3!$A$60,"Team","Red","Person","Ann")',
    'GETPIVOTDATA("Score",PT3!$J$60,"Person","Bob")',
    'GETPIVOTDATA("Score",PT3!$J$60,"Person","Ann")',
    'GETPIVOTDATA("Score",PT3!$J$60,"Team","Blue")',
    'GETPIVOTDATA("Score",PT3!$J$60)',
    'GETPIVOTDATA("Score",PT3!$A$75,"Person","Bob")',
    'GETPIVOTDATA("Score",PT3!$A$75,"Team","Red")',
    'GETPIVOTDATA("Score",PT3!$A$75,"Team","Red","Person","Bob")',
    'GETPIVOTDATA("Score",PT3!$A$75)',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Months (Day)","Feb")',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Months (Day)",2)',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",2023,"Months (Day)",1)',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",2024,"Months (Day)",3.5)',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",2024,"Months (Day)",12)',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",2024,"Months (Day)",13)',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",2024,"Months (Day)",0)',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",2022)',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",2024.9)',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",2025)',
    'GETPIVOTDATA("Qty",PT3!$A$3,"Years (Day)",-2023)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",-1)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",-50)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",-99)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",-100)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",-150)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",0.5)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",1)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",10)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",100.5)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",200)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",250)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",300)',
    'GETPIVOTDATA("Qty",PT3!$J$3,"Sales",-0.5)',
]

_PIVOT_DATA_TEMPLATE = r'''
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sh As Worksheet
    Dim q As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim data As Variant
    Dim i As Long

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Data"
    ws.Range("A1:F1").Value = Array("Region", "Product", "Year", "Qty", "Sales", "Day")
    data = Array( _
        Array("East", "Pen", 2023, 3, 30, DateSerial(2023, 1, 15)), _
        Array("East", "Ink", 2023, 5, 50, DateSerial(2023, 2, 20)), _
        Array("East", "Pen", 2024, 2, 24, DateSerial(2024, 3, 10)), _
        Array("West", "Pen", 2023, 4, 44, DateSerial(2023, 4, 5)), _
        Array("West", "Book", 2024, 1, 120, DateSerial(2024, 5, 6)), _
        Array("West", "Ink", 2024, 6, 66, DateSerial(2024, 6, 7)), _
        Array("North", "Book", 2023, 2, 200, DateSerial(2023, 7, 8)), _
        Array("North", "Pen", 2024, 8, 88, DateSerial(2024, 8, 9)), _
        Array("North", "Ink", 2023, 7, 77, DateSerial(2023, 9, 10)), _
        Array("South", "Book", 2024, 3, 330, DateSerial(2024, 10, 11)), _
        Array("South", "Pen", 2023, 9, 99, DateSerial(2023, 11, 12)), _
        Array("East", "Book", 2024, 2, 220, DateSerial(2024, 12, 13)))
    For i = 0 To UBound(data)
        ws.Range(ws.Cells(i + 2, 1), ws.Cells(i + 2, 6)).Value = data(i)
    Next i
    ws.Range("H1:J1").Value = Array("Team", "Person", "Score")
    ws.Range("H2:J2").Value = Array("Red", "Ann", 1)
    ws.Range("H3:J3").Value = Array("Red", "Bob", 2)
    ws.Range("H4:J4").Value = Array("Blue", "Cat", 4)
    ws.Range("H5:J5").Value = Array("Blue", "Ann", 8)

    Set sh = wb.Worksheets.Add(After:=ws)
    sh.Name = "PT1"
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data!R1C1:R13C6")
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("A3"), TableName:="Basic")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", xlSum
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("E3"), TableName:="Nested")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.PivotFields("Product").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", xlSum
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("J3"), TableName:="Cross")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.PivotFields("Year").Orientation = xlColumnField
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", xlSum
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("A30"), TableName:="Two")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.PivotFields("Year").Orientation = xlColumnField
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", xlSum
    pt.AddDataField pt.PivotFields("Qty"), "Sum of Qty", xlSum
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("J30"), TableName:="Tabular")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.PivotFields("Product").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", xlSum
    pt.RowAxisLayout xlTabularRow
    pt.PivotFields("Region").LayoutSubtotalLocation = xlAtBottom
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("A62"), TableName:="Filtered")
    pt.PivotFields("Year").Orientation = xlPageField
    pt.PivotFields("Product").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", xlSum
    pt.PivotFields("Year").CurrentPage = "2023"
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("J62"), TableName:="NoTotals")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.PivotFields("Year").Orientation = xlColumnField
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", xlSum
    pt.ColumnGrand = False
    pt.RowGrand = False
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("A80"), TableName:="ValuesDown")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", xlSum
    pt.AddDataField pt.PivotFields("Product"), "Count of Product", xlCount
    pt.DataPivotField.Orientation = xlRowField
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("J80"), TableName:="Renamed")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Revenue", xlSum
    pt.PivotFields("Region").PivotItems("East").Caption = "Orient"
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("A105"), TableName:="Numbers")
    pt.PivotFields("Year").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Qty"), "Sum of Qty", xlSum

    Set sh = wb.Worksheets.Add(After:=sh)
    sh.Name = "PT2"
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("A3"), TableName:="Holes")
    pt.PivotFields("Product").Orientation = xlRowField
    pt.PivotFields("Region").Orientation = xlColumnField
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", xlSum
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("J3"), TableName:="OutlineBottom")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.PivotFields("Product").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", xlSum
    pt.RowAxisLayout xlOutlineRow
    pt.PivotFields("Region").LayoutSubtotalLocation = xlAtBottom
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("A30"), TableName:="NoSubtotals")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.PivotFields("Product").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", xlSum
    pt.PivotFields("Region").Subtotals(1) = False
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("J32"), TableName:="AllPages")
    pt.PivotFields("Year").Orientation = xlPageField
    pt.PivotFields("Product").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", xlSum
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("A60"), TableName:="Dates")
    pt.PivotFields("Day").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Qty"), "Sum of Qty", xlSum
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("J60"), TableName:="SameSource")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", xlSum
    pt.AddDataField pt.PivotFields("Sales"), "Average of Sales", xlAverage
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("A95"), TableName:="Caption")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", xlSum
    pt.PivotFields("Region").Caption = "Area"
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("J95"), TableName:="Hidden")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", xlSum
    pt.PivotFields("Region").PivotItems("West").Visible = False
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("A110"), TableName:="ValuesBottom")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.PivotFields("Product").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", xlSum
    pt.AddDataField pt.PivotFields("Qty"), "Sum of Qty", xlSum
    pt.DataPivotField.Orientation = xlRowField
    pt.RowAxisLayout xlTabularRow
    pt.PivotFields("Region").LayoutSubtotalLocation = xlAtBottom

    Set sh = wb.Worksheets.Add(After:=sh)
    sh.Name = "PT3"
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data!R1C1:R13C6")
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("A3"), TableName:="Grouped")
    pt.PivotFields("Day").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Qty"), "Sum of Qty", xlSum
    pt.PivotFields("Day").DataRange.Cells(1).Group Start:=True, End:=True, _
        Periods:=Array(False, False, False, False, True, False, True)
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data!R1C1:R13C6")
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("J3"), TableName:="Bins")
    pt.PivotFields("Sales").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Qty"), "Sum of Qty", xlSum
    pt.PivotFields("Sales").DataRange.Cells(1).Group Start:=0, End:=399, By:=100
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data!R1C1:R13C6")
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("A40"), TableName:="Calculated")
    pt.CalculatedFields.Add "Profit", "=Sales*0.2"
    pt.PivotFields("Region").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Profit"), "Sum of Profit", xlSum
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("J40"), TableName:="Percent")
    pt.PivotFields("Region").Orientation = xlRowField
    With pt.AddDataField(pt.PivotFields("Sales"), "Share", xlSum)
        .Calculation = xlPercentOfTotal
    End With
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data!R1C8:R5C10")
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("A60"), TableName:="Teams")
    pt.PivotFields("Team").Orientation = xlRowField
    pt.PivotFields("Person").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Score"), "Sum of Score", xlSum
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("J60"), TableName:="TeamsAcross")
    pt.PivotFields("Team").Orientation = xlColumnField
    pt.PivotFields("Person").Orientation = xlColumnField
    pt.AddDataField pt.PivotFields("Score"), "Sum of Score", xlSum
    pt.ColumnGrand = False
    Set pt = pc.CreatePivotTable(TableDestination:=sh.Range("A75"), TableName:="TeamsBare")
    pt.PivotFields("Team").Orientation = xlRowField
    pt.PivotFields("Person").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Score"), "Sum of Score", xlSum
    pt.PivotFields("Team").Subtotals(1) = False
    pt.RowGrand = False

    Set q = wb.Worksheets.Add(After:=sh)
    q.Name = "Q"
    ' GETPIVOTDATA formulas
    Application.CalculateFull
    wb.RemovePersonalInformation = True
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = ""
End Function
'''

#: What-if data tables made by Range.Table over a small loan model: rates
#: down a column against two formulas, years along a row, both at once, a
#: chain through another sheet and back, a formula for an input cell, a
#: branch the input decides, and a blank, an error, text and a formula
#: among the values tried. The model is built twice, on Model with Side
#: beside it and on Moved with Beside, the second with other years and
#: principal: set Model's to Moved's and recalculate, and every table has
#: to come to what Excel cached on Moved. Excel removes personal
#: information as it saves.
_BUILD_DATA_TABLES = r'''
Private Sub Model(ByVal ws As Worksheet, ByVal sd As Worksheet, ByVal years As Double, ByVal principal As Double)
    Dim i As Long

    ws.Range("A1").Value = "Rate"
    ws.Range("B1").Value = 0.05
    ws.Range("A2").Value = "Years"
    ws.Range("B2").Value = years
    ws.Range("A3").Value = "Principal"
    ws.Range("B3").Value = principal
    ws.Range("A4").Value = "Double rate"
    ws.Range("B4").Formula = "=B1*2"
    ws.Range("A5").Value = "Payment"
    ws.Range("B5").Formula = "=PMT(B1/12,B2*12,-B3)"
    ws.Range("B6").Formula = "=B5*B2*12"
    ws.Range("B7").Formula = "=B6-B3"
    ws.Range("B8").Formula = "=" & sd.Name & "!A1+1"
    ws.Range("B9").Formula = "=B4*1000+B3/1000"
    ws.Range("B10").Formula = "=IF(B1>0.06,B6,B7)"
    sd.Range("A1").Formula = "=" & ws.Name & "!B7/2"

    ws.Range("E2").Formula = "=B5"
    ws.Range("F2").Formula = "=B6"
    For i = 0 To 5
        ws.Cells(3 + i, 4).Value = 0.03 + i * 0.01
    Next i
    ws.Range("D2:F8").Table ColumnInput:=ws.Range("B1")

    ws.Range("H3").Formula = "=B5"
    For i = 0 To 4
        ws.Cells(2, 9 + i).Value = 10 + i * 5
    Next i
    ws.Range("H2:M3").Table RowInput:=ws.Range("B2")

    ws.Range("D11").Formula = "=B5"
    For i = 0 To 4
        ws.Cells(12 + i, 4).Value = 0.03 + i * 0.01
        ws.Cells(11, 5 + i).Value = 10 + i * 5
    Next i
    ws.Range("D11:I16").Table RowInput:=ws.Range("B2"), ColumnInput:=ws.Range("B1")

    ws.Range("L11").Formula = "=B8"
    ws.Range("M11").Formula = "=B9"
    ws.Range("N11").Formula = "=B10"
    For i = 0 To 4
        ws.Cells(12 + i, 11).Value = 0.04 + i * 0.01
    Next i
    ws.Range("K11:N16").Table ColumnInput:=ws.Range("B1")
    ws.Range("Q11").Formula = "=B9"
    ws.Range("P12").Value = 0.1
    ws.Range("P13").Value = 0.2
    ws.Range("P11:Q13").Table ColumnInput:=ws.Range("B4")

    ws.Range("E20").Formula = "=B5"
    ws.Range("D21").Value = 0.04
    ws.Range("D23").Formula = "=NA()"
    ws.Range("D24").Value = "text"
    ws.Range("D25").Formula = "=0.01*ROW()/5"
    ws.Range("D20:E25").Table ColumnInput:=ws.Range("B1")
End Sub

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Model"
    wb.Worksheets.Add(After:=ws).Name = "Side"
    wb.Worksheets.Add(After:=wb.Worksheets("Side")).Name = "Moved"
    wb.Worksheets.Add(After:=wb.Worksheets("Moved")).Name = "Beside"
    Model wb.Worksheets("Model"), wb.Worksheets("Side"), 30, 250000
    Model wb.Worksheets("Moved"), wb.Worksheets("Beside"), 25, 300000
    Application.CalculateFull
    wb.RemovePersonalInformation = True
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = ""
End Function
'''

#: The recipe, each formula written into Q's column A a row at a time.
_BUILD_PIVOT_DATA = _PIVOT_DATA_TEMPLATE.replace(
    "    ' GETPIVOTDATA formulas\n",
    "".join(
        f'    q.Cells({row}, 1).Formula2 = "={formula.replace(chr(34), chr(34) * 2)}"\n'
        for row, formula in enumerate(_PIVOT_DATA_FORMULAS, start=1)
    ),
)

#: One chart of each kind Excel's Insert Chart makes, added the way it adds
#: one, with ``Shapes.AddChart2`` and its default style, all from
#: Data!A1:C6, and a column chart with a title typed in. They are the
#: markup a chart written here is held to. The reply is each chart's name,
#: its kind and its series formulas.
_BUILD_CHART_KINDS = r'''
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sh As Shape
    Dim kinds As Variant
    Dim labels As Variant
    Dim i As Long
    Dim r As Long
    Dim s As Series
    Dim out As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Data"
    ws.Range("A1:C1").Value = Array("Month", "Sales", "Costs")
    For r = 2 To 6
        ws.Cells(r, 1).Value = "M" & CStr(r - 1)
        ws.Cells(r, 2).Value = r * 10
        ws.Cells(r, 3).Value = r * 4
    Next r
    kinds = Array(xlColumnClustered, xlBarClustered, xlLine, xlLineMarkers, xlPie, xlDoughnut, xlXYScatter, _
                  xlArea, xlColumnClustered)
    labels = Array("column", "bar", "line", "lineMarkers", "pie", "doughnut", "scatter", "area", "column titled")
    For i = 0 To UBound(kinds)
        Set sh = ws.Shapes.AddChart2(-1, kinds(i), 250, 20 + 230 * i, 360, 216)
        sh.Chart.SetSourceData Source:=ws.Range("A1:C6")
        If labels(i) = "column titled" Then
            sh.Chart.HasTitle = True
            sh.Chart.ChartTitle.Text = "Sales by month"
        End If
        out = out & sh.Name & "|" & labels(i) & "|"
        For Each s In sh.Chart.SeriesCollection
            out = out & s.Formula & ";"
        Next s
        out = out & vbLf
    Next i

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = out
End Function
'''

#: What each field of a described cell is.
_STYLE_FIELDS = (
    "style", "number_format", "font_name", "font_size", "bold", "italic", "font_color",
    "pattern", "fill_color", "horizontal", "left", "top", "bottom", "right",
)

#: Which fields of the reported line mean what.
_CONTROL_FIELDS = ("type", "macro", "linked_cell", "list_range", "value")

#: The same for a picture, after its name.
_PICTURE_FIELDS = ("type", "left", "top", "width", "height", "description", "locks_aspect")

#: The same for a note, after its sheet-qualified address.
_COMMENT_FIELDS = ("text", "author", "visible", "left", "top", "width", "height")


#: Cells for each of the ten error-checking rules, and a report of what
#: Range.Errors says of every cell. Each sheet's cells are spaced so a
#: case meets no other: the rules about formulas beside formulas, ranges
#: beside numbers and empty cells read their neighbours.
_BUILD_ERROR_CHECKS = r'''
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    wb.Worksheets(1).Name = "Text"
    wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)).Name = "Values"
    wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)).Name = "Regions"
    wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)).Name = "Empty"
    wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)).Name = "Tables"
    wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)).Name = "Ignored"
    wb.Names.Add Name:="NameNA", RefersTo:="=NA()"
    wb.Names.Add Name:="Four", RefersTo:="=Regions!$A$15"
    FillText wb.Worksheets("Text")
    FillValues wb.Worksheets("Values")
    FillRegions wb.Worksheets("Regions")
    FillEmpty wb.Worksheets("Empty")
    FillTables wb.Worksheets("Tables")
    FillIgnored wb.Worksheets("Ignored")
    Application.CalculateFull
    Build = Report(wb)
    Application.DisplayAlerts = False
    wb.RemovePersonalInformation = True
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
End Function

Private Sub WriteDown(ws As Worksheet, ByVal col As Long, ByVal texts As String)
    Dim parts() As String, i As Long
    parts = Split(texts, "|")
    For i = 0 To UBound(parts)
        ws.Cells(i + 1, col).Formula = parts(i)
    Next i
End Sub

Private Sub FillText(ws As Worksheet)
    ' Numbers stored as text, and text that is not a number.
    WriteDown ws, 1, "'123|'1,000|'$5|'5%|'(5)|'1 1/2|'1e5|'0123|' 123|'1,0000|'0,123|'$5%|'5-|'1,2|'abc|'TRUE|'1/2|'12:30|'1/1/2020"
    ' Dates with a two-digit year, and near misses.
    WriteDown ws, 3, "'1/1/99|'1-Jan-99|'Jan 99|'99 Jan|'Jan 1 99|'1 1/0|'1/0|'12/31/29|' 1/1/99|'Feb Jan 99|'1/287/99|'Jan, 99|'1/1/99 |'1/1/99 12:00|'1.1.99|'99/1/1|'Jan 0|'1/1/099|'Jan,99|'1/ 1/99"
    ' Strings in formulas, two rows apart.
    ws.Range("E1").Formula = "=""1/1/99"""
    ws.Range("E3").Formula = "=YEAR(""1/1/31"")"
    ws.Range("E5").Formula = "=DATEVALUE(""1/1/1931"")"
    ws.Range("E7").Formula = "=""1 1/0"""
    ws.Range("E9").Formula = "=COUNTA({""1/1/99"",""x""})"
    ws.Range("E11").Formula = "=""123"""
End Sub

Private Sub FillValues(ws As Worksheet)
    Dim i As Long
    ws.Range("Z10").Value = 1
    ws.Range("Z11").Value = 2
    ' Formulas whose value is an error, NA() among them.
    Dim errors As Variant
    errors = Array("=1/0", "=NA()", "=#N/A", "=IF(TRUE,NA(),1)", "=VLOOKUP(99,Z10:Z11,1,FALSE)", _
        "=IFERROR(1/0,NA())", "=NameNA", "#N/A", "=IF(TRUE,1/0,NA())")
    For i = 0 To UBound(errors)
        ws.Cells(1 + 2 * i, 1).Formula = errors(i)
    Next i
    ' Unlocked cells.
    ws.Range("C1").Locked = False
    ws.Range("C1").Formula = "=1+1"
    ws.Range("C3").Locked = False
    ws.Range("C3").Value = 5
    ws.Range("C5").Formula = "=1+1"
    ' Formats that mislead: a date, a number, and a number shown as text.
    ws.Range("X1").Value = 45000
    ws.Range("X1").NumberFormat = "m/d/yyyy"
    ws.Range("X2").Value = 5
    ws.Range("X3").Value = 7
    ws.Range("X3").NumberFormat = "@"
    Dim shown As Variant, formats As Variant
    shown = Array("=X1", "=+X1", "=(X1)", "=X2", "=X3", "=-X1", "=X1", "=X1+0")
    formats = Array("General", "General", "General", "m/d/yyyy", "m/d/yyyy", "0.00", "h:mm", "General")
    For i = 0 To UBound(shown)
        ws.Cells(1 + 2 * i, 5).Formula = shown(i)
        ws.Cells(1 + 2 * i, 5).NumberFormat = formats(i)
    Next i
End Sub

Private Sub FillRegions(ws As Worksheet)
    Dim i As Long
    ' A formula unlike the two it sits between, down a column and along a row.
    For i = 1 To 5
        ws.Cells(i, 1).Value = i
        ws.Cells(i, 2).Formula = "=A" & i & "*2"
        ws.Cells(8, i).Value = i
        ws.Cells(9, i).Formula = "=" & Chr(64 + i) & "8*2"
    Next i
    ws.Range("B3").Formula = "=A3*3"
    ws.Range("C9").Formula = "=C8*3"
    ws.Range("D1").Formula = "=A1*2"
    ws.Range("D2").Formula = "=A2 *2"
    ws.Range("D3").Formula = "=A3*2"
    ' Ranges beside a number they leave out, and the ones spared.
    For i = 12 To 15
        ws.Cells(i, 1).Value = i
    Next i
    ws.Range("B12").Formula = "=SUM(A12:A14)"
    ws.Range("C13").Formula = "=SUMIF(A12:A14,"">0"")"
    ws.Range("D14").Formula = "=SUM(A$12:A$14)"
    ws.Range("E15").Formula = "=SUM(A12:A14)+A15"
    ws.Range("F16").Formula = "=SUM(A12:A14)+Four"
    ws.Range("A20").Value = 1
    ws.Range("A21").Formula = "'x"
    ws.Range("A22").Value = 3
    ws.Range("A23").Value = 4
    ws.Range("B20").Formula = "=SUM(A20:A22)"
End Sub

Private Sub FillEmpty(ws As Worksheet)
    Dim i As Long
    ws.Range("A1").Formula = "=Y100"
    ws.Range("A3").Formula = "=SUM(Y1:Y2)"
    ws.Range("A5").Formula = "=SUM(C:C)"
    ' A range with a gap in a row that holds nothing, and one whose row does.
    For i = 5 To 9
        If i <> 7 Then ws.Cells(i, 3).Value = i
    Next i
    ws.Range("E5").Formula = "=SUM(C5:C9)"
    For i = 21 To 25
        If i <> 23 Then ws.Cells(i, 3).Value = i
    Next i
    ws.Range("AA23").Value = 1
    ws.Range("E21").Formula = "=SUM(C21:C25)"
    ' A range running past the last row the sheet uses.
    For i = 40 To 42
        ws.Cells(i, 8).Value = i
    Next i
    ws.Range("G40").Formula = "=SUM(H40:H45)"
End Sub

Private Sub FillTables(ws As Worksheet)
    Dim lo As ListObject, i As Long
    ws.Range("A1").Value = "a"
    ws.Range("B1").Value = "b"
    ws.Range("C1").Value = "c"
    ws.Range("D1").Value = "d"
    For i = 2 To 6
        ws.Cells(i, 1).Value = i
        ws.Cells(i, 2).Value = i * 10
    Next i
    ws.Range("D2").Value = "x"
    ws.Range("D3").Value = "X"
    ws.Range("D4").Value = "z"
    Set lo = ws.ListObjects.Add(1, ws.Range("A1:D6"), , 1)
    lo.Name = "Checked"
    ' A calculated column with a value and two cells of another formula.
    lo.ListColumns(3).DataBodyRange.Formula = "=[@a]*2"
    ws.Range("C4").Value = 7
    ws.Range("C5:C6").Formula = "=[@a]*3"
    ' Validation in the table: a whole number, and a list.
    ws.Range("B4").Value = 500
    With ws.Range("B2:B6").Validation
        .Delete
        .Add Type:=1, AlertStyle:=1, Operator:=1, Formula1:="1", Formula2:="100"
    End With
    With ws.Range("D2:D5").Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Formula1:="x,y"
        .IgnoreBlank = True
    End With
    With ws.Range("D6").Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Formula1:="x,y"
        .IgnoreBlank = False
    End With
    ' The same validation outside a table is not this rule's business.
    ws.Range("G2").Value = 500
    With ws.Range("G2").Validation
        .Delete
        .Add Type:=1, AlertStyle:=1, Operator:=1, Formula1:="1", Formula2:="100"
    End With
    FillKinds ws
End Sub

Private Sub FillKinds(ws As Worksheet)
    ' Each kind of validation over a value of its own: type, the formulas,
    ' and the value typed. #R# is the value's row, for a relative formula.
    Dim spec As Variant, parts() As String, i As Long, r As Long, lo As ListObject
    spec = Array("1|1|10|5", "1|1|10|'5", "1|1|10|5.5", "1|1|10|TRUE", "3|a,b,c||A", "3|a,b,c|| b", _
        "3|=$K$1:$K$3||X", "3|=$K$1:$K$3||'7", "3|=$K$1:$K$3||TRUE", "4|43831|44196|43831.5", _
        "4|43831|44196|44197", "5|0.375|0.708333333333333|0.5", "5|0.375|0.708333333333333|45000.5", _
        "6|2|4|TRUE", "6|2|4|12345", "6|2|4|1.5", "7|=B#R#>0||-5", "7|=B#R#>0||abc", "7|=5||1", "7|=FALSE||1")
    ws.Range("K1").Value = "x"
    ws.Range("K2").Value = 7
    ws.Range("K3").Value = True
    ws.Range("A10").Value = "k"
    ws.Range("B10").Value = "v"
    ' The values first: a formula typed into a table may fill its column.
    For i = 0 To UBound(spec)
        parts = Split(spec(i), "|")
        ws.Cells(11 + i, 1).Value = i
        ws.Cells(11 + i, 2).Formula = parts(3)
    Next i
    Set lo = ws.ListObjects.Add(1, ws.Range(ws.Cells(10, 1), ws.Cells(11 + UBound(spec), 2)), , 1)
    lo.Name = "Kinds"
    ' A validation's relative references are read from the active cell.
    ws.Activate
    For i = 0 To UBound(spec)
        parts = Split(spec(i), "|")
        r = 11 + i
        ws.Cells(r, 2).Select
        With ws.Cells(r, 2).Validation
            .Delete
            If parts(2) = "" Then
                .Add Type:=CLng(parts(0)), AlertStyle:=1, Operator:=1, Formula1:=Replace(parts(1), "#R#", CStr(r))
            Else
                .Add Type:=CLng(parts(0)), AlertStyle:=1, Operator:=1, Formula1:=parts(1), Formula2:=parts(2)
            End If
        End With
    Next i
End Sub

Private Sub FillIgnored(ws As Worksheet)
    ws.Range("B2").Formula = "'5"
    ws.Range("B3").Formula = "'6"
    ws.Range("B2").Errors(3).Ignore = True
    ws.Range("B3").Errors(3).Ignore = True
    ws.Range("B5").Formula = "=1/0"
    ws.Range("B5").Errors(1).Ignore = True
    ws.Range("B7").Formula = "'1/1/99"
End Sub

Private Function Report(wb As Workbook) As String
    Dim ws As Worksheet, c As Range, i As Long, flags As String, out As String, flagged As Boolean
    For Each ws In wb.Worksheets
        For Each c In ws.UsedRange.Cells
            flags = ""
            flagged = False
            For i = 1 To 10
                If c.Errors(i).Value Then
                    flags = flags & "1"
                    flagged = True
                Else
                    flags = flags & "0"
                End If
            Next i
            If flagged Or Not IsEmpty(c.Value) Or c.HasFormula Then
                out = out & ws.Name & "!" & c.Address(False, False) & vbTab & flags & vbLf
            End If
        Next c
    Next ws
    Report = out
End Function
'''


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


#: The error-checking rules in the order Range.Errors numbers them, named
#: as a worksheet's <ignoredErrors> names them.
_ERROR_RULES = (
    "evalError", "twoDigitTextYear", "numberStoredAsText", "formula", "formulaRange",
    "unlockedFormula", "emptyCellReference", "listDataValidation", "calculatedColumn", "misleadingFormat",
)  # fmt: skip


def error_check_answers(reply: str) -> Answers:
    """One line per cell, ``Sheet!A1`` and then a 0 or 1 for each rule,
    recorded as the rules that catch the cell."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        cell, flags = line.split("\t")
        answers[cell] = {"rules": [rule for rule, flag in zip(_ERROR_RULES, flags, strict=True) if flag == "1"]}
    return answers


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


def chart_kind_answers(reply: str) -> Answers:
    """One line per chart: its name, the kind it was made as, and its
    series formulas."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        name, kind, series = line.split("|")
        answers[name] = {"kind": kind, "series": [formula for formula in series.split(";") if formula]}
    return answers


def pivot_answers(reply: str) -> Answers:
    """For the file as built and for each edit Excel made: whether Excel
    refused it, and otherwise each pivot table's location and the source of
    the cache it reads, with how many caches the workbook kept. Read from
    the saved files, which are removed once read."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        parts = line.split("|")
        if parts[0] == "refused":
            answers[parts[1]] = {"refused": True}
            continue
        _, state, where = parts
        path = Path(where)
        answers[state] = {"refused": False, **_pivot_state(path)}
        path.unlink()
    return answers


def _pivot_state(path: Path) -> dict[str, object]:
    """Each pivot table's location and its cache's source, and how many
    caches there are, read with nothing but the archive and a few
    patterns."""
    tables: dict[str, dict[str, str]] = {}
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        for name in names:
            if not re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name):
                continue
            text = package.read(name).decode("utf-8")
            table = re.search(r'<pivotTableDefinition\b[^>]*\bname="([^"]*)"', text)
            location = re.search(r'<location\b[^>]*\bref="([^"]*)"', text)
            rels = package.read(f"xl/pivotTables/_rels/{name.rpartition('/')[2]}.rels").decode("utf-8")
            target = re.search(r'Target="([^"]*pivotCacheDefinition[^"]*)"', rels)
            source = ""
            if target is not None:
                cache = posixpath.normpath(posixpath.join("xl/pivotTables", target.group(1)))
                found = re.search(r"<worksheetSource\b[^>]*>", package.read(cache).decode("utf-8"))
                if found is not None:
                    sheet = re.search(r'\bsheet="([^"]*)"', found.group(0))
                    ref = re.search(r'\bref="([^"]*)"', found.group(0))
                    source = f"{html.unescape(sheet.group(1)) if sheet else ''}!{ref.group(1) if ref else ''}"
            if table is not None and location is not None:
                tables[html.unescape(table.group(1))] = {"location": location.group(1), "source": source}
        caches = sum(1 for name in names if re.fullmatch(r"xl/pivotCache/pivotCacheDefinition\d+\.xml", name))
    return {"tables": tables, "caches": caches}


def chart_answers(reply: str) -> Answers:
    """Each chart's series formulas and title before any edit, as Excel's
    object model gave them, and each chart's references, as Excel wrote
    them, in the file saved before any edit and in each saved after one.
    The files are in the temporary folder and are removed once read."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        parts = line.split("|")
        if parts[0] == "file":
            _, state, where = parts
            path = Path(where)
            for chart, references in _chart_references(path).items():
                answers[f"{state}|{chart}"] = {"references": references}
            path.unlink()
            continue
        state, chart, series, title = parts
        answers[f"reported {state}|{chart}"] = {
            "series": [formula for formula in series.split(";") if formula],
            "title": title,
        }
    return answers


def _chart_references(path: Path) -> dict[str, list[str]]:
    """Every chart in a workbook, as ``Sheet/Name`` or a chart sheet's
    name, and the references in its part in the order the part holds them.

    Read with nothing but the archive and a few patterns, following each
    sheet's drawing to the chart parts it names, so the answers owe nothing
    to the library they are there to check.
    """

    def text(package: zipfile.ZipFile, part: str) -> str:
        return package.read(part).decode("utf-8")

    def targets(package: zipfile.ZipFile, part: str) -> dict[str, str]:
        folder, _, name = part.rpartition("/")
        rels = f"{folder}/_rels/{name}.rels"
        if rels not in package.namelist():
            return {}
        found: dict[str, str] = {}
        for entry in re.findall(r"<Relationship\b[^>]*>", text(package, rels)):
            identifier = re.search(r'\bId="([^"]*)"', entry)
            target = re.search(r'\bTarget="([^"]*)"', entry)
            if identifier and target:
                found[identifier.group(1)] = posixpath.normpath(posixpath.join(folder, target.group(1)))
        return found

    charts: dict[str, list[str]] = {}
    with zipfile.ZipFile(path) as package:
        books = targets(package, "xl/workbook.xml")
        for entry in re.findall(r"<sheet\b[^>]*>", text(package, "xl/workbook.xml")):
            name = html.unescape(re.search(r'\bname="([^"]*)"', entry).group(1))  # type: ignore[union-attr]
            sheet_part = books[re.search(r'\br:id="([^"]*)"', entry).group(1)]  # type: ignore[union-attr]
            for drawing in targets(package, sheet_part).values():
                if "/drawings/" not in drawing or not drawing.endswith(".xml"):
                    continue
                drawn = targets(package, drawing)
                for frame in re.findall(r"<xdr:graphicFrame\b.*?</xdr:graphicFrame>", text(package, drawing), re.S):
                    chart = re.search(r'<c:chart\b[^>]*\br:id="([^"]*)"', frame)
                    frame_name = re.search(r'<xdr:cNvPr\b[^>]*\bname="([^"]*)"', frame)
                    if chart is None or frame_name is None:
                        continue
                    references = re.findall(r"<(?:\w+:)?f>(.*?)</(?:\w+:)?f>", text(package, drawn[chart.group(1)]))
                    key = name if "/chartsheets/" in sheet_part else f"{name}/{html.unescape(frame_name.group(1))}"
                    charts[key] = [html.unescape(reference) for reference in references]
    return charts


def style_answers(reply: str) -> Answers:
    """A line per built-in style, ``includes|name|`` and six flags, and a
    line per cell, its address and then the fields of ``_STYLE_FIELDS``,
    kept as the text Excel gave them."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        parts = line.split("|")
        if parts[0] == "includes":
            names = ("number_format", "font", "fill", "border", "alignment", "protection")
            flags = [flag == "True" for flag in parts[2].split(",")]
            answers[f"includes {parts[1]}"] = dict(zip(names, flags, strict=True))
            continue
        answers[parts[0]] = dict(zip(_STYLE_FIELDS, parts[1:], strict=True))
    return answers


def escape_answers(reply: str) -> Answers:
    """One line per piece of text: what it is, then its UTF-16 codes in
    hex. A surrogate pair makes one character, as it does in Excel."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        key, _, codes = line.partition("|")
        units = "".join(chr(int(code, 16)) for code in codes.split())
        text = units.encode("utf-16-le", "surrogatepass").decode("utf-16-le", "surrogatepass")
        answers[key] = {"text": text}
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
        ("pivotdata.xlsx", _BUILD_PIVOT_DATA),
        ("datatables.xlsx", _BUILD_DATA_TABLES),
    ]
    #: Fixtures whose measurements are recorded beside them, each with what
    #: reads its reply back.
    measured = [
        ("controls.xlsm", _BUILD_CONTROLS, control_answers),
        ("filters.xlsx", _BUILD_FILTERS, filter_answers),
        ("comments.xlsx", _BUILD_COMMENTS, comment_answers),
        ("pictures.xlsx", _BUILD_PICTURES, picture_answers),
        ("richtext.xlsx", _BUILD_RICHTEXT, richtext_answers),
        ("escapes.xlsx", _BUILD_ESCAPES, escape_answers),
        ("styles.xlsx", _BUILD_STYLES, style_answers),
        ("charts.xlsx", _BUILD_CHARTS, chart_answers),
        ("pivots.xlsx", _BUILD_PIVOTS, pivot_answers),
        ("chartkinds.xlsx", _BUILD_CHART_KINDS, chart_kind_answers),
        ("errorchecks.xlsx", _BUILD_ERROR_CHECKS, error_check_answers),
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
