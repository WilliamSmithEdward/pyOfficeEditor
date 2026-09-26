"""Measure how Excel spells a sheet's name in a formula, and what it writes
when a sheet a 3D reference names is renamed.

Two things are measured. First, which names Excel quotes: one sheet, S,
is renamed through every name below in turn, and after each rename the
formulas ``=S!B3`` and ``=SUM(S:T!B3)`` on another sheet are read back
through ``Range.Formula``, so each name is seen alone and at the start of
a span whose other end needs no quotes. The names are the ones the rules
were worked out from, logicals and names that read as references in A1 or
R1C1, digits and punctuation, and characters past ASCII swept a Unicode
block at a time, each inside a name and at its start.

Second, renames: each case opens a workbook of its own, adds sheets in tab
order, writes formulas on a sheet called Summary and a defined name, reads
them back as entered, renames one sheet through ``Worksheet.Name``, reads
them again, and closes the workbook unsaved.

Run this on a Windows machine with Excel installed:

    python -m pip install -e ".[dev]" --group live
    python scripts/measure_sheet_names.py

It writes ``tests/fixtures/excel/sheet_names.json``, replacing what is
there, and nothing else.
"""

from __future__ import annotations

import datetime as dt
import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FIXTURE = ROOT / "tests" / "fixtures" / "excel" / "sheet_names.json"

#: Names the rules were worked out from.
NAMES = [
    # Logicals, and names that only start or end like one.
    "TRUE", "FALSE", "True", "false", "tRuE", "TRUEX", "XTRUE", "FALSE1", "TRUE_", "_TRUE", "TRU", "FALS",
    # R1C1: the reference alone, followed by more, and past its limits.
    "R", "C", "RC", "rc", "R1", "C1", "R1C1", "R2C3", "r2c3", "RC1", "R1C", "R1X", "C1X", "R1C1X", "R1CX",
    "RC1X", "R2x", "C2x", "R2D2", "R1C1R1C1", "R01", "C01", "R001", "C001", "R1C01", "R0", "C0", "R0C0",
    "RC0", "R0C", "R0C1", "R1C0", "RCX", "RCa", "RCC", "RRC", "Rx", "Cx", "rX", "cX", "R_", "C_", "R.", "C.",
    "XRC1", "R1048576", "R1048577", "R1048576C16384", "R1048577C1", "R1C16384", "R1C16385", "R1C99999",
    "R1C2147483647", "C16384", "C16385", "C65536", "C99999", "C1048576", "C1048577", "C2147483647",
    "RC16385", "RC1048576", "RC1048577",
    # A1: the reference, and past its limits.
    "A1", "a1", "A01", "A001", "B0010", "XFD1048576", "XFE1", "XFD1048577", "A0", "AAA1", "AAAA1", "ZZ99",
    "RX1", "Rx1", "Cx1", "RR1", "CC1", "A", "AB", "ZZZZ", "A1B", "AB12C",
    # LOG10 is a function's name as well as a cell.
    "LOG1", "LOG9", "LOG10", "log10", "LOG11", "LOG19", "LOG20", "LOG99", "LOG100", "log100", "LOG101",
    "LOG1000", "LOG10X", "LOGA10", "XLOG10", "LOF10", "LOH10", "LN10", "EXP10", "SUM1", "SUM10", "T1", "N1",
    # Function names and the like.
    "SUM", "IF", "LET", "NA", "REF", "x", "_xlfn", "ATAN2", "DAYS360", "IMLOG2", "DEC2",
    # Digits, and what a name may start and end with.
    "1", "3rd", "1.5", "1E3", ".a", "_a", "_1", "a.", "a_", "a1.", "a.b",
    # ASCII punctuation inside a name.
    "a b", "a-b", "a+b", "a$b", "a#b", "a&b", "a%b", "a@b", "a!b", "a(b", "a)b", "a,b", "a;b", "a=b",
    "a<b", "a>b", "a{b", "a}b", "a~b", "a^b", "a`b", 'a"b', "a'b", "a|b",
    # Letters and digits past ASCII.
    "M\u00e4r", "\u00c9t\u00e9", "\u00dc1", "\u00e41", "\u65e5\u672c", "\u03a9mega", "\u00f1", "\u00df",
    "\u0130", "\uff21\uff22", "\uff11", "\uff41\uff11",
]  # fmt: skip

#: Blocks of characters past ASCII swept a character at a time.
BLOCKS = [
    (0x00A0, 0x00FF),  # Latin-1 Supplement
    (0x2000, 0x206F),  # General Punctuation
    (0x20A0, 0x20C0),  # Currency Symbols
    (0x2100, 0x214F),  # Letterlike Symbols
    (0x2190, 0x219F),  # Arrows, the first of them
    (0x2200, 0x221F),  # Mathematical Operators, the first of them
    (0x3000, 0x303F),  # CJK Symbols and Punctuation
    (0xFF01, 0xFF65),  # Halfwidth and Fullwidth Forms, to the half-width ones
]


def swept() -> list[str]:
    """Each character of the blocks inside a name and at its start."""
    names: list[str] = []
    for low, high in BLOCKS:
        for point in range(low, high + 1):
            names += [f"a{chr(point)}b", f"{chr(point)}a"]
    return names


#: Renames: the sheets after Summary in tab order, the formulas written on
#: it, the sheet renamed and its new name. The first formula's argument is
#: also the defined name Span.
RENAMES = [
    (["Jan", "Feb", "Mar"], ["=SUM(Jan:Mar!B3)", "=Jan!B3"], "Jan", "January"),
    (["Jan", "Feb", "Mar"], ["=SUM(Jan:Mar!B3)", "=Mar!B3"], "Mar", "March"),
    (["Jan", "Feb", "Mar"], ["=SUM(Jan:Mar!B3)", "=Feb!B3"], "Feb", "February"),
    (["Jan", "Feb", "Mar"], ["=SUM(Jan:Mar!B3)", "=Mar!B3"], "Mar", "Q1 End"),
    (["Jan", "Feb", "Mar"], ["=SUM(Jan:Mar!B3)", "=Jan!B3"], "Jan", "Q1 Start"),
    (["Jan 1", "Feb", "Mar"], ["=SUM('Jan 1:Mar'!B3)", "='Jan 1'!B3"], "Jan 1", "Start"),
    (["Jan 1", "Feb", "Mar 1"], ["=SUM('Jan 1:Mar 1'!B3)", "='Mar 1'!B3"], "Jan 1", "Start"),
    (["Jan 1", "Feb", "Mar"], ["=SUM('Jan 1:Mar'!B3)", "=Mar!B3"], "Mar", "End"),
    (["Jan", "Feb", "Mar"], ["=SUM(Jan:Mar!B3:C4)", "=SUM(Jan!B3:C4)"], "Jan", "January"),
    (["Jan", "Feb", "Mar"], ["=SUM(Jan:Mar!B:B)", "=SUM(Jan!3:3)"], "Mar", "Q1 End"),
    (["Jan", "Feb", "Mar"], ["=SUM(Jan:Mar!B3)+Jan!B3+Mar!B3", "=Jan!B3&Mar!B3"], "Jan", "Q1 Start"),
    (["Jan", "Feb", "Mar"], ["=SUM(Jan:Mar!B3)", "=Mar!B3"], "Mar", "A1"),
    (["Jan", "Feb", "Mar"], ["=SUM(Jan:Mar!B3)", "=Mar!B3"], "Mar", "O'Brien"),
    (["Jan", "Feb", "Mar"], ["=SUM(Jan:Mar!B3)", "=Mar!B3"], "Mar", "M\u00e4r"),
    (["Jan", "Feb", "Mar"], ["=SUM(Jan:Mar!B3)", "=Mar!B3"], "Mar", "3rd"),
    (["Jan", "Feb", "Mar"], ["=SUM(Jan:Mar!B3)", "=Mar!B3"], "Mar", "R2C3"),
    (["Jan", "Feb", "Mar"], ["=SUM(Jan:Mar!B3)", "=Mar!B3"], "Mar", "TRUE"),
    (["Jan", "Feb", "Mar"], ["=SUM(Jan:Mar!B3)", "=Mar!B3"], "Mar", "a.b"),
    (["O'Brien", "Feb", "Mar"], ["=SUM('O''Brien:Mar'!B3)", "='O''Brien'!B3"], "O'Brien", "OBrien"),
    (["Jan", "Feb", "Mar"], ['=SUM(Jan:Mar!B3)&"Jan:Mar!B3"', '=COUNTIF(Jan!B3,"Jan")'], "Jan", "Q1 Start"),
    (["Jan", "Feb", "Mar", "MyJan"], ["=SUM(Jan:Mar!B3)+MyJan!B3", "=SUM(MyJan!B3)"], "Jan", "January"),
    (["M\u00e4r", "Apr", "May"], ["=SUM(M\u00e4r:May!B3)", "=M\u00e4r!B3"], "M\u00e4r", "Q2 Start"),
    (["A", "Feb", "B"], ["=SUM(A:B!B3)", "=A!B3"], "A", "R1"),
]

_MEASURE = r'''
Private Function Rename(ByVal Sheets_ As String, ByVal Formulas As String, ByVal Old As String, ByVal New_ As String) As String
    Dim wb As Workbook, names() As String, cells_() As String, i As Long, summary As Worksheet
    Dim entered As String, after As String
    Set wb = Workbooks.Add
    Set summary = wb.Worksheets(1)
    summary.Name = "Summary"
    names = Split(Sheets_, "{")
    For i = 0 To UBound(names)
        wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)).Name = names(i)
        wb.Worksheets(names(i)).Range("B3").Value = i + 1
    Next i
    cells_ = Split(Formulas, "{")
    For i = 0 To UBound(cells_)
        summary.Cells(i + 1, 1).Formula = cells_(i)
        entered = entered & summary.Cells(i + 1, 1).Formula & "{"
    Next i
    wb.Names.Add Name:="Span", RefersTo:="=" & Split(Split(cells_(0), "(")(1), ")")(0)
    entered = entered & wb.Names("Span").RefersTo
    wb.Worksheets(Old).Name = New_
    For i = 0 To UBound(cells_)
        after = after & summary.Cells(i + 1, 1).Formula & "{"
    Next i
    after = after & wb.Names("Span").RefersTo
    wb.Close SaveChanges:=False
    Rename = entered & "}" & after
End Function

Public Function Measure(ByVal Names As String, ByVal Renames As String) As String
    Dim wb As Workbook, summary As Worksheet, each_() As String, fields() As String, i As Long, out As String
    Set wb = ActiveWorkbook
    Set summary = wb.Worksheets(1)
    summary.Name = "Summary"
    wb.Worksheets.Add(After:=summary).Name = "S"
    wb.Worksheets.Add(After:=wb.Worksheets("S")).Name = "T"
    summary.Range("A1").Formula = "=S!B3"
    summary.Range("A2").Formula = "=SUM(S:T!B3)"
    each_ = Split(Names, "@@")
    For i = 0 To UBound(each_)
        On Error Resume Next
        wb.Worksheets(2).Name = each_(i)
        If Err.Number <> 0 Then
            out = out & "refused" & vbLf
            Err.Clear
        ElseIf wb.Worksheets(2).Name <> each_(i) Then
            out = out & "renamed" & vbLf
        Else
            out = out & summary.Range("A1").Formula & vbTab & summary.Range("A2").Formula & vbLf
        End If
        On Error GoTo 0
    Next i
    each_ = Split(Renames, "@@")
    For i = 0 To UBound(each_)
        fields = Split(each_(i), "`")
        out = out & Rename(fields(0), fields(1), fields(2), fields(3)) & vbLf
    Next i
    Measure = out
End Function
'''


def main() -> int:
    try:
        from pyvbaharness import ExcelSession
        from pyvbaharness.session import HarnessConfig
    except ImportError:
        print('pyvbaharness is not installed. Install the live extra:\n    python -m pip install -e ".[dev]" --group live',
              file=sys.stderr)
        return 2
    names = NAMES + swept()
    renames = "@@".join("`".join(["{".join(sheets), "{".join(formulas), old, new]) for sheets, formulas, old, new in RENAMES)
    wait = float(os.environ.get("LIVE_EXCEL_LOCK_WAIT", "0"))
    with ExcelSession(HarnessConfig(lock_wait_s=wait)) as excel:
        excel.new_workbook()
        today = dt.date.today()
        result = excel.run_vba(_MEASURE, proc="Measure", args=("@@".join(names), renames), timeout=900)
        if result.outcome != "passed":
            print(f"Excel refused the measurement: {result!r}", file=sys.stderr)
            return 1
    # Only a line feed ends a line: a name may hold U+2028, which
    # str.splitlines would take for one.
    lines = str(result.value or "").split("\n")
    if len(lines) != len(names) + len(RENAMES) + 1 or lines[-1]:
        print(f"expected {len(names) + len(RENAMES)} lines, Excel answered {len(lines) - 1}", file=sys.stderr)
        return 1
    quoting: list[list[str]] = []
    for name, line in zip(names, lines, strict=False):
        if line in ("refused", "renamed"):
            print(f"Excel {line} the name {name!r}; it is left out", file=sys.stderr)
            continue
        alone, span = line.split("\t")
        quoting.append([name, alone, span])
    renamed: list[dict[str, object]] = []
    for (sheets, _, old, new), line in zip(RENAMES, lines[len(names) :], strict=False):
        entered, after = (part.split("{") for part in line.split("}"))
        renamed.append({"sheets": sheets, "old": old, "new": new, "entered": entered, "after": after})
    record = {
        "measured": today.isoformat(),
        "fields": ["name", "alone", "span"],
        "renames": renamed,
    }
    head = json.dumps(record)[:-1]
    body = [f"  {json.dumps(case)}," for case in quoting[:-1]] + [f"  {json.dumps(quoting[-1])}"]
    FIXTURE.write_text("\n".join([head + ', "quoting": [', *body, "]}"]) + "\n", encoding="utf-8")
    print(f"wrote {FIXTURE.relative_to(ROOT)} ({len(quoting)} names, {len(renamed)} renames)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
