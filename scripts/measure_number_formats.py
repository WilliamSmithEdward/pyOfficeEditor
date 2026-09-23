"""Measure the text Excel shows for values under number formats.

The number-format renderer is held to this, not to the documentation: every
format below is applied to every value in its suite, in a column wide enough
that nothing is cut short, and ``Range.Text`` is read back. The builtin ids
are measured too, by writing a package whose cells carry the ids directly
and asking Excel what code each one means. Last, every used cell of the
committed Excel-authored workbooks is read, so ``Cell.text`` is held to
Excel from the file's bytes onward; rebuilding one of those workbooks makes
that part stale until this runs again.

Run this on a Windows machine with Excel installed:

    python -m pip install -e ".[dev]" --group live
    python scripts/measure_number_formats.py

It writes ``tests/fixtures/excel/number_formats.json``, replacing what is
there, and nothing else. The locale is part of the result: day names,
separators and the locale-dependent builtin ids come from the machine's
Excel, and the committed corpus was measured in en-US.

Runs of sixteen or more of one character, a fill or a column of ``#``, are
stored as ``\\x01`` + the character + the count + ``\\x02`` so the file stays
readable. ``decode_text`` in the tests reverses it.
"""

from __future__ import annotations

import json
import re
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / "tests" / "fixtures" / "excel"
FIXTURE = FIXTURES / "number_formats.json"
EMPTY = FIXTURES / "empty.xlsx"

#: Reply separators, between fields and between records: private-use
#: characters no measured text contains, so no rendered text can pass for
#: one. VBA makes them with the functions in SEPARATORS_VBA.
FS = "\ue001"
RS = "\ue002"

SEPARATORS_VBA = r"""
Private Function FS() As String
    FS = ChrW(&HE001)
End Function

Private Function RS() As String
    RS = ChrW(&HE002)
End Function
"""


#: A value: its label, the VBA expression that puts it in a cell (through
#: ``.Formula`` when it starts with ``"=``), and what it is in JSON.
Value = tuple[str, str, object]


@dataclass
class Measured:
    """One suite's result: the formats Excel refused, and per format the
    text of each value."""

    name: str
    date1904: bool
    refused: list[str]
    values: list[Value]
    texts: dict[str, list[str]]


def number(label: str, vba: str | None = None) -> Value:
    """A number, written in VBA as a Double literal."""
    if vba is None:
        vba = label if any(c in label for c in ".E") else label + "#"
    return label, vba, float(label)


def text(label: str, vba: str, content: str) -> Value:
    return label, vba, {"text": content}


def error(code: str, formula: str) -> Value:
    return code, f'"{formula}"', {"error": code}


BASE_VALUES: list[Value] = [
    *(number(n) for n in ["0", "1", "-1", "0.5", "-0.5", "2.5", "3.5", "-2.5", "0.125", "1.005",
                          "1234.5678", "-1234.5678"]),
    ("0.1+0.2", "0.1 + 0.2", 0.1 + 0.2),
    ("1/3", "1 / 3", 1 / 3),
    ("2/3", "2 / 3", 2 / 3),
    number("1E-7", "1E-07"),
    *(number(n) for n in ["0.000123456", "123456789012", "12345678901"]),
    number("1E15", "1E+15"),
    number("1E20", "1E+20"),
    *(number(n) for n in ["99999999999", "999999999999", "46027", "46027.75", "46027.999999",
                          "0.75", "1.5", "36.25", "60", "61", "0.0001", "1234567.891", "-0.0001",
                          "100", "10.1", "0.3333333333333", "7.0000001", "0.9999999999",
                          "123.456", "5", "-5"]),
    text("text abc", '"abc"', "abc"),
    text("text 12", "\"'12\"", "12"),
    ("TRUE", "True", True),
    error("#DIV/0!", "=1/0"),
    *(number(n) for n in ["2958465", "2958465.9999999", "2958466", "-0.25", "1E-10", "0.00005"]),
    number("-1E-7", "-1E-07"),
    number("1E300", "1E+300"),
    *(number(n) for n in ["12345.6789", "0.999994", "0.999995", "1462", "-1462"]),
    number("2.5E-6", "0.0000025"),
    *(number(n) for n in ["100000", "-100000", "1000000", "0.05", "-0.05", "59.99999",
                          "1234567890123450", "-46027.5"]),
    text("text empty", '"="""""', ""),
    text("text space", '" "', " "),
    text("text Mixed", '"Mixed Case"', "Mixed Case"),
    ("FALSE", "False", False),
    error("#N/A", "=NA()"),
]

CATALOG_FORMATS: list[str] = [
    "General", "0", "0.00", "#,##0", "#,##0.00", "0%", "0.00%", "0.00E+00", "##0.0E+0",
    "0.0E+0", "# ?/?", "# ??/??", "?/?", "# ?/8", "# ??/100", "$#,##0.00",
    "$#,##0_);($#,##0)", '"$"#,##0.00_);[Red]("$"#,##0.00)', "#,##0.00_);(#,##0.00)",
    '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
    '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    "m/d/yyyy", "d-mmm-yy", "d-mmm", "mmm-yy", "h:mm AM/PM", "h:mm:ss AM/PM", "h:mm",
    "h:mm:ss", "m/d/yyyy h:mm", "mm:ss", "[h]:mm:ss", "[m]:ss", "[s]", "mm:ss.0",
    "h:mm:ss.00", "yyyy-mm-dd", "dddd", "ddd", "mmmm", "mmmmm", "yy", "d", "dd",
    "mmmm d, yyyy", "dddd, mmmm dd, yyyy", "h AM/PM", "h:mm A/P", "@", '"Total: "0.00',
    '0.00" units"', "#,##0,,", "#,##0,", '0.0,,"M"', '[>=1000]#,##0,"K";0',
    '[<0]"neg";"pos"', "[Red]0.00;[Blue]-0.00", "0;-0;;@", ";;;", "00000", "000-00-0000",
    "(###) ###-####", "###,###.##", "#.##", "0.#", "#", "?.??", "0.0??", '#,##0.00 "USD"',
    "0.00_);(0.00)", "\\$0.00", "0.00\\%", "[$-409]mmmm d, yyyy", "[$€-407]#,##0.00",
    "[$¥-411]#,##0", "0.00;[Red]0.00", '#,##0;-#,##0;"zero"', '0%;-0%;"flat"',
    "[Blue][>100]0;[Red][<0]0;0", "0.000", "#,##0.000", "0.0", "#,##0.0", "0.0%",
    "yyyy", "m/d", "mm/dd/yy", "hh:mm:ss", "[hh]:mm", "d/m/yyyy",
    # Literal-only, empty and odd section layouts.
    '"x"', '"x";"y"', '"x";"y";"z"', "0.00;;", "0.00;", ";0.00", "0;@", "@;@",
    # Exponents.
    "0.00E-00", "0.0E-0", "#.##E+00", "0E+0", "00.0E+00", "###.0E+0", "0.00E+000",
    # Signs, percent, scaling and grouping.
    '#,##0.00;(#,##0.00);"-"', "0.0%%", "%0", '0.00" %"', "$#,##0.00;-$#,##0.00",
    "+0;-0;0", "0.00,", "#,#", "#,###,##0", "0,000", "0.0,,", '#,##0.0,"K"',
    # General with company.
    "General;-General", '"$"General', 'General" USD"', "[Red]General",
    "General;[Red]-General",
    # Spaces, escapes, currencies, fill.
    "0\\ 0", '0" "0', "0 0", "[$$-409]#,##0.00", "[$USD] #,##0.00", "#,##0.00 [$€-1]",
    "[$€-2] #,##0.00", '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-', "* #,##0",
    "#,##0*-",
    # Fractions.
    "0/0", "#/#", "# #/#", "0 ?/?", "?/10", "# ???/???", "?/?;-?/?",
    # Conditions.
    '[<=100]"small";[<=1000]"mid";"big"', "[>0]0;[<0]-0;0;@", '[=0]"zero";General',
    "[<>0]0.0", "[<100]0.0", "[>=100]0;0.00", "[>=100]0;[<0]0.00",
    # Text sections.
    "@ @", '"<"@">"', '0;0;0;"pre"@"post"', "General;General;General;@", ";;;@",
    '"text: "@',
    # Rarer date and time tokens.
    "y", "yyy", "yyyyy", "m", "mmmmmm", "ddddd", "h", "hhh", "s", "ss", "sss", "e",
    "h:mm am/pm", "h:mm a/p", "h:mm:ss.000", "[h]:mm:ss.0", "[ss].00", "[mm]:ss", "m:ss",
    "yyyy-m", "h m", "m/d/yyyy h:mm:ss AM/PM", "[$-F800]dddd, mmmm dd, yyyy",
    "[$-F400]h:mm:ss AM/PM", "[$-409]h:mm:ss AM/PM", 'd "days" h:mm', "mmm d",
    "dd/mm/yyyy hh:mm", "yyyy/mm/dd hh:mm:ss.0", "[h]", "[m]", "h:mm;@", "m/d/yy;@",
    "[$-409]m/d/yy h:mm AM/PM;@", "h:mm:ss;@", "mm:ss.00", "hh", "[hh]", "[Blue]h:mm",
    'h:mm AM/PM;"neg"', "0.00;m/d/yyyy", "m/d/yyyy;0.00",
]

#: The 1904 system: where Excel shows negative dates and times.
CALENDAR_1904_VALUES: list[Value] = [
    number(n) for n in ["0", "1", "-1", "0.5", "-0.5", "-0.25", "1462", "2957003",
                        "2957003.9999999", "2957004", "2958465", "-1462", "-1.5", "46027.75",
                        "60", "61", "-0.0001", "-2957003", "-2957004", "366", "-1.75"]
]

CALENDAR_1904_FORMATS: list[str] = [
    "General", "m/d/yyyy", "h:mm", "h:mm:ss", "[h]:mm:ss", "[s]", "mm:ss", "dddd", "yyyy",
    "d-mmm-yy", "m/d/yyyy h:mm", "0.00", "h:mm;-h:mm", "[h]:mm;-[h]:mm", "[m]", "h:mm AM/PM",
    "ddd d mmm yyyy",
]

DETAIL_VALUES: list[Value] = [
    *(number(n) for n in [
        "-200", "-60", "-20", "-5", "-1", "-0.5", "-0.25", "0", "0.25", "0.5", "1", "2.5", "5",
        "20", "60", "70", "150", "200", "0.3333333333333", "1234.5678", "-1234.5678",
        "3276.7", "3276.8", "6553.5", "6553.6", "9999.9", "10000", "10000.1", "8191.75",
        "8192", "16383.75", "16384", "24999.75", "25000", "327.67", "327.68", "655.35",
        "655.36", "999.99", "1000", "21474836.47", "21474836.48", "214748364.7",
        "214748364.8", "2147483647", "2147483648", "99999", "100000", "46027.75",
    ]),
    text("text abc", '"abc"', "abc"),
    ("TRUE", "True", True),
]

DETAIL_FORMATS: list[str] = [
    # Which sections show a magnitude.
    "[>100]0;[<50]0;0", "[<5]0;0", "[<=0]0;0", "[<-10]0;0", "[=-5]0;0", "[>-10]0;0",
    "[<0]0", '[<=-1]"n"0;"p"0', "[<0]0;[>0]0", "[>=0]0;[<0]0", "[<10]0;[<0]0;0",
    "[>100]0;[>50]0;0", '[>100]"big";[<-100]"small";"mid"', '[<1]"lt1"', "[<0]0;0",
    "[<=0]0.0;0.0", "[<0.5]0.0;0.0", "[<-0.5]0.0;0.0", "[>0]0;0", "[>=0]0;0",
    '[>0]"pos";[<0]"neg"', '[<0]"neg";[>0]"pos"', "[<>5]0;0", "[=0]0;0",
    "[<0]0;[<-10]0;0", "[>=100]0;[<=-100]0;0",
    # Fractions: # against ?, spacing, and overflow.
    "# #/#", "0 #/#", "# ?/#", "# #/?", '#" "?/?', "# - ?/?", "#-?/?", "#  ?/?", "?/#",
    "#/?", "# ##/##", "# #?/?#", "# ??/#", '# ?/? "in"', '"x"# ?/?', "# #/4", "# ?/4",
    "#/4", "0/4", "?/4", "# #/10", "?/10", "?/100", "#/1000", "?/?", "??/??", "0/0",
    "# ??/100", "0 ??/??",
    # Negative sections under dates, and elapsed counts on their own.
    "h:mm;-h:mm", "[h]:mm;-[h]:mm", "m/d/yyyy;-m/d/yyyy", "[h];-[h]", "[s];-[s]",
    '[h] "hours"', '[m] "min"', "[s].0", "[h]:mm", '"t"[h]', "[h]h",
    # AM/PM spellings.
    "h AM/pm", "h Am/Pm", "h A/p", "h a/P", "h am/PM",
    # System locales and the era year.
    "[$-F800]yyyy", "[$-F800]m/d/yyyy h:mm", "[$-F400]h:mm", "[$-F400]yyyy", "ee", "e/m/d",
    "yyyy e", "[$-F800]dddd, mmmm dd, yyyy;-0",
    # Text sections in odd places, which Excel refuses but for the last.
    "@;0", "0;0;@", "0;@;0", '"a"@;0', "0;0;0;0;@",
]


def _condition_formats() -> list[str]:
    """Every operator against limits either side of zero, in the four
    layouts a single condition can take, with literals marking which
    section rendered."""
    formats: list[str] = []
    for operator in ["<", "<=", ">", ">=", "=", "<>"]:
        for limit in ["-10", "-1", "-0.5", "0", "0.5", "1", "10"]:
            c = f"[{operator}{limit}]"
            formats += [f'{c}"a"0;"b"0', f'{c}"a"0;"b"0;"c"0', f'"a"0;{c}"b"0;"c"0', f'{c}"a"0']
    return formats


CONDITION_FORMATS: list[str] = [
    *_condition_formats(),
    "?/?", "??/??", "?/10", "# ?/?",
    # Small negatives against the precision each section rounds to.
    '[<=0]"a"0;"b"0', '[<=0]"a"0.0;"b"0.0', '[<=0]"a"0.00;"b"0.00', '[<=0]"a"General;"b"General',
    '[<=0]"a";"b"', '[<=0]"a"0.0;"b"0', '[<=0]"a"0;"b"0.0', '[<0]"a"0;"b"0', '[<=-0.1]"a"0;"b"0',
    '[<=0.1]"a"0;"b"0', '[<=0]"a"0%;"b"0%', '[<=0]"a"0.0E+0;"b"0', '[>=0]"a"0;"b"0',
    '[>0]"a"0;"b"0', '[=0]"a"0;"b"0', '[<>0]"a"0;"b"0', '[<=0]"a"0',
    '"a"0;[<-10]"b"0;"c"0', '"a"0.0;[<-10]"b"0;"c"0', '"a"0.00;[<-10]"b"0;"c"0',
    '"a"0;[<-10]"b"0.0;"c"0', '"a"0;[<-10]"b"0;"c"0.0', '"a"General;[<-10]"b"0;"c"0',
    '"a";[<-10]"b";"c"',
    '[<-10]"a"0;"b"0;"c"0', '[<-10]"a"0;"b"0.0;"c"0', '[<-10]"a"0;"b"0;"c"0.0',
    '[<-10]"a"0.0;"b"0;"c"0', '[>10]"a"0;"b"0;"c"0', '[>10]"a"0;"b"0.0;"c"0',
    '[<=0]"a"0;"b"0;"c"0', '[<=0]"a"0.0;"b"0;"c"0', '[<=0]"a"0;"b"0.0;"c"0',
    # Going round again: fractions, percent, scaling, and more layouts.
    '[<0.2]"a"0;"b"0', '[<0.3]"a"0;"b"0', '[>=-0.3]"a"0;"b"0', '[=-0.25]"a"0;"b"0',
    '[<>0.25]"a"0;"b"0', '[<=0]"a"?/?;"b"0', '[<=0]"a"# ?/?;"b"0', '[<=0]"a"#,##0,;"b"0',
    '[<=0]"a"0%;"b"0', '[<=0]"a"0.0E+0;"b"0', '[<=0]"a"h:mm;"b"0', '[<=0]"a"[h];"b"0',
    '"a"0;[<-5]"b"0', '[>5]"a"0;"b"0;[<-5]"c"0', '"a"0;"b"0;[<-5]"c"0',
    '[>5]"a"0;[<-5]"b"0;[=0]"c"0', '[>5]"a"0;[<-5]"b"0;[<=0]"c"0', '"a"0;"b"0;"c"0',
    '[>5]"a"0;"b"0.0;"c"0', '"a"0.0;[<-5]"b"0;"c"0', "[s]", '[<0]"a"0;[>0]"b"0',
    '[<-5]"a"0;[<0]"b"0', '[<-5]"a"0;[<0]"b"0;"c"0',
]


def _condition_values() -> list[Value]:
    labels = [
        "-600", "-400", "-20", "-10", "-5", "-3", "-1", "-0.6", "-0.5", "-0.49", "-0.45",
        "-0.25", "-0.1", "-0.05", "-0.04", "-0.01", "-0.005", "-0.004", "-0.001", "-0.0001",
        "-1E-10", "0", "1E-10", "0.0001", "0.004", "0.005", "0.25", "0.5", "1", "3", "5",
        "10", "20", "400", "600", "1073741823.5", "1073741824.5", "2147483646",
        "2147483646.5", "-2147483646", "-2147483647", "-3276.7", "-3276.8", "2147483647.5",
    ]
    return [number(label) for label in labels]


SUITES: list[tuple[str, bool, list[Value], list[str]]] = [
    ("catalog", False, BASE_VALUES, CATALOG_FORMATS),
    ("calendar_1904", True, CALENDAR_1904_VALUES, CALENDAR_1904_FORMATS),
    ("details", False, DETAIL_VALUES, DETAIL_FORMATS),
    ("conditions", False, _condition_values(), CONDITION_FORMATS),
]

#: Well under the two hundred or so custom formats one workbook can hold.
FORMATS_PER_WORKBOOK = 100

#: Builtin ids 0 to 163 and the values each is shown with.
BUILTIN_IDS = range(164)
BUILTIN_SAMPLES = [46027.5, 1234.5678, -1234.5678, 0.5, 0.0]


def vba_string(content: str) -> str:
    """A VBA expression evaluating to this string, ASCII only."""
    if not content:
        return '""'
    pieces: list[str] = []
    run = ""
    for char in content:
        if 32 <= ord(char) < 127 and char != '"':
            run += char
            continue
        if run:
            pieces.append(f'"{run}"')
            run = ""
        pieces.append(f"ChrW({ord(char)})")
    if run:
        pieces.append(f'"{run}"')
    return " & ".join(pieces)


def suite_source(values: list[Value], formats: list[str], date1904: bool) -> str:
    fill = ["Private Sub Fill(ByVal ws As Worksheet, ByVal col As Long)"]
    for row, (_, expression, _) in enumerate(values, start=1):
        target = "Formula" if expression.startswith('"=') else "Value"
        fill.append(f"    ws.Cells({row}, col).{target} = {expression}")
    fill.append("End Sub")

    codes = ["Private Function FormatCode(ByVal i As Long) As String", "    Select Case i"]
    for index, code in enumerate(formats):
        codes.append(f"        Case {index}: FormatCode = {vba_string(code)}")
    codes += ["    End Select", "End Function"]

    return SEPARATORS_VBA + "\n".join(fill) + "\n" + "\n".join(codes) + rf"""
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim f As Long
    Dim r As Long
    Dim out As String
    Dim chunk As String
    Dim nFormats As Long
    Dim nValues As Long

    nFormats = {len(formats)}
    nValues = {len(values)}
    Set wb = ActiveWorkbook
    wb.Date1904 = {"True" if date1904 else "False"}
    Set ws = wb.Worksheets(1)
    out = Application.Version & " " & Application.Build & RS()

    For f = 0 To nFormats - 1
        Fill ws, f + 2
        On Error Resume Next
        ws.Range(ws.Cells(1, f + 2), ws.Cells(nValues, f + 2)).NumberFormat = FormatCode(f)
        If Err.Number <> 0 Then out = out & "REFUSED" & FS() & f & RS()
        Err.Clear
        On Error GoTo 0
    Next f
    ws.Range(ws.Columns(2), ws.Columns(nFormats + 1)).ColumnWidth = 255

    For f = 0 To nFormats - 1
        chunk = f & FS() & ws.Cells(1, f + 2).NumberFormat
        For r = 1 To nValues
            chunk = chunk & FS() & ws.Cells(r, f + 2).Text
        Next r
        out = out & chunk & RS()
    Next f

    Application.DisplayAlerts = False
    wb.SaveAs Target, 51
    Application.DisplayAlerts = True
    Build = out
End Function
"""


_READ_IDS = r"""
Public Function ReadIds(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim r As Long
    Dim c As Long
    Dim out As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets(1)
    ws.Columns("B:F").ColumnWidth = 255
    For r = 1 To IDS_COUNT
        out = out & ws.Cells(r, 1).Value & FS() & ws.Cells(r, 2).NumberFormat
        For c = 2 To 6
            out = out & FS() & ws.Cells(r, c).Text
        Next c
        out = out & RS()
    Next r
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ReadIds = out
End Function
""".replace("IDS_COUNT", str(len(BUILTIN_IDS)))


#: The committed Excel-authored workbooks whose every cell's text is
#: recorded, so ``Cell.text`` is held to Excel end to end: style lookup,
#: builtin ids and rendering together. Rebuild a fixture and this goes
#: stale; measure again.
CELL_WORKBOOKS = [
    "sample.xlsx", "structures.xlsx", "links.xlsx", "settings.xlsx", "filters.xlsx",
    "refused.xlsx", "geometry.xlsx", "controls.xlsm", "shapes.xlsm",
    "excel_authored_minimal.xlsm", "excel_authored_powerquery.xlsx",
]

_READ_CELLS = r"""
Public Function ReadCells(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim c As Range
    Dim out As String
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Set wb = Workbooks.Open(Target, UpdateLinks:=0, ReadOnly:=True)
    For Each ws In wb.Worksheets
        ' A protected sheet will not take the column width that keeps Text
        ' from being cut short, and unprotecting one asks for its password.
        If Not ws.ProtectContents Then
            ws.Cells.ColumnWidth = 255
            For Each c In ws.UsedRange.Cells
                If Not IsEmpty(c.Value) Or c.HasFormula Then
                    out = out & ws.Name & FS() & c.Address(False, False) & FS() & c.Text & RS()
                End If
            Next c
        End If
    Next ws
    wb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    ReadCells = out
End Function
"""


def builtin_package(target: Path) -> None:
    """``empty.xlsx`` with one row per builtin id, each cell styled by an
    ``xf`` naming that id and nothing else."""
    with zipfile.ZipFile(EMPTY) as archive:
        members = {name: archive.read(name) for name in archive.namelist()}

    styles = members["xl/styles.xml"].decode("utf-8")
    match = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', styles, re.S)
    if match is None:
        raise SystemExit("empty.xlsx has no cellXfs to extend")
    base = int(match.group(1))
    extra = "".join(
        f'<xf numFmtId="{format_id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        for format_id in BUILTIN_IDS
    )
    styles = (
        styles[: match.start()]
        + f'<cellXfs count="{base + len(BUILTIN_IDS)}">{match.group(2)}{extra}</cellXfs>'
        + styles[match.end() :]
    )
    members["xl/styles.xml"] = styles.encode("utf-8")

    rows: list[str] = []
    for row, format_id in enumerate(BUILTIN_IDS, start=1):
        cells = [f'<c r="A{row}"><v>{format_id}</v></c>']
        for column, sample in zip("BCDEF", BUILTIN_SAMPLES, strict=True):
            cells.append(f'<c r="{column}{row}" s="{base + row - 1}"><v>{sample!r}</v></c>')
        rows.append(f'<row r="{row}">{"".join(cells)}</row>')
    sheet = members["xl/worksheets/sheet1.xml"].decode("utf-8")
    sheet = re.sub(r"<sheetData\s*/>|<sheetData>.*?</sheetData>",
                   "<sheetData>" + "".join(rows) + "</sheetData>", sheet, flags=re.S)
    members["xl/worksheets/sheet1.xml"] = sheet.encode("utf-8")

    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for name, body in members.items():
            out.writestr(name, body)


def squeeze(content: str) -> str:
    """Runs of sixteen or more of a character, stored as a count."""
    return re.sub(r"(.)\1{15,}", lambda m: f"\x01{m.group(1)}{len(m.group(0))}\x02", content)


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

    suites: list[Measured] = []
    excel_version = ""
    # Hidden macro runs only, which the harness allows beside other sessions.
    config = HarnessConfig(exclusive=False)
    with tempfile.TemporaryDirectory() as scratch, ExcelSession(config) as excel:
        for name, date1904, values, formats in SUITES:
            refused: list[str] = []
            texts: dict[str, list[str]] = {}
            # A workbook holds only two hundred or so custom formats, and past
            # that Excel refuses valid ones, so each batch gets a fresh one.
            for start in range(0, len(formats), FORMATS_PER_WORKBOOK):
                batch = formats[start : start + FORMATS_PER_WORKBOOK]
                excel.new_workbook()
                excel.reset_sheets()
                result = excel.run_vba(
                    suite_source(values, batch, date1904), proc="Build",
                    args=(str(Path(scratch) / f"{name}_{start}.xlsx"),), timeout=900,
                )
                if result.outcome != "passed":
                    raise SystemExit(f"{name}: Excel refused the measurement: {result!r}")
                records = str(result.value or "").split(RS)
                excel_version = records[0]
                for record in records[1:]:
                    if not record:
                        continue
                    fields = record.split(FS)
                    if fields[0] == "REFUSED":
                        refused.append(batch[int(fields[1])])
                        continue
                    code = batch[int(fields[0])]
                    if code not in refused:
                        texts[code] = [squeeze(cell) for cell in fields[2:]]
            suites.append(Measured(name, date1904, refused, values, texts))
            print(f"  {name}: {len(texts)} formats x {len(values)} values; refused {refused}")

        package = Path(scratch) / "builtin_ids.xlsx"
        builtin_package(package)
        excel.new_workbook()
        result = excel.run_vba(SEPARATORS_VBA + _READ_IDS, proc="ReadIds", args=(str(package),), timeout=300,
                               module_name="ReadIds")
        if result.outcome != "passed":
            raise SystemExit(f"builtin ids: Excel refused the measurement: {result!r}")
        builtins: dict[str, dict[str, object]] = {}
        for record in str(result.value or "").split(RS):
            if record:
                fields = record.split(FS)
                builtins[str(int(float(fields[0])))] = {
                    "code": fields[1], "texts": [squeeze(cell) for cell in fields[2:]]}

        cells: dict[str, dict[str, str]] = {}
        for name in CELL_WORKBOOKS:
            excel.new_workbook()
            result = excel.run_vba(SEPARATORS_VBA + _READ_CELLS, proc="ReadCells", args=(str(FIXTURES / name),),
                                   timeout=300, module_name="ReadCells")
            if result.outcome != "passed":
                raise SystemExit(f"{name}: Excel refused to read its cells: {result!r}")
            shown: dict[str, str] = {}
            for record in str(result.value or "").split(RS):
                if record:
                    sheet, address, content = record.split(FS)
                    shown[f"{sheet}!{address}"] = squeeze(content)
            cells[name] = shown
            print(f"  {name}: {len(shown)} cells")

    write(excel_version, suites, builtins, cells)
    print(f"wrote {FIXTURE.relative_to(ROOT)} ({FIXTURE.stat().st_size} bytes)")
    return 0


def write(
    excel_version: str, suites: list[Measured], builtins: dict[str, dict[str, object]],
    cells: dict[str, dict[str, str]],
) -> None:
    """One format per line, so a change to the corpus diffs by format."""
    dump = json.dumps
    lines = ["{", f'"excel": {dump(excel_version)},', f'"builtin_samples": {dump(BUILTIN_SAMPLES)},',
             '"builtins": {']
    lines.append(",\n".join(f"{dump(key)}: {dump(entry)}" for key, entry in builtins.items()))
    lines.append("},")
    lines.append('"cells": {')
    lines.append(",\n".join(f"{dump(name)}: {dump(texts)}" for name, texts in cells.items()))
    lines.append("},")
    lines.append('"suites": [')
    for position, suite in enumerate(suites):
        head = {
            "name": suite.name, "date1904": suite.date1904, "refused": suite.refused,
            "values": [[label, typed] for label, _, typed in suite.values],
        }
        lines.append(dump(head)[:-1] + ', "texts": {')
        lines.append(",\n".join(f"{dump(code)}: {dump(row)}" for code, row in suite.texts.items()))
        lines.append("}}" + ("," if position + 1 < len(suites) else ""))
    lines.append("]")
    lines.append("}")
    FIXTURE.write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    raise SystemExit(main())
