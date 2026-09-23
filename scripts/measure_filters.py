"""Measure what Excel's autofilter keeps, criterion by criterion.

A filter is only as good as its evaluator, and the evaluator is held to this
rather than to the documentation. Every case below is a column of cells and
one criterion; Excel applies the criterion and the rows it hid are recorded,
with every cell's value and displayed text and the markup Excel stored.

Three ways a criterion reaches Excel, because they answer different things:

- **object model**: ``Range.AutoFilter`` applies it, which is what a user
  does, and the saved markup shows what Excel stores for it;
- **file**: markup written straight into a package Excel then opens and
  re-applies with ``ApplyFilter``, for criteria the object model never
  writes itself but a file, or this library, can carry;
- **collation**: a ``>`` criterion per word against a column of all of
  them, which gives Excel's text order one pair at a time.

Run this on a Windows machine with Excel installed:

    python -m pip install -e ".[dev]" --group live
    python scripts/measure_filters.py

It writes ``tests/fixtures/excel/filter_semantics.json``, replacing what is
there. Dates are relative to the day it runs, and the stored date-period
bounds with them, so each run is self-consistent but not identical to the
last.
"""

from __future__ import annotations

import json
import re
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FIXTURE = ROOT / "tests" / "fixtures" / "excel" / "filter_semantics.json"


#: Reply separators: private-use characters no measured text contains,
#: so a cell reading "B" cannot pass for a separator. VBA makes them
#: with the functions in SEPARATORS_VBA.
FS = "\ue001"
RS = "\ue002"
CS = "\ue003"
NS = "\ue004"
BS = "\ue005"

SEPARATORS_VBA = r"""
Private Function FS() As String
    FS = ChrW(&HE001)
End Function

Private Function RS() As String
    RS = ChrW(&HE002)
End Function

Private Function CS() As String
    CS = ChrW(&HE003)
End Function

Private Function NS() As String
    NS = ChrW(&HE004)
End Function

Private Function BS() As String
    BS = ChrW(&HE005)
End Function
"""


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


# ---------------------------------------------------------------------------
# Columns
# ---------------------------------------------------------------------------

#: The mixed column: blanks, whitespace, an empty-string formula, numbers,
#: numeric-looking text, text in both cases and with accents, both
#: booleans, two errors, a date and a literal asterisk. Rows 2 to 30.
# Single-quoted: the ="" formula is five double quotes in a row.
FILL_MIXED = r'''
Private Sub FillMixed(ByVal ws As Worksheet)
    ws.Range("A1").Value = "V"
    ws.Range("A3").Value = "' "
    ws.Range("A4").Value = "'  "
    ws.Range("A5").Formula = "="""""
    ws.Range("A6").Value = 0
    ws.Range("A7").Value = 5
    ws.Range("A8").Value = 50
    ws.Range("A9").Value = 150
    ws.Range("A10").Value = -3
    ws.Range("A11").Value = "'12"
    ws.Range("A12").Value = "pending"
    ws.Range("A13").Value = "Zebra"
    ws.Range("A14").Value = "apple"
    ws.Range("A15").Value = True
    ws.Range("A16").Value = False
    ws.Range("A17").Formula = "=1/0"
    ws.Range("A18").Formula = "=NA()"
    ws.Range("A19").Value = DateSerial(2026, 1, 5)
    ws.Range("A20").Value = ChrW(201) & "mile"
    ws.Range("A21").Value = ChrW(225) & "bc"
    ws.Range("A22").Value = "Mango"
    ws.Range("A23").Value = "M"
    ws.Range("A24").Value = "m"
    ws.Range("A25").Value = "a*b"
    ws.Range("A26").Value = 12
    ws.Range("A27").Value = "'5"
    ws.Range("A28").Value = 5.5
    ws.Range("A29").Value = "Apple"
    ws.Range("A30").Value = "banana"
End Sub
'''

#: Dates either side of today, the fifteenth of every month this year, one
#: date last year, and a text cell.
FILL_DATES = r"""
Private Sub FillDates(ByVal ws As Worksheet)
    Dim offsets As Variant
    Dim i As Long
    ws.Range("A1").Value = "D"
    offsets = Array(-800, -400, -120, -95, -40, -31, -8, -7, -2, -1, 0, 1, 2, 7, 8, 31, 40, 95, 120, 400, 800)
    For i = LBound(offsets) To UBound(offsets)
        ws.Cells(i + 2, 1).Value = Date + offsets(i)
    Next i
    For i = 1 To 12
        ws.Cells(i + 23, 1).Value = DateSerial(Year(Date), i, 15)
    Next i
    ws.Cells(36, 1).Value = DateSerial(Year(Date) - 1, 6, 15)
    ws.Cells(37, 1).Value = "not a date"
End Sub
"""

#: Date-times to group at every level, a bare date, text, and a date serial
#: under General.
FILL_TIMES = r"""
Private Sub FillTimes(ByVal ws As Worksheet)
    ws.Range("A1").Value = "T"
    ws.Range("A2").Value = DateSerial(2026, 1, 5) + TimeSerial(10, 0, 0)
    ws.Range("A3").Value = DateSerial(2026, 1, 5) + TimeSerial(15, 30, 0)
    ws.Range("A4").Value = DateSerial(2026, 1, 5)
    ws.Range("A5").Value = DateSerial(2026, 1, 6) + TimeSerial(10, 0, 0)
    ws.Range("A6").Value = DateSerial(2026, 2, 5) + TimeSerial(10, 0, 0)
    ws.Range("A7").Value = DateSerial(2025, 1, 5) + TimeSerial(10, 0, 0)
    ws.Range("A8").Value = DateSerial(2026, 1, 5) + TimeSerial(10, 30, 0)
    ws.Range("A9").Value = DateSerial(2026, 1, 5) + TimeSerial(10, 0, 45)
    ws.Range("A2:A9").NumberFormat = "m/d/yyyy h:mm:ss"
    ws.Range("A10").Value = "text"
    ws.Range("A12").Value = 46027
End Sub
"""

#: Numbers with text, a boolean and numeric text among them, and the same
#: with an error, which is what top ten and averages refuse.
FILL_NUMS = r"""
Private Sub FillNums(ByVal ws As Worksheet)
    Dim values As Variant
    Dim i As Long
    ws.Range("A1").Value = "N"
    values = Array(5, 50, 150, -3, 0, 12, 5.5, 99, 7, 42)
    For i = LBound(values) To UBound(values)
        ws.Cells(i + 2, 1).Value = values(i)
    Next i
    ws.Range("A13").Value = "pending"
    ws.Range("A14").Value = True
    ws.Range("A15").Value = "'12"
End Sub

Private Sub FillNumsErr(ByVal ws As Worksheet)
    FillNums ws
    ws.Range("A16").Formula = "=1/0"
End Sub
"""

#: Words whose order is measured pairwise.
WORDS: list[str] = [
    "a", "A", "b", "B", "z", "Z", "0", "1", "10", "9", "2",
    "résumé", "resume", "Resume", "résume",
    "co-op", "coop", "co op", "coop2", "cop",
    "_x", "-x", "!x", "~x", "x",
    "ß", "ss", "st", "sr",
    "æ", "ae", "af", "ad",
    "ä", "az", "Ä", "ab",
    "ø", "o", "oz", "ö", "p",
    "é", "e", "f", "E",
    "Ω", "ω",
    "ž", "zz",
    "a b", "a-b", "a'b", "a.b", "a_b",
    "1a", "a1", "abc", "ABD", "abd",
]

#: Every ASCII symbol and some common others, each before an x.
SYMBOL_WORDS: list[str] = [
    symbol + "x"
    for symbol in [chr(code) for code in range(32, 127) if not chr(code).isalnum()]
    + ["\u00a7", "\u00b0", "\u00b1", "\u20ac", "\u00a3", "\u00a9", "\u00ab", "\u00bf", "\u2013", "\u2014", "\u2019", "\u2026", "\u00a0"]
] + ["0x", "9x", "x", "ax"]


def fill_text(name: str, words: list[str]) -> str:
    """A fill procedure writing each word as text, one per row from 2."""
    lines = [f"Private Sub {name}(ByVal ws As Worksheet)", '    ws.Range("A1").Value = "W"']
    for row, word in enumerate(words, start=2):
        lines.append(f"    ws.Cells({row}, 1).Value = \"'\" & {vba_string(word)}")
    lines.append("End Sub")
    return "\n".join(lines)


#: Equality and wildcards under collation: (column name, words, criterion).
EQUALITY: list[tuple[str, list[str], str]] = [
    ("eq_ae_vs_ligature", ["æ", "ae", "a e", "AE"], 'Criteria1:=Array("ae"), Operator:=xlFilterValues'),
    ("eq_ss_vs_eszett", ["ß", "ss", "SS", "s"], 'Criteria1:=Array("ss"), Operator:=xlFilterValues'),
    ("eq_coop_vs_hyphen", ["co-op", "coop", "co op", "co'op"],
     'Criteria1:=Array("coop"), Operator:=xlFilterValues'),
    ("eq_resume_vs_accent", ["résumé", "resume", "RESUME", "résume"],
     'Criteria1:=Array("resume"), Operator:=xlFilterValues'),
    ("cf_ae_vs_ligature", ["æ", "ae", "a e", "AE"], 'Criteria1:="=a*e"'),
    ("cf_coop_wild", ["co-op", "coop", "co op", "co'op"], 'Criteria1:="=co*op"'),
    ("ne_coop", ["co-op", "coop", "co op", "co'op"], 'Criteria1:="<>coop"'),
    ("ne_ae", ["æ", "ae", "a e", "AE"], 'Criteria1:="<>ae"'),
    ("eq_nbsp", ["a b", "a\u00a0b", "ab"], 'Criteria1:=Array("a b"), Operator:=xlFilterValues'),
]

#: Stored values against padded, filled and fraction formats: (column
#: name, format, value expression, candidate values).
ACCOUNTING = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
FORMATTED: list[tuple[str, str, str, list[str]]] = [
    ("fmt_accounting", ACCOUNTING, "1235",
     [" $1,235.00 ", "$1,235.00", " $ 1,235.00 ", "1,235.00", "$1,235.00 "]),
    ("fmt_accounting_zero", ACCOUNTING, "0", [" $-   ", "$-", " $ - ", "-"]),
    ("fmt_currency", "$#,##0.00_);($#,##0.00)", "1234.57", ["$1,234.57 ", "$1,234.57"]),
    ("fmt_fill_before", "*x0", "5", ["xxxxxxxx5", "5", "x5", " 5"]),
    ("fmt_fill_after", "0*-", "5", ["5", "5-", "5----------"]),
    ("fmt_padding", "_W0", "5", [" 5", "5"]),
    ("fmt_fraction", "# ??/??", "0.5", ["  1/2 ", "1/2"]),
    ("fmt_text", "General", '"apple"', ["apple", " apple", "apple "]),
    ("fmt_text_padded", "General", '"\' apple"', [" apple", "apple"]),
    ("fmt_at", "@", "5", ["5", " 5"]),
    ("fmt_decimals", "0.00", "5", ["5.00", "5", " 5.00"]),
    ("fmt_general_long", "General", "1234.5678", ["1234.5678", "1234.568"]),
    ("fmt_third", "General", "1 / 3", ["0.333333333", "0.333333"]),
    ("fmt_negative_third", "General", "-1 / 3", ["-0.333333333", "-0.33333"]),
]


def fill_formatted(name: str, code: str, value: str) -> str:
    """One formatted cell in A2 and a sentinel number in A3."""
    return "\n".join([
        f"Private Sub {name}(ByVal ws As Worksheet)",
        '    ws.Range("A1").Value = "V"',
        f"    ws.Range(\"A2\").NumberFormat = {vba_string(code)}",
        f"    ws.Range(\"A2\").Value = {value}",
        '    ws.Range("A3").Value = 424242',
        "End Sub",
    ])


# ---------------------------------------------------------------------------
# Cases
# ---------------------------------------------------------------------------

#: Criteria the object model applies to the mixed column.
MIXED_CASES: list[tuple[str, str]] = [
    ("num_eq_5", 'Criteria1:="=5"'),
    ("num_ne_5", 'Criteria1:="<>5"'),
    ("num_gt_5", 'Criteria1:=">5"'),
    ("num_ge_5", 'Criteria1:=">=5"'),
    ("num_lt_5", 'Criteria1:="<5"'),
    ("num_le_5", 'Criteria1:="<=5"'),
    ("num_lt_100", 'Criteria1:="<100"'),
    ("num_gt_40", 'Criteria1:=">40"'),
    ("num_eq_12", 'Criteria1:="=12"'),
    ("num_eq_0", 'Criteria1:="=0"'),
    ("num_lt_0", 'Criteria1:="<0"'),
    ("num_gt_neg10", 'Criteria1:=">-10"'),
    ("num_eq_5_5", 'Criteria1:="=5.5"'),
    ("txt_eq_apple", 'Criteria1:="=apple"'),
    ("txt_ne_apple", 'Criteria1:="<>apple"'),
    ("txt_gt_M", 'Criteria1:=">M"'),
    ("txt_ge_M", 'Criteria1:=">=M"'),
    ("txt_lt_M", 'Criteria1:="<M"'),
    ("txt_le_M", 'Criteria1:="<=M"'),
    ("txt_eq_M", 'Criteria1:="=M"'),
    ("txt_gt_m_lower", 'Criteria1:=">m"'),
    ("txt_eq_pending", 'Criteria1:="=pending"'),
    ("txt_eq_space", 'Criteria1:="= "'),
    ("txt_gt_empty_a", 'Criteria1:=">a"'),
    ("blanks", 'Criteria1:="="'),
    ("nonblanks", 'Criteria1:="<>"'),
    ("eq_star", 'Criteria1:="=*"'),
    ("ne_star", 'Criteria1:="<>*"'),
    ("wild_begins_a", 'Criteria1:="=a*"'),
    ("wild_ends_e", 'Criteria1:="=*e"'),
    ("wild_contains_an", 'Criteria1:="=*an*"'),
    ("wild_notcontains_an", 'Criteria1:="<>*an*"'),
    ("wild_q_pple", 'Criteria1:="=?pple"'),
    ("wild_escaped_star", 'Criteria1:="=a~*b"'),
    ("wild_q_mango", 'Criteria1:="=M?ngo"'),
    ("wild_begins_1", 'Criteria1:="=1*"'),
    ("wild_gt_a_star", 'Criteria1:=">a*"'),
    ("wild_begins_upper_A", 'Criteria1:="=A*"'),
    ("date_ge", 'Criteria1:=">=1/1/2026"'),
    ("date_eq", 'Criteria1:="=1/5/2026"'),
    ("bool_eq_true", 'Criteria1:="=TRUE"'),
    ("bool_eq_false", 'Criteria1:="=FALSE"'),
    ("err_eq_div0", 'Criteria1:="=#DIV/0!"'),
    ("and_gt5_lt100", 'Criteria1:=">5", Operator:=xlAnd, Criteria2:="<100"'),
    ("or_lt0_gt100", 'Criteria1:="<0", Operator:=xlOr, Criteria2:=">100"'),
    ("and_text_a_n", 'Criteria1:=">=a", Operator:=xlAnd, Criteria2:="<=n"'),
    ("or_blank_5", 'Criteria1:="=", Operator:=xlOr, Criteria2:="=5"'),
    ("vals_5_apple", 'Criteria1:=Array("5", "apple"), Operator:=xlFilterValues'),
    ("vals_true", 'Criteria1:=Array("TRUE"), Operator:=xlFilterValues'),
    ("vals_div0", 'Criteria1:=Array("#DIV/0!"), Operator:=xlFilterValues'),
    ("vals_12", 'Criteria1:=Array("12"), Operator:=xlFilterValues'),
    ("vals_date", 'Criteria1:=Array("1/5/2026"), Operator:=xlFilterValues'),
    ("vals_5_blank", 'Criteria1:=Array("5", "="), Operator:=xlFilterValues'),
    ("vals_apple_upper", 'Criteria1:=Array("APPLE"), Operator:=xlFilterValues'),
    ("vals_space", 'Criteria1:=Array(" "), Operator:=xlFilterValues'),
    ("vals_5_5", 'Criteria1:=Array("5.5"), Operator:=xlFilterValues'),
    ("top3_items", 'Criteria1:="3", Operator:=xlTop10Items'),
    ("bottom3_items", 'Criteria1:="3", Operator:=xlBottom10Items'),
    ("top10_percent", 'Criteria1:="10", Operator:=xlTop10Percent'),
    ("bottom20_percent", 'Criteria1:="20", Operator:=xlBottom10Percent'),
    ("above_avg", "Criteria1:=xlFilterAboveAverage, Operator:=xlFilterDynamic"),
    ("below_avg", "Criteria1:=xlFilterBelowAverage, Operator:=xlFilterDynamic"),
]

#: The dynamic date periods by number, 1 for today through 32 for December.
#: Not by name: Excel's type library spells February
#: "xlFilterAllDatesInPeriodFebruray", and an undeclared name is Empty, 0.
DATE_PERIODS = [
    "Today", "Yesterday", "Tomorrow", "ThisWeek", "LastWeek", "NextWeek", "ThisMonth",
    "LastMonth", "NextMonth", "ThisQuarter", "LastQuarter", "NextQuarter", "ThisYear",
    "LastYear", "NextYear", "YearToDate", "Quarter1", "Quarter2", "Quarter3", "Quarter4",
    "January", "February", "March", "April", "May", "June", "July", "August", "September",
    "October", "November", "December",
]
DATE_CASES = [
    (f"per_{name}", f"Criteria1:={number}, Operator:=xlFilterDynamic")
    for number, name in enumerate(DATE_PERIODS, start=1)
]

#: Date groups at every level, the way VBA sets them: (level, date) pairs.
GROUP_CASES: list[tuple[str, str]] = [
    ("grp_year", 'Operator:=xlFilterValues, Criteria2:=Array(0, "1/5/2026")'),
    ("grp_month", 'Operator:=xlFilterValues, Criteria2:=Array(1, "1/5/2026")'),
    ("grp_day", 'Operator:=xlFilterValues, Criteria2:=Array(2, "1/5/2026")'),
    ("grp_hour", 'Operator:=xlFilterValues, Criteria2:=Array(3, "1/5/2026 10:00:00")'),
    ("grp_minute", 'Operator:=xlFilterValues, Criteria2:=Array(4, "1/5/2026 10:00:00")'),
    ("grp_second", 'Operator:=xlFilterValues, Criteria2:=Array(5, "1/5/2026 10:00:45")'),
    ("grp_two_days", 'Operator:=xlFilterValues, Criteria2:=Array(2, "1/5/2026", 2, "2/5/2026")'),
    ("grp_day_and_text",
     'Operator:=xlFilterValues, Criteria1:=Array("text"), Criteria2:=Array(2, "1/6/2026")'),
]

NUM_CASES: list[tuple[str, str]] = [
    ("n_top3_items", 'Criteria1:="3", Operator:=xlTop10Items'),
    ("n_bottom3_items", 'Criteria1:="3", Operator:=xlBottom10Items'),
    ("n_top10_percent", 'Criteria1:="10", Operator:=xlTop10Percent'),
    ("n_top25_percent", 'Criteria1:="25", Operator:=xlTop10Percent'),
    ("n_bottom20_percent", 'Criteria1:="20", Operator:=xlBottom10Percent'),
    ("n_above_avg", "Criteria1:=xlFilterAboveAverage, Operator:=xlFilterDynamic"),
    ("n_below_avg", "Criteria1:=xlFilterBelowAverage, Operator:=xlFilterDynamic"),
    ("n_top1_items", 'Criteria1:="1", Operator:=xlTop10Items'),
    ("n_top500_items", 'Criteria1:="500", Operator:=xlTop10Items'),
    ("n_top100_percent", 'Criteria1:="100", Operator:=xlTop10Percent'),
]

NUM_ERR_CASES: list[tuple[str, str]] = [
    ("e_top3_items", 'Criteria1:="3", Operator:=xlTop10Items'),
    ("e_above_avg", "Criteria1:=xlFilterAboveAverage, Operator:=xlFilterDynamic"),
    ("e_gt_10", 'Criteria1:=">10"'),
]

#: filterColumn bodies the object model never writes, re-applied from a file
#: over the mixed column.
FILE_CASES: dict[str, str] = {
    "cf_eq_5": '<customFilters><customFilter val="5"/></customFilters>',
    "cf_eq_5_explicit": '<customFilters><customFilter operator="equal" val="5"/></customFilters>',
    "cf_eq_apple": '<customFilters><customFilter val="apple"/></customFilters>',
    "cf_eq_12": '<customFilters><customFilter val="12"/></customFilters>',
    "cf_eq_space": '<customFilters><customFilter val=" "/></customFilters>',
    "cf_eq_empty": '<customFilters><customFilter val=""/></customFilters>',
    "cf_ne_empty": '<customFilters><customFilter operator="notEqual" val=""/></customFilters>',
    "cf_eq_true": '<customFilters><customFilter val="TRUE"/></customFilters>',
    "cf_eq_div0": '<customFilters><customFilter val="#DIV/0!"/></customFilters>',
    "cf_eq_date_text": '<customFilters><customFilter val="1/5/2026"/></customFilters>',
    "cf_eq_date_serial": '<customFilters><customFilter val="46027"/></customFilters>',
    "cf_ne_12": '<customFilters><customFilter operator="notEqual" val="12"/></customFilters>',
    "cf_ne_true": '<customFilters><customFilter operator="notEqual" val="TRUE"/></customFilters>',
    "cf_lt_a_star": '<customFilters><customFilter operator="lessThan" val="a*"/></customFilters>',
    "cf_ge_a_star": '<customFilters><customFilter operator="greaterThanOrEqual" val="a*"/></customFilters>',
    "cf_le_a_star": '<customFilters><customFilter operator="lessThanOrEqual" val="a*"/></customFilters>',
    "cf_gt_star_an": '<customFilters><customFilter operator="greaterThan" val="*an*"/></customFilters>',
    "cf_gt_5_text": '<customFilters><customFilter operator="greaterThan" val="5x"/></customFilters>',
    "cf_gt_true": '<customFilters><customFilter operator="greaterThan" val="TRUE"/></customFilters>',
    "cf_three": (
        '<customFilters><customFilter operator="greaterThan" val="0"/>'
        '<customFilter operator="lessThan" val="100"/>'
        '<customFilter operator="notEqual" val="50"/></customFilters>'
    ),
    "cf_empty": "<customFilters/>",
    "vf_05": '<filters><filter val="05"/></filters>',
    "vf_5_0": '<filters><filter val="5.0"/></filters>',
    "vf_space_apple": '<filters><filter val=" apple"/></filters>',
    "vf_empty_value": '<filters><filter val=""/></filters>',
    "vf_none": "<filters/>",
    "vf_46027": '<filters><filter val="46027"/></filters>',
    "vf_a_star": '<filters><filter val="a*"/></filters>',
    "vf_true_lower": '<filters><filter val="true"/></filters>',
    "vf_emile_plain": '<filters><filter val="emile"/></filters>',
    "top_on_text": '<top10 val="3"/>',
    "top_no_filterval": '<top10 val="2"/>',
    "top_stale_filterval": '<top10 val="2" filterVal="10"/>',
    "top_zero": '<top10 val="0"/>',
    "top_501": '<top10 val="501"/>',
    "dyn_avg_no_val": '<dynamicFilter type="aboveAverage"/>',
    "dyn_avg_stale_val": '<dynamicFilter type="aboveAverage" val="1"/>',
    "dyn_today_no_val": '<dynamicFilter type="today"/>',
    "dyn_unknown": '<dynamicFilter type="someday"/>',
    "legacy_group_day": (
        '<filters><dateGroupItem year="2026" month="1" day="5" dateTimeGrouping="day"/></filters>'
    ),
    "legacy_group_year_and_value": (
        '<filters><filter val="apple"/><dateGroupItem year="2026" dateTimeGrouping="year"/></filters>'
    ),
}


#: Stored top-ten and average filters over numbers with no error among
#: them: whether Excel ranks again or trusts what the file says.
NUM_FILE_CASES: dict[str, str] = {
    "top_stale_on_nums": '<top10 val="3" filterVal="10"/>',
    "top_no_filterval_on_nums": '<top10 val="3"/>',
    "top_pct_no_filterval_on_nums": '<top10 val="25" percent="1"/>',
    "avg_below_no_val_on_nums": '<dynamicFilter type="belowAverage"/>',
    "avg_above_stale_on_nums": '<dynamicFilter type="aboveAverage" val="60"/>',
}

#: Stored date periods whose bounds are stale or missing: whether Excel
#: works the period out again from today or trusts the file.
DATE_FILE_CASES: dict[str, str] = {
    "today_stale_on_dates": '<dynamicFilter type="today" val="46256" maxVal="46257"/>',
    "month_stale_on_dates": '<dynamicFilter type="thisMonth" val="46023" maxVal="46054"/>',
    "today_no_val_on_dates": '<dynamicFilter type="today"/>',
    "month_no_val_on_dates": '<dynamicFilter type="thisMonth"/>',
}

#: Fixed periods over the date-time column, whose row 12 is a January date
#: serial shown as General: whether a period asks for a date format.
TIME_FILE_CASES: dict[str, str] = {
    "m1_on_times": '<dynamicFilter type="M1"/>',
    "q1_on_times": '<dynamicFilter type="Q1"/>',
    # With a date group in the list, is a date still matched by its text?
    "group_and_date_text_on_times": (
        '<filters><filter val="1/6/2026 10:00:00"/>'
        '<dateGroupItem year="2025" dateTimeGrouping="year"/></filters>'
    ),
    "date_text_alone_on_times": '<filters><filter val="1/6/2026 10:00:00"/></filters>',
    "group_and_number_text_on_times": (
        '<filters><filter val="46027"/><dateGroupItem year="2025" dateTimeGrouping="year"/></filters>'
    ),
}


def file_cases() -> list[tuple[str, str, str]]:
    """Every file case: (sheet name, column kind, filterColumn body). The
    formatted columns get each candidate stored exactly as written, which
    the object model would have trimmed first."""
    out = [(name, "mixed", body) for name, body in FILE_CASES.items()]
    out += [(name, "nums", body) for name, body in NUM_FILE_CASES.items()]
    out += [(name, "dates", body) for name, body in DATE_FILE_CASES.items()]
    out += [(name, "times", body) for name, body in TIME_FILE_CASES.items()]
    for name, _, _, candidates in FORMATTED:
        for index, candidate in enumerate(candidates):
            escaped = candidate.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;")
            out.append((f"raw_{name}_{index}", name, f'<filters><filter val="{escaped}"/></filters>'))
    return out

COMMON = r"""
Private Function NewSheet(ByVal wb As Workbook, ByVal name As String) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    Set NewSheet = ws
End Function

Private Function SafeCrit(ByVal f As Object, ByVal which As Long) As String
    Dim v As Variant
    Dim i As Long
    Dim parts As String
    On Error Resume Next
    If which = 1 Then v = f.Criteria1
    If which = 2 Then v = f.Criteria2
    If which = 3 Then v = f.Operator
    If Err.Number <> 0 Then
        SafeCrit = "n/a"
        Err.Clear
        Exit Function
    End If
    If IsArray(v) Then
        For i = LBound(v) To UBound(v)
            parts = parts & CStr(v(i)) & "|"
        Next i
        SafeCrit = "[" & parts & "]"
    Else
        SafeCrit = CStr(v)
    End If
    If Err.Number <> 0 Then SafeCrit = "n/a"
    Err.Clear
End Function

Private Function Report(ByVal ws As Worksheet, ByVal code As Long, ByVal last As Long) As String
    Dim i As Long
    Dim hidden As String
    Dim f As Object
    Dim first As String
    Dim second As String
    Dim op As String
    Dim k As Long

    For i = 2 To last
        If ws.Rows(i).Hidden Then hidden = hidden & i & ","
    Next i
    If ws.AutoFilterMode Then
        For k = 1 To ws.AutoFilter.Filters.Count
            Set f = ws.AutoFilter.Filters(k)
            If f.On Then
                first = first & SafeCrit(f, 1) & ";"
                second = second & SafeCrit(f, 2) & ";"
                op = op & SafeCrit(f, 3) & ";"
            End If
        Next k
    End If
    Report = ws.Name & FS() & CStr(code) & FS() & hidden & FS() & _
             first & FS() & second & FS() & op & FS() & _
             CStr(ws.AutoFilterMode) & FS() & CStr(ws.FilterMode)
End Function

Private Function CellRecord(ByVal c As Range) As String
    Dim v As Variant
    v = c.Value2
    Select Case VarType(v)
        Case vbEmpty
            CellRecord = "empty" & FS()
        Case vbString
            CellRecord = "text" & FS() & v
        Case vbBoolean
            CellRecord = "bool" & FS() & IIf(v, "1", "0")
        Case vbError
            CellRecord = "error" & FS() & c.Text
        Case Else
            CellRecord = "number" & FS() & CStr(v)
    End Select
    CellRecord = CellRecord & FS() & c.Text & FS() & c.NumberFormat
End Function

Private Function ColumnCells(ByVal ws As Worksheet, ByVal first As Long, ByVal last As Long, _
                       ByVal column As Long) As String
    Dim r As Long
    Dim out As String
    ws.Columns(column).ColumnWidth = 255
    For r = first To last
        out = out & r & FS() & CellRecord(ws.Cells(r, column)) & CS()
    Next r
    ColumnCells = out
End Function
"""

#: Each column kind: its fill procedure and the last row the filter spans.
#: The filter range can run past the data, as ``beyond_data`` measures.
COLUMNS: dict[str, tuple[str, int]] = {
    "mixed": ("FillMixed", 30),
    "dates": ("FillDates", 38),
    "times": ("FillTimes", 12),
    "nums": ("FillNums", 16),
    "nums_err": ("FillNumsErr", 16),
    "words": ("FillWords", len(WORDS) + 1),
    "symbols": ("FillSymbols", len(SYMBOL_WORDS) + 1),
}


def cases() -> list[tuple[str, str, str]]:
    """Every object-model case: (sheet name, column kind, AutoFilter arguments)."""
    out: list[tuple[str, str, str]] = []
    out += [(name, "mixed", args) for name, args in MIXED_CASES]
    out += [(name, "dates", args) for name, args in DATE_CASES]
    out += [(name, "times", args) for name, args in GROUP_CASES]
    out += [(name, "nums", args) for name, args in NUM_CASES]
    out += [(name, "nums_err", args) for name, args in NUM_ERR_CASES]
    # A ">" pivot per word, skipping those that would read as an operator,
    # a wildcard or a number, which are not text comparisons.
    for index, word in enumerate(WORDS):
        if not _numeric(word):
            out.append((f"w{index:02d}", "words", f'Criteria1:=">" & {vba_string(word)}'))
    for index, word in enumerate(SYMBOL_WORDS):
        if word[0] not in "=<>*?~":
            out.append((f"s{index:02d}", "symbols", f'Criteria1:=">" & {vba_string(word)}'))
    out += [(name, name, args) for name, _, args in EQUALITY]
    for name, _, _, candidates in FORMATTED:
        out += [
            (f"{name}_{index}", name, f"Criteria1:=Array({vba_string(candidate)}), Operator:=xlFilterValues")
            for index, candidate in enumerate(candidates)
        ]
    return out


def _numeric(word: str) -> bool:
    try:
        float(word)
    except ValueError:
        return False
    return True


def last_row(column: str) -> int:
    if column in COLUMNS:
        return COLUMNS[column][1]
    for name, words, _ in EQUALITY:
        if name == column:
            return len(words) + 1
    return 3


def fill_procedure(column: str) -> str:
    return COLUMNS[column][0] if column in COLUMNS else f"Fill_{column}"


def build_source() -> str:
    """The VBA that builds one sheet per case, split into procedures small
    enough for VBA to compile, and reads back each column kind's cells."""
    parts = [SEPARATORS_VBA, COMMON, FILL_MIXED, FILL_DATES, FILL_TIMES, FILL_NUMS,
             fill_text("FillWords", WORDS), fill_text("FillSymbols", SYMBOL_WORDS)]
    parts += [fill_text(f"Fill_{name}", words) for name, words, _ in EQUALITY]
    parts += [fill_formatted(f"Fill_{name}", code, value) for name, code, value, _ in FORMATTED]

    all_cases = cases()
    chunk = 40
    procedures: list[str] = []
    for start in range(0, len(all_cases), chunk):
        body: list[str] = []
        for name, column, args in all_cases[start : start + chunk]:
            last = last_row(column)
            body.append(
                f'    Set ws = NewSheet(wb, "{name}")\n'
                f"    {fill_procedure(column)} ws\n"
                "    On Error Resume Next\n"
                f'    ws.Range("A1:A{last}").AutoFilter Field:=1, {args}\n'
                f'    out = out & Report(ws, Err.Number, {last}) & RS()\n'
                "    Err.Clear\n"
                "    On Error GoTo 0\n"
            )
        procedure = f"Cases{start // chunk}"
        procedures.append(procedure)
        parts.append(
            f"Private Sub {procedure}(ByVal wb As Workbook, ByRef out As String)\n"
            "    Dim ws As Worksheet\n" + "".join(body) + "End Sub\n"
        )

    # A file case starts from its column with a placeholder value filter,
    # whose markup is then replaced.
    file_body = "".join(
        f'    Set ws = NewSheet(wb, "{name}")\n'
        f"    {fill_procedure(column)} ws\n"
        f'    ws.Range("A1:A{last_row(column)}").AutoFilter Field:=1, '
        f'Criteria1:="={5 if column in ("mixed", "nums") else 424242}"\n'
        for name, column, _ in file_cases()
    )
    parts.append(
        "Private Sub FileCases(ByVal wb As Workbook)\n    Dim ws As Worksheet\n"
        + file_body + "End Sub\n"
    )

    columns = list(COLUMNS) + [name for name, _, _ in EQUALITY] + [name for name, _, _, _ in FORMATTED]
    reads = "".join(
        f'    Set ws = NewSheet(wb, "col_{name}")\n'
        f"    {fill_procedure(name)} ws\n"
        f'    out = out & "{name}" & NS() & ColumnCells(ws, 2, {last_row(name)}, 1) & RS()\n'
        for name in columns
    )
    parts.append(
        "Private Sub ReadColumns(ByVal wb As Workbook, ByRef out As String)\n    Dim ws As Worksheet\n"
        + reads
        # The two-column case's second column: a formula giving even or odd.
        + '    Set ws = NewSheet(wb, "col_parity")\n'
        + '    ws.Range("B2:B30").Formula = "=IF(MOD(ROW(),2)=0,""even"",""odd"")"\n'
        + '    out = out & "parity" & NS() & ColumnCells(ws, 2, 30, 2) & RS()\n'
        + "End Sub\n"
    )

    calls = "".join(f"    {procedure} wb, out\n" for procedure in procedures)
    parts.append(
        rf"""
Private Sub SpecialCases(ByVal wb As Workbook, ByRef out As String)
    Dim ws As Worksheet

    ' A filter range running past the data.
    Set ws = NewSheet(wb, "beyond_data")
    FillMixed ws
    On Error Resume Next
    ws.Range("A1:A40").AutoFilter Field:=1, Criteria1:="=5"
    out = out & Report(ws, Err.Number, 40) & RS()
    Err.Clear
    On Error GoTo 0

    ' Arrows only, with a row already hidden by hand.
    Set ws = NewSheet(wb, "no_criteria")
    FillMixed ws
    ws.Rows(7).Hidden = True
    On Error Resume Next
    ws.Range("A1:A30").AutoFilter
    out = out & Report(ws, Err.Number, 30) & RS()
    Err.Clear
    On Error GoTo 0

    ' A row hidden by hand that the criterion keeps.
    Set ws = NewSheet(wb, "hand_hidden_kept")
    FillMixed ws
    ws.Rows(7).Hidden = True
    On Error Resume Next
    ws.Range("A1:A30").AutoFilter Field:=1, Criteria1:="=5"
    out = out & Report(ws, Err.Number, 30) & RS()
    Err.Clear
    On Error GoTo 0

    ' Two columns, one criterion each.
    Set ws = NewSheet(wb, "two_columns")
    FillMixed ws
    ws.Range("B1").Value = "W"
    ws.Range("B2:B30").Formula = "=IF(MOD(ROW(),2)=0,""even"",""odd"")"
    On Error Resume Next
    ws.Range("A1:B30").AutoFilter Field:=1, Criteria1:=">0"
    ws.Range("A1:B30").AutoFilter Field:=2, Criteria1:="=even"
    out = out & Report(ws, Err.Number, 30) & RS()
    Err.Clear
    On Error GoTo 0
End Sub

Public Function Build(ByVal Target As String, ByVal FileTarget As String) As String
    Dim wb As Workbook
    Dim files As Workbook
    Dim out As String
    Dim columns As String

    Set wb = ActiveWorkbook
    wb.Worksheets(1).Name = "About"
{calls}    SpecialCases wb, out
    ReadColumns wb, columns

    ' The file cases get a workbook of their own, which is reopened once per
    ' case and is quicker to open small.
    Set files = Workbooks.Add
    files.Worksheets(1).Name = "About"
    FileCases files

    Application.DisplayAlerts = False
    wb.SaveAs Target, 51
    files.SaveAs FileTarget, 51
    files.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Build = Application.Version & " " & Application.Build & BS() & Format(Date, "yyyy-mm-dd") & _
            BS() & out & BS() & columns
End Function
"""
    )
    return "\n".join(parts)


APPLY = r"""
Public Function Apply(ByVal Target As String, ByVal SheetName As String, ByVal LastRow As Long) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim hidden As String
    Dim i As Long
    Dim code As Long
    Dim onOpen As String

    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(Target)
    If Err.Number <> 0 Then
        Apply = "open-refused" & FS() & CStr(Err.Number)
        Err.Clear
        Application.DisplayAlerts = True
        Exit Function
    End If
    On Error GoTo 0
    Set ws = wb.Worksheets(SheetName)
    onOpen = CStr(ws.AutoFilterMode) & "/" & CStr(ws.FilterMode)
    On Error Resume Next
    ws.AutoFilter.ApplyFilter
    code = Err.Number
    Err.Clear
    On Error GoTo 0
    For i = 2 To LastRow
        If ws.Rows(i).Hidden Then hidden = hidden & i & ","
    Next i
    Apply = "ok" & FS() & CStr(code) & FS() & hidden & FS() & onOpen & FS() & CStr(ws.FilterMode)
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Function
"""


# ---------------------------------------------------------------------------
# Packages
# ---------------------------------------------------------------------------


def sheet_parts(members: dict[str, bytes]) -> dict[str, str]:
    """Sheet name to its worksheet part, following the workbook's own map."""
    book = members["xl/workbook.xml"].decode("utf-8")
    rels = members["xl/_rels/workbook.xml.rels"].decode("utf-8")
    targets = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    return {
        name: "xl/" + targets[rid].removeprefix("/xl/")
        for name, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="([^"]+)"', book)
    }


def auto_filter_markup(xml: str) -> str:
    """The sheet's ``autoFilter`` element as Excel wrote it, less the
    revision ids that change on every save."""
    found = re.search(r"<autoFilter\b[^>]*/>|<autoFilter\b.*?</autoFilter>", xml, re.S)
    return re.sub(r'\s*xr:uid="[^"]*"', "", found.group(0)) if found else ""


def with_filter(members: dict[str, bytes], part: str, body: str) -> dict[str, bytes]:
    """A copy of the package with one sheet's first filter column replaced
    and every hidden flag on it cleared, so ApplyFilter alone decides."""
    xml = members[part].decode("utf-8")
    xml = re.sub(
        r'(<filterColumn colId="0">).*?(</filterColumn>)',
        lambda match: match.group(1) + body + match.group(2), xml, count=1, flags=re.S,
    )
    changed = dict(members)
    changed[part] = xml.replace(' hidden="1"', "").encode("utf-8")
    return changed


def write_package(members: dict[str, bytes], target: Path) -> None:
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for name, body in members.items():
            out.writestr(name, body)


# ---------------------------------------------------------------------------
# Replies
# ---------------------------------------------------------------------------


def parse_cell(fields: list[str]) -> list[object]:
    """``[row, value, text, format]``, the value typed as JSON can carry it."""
    row, kind, raw, text, code = (fields + [""] * 5)[:5]
    value: object
    if kind == "empty":
        value = None
    elif kind == "text":
        value = {"text": raw}
    elif kind == "bool":
        value = raw == "1"
    elif kind == "error":
        value = {"error": raw}
    else:
        value = float(raw)
    return [int(row), value, text, code]


def parse_report(record: str) -> tuple[str, dict[str, object]]:
    fields = (record.split(FS) + [""] * 8)[:8]
    name, code, hidden, first, second, operator, mode, filter_mode = fields
    return name, {
        "error": int(code or 0),
        "hidden": [int(row) for row in hidden.split(",") if row],
        "criteria1": first.rstrip(";"),
        "criteria2": second.rstrip(";"),
        "operator": operator.rstrip(";"),
        "auto_filter_mode": mode == "True",
        "filter_mode": filter_mode == "True",
    }


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

    kinds = {name: column for name, column, _ in cases()}
    specials = {"beyond_data": "mixed", "no_criteria": "mixed", "hand_hidden_kept": "mixed",
                "two_columns": "mixed+parity"}
    # Hidden macro runs only, which the harness allows beside other sessions.
    config = HarnessConfig(exclusive=False)
    with tempfile.TemporaryDirectory() as scratch, ExcelSession(config) as excel:
        base = Path(scratch) / "cases.xlsx"
        file_base = Path(scratch) / "file_cases.xlsx"
        excel.new_workbook()
        excel.reset_sheets()
        built = excel.run_vba(build_source(), proc="Build", args=(str(base), str(file_base)),
                              timeout=900)
        if built.outcome != "passed":
            raise SystemExit(f"Excel refused the build: {built!r}")
        pieces = str(built.value or "").split(BS)
        if len(pieces) != 4:
            raise SystemExit(f"the build's reply has {len(pieces)} blocks, not 4")
        version, date, reports, column_dump = pieces

        with zipfile.ZipFile(base) as archive:
            members = {name: archive.read(name) for name in archive.namelist()}
        parts = sheet_parts(members)
        with zipfile.ZipFile(file_base) as archive:
            file_members = {name: archive.read(name) for name in archive.namelist()}
        file_parts = sheet_parts(file_members)

        results: dict[str, dict[str, object]] = {}
        for record in reports.split(RS):
            if not record.strip():
                continue
            name, result = parse_report(record)
            result["kind"] = "special" if name in specials else "object model"
            result["column"] = specials.get(name) or kinds[name]
            result["markup"] = auto_filter_markup(members[parts[name]].decode("utf-8"))
            results[name] = result

        excel.new_workbook()
        for number, (name, column, body) in enumerate(file_cases()):
            # One package per case: a criterion Excel refuses takes the
            # whole workbook down with it.
            target = Path(scratch) / f"{name}.xlsx"
            write_package(with_filter(file_members, file_parts[name], body), target)
            applied = excel.run_vba(SEPARATORS_VBA + APPLY, proc="Apply",
                                    args=(str(target), name, last_row(column)), timeout=300,
                                    module_name=f"Apply{number}")
            if applied.outcome != "passed":
                raise SystemExit(f"{name}: Excel refused to apply: {applied!r}")
            fields = str(applied.value or "").split(FS)
            # The range as Excel stored it for the placeholder, which it
            # trims to the last row holding data.
            placeholder = auto_filter_markup(file_members[file_parts[name]].decode("utf-8"))
            found = re.search(r'ref="([^"]+)"', placeholder)
            ref = found.group(1) if found else f"A1:A{last_row(column)}"
            markup = f'<autoFilter ref="{ref}"><filterColumn colId="0">{body}</filterColumn></autoFilter>'
            if fields[0] == "open-refused":
                results[name] = {"kind": "file", "column": column, "markup": markup,
                                 "refused_to_open": True, "error": int(fields[1] or 0)}
                continue
            _, code, hidden, on_open, filter_mode = (fields + [""] * 5)[:5]
            results[name] = {
                "kind": "file", "column": column, "markup": markup, "refused_to_open": False,
                "error": int(code or 0), "hidden": [int(row) for row in hidden.split(",") if row],
                "filter_mode": filter_mode == "True", "on_open": on_open,
            }

    columns: dict[str, list[list[object]]] = {}
    for record in column_dump.split(RS):
        if record.strip():
            name, _, cells = record.partition(NS)
            columns[name] = [parse_cell(cell.split(FS)) for cell in cells.split(CS) if cell]

    write(version, date, columns, results)
    print(f"wrote {FIXTURE.relative_to(ROOT)} ({FIXTURE.stat().st_size} bytes): "
          f"{len(results)} cases over {len(columns)} columns")
    return 0


def write(version: str, date: str, columns: dict[str, list[list[object]]],
          results: dict[str, dict[str, object]]) -> None:
    """One case and one cell per line, so a change diffs by case."""
    dump = json.dumps
    lines = ["{", f'"excel": {dump(version)},', f'"built": {dump(date)},', '"columns": {']
    blocks: list[str] = []
    for name, cells in columns.items():
        rows = ",\n".join(f"  {dump(cell)}" for cell in cells)
        blocks.append(f"{dump(name)}: [\n{rows}\n]")
    lines.append(",\n".join(blocks))
    lines.append("},")
    lines.append('"cases": {')
    lines.append(",\n".join(f"{dump(name)}: {dump(result, sort_keys=True)}" for name, result in results.items()))
    lines.append("}")
    lines.append("}")
    FIXTURE.write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    raise SystemExit(main())
