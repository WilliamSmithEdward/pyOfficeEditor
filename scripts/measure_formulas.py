"""Measure what Excel's formulas give, for the formula engine to be held to.

The inputs are written by this library into a copy of ``empty.xlsx``, so a
number is the exact double meant and not what Excel's parser makes of its
digits: ``Numbers`` holds a column of them, ``Pairs`` two numbers a few
units in the last place apart per row, and ``Data`` a small table of every
kind of value. Excel then opens the file, puts every formula below in a cell
of its own, on a sheet named for its group, calculates, and saves. The
result each formula gave is in the saved file beside it, at the full
precision Excel keeps, so the workbook is the corpus.

Run this on a Windows machine with Excel installed:

    python -m pip install -e ".[dev]" --group live
    python scripts/measure_formulas.py

It writes ``tests/fixtures/excel/formulas.xlsx``, replacing what is there,
and nothing else. The locale is part of the result, as it is for the number
formats: the corpus was measured in en-US. The sheet ``About`` records the
Excel build and the day, which matters to a text such as ``"1/15"`` that
names a day in the current year.
"""

from __future__ import annotations

import datetime as dt
import math
import os
import random
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pyofficeeditor.excel import Workbook  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures" / "excel"
FIXTURE = FIXTURES / "formulas.xlsx"

#: Separators between fields and between groups: private-use characters no
#: formula contains.
FS = "\ue001"
RS = "\ue002"


def _numbers() -> list[float]:
    """Doubles whose text Excel is asked for: round ones, ones with more
    digits than it shows, and ones on either side of where it would round
    or switch to an exponent."""
    values = [
        0.0, 1.0, -1.0, 0.5, 1.5, 2.5, -2.5, 0.1, 0.2, 0.3, 0.1 + 0.2, 1 / 3, 2 / 3, -1 / 3, 1 / 7,
        math.pi, math.e, math.sqrt(2), 123456789012345.0, 1234567890123456.0, 12345678901234567.0,
        999999999999999.0, 9999999999999998.0, 99999999999999.99, 0.1234567890123456,
        1.2345678901234567e-5, 2.0**53, 2.0**53 + 2, 2.0**63, 1.7976931348623157e308,
        2.2250738585072014e-308, 1e-307, 0.99999999999999994, 0.9999999999999999, 9.9999999999999995,
        0.95, 0.995, 0.9995, 0.99995, 0.999995, 0.9999995, 0.99999995, 0.999999995, 0.9999999995,
        0.99999999999999, 0.999999999999999, 0.9999999999999995, 1.234567890123455, 1.2345678901234549,
        1.2345678901234551, 123456.7890123455, 0.000123456789012345, 0.0001234567890123455,
        1e21 - 2**17, 1e21, 1e20 + 2**14, 5e-324, 1e-310, 100.0, 1000000.0, 0.001, 0.0001,
        12345678901.0, 123456789012.0, 1234567890123.0, 0.00001, 0.000001, 0.0000001, 1.5e-7,
    ]  # fmt: skip
    values += [10.0**power for power in range(-30, 31)]
    values += [1.5 * 10.0**power for power in range(-12, 23)]
    values += [1.2345678901234567 * 10.0**power for power in range(-12, 23)]
    values += [-value for value in (0.1 + 0.2, 1e-5, 1.5e15, 1e20, 1e-20, 123.456)]
    unique: list[float] = []
    for value in values:
        if value not in unique:
            unique.append(value)
    return unique


NUMBERS = _numbers()


def _step(value: float, units: int) -> float:
    """``value`` moved ``units`` places in the last digit, up or down."""
    direction = math.inf if units > 0 else -math.inf
    for _ in range(abs(units)):
        value = math.nextafter(value, direction)
    return value


def _pairs() -> list[tuple[float, float]]:
    """Two numbers a few units in the last place apart, at several sizes,
    for where Excel starts treating a difference as zero."""
    bases = (1.0, 3.0, 0.1, 123.456, 1e10, 1e-10, 1e300)
    steps = (1, 2, 3, 4, 5, 6, 7, 8, 10, 12, 16, 20, 24, 32, 48, 64, 100, 128, 256, 512, 1024, 2048, 4096)
    pairs: list[tuple[float, float]] = []
    for base in bases:
        for step in steps:
            pairs.append((base, _step(base, step)))
            pairs.append((base, _step(base, -step)))
    return pairs


PAIRS = _pairs()

#: Texts put through arithmetic, VALUE, DATEVALUE and TIMEVALUE. Each is
#: written into a formula as a string, so a quote in one is doubled there.
TEXTS = [
    "1", "-1", "+1", "1.5", ".5", "5.", "1e3", "1E3", "1e+3", "1e-3", "1E", "e3", "1.5.2", "1,000",
    "1,000.5", "1,00", "10,00,000", ",100", "100,", "1 000", " 1", "1 ", "01", "00.5", "1_000",
    "0x1F", "&H1F", "1d2", "Infinity", "NaN", "1e308", "1e309", "-0", "1e-400", "50%", "%50", "50 %",
    "5%%", "-50%", "1.5%", "1e2%", "(50%)", "$5", "-$5", "$-5", "($5)", "$ 5", "5$", "\u20ac5",
    "\u00a35", "$1,000.50", "USD5", "(5)", "(-5)", "-(5)", "( 5 )", "(5", "5)", "5-", "5+", "--5",
    "+-5", "- 5", "1/2", "0 1/2", "1 1/2", "1 1/3", "-1 1/2", "1 1/2/3", "1 3/2", "3/2", " 1 1/2 ",
    "1  1/2", "1/2/2020", "1/2/20", "2020-01-15", "2020/01/15", "15-Jan-2020", "15 Jan 2020",
    "Jan 15, 2020", "January 15, 2020", "Jan 2020", "Jan-20", "15-Jan", "1/15", "13/1/2020",
    "2/29/2021", "2/29/2020", "1/1/1900", "12/31/1899", "1/1/9999", "1/1/10000", "1/1/0", "1/1/29",
    "1/1/30", "Monday, January 15, 2020", "2020-1-15", "15.1.2020", "1-15-2020", "20200115", "12:30",
    "12:30:45", "12:30:45.5", "12:30 PM", "12:30 AM", "12:30PM", "0:00", "24:00", "25:00", "12:60",
    "1:2:3", "12:30:45 PM", "12 PM", "12pm", "1:30 a", "12:30:45.123", "100:00", "9999:00",
    "1/2/2020 12:30", "2020-01-15 12:30:45", "1/2/2020 12:30 PM", "TRUE", "FALSE", "", " ", "abc",
    "#N/A", "e", ".", "-", "$", "%", "()", "1-1", "1+1", "1.5e3%", "$1e3", "1,5", "1,234,567",
    "12,34", "1.234,5", "0.5e1", "5e-1", "2 1/4%", "$(5)", "(-$5)", "1:30:00 PM", "2:00 AM",
    "1/2/2020 25:00", "Jan", "January", "1 Jan", "Jan 1", "2020 Jan 1", "1-Jan-2020 1:00",
]  # fmt: skip

#: Words compared with each other by "=" and "<", both ways round.
WORDS = [
    "a", "A", "b", "B", "", " ", "ab", "a b", "a-b", "a'b", "co-op", "coop", "cop", "\u00df", "ss",
    "\u00e6", "ae", "\u00e9", "e", "f", "\u00c9", "1", "10", "9", "a1", "a10", "a2", "_", "-", "~",
    "z", "\u03a9", "\u03c9",
]  # fmt: skip


def _quoted(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


#: The formulas, by the sheet they go on.
GROUPS: dict[str, list[str]] = {
    "Operators": [
        "=1+2*3", "=(1+2)*3", "=-2^2", "=2^3^2", "=-(2^2)", "=2^-1", "=10/4", "=1/0", "=0/0", "=7%",
        "=50%*2", "=5%%", "=-(-3)", "=+3", "=--3", "=3-", "=1-2-3", "=12/3/2", "=2*3^2", "=(2*3)^2",
        "=0.1+0.2", "=0.1+0.2-0.3", "=1-0.9-0.1", "=(0.1+0.2)-0.3", "=0.3-0.2-0.1", "=1*(0.5-0.4-0.1)",
        "=1E+308*10", "=2^1024", "=2^1023", "=-2^0.5", "=(-8)^(1/3)", "=0^0", "=0^-1", "=10^-400",
        "=1&2", "=1.5&\"\"", "=(0.1+0.2)&\"\"", "=TRUE&\"\"", "=FALSE&1", "=1E+20&\"\"", "=123456789012&\"\"",
        "=1234567890123456&\"\"", "=0.000001&\"\"", "=0.0000001&\"\"", "=1/3&\"\"", "=-1/3&\"\"",
        "=100000000000000000000&\"\"", "=Data!B4&\"x\"", "=Data!B4+1", "=Data!B4*2", "=Data!B4-Data!B4",
        "=\"3\"+4", "=\"3\"&4", "=TRUE+1", "=FALSE*5", "=\"TRUE\"+1", "=\" 3 \"+1", "=\"1,000\"+1",
        "=\"$5\"+1", "=\"1e3\"+1", "=\"50%\"+0", "=\"1/2/2020\"+0", "=\"12:30\"+0", "=\"abc\"+1",
        "=\"\"+1", "=\"(5)\"+0", "=\"-5\"+0", "=\"+5\"+0", "=\"5-\"+0", "=\"1 1/2\"+0", "=\"0x10\"+0",
        "=Data!B2+1", "=Data!B2*1", "=Data!B3+1", "=Data!B1+1", "=Data!B5+1", "=1+#N/A", "=#DIV/0!+#N/A",
        "=#N/A+#DIV/0!", "=1/0&\"x\"",
    ],
    "Compare": [
        "=1<2", "=2<1", "=1=1", "=1<>1", "=1<\"a\"", "=\"a\"<TRUE", "=1<TRUE", "=\"a\"=\"A\"", "=\"a\"<\"B\"",
        "=\"B\"<\"a\"", "=\"10\"<\"9\"", "=10<9", "=\"apple\"<\"Apple\"", "=\"Apple\"<\"apple\"",
        "=\"apple\"=\"Apple\"", "=Data!B4=0", "=Data!B4=\"\"", "=Data!B4=FALSE", "=Data!B4<1", "=Data!B4>-1",
        "=Data!B4<\"a\"", "=\"\"<\"a\"", "=\"\"=0", "=0=FALSE", "=1=TRUE", "=\"1\"=1", "=Data!B2=10",
        "=Data!B2=\"10\"", "=TRUE>FALSE", "=\"z\"<\"aa\"", "=\"a b\"<\"ab\"", "=\"a-b\"<\"ab\"", "=\"\u00e9\"<\"f\"",
        "=\"e\"<\"\u00e9\"", "=\"\u00df\"=\"ss\"", "=#N/A=1", "=1=#N/A", "=0.1+0.2=0.3", "=1/3*3=1", "=\"a\"<=\"A\"",
        "=\"A\">=\"a\"",
    ],
    "Math": [
        "=ROUND(2.5,0)", "=ROUND(-2.5,0)", "=ROUND(2.45,1)", "=ROUND(1.005,2)", "=ROUND(1234.5678,-2)",
        "=ROUND(0.285,2)", "=ROUNDUP(3.14159,2)", "=ROUNDUP(-3.14159,2)", "=ROUNDDOWN(3.99,0)",
        "=ROUNDDOWN(-3.99,0)", "=INT(3.7)", "=INT(-3.2)", "=TRUNC(-3.7)", "=TRUNC(3.14159,3)", "=MOD(5,3)",
        "=MOD(-5,3)", "=MOD(5,-3)", "=MOD(-5,-3)", "=MOD(5.5,2)", "=MOD(5,0)", "=ABS(-4.5)", "=SIGN(-2)",
        "=SIGN(0)", "=SQRT(16)", "=SQRT(-1)", "=POWER(2,10)", "=POWER(-8,1/3)", "=EXP(1)", "=LN(10)",
        "=LOG(100)", "=LOG(8,2)", "=LOG10(1000)", "=LN(0)", "=PI()", "=CEILING(4.3,1)", "=CEILING(-4.3,-1)",
        "=CEILING(4.3,0.5)", "=FLOOR(4.7,1)", "=FLOOR(-4.7,-1)", "=CEILING.MATH(-4.3)", "=FLOOR.MATH(-4.3)",
        "=MROUND(10,3)", "=MROUND(-10,-3)", "=EVEN(3)", "=ODD(4)", "=EVEN(-1.5)", "=FACT(5)", "=FACT(0)",
        "=COMBIN(5,2)", "=PERMUT(5,2)", "=GCD(12,18)", "=LCM(4,6)", "=QUOTIENT(7,2)", "=QUOTIENT(-7,2)",
        "=SUM(1,2,3)", "=SUM(Data!A1:A10)", "=SUM(1,\"2\",TRUE)", "=SUM(Data!B1:B3)", "=SUM(Data!A1:A3,5)",
        "=SUM(Data!B1:B10)", "=SUMPRODUCT(Data!A1:A3,Data!A4:A6)", "=SUMSQ(1,2,3)", "=PRODUCT(2,3,4)",
        "=PRODUCT(Data!A1:A5)", "=SUM()",
    ],
    "Aggregate": [
        "=AVERAGE(Data!A1:A10)", "=AVERAGE(1,\"2\",TRUE)", "=AVERAGE(Data!B1:B4)", "=AVERAGEA(Data!B1:B4)",
        "=MIN(Data!A1:A10)", "=MAX(Data!A1:A10)", "=MIN(Data!B1:B4)", "=MAX(Data!B1:B4)", "=MIN(\"3\",5)",
        "=MAX(Data!B6:B7)", "=MINA(Data!B1:B4)", "=MAXA(Data!B1:B4)", "=COUNT(Data!A1:B10)",
        "=COUNT(1,\"2\",TRUE,\"x\")", "=COUNTA(Data!B1:B10)", "=COUNTBLANK(Data!B1:B10)", "=MEDIAN(1,3,2,4)",
        "=MEDIAN(Data!A1:A10)", "=MODE(1,2,2,3)", "=LARGE(Data!A1:A10,2)", "=SMALL(Data!A1:A10,3)",
        "=LARGE(Data!A1:A10,11)", "=STDEV(Data!A1:A10)", "=STDEVP(Data!A1:A10)", "=VAR(Data!A1:A10)",
        "=VARP(Data!A1:A10)", "=STDEV.S(2,4,4,4,5,5,7,9)", "=STDEV.P(2,4,4,4,5,5,7,9)", "=RANK(3,Data!A1:A10)",
        "=RANK(3,Data!A1:A10,1)", "=AVERAGE(Data!B1)", "=MAX(Data!B5:B6)", "=COUNT(Data!B5)",
    ],
    "Conditional": [
        "=SUMIF(Data!A1:A10,\">5\")", "=SUMIF(Data!C1:C5,\"x\",Data!D1:D5)", "=SUMIF(Data!C1:C5,\"<>x\",Data!D1:D5)",
        "=COUNTIF(Data!A1:A10,\">=3\")", "=COUNTIF(Data!C1:C5,\"x\")", "=COUNTIF(Data!B1:B10,\"apple\")",
        "=COUNTIF(Data!B1:B10,\"a*\")", "=COUNTIF(Data!B1:B10,\"?pple\")", "=COUNTIF(Data!B1:B10,\"*\")",
        "=COUNTIF(Data!B1:B10,\"<>\")", "=COUNTIF(Data!B1:B10,\"\")", "=COUNTIF(Data!B1:B10,10)",
        "=COUNTIF(Data!B1:B10,\"10\")", "=COUNTIF(Data!B1:B10,TRUE)", "=COUNTIF(Data!A1:A10,\"<5\")",
        "=COUNTIF(Data!A1:A10,\"=5\")", "=COUNTIF(Data!A1:A10,5)", "=COUNTIF(Data!A1:A10,\"5\")",
        "=AVERAGEIF(Data!A1:A10,\">5\")", "=AVERAGEIF(Data!C1:C5,\"x\",Data!D1:D5)",
        "=SUMIFS(Data!D1:D5,Data!C1:C5,\"x\",Data!D1:D5,\">15\")", "=COUNTIFS(Data!C1:C5,\"x\",Data!D1:D5,\">15\")",
        "=AVERAGEIFS(Data!D1:D5,Data!C1:C5,\"<>y\")", "=MAXIFS(Data!D1:D5,Data!C1:C5,\"x\")",
        "=MINIFS(Data!D1:D5,Data!C1:C5,\"x\")", "=COUNTIF(Data!C1:C5,\"~*\")", "=SUMIF(Data!A1:A10,\">\"&5)",
        "=COUNTIF(Data!B1:B10,\">b\")", "=COUNTIF(Data!B1:B10,\"<b\")", "=AVERAGEIF(Data!C1:C5,\"q\",Data!D1:D5)",
    ],
    "Logic": [
        "=IF(1>0,\"yes\",\"no\")", "=IF(0,\"yes\",\"no\")", "=IF(\"TRUE\",1,2)", "=IF(\"x\",1,2)", "=IF(1,,2)",
        "=IF(0,1)", "=IF(Data!B4,1,2)", "=AND(TRUE,1,\"TRUE\")", "=AND(TRUE,0)", "=OR(FALSE,0)",
        "=OR(Data!B1:B4)", "=AND(Data!B1)", "=NOT(0)", "=NOT(\"FALSE\")", "=XOR(TRUE,TRUE,TRUE)",
        "=IFERROR(1/0,\"div\")", "=IFERROR(5,\"x\")", "=IFNA(Data!B5,\"na\")", "=IFNA(1/0,\"na\")",
        "=IFS(1>2,\"a\",2>1,\"b\")", "=IFS(1>2,\"a\")", "=SWITCH(2,1,\"one\",2,\"two\",\"other\")",
        "=SWITCH(9,1,\"one\",\"other\")", "=SWITCH(9,1,\"one\")", "=CHOOSE(2,\"a\",\"b\",\"c\")",
        "=CHOOSE(4,\"a\",\"b\",\"c\")", "=CHOOSE(2.9,\"a\",\"b\",\"c\")", "=TRUE()", "=FALSE()",
        "=IF(1/0>1,1,2)", "=IF(TRUE,1/0,2)", "=AND()",
    ],
    "Text": [
        "=LEFT(\"hello\",2)", "=LEFT(\"hello\")", "=RIGHT(\"hello\",3)", "=MID(\"hello\",2,3)", "=MID(\"hello\",9,2)",
        "=MID(\"hello\",0,2)", "=LEN(\"hello\")", "=LEN(1234.5)", "=LEN(Data!B4)", "=LEN(TRUE)",
        "=UPPER(\"Mixed Case\")", "=LOWER(\"Mixed Case\")", "=PROPER(\"hello wORLD o'neil\")",
        "=TRIM(\"  a   b  \")", "=FIND(\"l\",\"hello\")", "=FIND(\"L\",\"hello\")", "=SEARCH(\"L\",\"hello\")",
        "=SEARCH(\"l?o\",\"hello\")", "=SEARCH(\"e*o\",\"hello\")", "=FIND(\"l\",\"hello\",4)",
        "=SUBSTITUTE(\"a-b-c\",\"-\",\"+\")", "=SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)", "=REPLACE(\"abcdef\",2,3,\"X\")",
        "=CONCATENATE(\"a\",1,TRUE)", "=CONCAT(\"a\",Data!A1:A3)", "=TEXTJOIN(\",\",TRUE,Data!B1:B4)",
        "=TEXTJOIN(\",\",FALSE,Data!B1:B4)", "=REPT(\"ab\",3)", "=EXACT(\"a\",\"A\")", "=EXACT(\"a\",\"a\")",
        "=TEXT(1234.567,\"#,##0.00\")", "=TEXT(0.25,\"0%\")", "=TEXT(Data!B10,\"yyyy-mm-dd\")",
        "=TEXT(-5,\"0;(0)\")", "=VALUE(\"12.5\")", "=VALUE(\"abc\")", "=VALUE(\"1,234\")", "=VALUE(\"25%\")",
        "=CHAR(65)", "=CODE(\"A\")", "=UNICODE(\"\u00e9\")", "=UNICHAR(8364)", "=CLEAN(\"a\"&CHAR(7)&\"b\")",
        "=T(\"x\")", "=T(1)", "=N(\"x\")", "=N(TRUE)", "=DOLLAR(1234.567,2)", "=FIXED(1234.567,1)",
        "=FIXED(1234.567,1,TRUE)", "=NUMBERVALUE(\"1.234,5\",\",\",\".\")", "=LEFT(1234,2)",
    ],
    "Info": [
        "=ISBLANK(Data!B4)", "=ISBLANK(Data!B1)", "=ISNUMBER(Data!B2)", "=ISNUMBER(Data!A1)", "=ISTEXT(Data!B2)",
        "=ISTEXT(Data!B4)", "=ISLOGICAL(Data!B3)", "=ISERROR(Data!B5)", "=ISERR(Data!B5)", "=ISNA(Data!B5)",
        "=ISERR(1/0)", "=ISNONTEXT(Data!B4)", "=ISEVEN(4)", "=ISODD(4)", "=TYPE(1)", "=TYPE(\"a\")",
        "=TYPE(TRUE)", "=TYPE(1/0)", "=TYPE(Data!B4)", "=NA()", "=ERROR.TYPE(1/0)", "=ERROR.TYPE(Data!B5)",
        "=ROW(Data!C3)", "=COLUMN(Data!C3)", "=ROWS(Data!A1:B4)", "=COLUMNS(Data!A1:B4)", "=ISREF(Data!A1)",
        "=ISFORMULA(Data!B5)", "=ISFORMULA(Data!B1)",
    ],
    "Lookup": [
        "=VLOOKUP(\"book\",Data!F2:H4,3,FALSE)", "=VLOOKUP(\"BOOK\",Data!F2:H4,2,FALSE)",
        "=VLOOKUP(\"cup\",Data!F2:H4,2,FALSE)", "=VLOOKUP(4,Data!A1:A10,1)", "=VLOOKUP(4.5,Data!A1:A10,1,TRUE)",
        "=VLOOKUP(0,Data!A1:A10,1)", "=VLOOKUP(\"b*\",Data!F2:H4,3,FALSE)", "=VLOOKUP(\"book\",Data!F2:H4,4,FALSE)",
        "=HLOOKUP(\"Qty\",Data!F1:H4,3,FALSE)", "=INDEX(Data!F2:H4,2,3)", "=INDEX(Data!A1:A10,4)",
        "=INDEX(Data!F2:H4,4,1)", "=MATCH(\"ink\",Data!F2:F4,0)", "=MATCH(5.5,Data!A1:A10,1)",
        "=MATCH(5.5,Data!A1:A10)", "=MATCH(\"zzz\",Data!F2:F4,0)", "=MATCH(\"i*\",Data!F2:F4,0)",
        "=XLOOKUP(\"ink\",Data!F2:F4,Data!H2:H4)", "=XLOOKUP(\"cup\",Data!F2:F4,Data!H2:H4,\"none\")",
        "=XLOOKUP(\"cup\",Data!F2:F4,Data!H2:H4)", "=XLOOKUP(4.5,Data!A1:A10,Data!A1:A10,,-1)",
        "=XLOOKUP(4.5,Data!A1:A10,Data!A1:A10,,1)", "=LOOKUP(4.5,Data!A1:A10)", "=CHOOSE(MATCH(\"y\",Data!C1:C5,0),\"a\",\"b\",\"c\")",
        "=INDEX(Data!D1:D5,MATCH(\"z\",Data!C1:C5,0))", "=SUMPRODUCT((Data!C1:C5=\"x\")*Data!D1:D5)",
    ],
    "Dates": [
        "=DATE(2020,1,15)", "=DATE(2020,13,1)", "=DATE(2020,1,0)", "=DATE(1900,2,29)", "=DATE(99,1,1)",
        "=YEAR(Data!B10)", "=MONTH(Data!B10)", "=DAY(Data!B10)", "=WEEKDAY(Data!B10)", "=WEEKDAY(Data!B10,2)",
        "=EDATE(Data!B10,1)", "=EDATE(DATE(2020,1,31),1)", "=EOMONTH(Data!B10,0)", "=EOMONTH(Data!B10,-1)",
        "=DATEDIF(DATE(2020,1,15),DATE(2021,3,1),\"m\")", "=DATEDIF(DATE(2020,1,15),DATE(2021,3,1),\"d\")",
        "=DATEDIF(DATE(2020,1,15),DATE(2021,3,1),\"y\")", "=DATEDIF(DATE(2020,1,15),DATE(2021,3,1),\"md\")",
        "=DATEVALUE(\"2020-01-15\")", "=DATEVALUE(\"1/15/2020\")", "=TIME(12,30,0)", "=TIME(25,0,0)",
        "=HOUR(0.75)", "=MINUTE(0.7534)", "=SECOND(0.75342)", "=TIMEVALUE(\"6:45 PM\")", "=DAYS(DATE(2021,1,1),DATE(2020,1,1))",
        "=NETWORKDAYS(DATE(2020,1,1),DATE(2020,1,31))", "=WORKDAY(DATE(2020,1,1),10)", "=YEARFRAC(DATE(2020,1,1),DATE(2020,7,1))",
        "=WEEKNUM(Data!B10)", "=ISOWEEKNUM(DATE(2021,1,1))", "=DAY(0)", "=YEAR(0)", "=MONTH(-1)",
    ],
    # Which operation a formula ends with decides whether a result a
    # rounding error from zero is shown as zero.
    "NearZero": [
        "=0.1+0.2-0.3", "=(0.1+0.2-0.3)", "=(0.1+0.2)-0.3", "=0.1+(0.2-0.3)", "=-(0.1+0.2)+0.3",
        "=0.3-(0.1+0.2)", "=1*(0.1+0.2-0.3)", "=(0.1+0.2-0.3)*1", "=(0.1+0.2-0.3)/1", "=+(0.1+0.2-0.3)",
        "=-(0.1+0.2-0.3)", "=--(0.1+0.2-0.3)", "=(0.1+0.2-0.3)%", "=SUM(0.1+0.2-0.3)", "=SUM(0.1,0.2,-0.3)",
        "=SUM(0.3,-0.1,-0.2)", "=ABS(0.1+0.2-0.3)", "=IF(TRUE,0.1+0.2-0.3)", "=(0.1+0.2-0.3)&\"\"",
        "=0.1+0.2-0.3=0", "=0.1+0.2=0.3", "=0.3=0.1+0.2", "=0.1+0.2-0.3>0", "=0.1+0.2>0.3", "=0.1+0.2<>0.3",
        "=MAX(0.1+0.2-0.3)", "=N(0.1+0.2-0.3)", "=0.1+0.2-0.3+0", "=0+0.1+0.2-0.3", "=0.1-0.3+0.2",
        "=(0.1+0.2)-(0.3)", "=0.3-0.2-0.1", "=1*(1-0.9-0.1)", "=1-0.9-0.1", "=(1-0.9-0.1)",
        "=0.3-0.1-0.2", "=0.7-0.6-0.1", "=1.1-1-0.1", "=4.35-4.34-0.01", "=100.1-100-0.1",
        "=SUM(0.7,-0.6,-0.1)", "=0.1*3-0.3", "=0.1*3=0.3", "=(0.1*3)-0.3", "=1/3+1/3+1/3-1",
        "=SUM(1/3,1/3,1/3,-1)", "=10.1-10-0.1", "=1E+15+0.3-1E+15", "=1E+16+2-1E+16",
    ],
    "Power": [
        "=2^0.5", "=10^0.3", "=1.1^100", "=3^-3", "=(-8)^(1/3)", "=(-8)^(2/3)", "=(-27)^(1/3)",
        "=(-2)^(1/5)", "=(-32)^0.2", "=(-2)^0.5", "=8^(1/3)", "=1000^(1/3)", "=7^(1/7)",
        "=1.0000001^10000000", "=0.5^1074", "=2^-1022", "=2^-1023", "=2^-1074", "=2^-1075",
        "=(-8)^(1/3)*1", "=(-1)^(1/3)", "=(-64)^(1/6)", "=(-64)^(1/3)", "=(-0.001)^(1/3)",
        "=POWER(-8,1/3)", "=POWER(10,-2)", "=POWER(0,0)", "=POWER(0,-1)", "=POWER(-1,0.5)", "=10^15",
        "=10^16", "=10^-15", "=2^53", "=3^40", "=1.5^2.5", "=(-1.5)^3", "=(-2)^-3", "=(-8)^-(1/3)",
        "=(-8)^0.333333333333333", "=(-8)^(1/3+0)", "=0^0.5", "=0^-0.5", "=(-0)^2", "=EXP(LN(8)/3)",
        "=SQRT(2)", "=SQRT(2)^2", "=EXP(1)^2", "=EXP(2)", "=EXP(709)", "=EXP(710)", "=EXP(-745)",
        "=EXP(-746)", "=LN(2)", "=LOG(2)", "=LOG10(2)", "=LOG(8,2)", "=LOG(1000,10)", "=LOG(10,100)",
        "=SIN(PI())", "=COS(PI()/2)", "=TAN(PI()/4)", "=ASIN(1)", "=ACOS(-1)", "=ATAN(1)", "=ATAN2(1,1)",
        "=SINH(1)", "=COSH(1)", "=TANH(0.5)", "=DEGREES(PI())", "=RADIANS(180)", "=SIN(1E+8)",
        "=SIN(2^27)", "=SIN(2^28)", "=COS(1E+9)", "=TAN(1E+10)",
    ],
    "Errors": [
        "=SUM(#N/A,#DIV/0!)", "=SUM(#DIV/0!,#N/A)", "=#N/A&#DIV/0!", "=MAX(1/0,NA())", "=IF(#N/A,1,2)",
        "=AND(#N/A,FALSE)", "=AND(FALSE,#N/A)", "=1/0=NA()", "=NA()=1/0", "=-#N/A", "=#N/A%",
        "=#REF!", "=#NULL!", "=#NAME?", "=#NUM!", "=#VALUE!", "=#GETTING_DATA", "=ABS(#N/A)",
        "=LEN(#DIV/0!)", "=\"a\"+#N/A", "=#N/A+\"a\"", "=\"a\"*1/0", "=1/0*\"a\"", "=Data!B1*Data!B5",
        "=Data!B5*Data!B1", "=NOSUCHFUNCTION(1)", "=nosuchname", "=nosuchname+1", "=SUM(nosuchname)",
    ],
    "Literals": [
        "=1.23456789012345678", "=2.9999999999999996", "=0.30000000000000004", "=1234567890123456789",
        "=0.1234567890123456789", "=123456789012345.6", "=9.99999999999999999", "=1.000000000000001",
        "=1.0000000000000001", "=0.1+0.2", "=1234567890123456", "=1234567890123459", "=9999999999999999",
        "=99999999999999999", "=1.5E+308", "=1.8E+308", "=1E-307", "=1E-308", "=1E-309", "=2.2250738585072014E-308",
        "=1E-320", "=.5", "=5.", "=1e5", "=1E5", "=1.E5", "=0.000000000000000000001",
    ],
}  # fmt: skip


def _generated() -> dict[str, list[str]]:
    count = len(NUMBERS)
    pairs = len(PAIRS)
    return {
        "ToText": [f'=Numbers!A{row}&""' for row in range(1, count + 1)],
        "General": [f'=TEXT(Numbers!A{row},"General")' for row in range(1, count + 1)],
        "Difference": [f"=Pairs!B{row}-Pairs!A{row}" for row in range(1, pairs + 1)],
        "DifferenceSum": [f"=SUM(Pairs!B{row},-Pairs!A{row})" for row in range(1, pairs + 1)],
        "Equal": [f"=Pairs!A{row}=Pairs!B{row}" for row in range(1, pairs + 1)],
        "Less": [f"=Pairs!A{row}<Pairs!B{row}" for row in range(1, pairs + 1)],
        "Coerce": [f"={_quoted(text)}+0" for text in TEXTS]
        + ['=(CHAR(9)&"1")+0', '=("1"&CHAR(9))+0', '=(CHAR(10)&"1")+0', '=(CHAR(160)&"1")+0'],
        "Value": [f"=VALUE({_quoted(text)})" for text in TEXTS],
        "DateValue": [f"=DATEVALUE({_quoted(text)})" for text in TEXTS],
        "TimeValue": [f"=TIMEVALUE({_quoted(text)})" for text in TEXTS],
        "WordsEqual": [f"={_quoted(a)}={_quoted(b)}" for a in WORDS for b in WORDS],
        "WordsLess": [f"={_quoted(a)}<{_quoted(b)}" for a in WORDS for b in WORDS],
    }


def _powers() -> list[tuple[float, float]]:
    """Bases and exponents for ``^``: random, round, negative, and whole
    exponents small and large, where the method Excel uses shows in the
    last bit."""
    chance = random.Random(20260922)
    cases = [(chance.uniform(0.01, 100), chance.uniform(-4, 4)) for _ in range(120)]
    fractions = (1 / 3, -1 / 3, 0.5, -0.5, 1 / 7, -1 / 7, 0.25, -0.25, 1.5, -1.5, 2.5, 0.1, -0.1, 2 / 3, 1 / 9)
    for base in (2.0, 3.0, 5.0, 7.0, 8.0, 10.0, 27.0, 64.0, 100.0, 1000.0, 0.5, 0.1):
        cases += [(base, exponent) for exponent in fractions]
    roots = (1 / 3, -1 / 3, 1 / 5, -1 / 5, 2 / 3, 5 / 3, -5 / 3, 3 / 5, 1 / 7, 7 / 3, -7 / 3, 0.2, -0.2, 1 / 9)
    for base in (-8.0, -27.0, -2.0, -32.0, -64.0, -1000.0, -0.125, -3.0):
        cases += [(base, exponent) for exponent in roots]
    for _ in range(20):
        base = chance.uniform(0.5, 2)
        cases += [(base, float(power)) for power in range(-30, 31, 6)]
    cases += [
        (1.0000001, 1e7), (0.9999999, 1e7), (1.000000001, 2.0**31), (1.000000001, 2.0**31 + 1),
        (0.999999999, 1e9), (2.0, 1023.0), (2.0, -1022.0), (-2.0, 3.0), (-0.5, -3.0), (10.0, 22.0),
        (10.0, -22.0), (3.0, 40.0), (7.0, 20.0), (1.1, 1000.0), (-1.1, 99.0), (1.0001, 2.0**40),
        (2.0, 0.5 + 1), (9.0, 1.5), (16.0, 0.75), (1e10, 0.5), (1e-10, 0.5), (2.0, 10.5),
    ]  # fmt: skip
    return cases


POWERS = _powers()


def _reals() -> list[float]:
    chance = random.Random(314159)
    values = [chance.uniform(-10, 10) for _ in range(60)]
    values += [0.5, 1.0, 2.0, -0.5, 1e-8, math.pi / 4, math.pi / 3, math.pi / 6, 3 * math.pi / 4, 100.0, 1e5, 1e7]
    return values


REALS = _reals()


def _near(value: float, steps: int = 2) -> list[float]:
    """``value`` and its neighbours a few units in the last place away."""
    return [_step(value, units) for units in range(-steps, steps + 1)]


#: Numbers and places to round them to, at the boundaries where reading a
#: number to fifteen digits and reading its exact binary value differ.
ROUNDINGS: list[tuple[float, float]] = [
    (value, float(places))
    for base, places in (
        (2.675, 2), (1.005, 2), (0.285, 2), (1234.5, 0), (0.5, 0), (2.5, 0), (-2.5, 0), (4.35, 1),
        (8.345, 2), (0.045, 2), (1.0049999999999999, 2), (2.345, 2), (5.015, 2), (123.4565, 3),
        (0.15, 1), (0.25, 1), (0.35, 1), (1e15 + 0.5, 0), (5e-5, 4), (0.1 + 0.2, 1), (0.1 * 3, 1),
        (1234.5678, -2), (1250.0, -2), (-1250.0, -2), (1.23456789012345, 14), (0.3, 20),
    )
    for value in _near(base)
]  # fmt: skip

#: Numbers just around a whole number, for INT and TRUNC.
WHOLES = [
    value
    for base in (3.0, -3.0, 1.0, 0.0, 100.0, 1e15, 2.0**52, 0.1 * 3 * 10, 4.35 * 100, 1.1 * 3)
    for value in _near(base, 4)
]

#: Dividends and divisors, for MOD, QUOTIENT, CEILING, FLOOR and MROUND.
DIVISIONS = [
    (10.0, 3.3), (5.5, 2.0), (1e10, 3.0), (2.0**27, 3.0), (1.5e8, 7.0), (1.0, 0.1), (0.3, 0.1), (7.0, 0.7),
    (-7.0, 0.7), (7.0, -0.7), (1e15, 7.0), (2.5e15, 3.0), (4.3, 0.1), (4.3, 0.5), (-4.3, 0.1), (2.0**40, 3.0),
    (123.456, 0.01), (1.25, 0.05), (99.99, 0.1), (0.6, 0.2), (3.0, 1.5), (1e300, 7.0), (5.0, 1e-300),
    (2.0**53, 3.0), (2.0**53 + 2, 3.0), (6.0, 0.3), (0.7, 0.1), (11.0, 2.2),
]  # fmt: skip

#: Fractions of a day around whole seconds, minutes and hours.
TIMES = [
    value
    for base in (0.5, 0.75, 45000 / 86400, 1 / 86400, 59 / 86400, 3599 / 86400, 86399 / 86400, 0.999999999, 1.5)
    for value in (base, base - 0.4 / 86400, base - 0.5 / 86400, base - 0.6 / 86400, base + 0.4 / 86400, base + 0.5 / 86400)
]  # fmt: skip

#: Start and end dates for DAYS360 and YEARFRAC, round the ends of months.
DATE_PAIRS = [
    (dt.date(2020, 1, 31), dt.date(2020, 2, 29)), (dt.date(2020, 2, 29), dt.date(2020, 3, 31)),
    (dt.date(2019, 2, 28), dt.date(2019, 3, 31)), (dt.date(2020, 1, 30), dt.date(2020, 3, 31)),
    (dt.date(2020, 3, 31), dt.date(2021, 3, 30)), (dt.date(2020, 1, 1), dt.date(2020, 12, 31)),
    (dt.date(2019, 6, 15), dt.date(2021, 2, 28)), (dt.date(2019, 2, 28), dt.date(2020, 2, 29)),
    (dt.date(2020, 2, 29), dt.date(2021, 2, 28)), (dt.date(2021, 2, 28), dt.date(2021, 3, 31)),
    (dt.date(2020, 5, 31), dt.date(2020, 8, 31)), (dt.date(2020, 5, 30), dt.date(2020, 8, 31)),
    (dt.date(2020, 12, 31), dt.date(2020, 1, 1)), (dt.date(2019, 1, 1), dt.date(2023, 1, 1)),
    (dt.date(2020, 7, 1), dt.date(2021, 6, 30)), (dt.date(2019, 3, 1), dt.date(2020, 2, 29)),
    (dt.date(2020, 2, 28), dt.date(2020, 2, 29)), (dt.date(2019, 12, 31), dt.date(2020, 3, 1)),
]  # fmt: skip

#: What the criteria range of the COUNTIF edge cases holds, one per row:
#: a formula where it starts with "=".
CRITERIA_CELLS: list[object] = [
    10.0, "10", "10.0", "abc", "Apple", "apple", "\u00df", "ss", "SS", "\u00e6", "ae", "\u00e9", "e", True,
    False, "TRUE", "=NA()", '=""', None, 5.0, "5", dt.date(2020, 1, 2), 43832.0, "*", "a*e", " 10", 0.0,
    -1.0, 0.5, "1/2/2020",
]  # fmt: skip

#: The criteria tried against them, as formula text.
CRITERIA = [
    '""', '"="', '"<>"', '"=10"', '"10"', "10", '">5"', '"<5"', '">=10"', '"<>10"', '"a*"', '"*"', '"?"',
    '"\u00df"', '"SS"', '"\u00e6"', '"ae"', '"\u00e9"', '"e"', "TRUE", '"TRUE"', '"#N/A"', '"<>#N/A"',
    '">a"', '"<z"', '"1/2/2020"', '">1/1/2020"', '"12:00"', '" 10"', '"10 "', '"~*"', '"=a*e"', '"<>abc"',
    "Crit!D1", '"=5"', '">4.5"', '"<>5"',
]  # fmt: skip

#: Exact halves at the sixteenth digit, for the rounding of number text.
TIES = [1234567890123445.0, 1234567890123455.0, 1234567890123465.0, 1234567890123475.0, 1000000000000005.0,
        9999999999999995.0, 2.0**53 - 1, 1125899906842624.5]  # fmt: skip

EXTRA_GROUPS: dict[str, list[str]] = {
    "ValueLong": [
        '=VALUE("1.23456789012345678")', '=VALUE("2.9999999999999996")', '=VALUE("1e-308")',
        '=VALUE("1e-307")', '=VALUE("2.2250738585072014e-308")', '=VALUE("9.99999999999999e307")',
        '=VALUE("9.999999999999995e307")', '=VALUE("1234567890123456789")', '=VALUE("0.30000000000000004")',
        '="1.23456789012345678"+0', '="12345678901234567"+0', '=VALUE("1.5e-308")',
    ],
    "Lookups2": [
        "=MATCH(Crit!D1,Crit!A1:A30,0)", "=VLOOKUP(10,Crit!A1:B30,2,FALSE)", '=VLOOKUP("10",Crit!A1:B30,2,FALSE)',
        '=MATCH("a?e",Crit!A1:A30,0)', "=MATCH(5,{1,3,5,7},1)", "=MATCH(6,{1,3,5,7},1)", "=MATCH(0,{1,3,5,7},1)",
        "=MATCH(6,{7,5,3,1},-1)", "=MATCH(8,{7,5,3,1},-1)", "=MATCH(4,{5,1,9,2,7},1)", "=MATCH(6,{5,1,9,2,7},1)",
        '=MATCH("b",{"a","c","b"},1)', '=LOOKUP(2.5,{1,2,3},{"a","b","c"})', "=LOOKUP(0,{1,2,3})",
        "=XLOOKUP(4,{1,3,5,7},{10,30,50,70},,-1)", "=XLOOKUP(4,{1,3,5,7},{10,30,50,70},,1)",
        '=XLOOKUP("b*",{"apple","banana"},{1,2},,2)', '=XLOOKUP("b*",{"apple","banana"},{1,2})',
        "=XLOOKUP(5,{1,3,5,7},{10,30,50,70},,0,2)", "=XLOOKUP(5,{7,5,3,1},{70,50,30,10},,0,-2)",
        "=XLOOKUP(3,{1,3,3,5},{1,2,3,4},,0,-1)", "=XMATCH(4,{1,3,5,7},-1)", "=XMATCH(4,{1,3,5,7},1)",
        '=INDEX({1,2;3,4},0,2)', '=SUM(INDEX({1,2;3,4},0,2))', '=INDEX({1,2;3,4},2)', "=INDEX({1,2,3},2)",
        "=VLOOKUP(2,{1,\"a\";2,\"b\";3,\"c\"},2)", "=HLOOKUP(2,{1,2,3;\"a\",\"b\",\"c\"},2)",
    ],
    "Stats2": [
        "=LARGE({5,1,9,3},1.5)", "=SMALL({5,1,9,3},2.1)", "=LARGE({5,1,9,3},0.5)", "=MEDIAN(5,1,9,3)",
        "=MODE(1,2,2,3,3)", "=MODE(3,3,2,2)", "=RANK(2,{1,2,2,3})", "=RANK.AVG(2,{1,2,2,3})",
        "=RANK(2,{1,2,2,3},1)", "=AVERAGE(1,)", "=COUNT(1,)", "=COUNTA(1,)", "=SUM(1,)", "=MAX(-1,)",
        "=SUM(0.1,0.2,-0.3,1E-20)", "=1*SUM(0.1,0.2,-0.3)", "=SUM(0.1,0.2,-0.3)*1", "=SUM(1,1E-16,-1)",
        "=SUM(1E+16,1,-1E+16)", "=SUM({0.1,0.2,-0.3})", "=AVERAGE(0.1,0.2,-0.3)", "=PRODUCT()*1",
        "=STDEV(1)", "=VAR(1,1)", "=STDEVA(1,TRUE,\"a\")", "=AVEDEV(1,2,3,4)", "=DEVSQ(1,2,3,4)",
        "=GEOMEAN(1,2,4)", "=HARMEAN(1,2,4)", "=COUNTBLANK(Crit!A1:A30)", "=COUNTA(Crit!A1:A30)",
        "=COUNT(Crit!A1:A30)", "=SUMPRODUCT(Crit!A1:A30)", '=SUMPRODUCT(--(Crit!A1:A30="10"))',
    ],
    "TextFns2": [
        '=UPPER("stra\u00dfe")', '=PROPER("2nd place")', '=PROPER("\u00e9t\u00e9 D\'\u00c9T\u00c9")', "=LEN(CLEAN(CHAR(127)))",
        "=CODE(CHAR(129))", '=CODE("\u20ac")', "=CHAR(128)", '=FIND("","abc")', '=SEARCH("","abc")',
        '=FIND("","abc",4)', '=FIND("","abc",5)', '=SUBSTITUTE("aaa","a","b",2)', '=SUBSTITUTE("aaa","aa","b")',
        '=TRIM(CHAR(160)&"a"&CHAR(160))', '=LEN(TRIM(" a  b "))', '=LEFT("abc",-1)', '=RIGHT("abc",0)',
        '=MID("abc",2,0)', '=REPT("ab",0)', '=TEXT(0.5,"hh:mm")', '=TEXT(1234.5,"0.00E+00")',
        '=TEXT("abc","@@")', '=TEXT(-1,"0;-0;zero")', '=TEXT(0,"0;-0;zero")', '=TEXT(TRUE,"0")',
        '=DOLLAR(-1234.567,1)', '=DOLLAR(1234.567,-2)', '=FIXED(-1234.567,-1,TRUE)', '=FIXED(0.5,0)',
        '=VALUE("$1,234.5")', '=NUMBERVALUE("12%")', '=NUMBERVALUE("1 234,5",",")', '=T(Crit!A1)',
        '=EXACT(1,"1")', '=CONCATENATE(1.5,TRUE,"x")', '=UNICODE("\U0001f600")', "=UNICHAR(128512)",
        '=SEARCH("~*","a*b")', '=SEARCH("a~?","xa?")', '=LOWER("\u0130")', '=UPPER("\u0131")',
    ],
    "DateFns2": [
        "=DATE(1900,1,0)", "=DATE(1900,3,0)", "=DATE(-1,1,1)", "=DATE(10000,1,1)", "=DATE(9999,12,31)",
        "=DATE(1899,12,31)", "=DATE(2020,0,0)", "=DATE(2020,-13,1)", "=EDATE(DATE(1900,1,31),1)",
        "=WEEKDAY(0)", "=WEEKDAY(60)", "=WEEKDAY(61)", "=WEEKDAY(DATE(2020,1,15),3)", "=WEEKDAY(DATE(2020,1,15),17)",
        "=WEEKNUM(DATE(2020,1,1),2)", "=WEEKNUM(DATE(2020,12,31),21)", "=WEEKNUM(DATE(2021,1,3),1)",
        "=YEAR(2958465.99)", "=YEAR(2958466)", "=NETWORKDAYS(DATE(2020,1,31),DATE(2020,1,1))",
        "=WORKDAY(DATE(2020,1,3),1,DATE(2020,1,6))", "=WORKDAY(DATE(2020,1,6),-1)", "=EOMONTH(DATE(2020,1,31),-13)",
        '=DATEDIF(DATE(2020,1,31),DATE(2020,3,1),"md")', '=DATEDIF(DATE(2020,1,31),DATE(2020,3,1),"yd")',
        '=DATEDIF(DATE(2019,3,1),DATE(2020,2,29),"yd")', '=DATEDIF(DATE(2020,2,29),DATE(2021,2,28),"y")',
        '=DATEDIF(DATE(2020,1,15),DATE(2021,3,1),"ym")', "=TIME(0,0,-1)", "=TIME(0,-1,120)", "=TIME(12.9,0,0)",
        "=TIME(32767,0,0)", "=TIME(32768,0,0)", "=HOUR(-0.5)", "=HOUR(1.5)", "=SECOND(1/86400*0.5)",
        '=DATEVALUE("29-Feb-1900")', '=DATEVALUE("2/29/1900")', '=DAY("2/29/1900")', "=DAYS(10,5.9)",
        "=ISOWEEKNUM(DATE(2020,12,31))", "=ISOWEEKNUM(1)", "=MONTH(59)", "=DAY(59)", "=DAY(61)",
    ],
}  # fmt: skip


#: The families of functions the fourth batch measures, over the inputs on
#: ``Stats`` (x, y and growth columns, two rows hidden) and ``Db`` (a small
#: database with criteria beside it).
FAMILIES: dict[str, list[str]] = {
    "Financial": [
        "=PMT(0.08/12,10,10000)", "=PMT(0.08/12,10,10000,0,1)", "=PMT(0,10,10000)", "=PMT(0.05,5,-1000,500)",
        "=PMT(-0.1,5,1000)", "=PMT(0.08,0,1000)", "=PV(0.08/12,240,500)", "=PV(0,10,-100)",
        "=PV(0.05,10,-100,1000,1)", "=FV(0.06/12,10,-200,-500,1)", "=FV(0.12/12,12,-1000)", "=FV(0,10,-100,-50)",
        "=FV(0.05,10,0,-1000)", "=NPER(0.12/12,-100,-1000,10000,1)", "=NPER(0.01,-100,1000)", "=NPER(0,-100,1000)",
        "=NPER(0.01,-1,1000)", "=RATE(48,-200,8000)", "=RATE(10,-100,1000,0,0,0.1)", "=RATE(10,100,1000)",
        "=RATE(360,-1000,150000)", "=RATE(12,-100,1000,0,1)", "=IPMT(0.1/12,1,3*12,8000)", "=IPMT(0.1,3,3,8000)",
        "=IPMT(0.1,1,3,8000,0,1)", "=IPMT(0.1,2,3,8000,0,1)", "=PPMT(0.1/12,1,24,2000)", "=PPMT(0.08,10,10,200000)",
        "=CUMIPMT(0.09/12,30*12,125000,13,24,0)", "=CUMIPMT(0.09/12,30*12,125000,1,1,0)",
        "=CUMPRINC(0.09/12,30*12,125000,13,24,0)", "=CUMPRINC(0.09/12,30*12,125000,1,1,1)",
        "=NPV(0.1,-10000,3000,4200,6800)", "=NPV(0.08,Stats!A1:A5)", "=NPV(0.1,1,,2)",
        "=XNPV(0.09,{-10000,2750,4250,3250,2750},{39448,39508,39751,39859,39904})",
        "=IRR({-70000,12000,15000,18000,21000})", "=IRR({-70000,12000,15000,18000,21000,26000})",
        "=IRR({-70000,12000,15000},-0.1)", "=IRR({1,2,3})",
        "=XIRR({-10000,2750,4250,3250,2750},{39448,39508,39751,39859,39904})",
        "=MIRR({-120000,39000,30000,21000,37000,46000},0.1,0.12)", "=SLN(30000,7500,10)", "=SYD(30000,7500,10,1)",
        "=SYD(30000,7500,10,10)", "=DB(1000000,100000,6,1,7)", "=DB(1000000,100000,6,2,7)", "=DB(1000000,100000,6,7,7)",
        "=DB(1000000,100000,6,6)", "=DDB(2400,300,10*365,1)", "=DDB(2400,300,10,1,2)", "=DDB(2400,300,10,10)",
        "=DDB(2400,300,10,2,1.5)", "=VDB(2400,300,10*365,0,1)", "=VDB(2400,300,10,0,0.875,1.5)",
        "=VDB(2400,300,10,6,10,2,TRUE)", "=VDB(2400,300,10,6,10)", "=EFFECT(0.0525,4)", "=NOMINAL(0.053543,4)",
        "=FVSCHEDULE(1,{0.09,0.11,0.1})", "=PDURATION(0.025,2000,2200)", "=RRI(96,10000,11000)",
        "=ISPMT(0.1/12,1,36,8000000)", "=DOLLARDE(1.02,16)", "=DOLLARFR(1.125,16)",
        "=TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)", "=TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)",
        "=TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)", "=PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)",
        "=YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)",
        "=DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1)", "=MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)",
        "=ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0)",
        "=ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,3)", "=COUPDAYS(DATE(2011,1,25),DATE(2011,11,15),2,1)",
        "=COUPDAYBS(DATE(2011,1,25),DATE(2011,11,15),2,1)", "=COUPDAYSNC(DATE(2011,1,25),DATE(2011,11,15),2,1)",
        "=COUPNCD(DATE(2011,1,25),DATE(2011,11,15),2,1)", "=COUPNUM(DATE(2007,1,25),DATE(2008,11,15),2,1)",
        "=COUPPCD(DATE(2011,1,25),DATE(2011,11,15),2,1)", "=DISC(DATE(2018,1,25),DATE(2018,6,15),97.975,100,1)",
        "=INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,2)",
        "=RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,0.0575,2)",
        "=PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)",
        "=PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)",
        "=YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)",
        "=YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)",
        "=AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)",
        "=AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)",
    ],
    "Distributions": [
        "=NORM.DIST(42,40,1.5,TRUE)", "=NORM.DIST(42,40,1.5,FALSE)", "=NORMDIST(42,40,1.5,TRUE)",
        "=NORM.DIST(-5,0,1,TRUE)", "=NORM.DIST(8,0,1,TRUE)", "=NORM.DIST(-40,0,1,TRUE)", "=NORM.S.DIST(1.333333,TRUE)",
        "=NORM.S.DIST(1.333333,FALSE)", "=NORMSDIST(1.333333)", "=NORM.S.DIST(-10,TRUE)", "=NORM.S.DIST(0,TRUE)",
        "=NORM.INV(0.908789,40,1.5)", "=NORMINV(0.908789,40,1.5)", "=NORM.S.INV(0.908789)", "=NORMSINV(0.5)",
        "=NORM.S.INV(1E-10)", "=NORM.S.INV(0.999999)", "=NORM.S.INV(0)", "=STANDARDIZE(42,40,1.5)",
        "=T.DIST(60,1,TRUE)", "=T.DIST(8,3,FALSE)", "=T.DIST.2T(1.959999998,60)", "=T.DIST.RT(1.959999998,60)",
        "=TDIST(1.959999998,60,2)", "=TDIST(1.959999998,60,1)", "=T.DIST(-1.5,10,TRUE)", "=T.INV(0.75,2)",
        "=T.INV.2T(0.546449,60)", "=TINV(0.546449,60)", "=CHISQ.DIST(0.5,1,TRUE)", "=CHISQ.DIST(2,3,FALSE)",
        "=CHISQ.DIST.RT(18.307,10)", "=CHIDIST(18.307,10)", "=CHISQ.INV(0.93,1)", "=CHISQ.INV.RT(0.050001,10)",
        "=CHIINV(0.050001,10)", "=F.DIST(15.2069,6,4,TRUE)", "=F.DIST(15.2069,6,4,FALSE)", "=F.DIST.RT(15.2069,6,4)",
        "=FDIST(15.2069,6,4)", "=F.INV(0.01,6,4)", "=F.INV.RT(0.01,6,4)", "=FINV(0.01,6,4)",
        "=BINOM.DIST(6,10,0.5,FALSE)", "=BINOM.DIST(6,10,0.5,TRUE)", "=BINOMDIST(6,10,0.5,FALSE)",
        "=BINOM.INV(6,0.5,0.75)", "=CRITBINOM(6,0.5,0.75)", "=BINOM.DIST.RANGE(60,0.75,48)",
        "=BINOM.DIST.RANGE(60,0.75,45,50)", "=POISSON.DIST(2,5,TRUE)", "=POISSON.DIST(2,5,FALSE)",
        "=POISSON(2,5,TRUE)", "=EXPON.DIST(0.2,10,TRUE)", "=EXPON.DIST(0.2,10,FALSE)", "=EXPONDIST(0.2,10,TRUE)",
        "=GAMMA(2.5)", "=GAMMA(-3.75)", "=GAMMA(0)", "=GAMMALN(4)", "=GAMMALN.PRECISE(4.5)",
        "=GAMMA.DIST(10.00001131,9,2,FALSE)", "=GAMMA.DIST(10.00001131,9,2,TRUE)", "=GAMMADIST(10.00001131,9,2,TRUE)",
        "=GAMMA.INV(0.068094,9,2)", "=GAMMAINV(0.068094,9,2)", "=BETA.DIST(2,8,10,TRUE,1,3)",
        "=BETA.DIST(2,8,10,FALSE,1,3)", "=BETADIST(2,8,10,1,3)", "=BETA.INV(0.685470581,8,10,1,3)",
        "=BETAINV(0.685470581,8,10,1,3)", "=LOGNORM.DIST(4,3.5,1.2,TRUE)", "=LOGNORM.DIST(4,3.5,1.2,FALSE)",
        "=LOGNORMDIST(4,3.5,1.2)", "=LOGNORM.INV(0.039084,3.5,1.2)", "=LOGINV(0.039084,3.5,1.2)",
        "=WEIBULL.DIST(105,20,100,TRUE)", "=WEIBULL.DIST(105,20,100,FALSE)", "=WEIBULL(105,20,100,TRUE)",
        "=HYPGEOM.DIST(1,4,8,20,TRUE)", "=HYPGEOM.DIST(1,4,8,20,FALSE)", "=HYPGEOMDIST(1,4,8,20)",
        "=NEGBINOM.DIST(10,5,0.25,TRUE)", "=NEGBINOM.DIST(10,5,0.25,FALSE)", "=NEGBINOMDIST(10,5,0.25)",
        "=CONFIDENCE(0.05,2.5,50)", "=CONFIDENCE.NORM(0.05,2.5,50)", "=CONFIDENCE.T(0.05,1,50)", "=FISHER(0.75)",
        "=FISHERINV(0.972955)", "=PHI(0.75)", "=GAUSS(2)", "=ERF(0.745)", "=ERF(1,2)", "=ERF.PRECISE(0.745)",
        "=ERFC(1)", "=ERFC.PRECISE(1)", "=T.TEST(Stats!A1:A10,Stats!B1:B10,2,1)", "=T.TEST(Stats!A1:A10,Stats!B1:B10,2,2)",
        "=T.TEST(Stats!A1:A10,Stats!B1:B10,1,3)", "=TTEST(Stats!A1:A10,Stats!B1:B10,2,2)",
        "=F.TEST(Stats!A1:A10,Stats!B1:B10)", "=Z.TEST(Stats!A1:A10,4)", "=Z.TEST(Stats!A1:A10,4,2)",
        "=CHISQ.TEST(Stats!A1:B3,Stats!C1:D3)", "=PERMUTATIONA(3,2)", "=COMBINA(4,3)", "=MULTINOMIAL(2,3,4)",
        "=SERIESSUM(PI()/4,0,2,{1,-0.5,0.041666667,-0.001388889})", "=GAMMA(171)", "=GAMMA(172)",
        "=NORM.S.DIST(38,TRUE)", "=NORM.S.INV(0.02425)", "=T.DIST(0.5,1.5,TRUE)", "=CHISQ.DIST(3,2.5,TRUE)",
    ],
    "Descriptive": [
        "=CORREL(Stats!A1:A20,Stats!B1:B20)", "=PEARSON(Stats!A1:A20,Stats!B1:B20)", "=RSQ(Stats!B1:B20,Stats!A1:A20)",
        "=SLOPE(Stats!B1:B20,Stats!A1:A20)", "=INTERCEPT(Stats!B1:B20,Stats!A1:A20)", "=STEYX(Stats!B1:B20,Stats!A1:A20)",
        "=FORECAST(30,Stats!B1:B20,Stats!A1:A20)", "=FORECAST.LINEAR(30,Stats!B1:B20,Stats!A1:A20)",
        "=COVAR(Stats!A1:A20,Stats!B1:B20)", "=COVARIANCE.P(Stats!A1:A20,Stats!B1:B20)",
        "=COVARIANCE.S(Stats!A1:A20,Stats!B1:B20)", "=PERCENTILE(Stats!A1:A20,0.3)", "=PERCENTILE.INC(Stats!A1:A20,0.3)",
        "=PERCENTILE.EXC(Stats!A1:A20,0.3)", "=PERCENTILE.EXC(Stats!A1:A20,0.01)", "=QUARTILE(Stats!A1:A20,1)",
        "=QUARTILE.INC(Stats!A1:A20,3)", "=QUARTILE.EXC(Stats!A1:A20,1)", "=PERCENTRANK(Stats!A1:A20,7)",
        "=PERCENTRANK.INC(Stats!A1:A20,7.5)", "=PERCENTRANK.EXC(Stats!A1:A20,7.5)", "=PERCENTRANK(Stats!A1:A20,7.5,5)",
        "=TRIMMEAN(Stats!A1:A20,0.2)", "=KURT(Stats!A1:A20)", "=SKEW(Stats!A1:A20)", "=SKEW.P(Stats!A1:A20)",
        "=INDEX(FREQUENCY(Stats!A1:A20,{5,10,15}),2)", "=SUM(FREQUENCY(Stats!A1:A20,{5,10,15}))",
        "=INDEX(MODE.MULT(1,2,2,3,3),2)", "=PROB({0,1,2,3},{0.2,0.3,0.1,0.4},2)", "=PROB({0,1,2,3},{0.2,0.3,0.1,0.4},1,3)",
        "=INDEX(LINEST(Stats!B1:B20,Stats!A1:A20),1)", "=INDEX(LINEST(Stats!B1:B20,Stats!A1:A20,TRUE,TRUE),3,1)",
        "=INDEX(TREND(Stats!B1:B20,Stats!A1:A20,{21;22}),2)", "=INDEX(GROWTH(Stats!C1:C10,Stats!A1:A10,{11}),1)",
        "=INDEX(LOGEST(Stats!C1:C10,Stats!A1:A10),1)", "=RANK.EQ(Stats!A5,Stats!A1:A20)", "=RANK.AVG(Stats!A5,Stats!A1:A20)",
        "=AVEDEV(Stats!A1:A20)", "=DEVSQ(Stats!A1:A20)", "=MEDIAN(Stats!A1:A20)", "=MODE.SNGL(Stats!A1:A20)",
        "=STDEV.S(Stats!A1:A20)", "=VAR.P(Stats!A1:A20)", "=GEOMEAN(Stats!A1:A10)", "=HARMEAN(Stats!A1:A10)",
    ],
    "Engineering": [
        "=DEC2BIN(9,4)", "=DEC2BIN(-100)", "=DEC2BIN(512)", "=DEC2OCT(58,3)", "=DEC2OCT(-100)", "=DEC2HEX(100,4)",
        "=DEC2HEX(-54)", "=BIN2DEC(1100100)", "=BIN2DEC(1111111111)", "=BIN2OCT(1001,3)", "=BIN2HEX(11111011,4)",
        "=OCT2DEC(54)", "=OCT2DEC(7777777533)", "=OCT2BIN(3,3)", "=OCT2HEX(100,4)", "=HEX2DEC(\"A5\")",
        "=HEX2DEC(\"FFFFFFFF5B\")", "=HEX2BIN(\"F\",8)", "=HEX2OCT(\"F\",3)", "=BITAND(1,5)", "=BITOR(23,10)",
        "=BITXOR(5,3)", "=BITLSHIFT(4,2)", "=BITRSHIFT(13,2)", "=DELTA(5,4)", "=DELTA(5,5)", "=GESTEP(5,4)",
        "=GESTEP(-4,-5)", "=BASE(7,2)", "=BASE(100,16,4)", "=DECIMAL(\"FF\",16)", "=DECIMAL(\"111\",2)",
        "=ROMAN(499)", "=ROMAN(499,1)", "=ROMAN(499,4)", "=ARABIC(\"LVII\")", "=ARABIC(\"mcmxii\")",
        "=CONVERT(1,\"lbm\",\"kg\")", "=CONVERT(68,\"F\",\"C\")", "=CONVERT(2.5,\"ft\",\"sec\")",
        "=CONVERT(6,\"tsp\",\"tbs\")", "=CONVERT(1,\"mi\",\"km\")", "=CONVERT(100,\"kW\",\"W\")",
        "=CONVERT(1,\"hr\",\"mn\")", "=CONVERT(1,\"in\",\"cm\")", "=CONVERT(1,\"gal\",\"l\")", "=CONVERT(1,\"atm\",\"Pa\")",
        "=COMPLEX(3,4)", "=COMPLEX(3,4,\"j\")", "=COMPLEX(0,1)", "=COMPLEX(0,-1)", "=IMABS(\"5+12i\")",
        "=IMREAL(\"6-9i\")", "=IMAGINARY(\"3+4i\")", "=IMARGUMENT(\"3+4i\")", "=IMCONJUGATE(\"3+4i\")",
        "=IMSUM(\"3+4i\",\"5-3i\")", "=IMSUB(\"13+4i\",\"5+3i\")", "=IMPRODUCT(\"3+4i\",\"5-3i\")",
        "=IMDIV(\"-238+240i\",\"10+24i\")", "=IMPOWER(\"2+3i\",3)", "=IMSQRT(\"1+i\")", "=IMEXP(\"1+i\")",
        "=IMLN(\"3+4i\")", "=IMLOG10(\"3+4i\")", "=IMLOG2(\"3+4i\")", "=IMSIN(\"4+3i\")", "=IMCOS(\"1+i\")",
        "=IMTAN(\"4+3i\")", "=BESSELJ(1.9,2)", "=BESSELY(2.5,1)", "=BESSELI(1.5,1)", "=BESSELK(1.5,1)",
        "=ERF(0,1)", "=ERFC(-1)",
    ],
    "Matrix": [
        "=MDETERM({1,3,8,5;1,3,6,1;1,1,1,0;7,3,10,2})", "=MDETERM({3,6;1,1})", "=INDEX(MINVERSE({4,-1;2,0}),1,2)",
        "=INDEX(MINVERSE({1,2,1;3,4,-1;0,2,0}),2,3)", "=INDEX(MMULT({1,3;7,2},{2,0;0,2}),2,1)", "=SUM(MUNIT(3))",
        "=INDEX(TRANSPOSE({1,2,3}),2)", "=SUMPRODUCT(MMULT({1,2},{3;4}))", "=MDETERM({1,2;2,4})",
        "=MINVERSE({1,2;2,4})", "=INDEX(MMULT(Stats!A1:B2,Stats!A1:B2),2,2)",
    ],
    "Database": [
        '=DSUM(Db!A1:E11,"Profit",Db!G1:H2)', '=DSUM(Db!A1:E11,5,Db!G1:H2)', '=DAVERAGE(Db!A1:E11,"Yield",Db!G1:G2)',
        '=DCOUNT(Db!A1:E11,"Age",Db!G1:H2)', '=DCOUNTA(Db!A1:E11,"Profit",Db!G1:H2)', '=DGET(Db!A1:E11,"Yield",Db!J1:J2)',
        '=DGET(Db!A1:E11,"Yield",Db!G1:G2)', '=DMAX(Db!A1:E11,"Profit",Db!G1:G3)', '=DMIN(Db!A1:E11,"Profit",Db!G1:H2)',
        '=DPRODUCT(Db!A1:E11,"Yield",Db!G1:G2)', '=DSTDEV(Db!A1:E11,"Yield",Db!G1:G3)',
        '=DSTDEVP(Db!A1:E11,"Yield",Db!G1:G3)', '=DVAR(Db!A1:E11,"Yield",Db!G1:G3)', '=DVARP(Db!A1:E11,"Yield",Db!G1:G3)',
        '=DSUM(Db!A1:E11,"Profit",Db!L1:M2)', '=DCOUNT(Db!A1:E11,,Db!L1:L2)', '=DSUM(Db!A1:E11,"Nothing",Db!G1:G2)',
        '=DAVERAGE(Db!A1:E11,3,Db!O1:O2)',
    ],
    "Subtotals": [
        "=SUBTOTAL(9,Stats!A1:A20)", "=SUBTOTAL(109,Stats!A1:A20)", "=SUBTOTAL(1,Stats!A1:A20)",
        "=SUBTOTAL(101,Stats!A1:A20)", "=SUBTOTAL(2,Stats!A1:A20)", "=SUBTOTAL(3,Stats!A1:A20)",
        "=SUBTOTAL(4,Stats!A1:A20)", "=SUBTOTAL(5,Stats!A1:A20)", "=SUBTOTAL(6,Stats!A1:A5)",
        "=SUBTOTAL(7,Stats!A1:A20)", "=SUBTOTAL(8,Stats!A1:A20)", "=SUBTOTAL(10,Stats!A1:A20)",
        "=SUBTOTAL(11,Stats!A1:A20)", "=SUBTOTAL(104,Stats!A1:A20)", "=SUBTOTAL(9,Stats!A1:A20,Stats!B1:B3)",
        "=SUBTOTAL(9,Stats!E1:E4)", "=SUBTOTAL(12,Stats!A1:A20)", "=AGGREGATE(9,5,Stats!A1:A20)",
        "=AGGREGATE(9,0,Stats!A1:A20)", "=AGGREGATE(14,6,Stats!A1:A20,2)", "=AGGREGATE(4,0,Stats!A1:A20)",
        "=AGGREGATE(12,1,Stats!A1:A20)", "=AGGREGATE(9,6,Stats!F1:F5)", "=AGGREGATE(9,4,Stats!F1:F5)",
        "=AGGREGATE(15,6,Stats!A1:A20,3)", "=AGGREGATE(16,4,Stats!A1:A20,0.9)", "=AGGREGATE(19,5,Stats!A1:A20,1)",
        "=AGGREGATE(9,3,Stats!E1:E4)", "=AGGREGATE(9,1,Stats!E1:E4)",
    ],
    "Arrays": [
        "=INDEX(SORT({3,1,2},1,1,TRUE),1,2)", "=INDEX(SORT(Stats!A1:A20),3)", "=INDEX(SORT(Stats!A1:A20,1,-1),1)",
        "=INDEX(SORTBY(Db!A2:A11,Db!B2:B11,-1),1)", "=ROWS(UNIQUE(Db!A2:A11))", "=INDEX(UNIQUE(Db!A2:A11),2)",
        '=SUM(FILTER(Db!B2:B11,Db!A2:A11="Apple"))', "=ROWS(FILTER(Db!B2:B11,Db!B2:B11>10))",
        '=IFERROR(FILTER(Db!B2:B11,Db!B2:B11>1000),"none")', "=INDEX(SEQUENCE(3,2,10,5),3,2)", "=SUM(SEQUENCE(10))",
        "=ROWS(SEQUENCE(5))", "=INDEX(TAKE(Stats!A1:A20,3),3)", "=INDEX(DROP(Stats!A1:A20,2),1)",
        "=INDEX(CHOOSEROWS(Stats!A1:A20,4,2),2)", "=INDEX(CHOOSECOLS(Stats!A1:C2,3,1),1,2)", "=INDEX(VSTACK({1,2},{3,4}),2,1)",
        "=INDEX(HSTACK({1;2},{3;4}),2,2)", "=ROWS(TOCOL({1,2;3,4}))", "=INDEX(TOROW({1,2;3,4}),3)",
        "=INDEX(WRAPROWS({1,2,3,4,5},2),3,1)", "=INDEX(WRAPCOLS({1,2,3,4,5},2),1,3)", "=INDEX(EXPAND({1,2},2,3,0),2,3)",
        '=INDEX(TEXTSPLIT("a,b,c",","),2)', '=TEXTBEFORE("a-b-c","-")', '=TEXTBEFORE("a-b-c","-",2)',
        '=TEXTAFTER("a-b-c","-",-1)', '=TEXTAFTER("abc","x")', '=TEXTAFTER("abc","x",,,,"none")', "=LET(x,2,y,3,x*y)",
        "=LET(x,SEQUENCE(3),SUM(x))", "=LAMBDA(x,x*2)(4)", "=SUM(MAP({1,2,3},LAMBDA(x,x*x)))",
        "=REDUCE(0,{1,2,3},LAMBDA(a,b,a+b))", "=INDEX(SCAN(0,{1,2,3},LAMBDA(a,b,a+b)),3)",
        "=INDEX(BYROW({1,2;3,4},LAMBDA(r,SUM(r))),2)", "=INDEX(BYCOL({1,2;3,4},LAMBDA(c,SUM(c))),2)",
        "=INDEX(MAKEARRAY(2,3,LAMBDA(r,c,r*c)),2,3)", "=ARRAYTOTEXT({1,2;3,4})", '=ARRAYTOTEXT({1,"a"},1)',
        "=INDEX(RANDARRAY(2,2,5,5),1,1)", "=XMATCH(3,{1,2,3})", "=ROWS(Stats!A1:A20#)", "=SUM(Db!B2:B11*(Db!A2:A11=\"Apple\"))",
    ],
    "Misc": [
        '=CELL("row",Stats!B5)', '=CELL("address",Stats!B5)', '=CELL("col",Stats!B5)', '=CELL("contents",Stats!A1)',
        '=CELL("type",Stats!A1)', '=CELL("type",Stats!D1)', '=HYPERLINK("http://example.com","text")', "=SHEET()",
        "=SHEETS()", "=AREAS((Stats!A1:A2,Stats!B1))", "=N(DATE(2020,1,1))", '=FIXED(1234567.891,-3)',
        "=ISOMITTED(1)", "=SUM(Stats!A1:A20)", "=COUNT(Stats!A:A)", "=MAX(Stats!A:A)", '=COUNTIF(Db!A:A,"Apple")',
        '=SUMIF(Db!A:A,"Pear",Db!E:E)', "=SUMPRODUCT((Db!A2:A11=\"Apple\")*(Db!B2:B11>10),Db!E2:E11)",
    ],
}  # fmt: skip


#: The fifth batch: the functions added after the fourth, on grids, and a
#: second look, over ten fresh samples, at formulas the fourth pinned from
#: one sample each.
BESSEL_X = (0.1, 0.5, 1, 1.9, 2.5, 3.75, 5, 7.9, 8, 8.1, 12, 20, 35)
CONVERSIONS = [
    ("1", "ft", "m"), ("1", "m", "ft"), ("100", "C", "F"), ("0", "C", "K"), ("1", "Rank", "K"), ("1", "Reau", "C"),
    ("98.6", "F", "C"), ("300", "K", "F"), ("1", "ozm", "g"), ("1", "ton", "kg"), ("1", "stone", "lbm"),
    ("1", "grain", "g"), ("1", "sg", "kg"), ("1", "u", "g"), ("1", "lbf", "N"), ("1", "dyn", "N"), ("1", "pond", "N"),
    ("1", "HP", "W"), ("1", "PS", "W"), ("1", "BTU", "J"), ("1", "eV", "J"), ("1", "cal", "J"), ("1", "c", "J"),
    ("1", "Wh", "J"), ("1", "flb", "J"), ("1", "HPh", "J"), ("1", "atm", "mmHg"), ("1", "psi", "Pa"),
    ("1", "Torr", "Pa"), ("1", "yr", "day"), ("1", "day", "sec"), ("90", "mn", "hr"), ("1", "mph", "m/s"),
    ("1", "kn", "m/s"), ("1", "admkn", "kn"), ("1", "ly", "m"), ("1", "pc", "ly"), ("1", "Nmi", "km"),
    ("1", "mi2", "km2"), ("1", "ha", "m2"), ("1", "us_acre", "m2"), ("1", "uk_acre", "ha"), ("1", "m3", "l"),
    ("1", "ft3", "l"), ("1", "barrel", "gal"), ("1", "cup", "tsp"), ("1", "uk_gal", "gal"), ("1", "uk_pt", "pt"),
    ("1", "bushel", "l"), ("1", "byte", "bit"), ("1", "kibyte", "byte"), ("1", "Mibyte", "kbyte"), ("1", "kg", "lbm"),
    ("1", "cm", "in"), ("1", "mm", "ang"), ("1", "dam", "m"), ("1", "hm", "km"), ("1", "T", "ga"), ("1", "pica", "in"),
    ("1", "Pica", "mm"), ("1", "survey_mi", "mi"), ("1", "ell", "m"), ("1", "Morgen", "m2"), ("1", "GRT", "m3"),
]  # fmt: skip


def _fifth() -> dict[str, list[str]]:
    reals = len(REALS)
    divisions = len(DIVISIONS)
    groups: dict[str, list[str]] = {}
    for name, template in (
        ("Cot", "=COT(Reals!A{row})"), ("Csc", "=CSC(Reals!A{row})"), ("Sec", "=SEC(Reals!A{row})"),
        ("Coth", "=COTH(Reals!A{row})"), ("Csch", "=CSCH(Reals!A{row})"), ("Sech", "=SECH(Reals!A{row})"),
        ("Acot", "=ACOT(Reals!A{row})"), ("Acoth", "=ACOTH(ABS(Reals!A{row})+1)"),
    ):
        groups[name] = [template.format(row=row) for row in range(1, reals + 1)]
    groups["EcmaCeiling"] = [f"=ECMA.CEILING(Divisions!A{row},Divisions!B{row})" for row in range(1, divisions + 1)]
    start, end = "DATE(2024,1,1)", "DATE(2024,3,31)"
    workdays = [f"=NETWORKDAYS.INTL({start},{end},{code})" for code in (*range(1, 8), *range(11, 18))]
    workdays += [f'=NETWORKDAYS.INTL({start},{end},"{mask}")' for mask in ("0000011", "1000001", "0101010", "1111111", "000001")]
    workdays += [f"=WORKDAY.INTL({start},30,{code})" for code in (1, 7, 11, 17)]
    workdays += [
        f'=WORKDAY.INTL({start},30,"0000110")', f"=WORKDAY.INTL({start},-30,1)",
        f"=WORKDAY.INTL({start},10,1,{{45306,45307}})", f"=NETWORKDAYS.INTL({end},{start},1)",
        f"=NETWORKDAYS.INTL({start},{end},99)", f"=NETWORKDAYS.INTL({start},{end},1,DATE(2024,1,15))",
    ]
    groups["Workdays"] = workdays
    groups["TextBytes"] = [
        '=LEFTB("hello",2)', '=RIGHTB("hello",3)', '=MIDB("hello",2,3)', '=LENB("hello")', '=LENB("\u65e5\u672c")',
        '=FINDB("l","hello")', '=SEARCHB("L","hello")', '=REPLACEB("hello",2,3,"EY")', "=USDOLLAR(1234.567,2)",
        "=USDOLLAR(-5)", '=JIS("abc")', '=ENCODEURL("a b&c/d?e=f")', '=ENCODEURL("na\u00efve \u2713")',
        "=ENCODEURL(\"~-_.!*'()\")", '=ENCODEURL("100%")',
    ]  # fmt: skip
    groups["Regex"] = [
        '=REGEXTEST("abc123","\\d+")', '=REGEXTEST("ABC","abc")', '=REGEXTEST("ABC","abc",1)',
        '=REGEXEXTRACT("abc123def456","\\d+")', '=INDEX(REGEXEXTRACT("abc123def456","\\d+",1),2)',
        '=INDEX(REGEXEXTRACT("John Smith","(\\w+) (\\w+)",2),2)', '=REGEXEXTRACT("abc","\\d")',
        '=REGEXREPLACE("abc123","\\d","#")', '=REGEXREPLACE("a1b2c3","\\d","#",2)', '=REGEXREPLACE("a1b2c3","\\d","#",-1)',
        '=REGEXREPLACE("John Smith","(\\w+) (\\w+)","$2, $1")', '=REGEXTEST("x","[")',
    ]  # fmt: skip
    groups["Ranges"] = [
        "=PERCENTOF(Stats!A1:A5,Stats!A1:A20)", "=ROWS(TRIMRANGE(Stats!A1:A40))", "=COLUMNS(TRIMRANGE(Stats!A1:Z3))",
        "=ROWS(TRIMRANGE(Stats!A1:A40,2))", "=ROWS(TRIMRANGE(Stats!A1:A40,0))", "=ROWS(TRIMRANGE(Stats!Z1:Z40))",
        "=SUM(TRIMRANGE(Stats!A1:A40))", "=PERCENTOF(0,0)",
    ]  # fmt: skip
    for function in ("BESSELJ", "BESSELY", "BESSELI", "BESSELK"):
        groups[function.title()] = [f"={function}({x},{n})" for x in BESSEL_X for n in range(5)]
    groups["Convert"] = [f'=CONVERT({value},"{source}","{target}")' for value, source, target in CONVERSIONS]
    columns = "ABCDEFGHIJ"
    spread: list[str] = []
    for index, column in enumerate(columns):
        data = f"Samples!{column}2:{column}13"
        other = f"Samples!{columns[(index + 1) % 10]}2:{columns[(index + 1) % 10]}13"
        spread += [
            f"=VAR.P({data})", f"=STDEV.P({data})", f"=VAR.S({data})", f"=STDEV.S({data})", f"=DEVSQ({data})",
            f"=AVEDEV({data})", f"=GEOMEAN({data})", f"=HARMEAN({data})", f"=SKEW({data})", f"=SKEW.P({data})",
            f"=KURT({data})", f"=NPV(0.07,{data})", f"=COVARIANCE.P({data},{other})", f"=COVARIANCE.S({data},{other})",
            f"=CORREL({data},{other})", f"=SLOPE({data},{other})", f"=INTERCEPT({data},{other})",
            f"=RSQ({data},{other})", f"=STEYX({data},{other})", f"=INDEX(LINEST({data},{other}),1)",
            f'=DVAR(Samples!A1:J13,"V{index + 1}",Samples!L1:L2)', f'=DVARP(Samples!A1:J13,"V{index + 1}",Samples!L1:L2)',
            f'=DSTDEV(Samples!A1:J13,"V{index + 1}",Samples!L1:L2)',
            f'=DSTDEVP(Samples!A1:J13,"V{index + 1}",Samples!L1:L2)',
        ]  # fmt: skip
    groups["Spread"] = spread
    groups["NormGrid"] = [f"=NORM.S.DIST({-9 + 0.36 * step:.2f},TRUE)" for step in range(51)]
    groups["NormInvGrid"] = [
        f"=NORM.S.INV({p})"
        for p in ("1E-300", "1E-100", "1E-20", "1E-10", "1E-5", "0.001", "0.01", "0.02425", "0.05", "0.1", "0.2", "0.3",
                  "0.4", "0.49", "0.51", "0.6", "0.7", "0.8", "0.9", "0.95", "0.975", "0.99", "0.999", "0.99999",
                  "0.9999999")
    ]  # fmt: skip
    groups["TGrid"] = [f"=T.DIST({x},{df},TRUE)" for x in (-5, -2, -1, -0.5, 0.3, 1, 2, 5) for df in (1, 2, 5, 30)]
    groups["ChiGrid"] = [f"=CHISQ.DIST({x},{k},TRUE)" for x in (0.1, 1, 5, 20) for k in (1, 2, 5, 10, 50)]
    groups["GammaGrid"] = [f"=GAMMA.DIST({x},{a},2,TRUE)" for x in (0.5, 3, 9, 30) for a in (0.5, 2, 7)]
    groups["BinomGrid"] = [f"=BINOM.DIST({k},20,{p},FALSE)" for k in (0, 3, 10, 17, 20) for p in (0.1, 0.5, 0.73)]
    groups["PoissonGrid"] = [f"=POISSON.DIST({k},{mean},{flag})" for k in (0, 4, 15) for mean in (0.5, 6) for flag in ("TRUE", "FALSE")]
    groups["WeibullGrid"] = [f"=WEIBULL.DIST({x},{a},{b},FALSE)" for x in (0.5, 2, 9) for a in (0.7, 1.5, 3) for b in (1, 4)]
    groups["GammaFunction"] = [
        f"=GAMMA({x})" for x in (0.1, 0.5, 1.5, 3, 4.5, 10.3, 25, 70.5, 120, 170.6, -0.5, -1.5, -7.25, -33.5)
    ] + [f"=GAMMALN({x})" for x in (0.1, 0.5, 1.5, 3, 4.5, 10.3, 25, 70.5, 1000, 1e10)]
    return groups  # fmt: skip


def _more() -> dict[str, list[str]]:
    powers = len(POWERS)
    reals = len(REALS)
    roundings = len(ROUNDINGS)
    wholes = len(WHOLES)
    divisions = len(DIVISIONS)
    times = len(TIMES)
    pairs = len(PAIRS)
    dated = len(DATE_PAIRS)
    rows = len(CRITERIA_CELLS)
    groups: dict[str, list[str]] = {
        "PowerData": [f"=Powers!A{row}^Powers!B{row}" for row in range(1, powers + 1)],
        "PowerFunction": [f"=POWER(Powers!A{row},Powers!B{row})" for row in range(1, 61)],
        "DifferenceReversed": [f"=Pairs!A{row}-Pairs!B{row}" for row in range(1, pairs + 1)],
        "SumNested": [f"=1*SUM(Pairs!B{row},-Pairs!A{row})" for row in range(1, pairs + 1)],
        "TiesText": [f'=Ties!A{row}&""' for row in range(1, len(TIES) + 1)],
    }
    for name, template in (
        ("Exp", "=EXP(Reals!A{row})"), ("Ln", "=LN(ABS(Reals!A{row}))"), ("Sin", "=SIN(Reals!A{row})"),
        ("Cos", "=COS(Reals!A{row})"), ("Tan", "=TAN(Reals!A{row})"), ("Atan", "=ATAN(Reals!A{row})"),
        ("Sinh", "=SINH(Reals!A{row})"), ("Cosh", "=COSH(Reals!A{row})"), ("Tanh", "=TANH(Reals!A{row})"),
        ("Sqrt", "=SQRT(ABS(Reals!A{row}))"), ("Asin", "=ASIN(Reals!A{row}/10)"), ("Acos", "=ACOS(Reals!A{row}/10)"),
        ("Log10", "=LOG10(ABS(Reals!A{row}))"), ("Log2", "=LOG(ABS(Reals!A{row}),2)"),
        ("Degrees", "=DEGREES(Reals!A{row})"), ("Radians", "=RADIANS(Reals!A{row})"),
        ("Atan2", "=ATAN2(Reals!A{row},Reals!A{next})"),
    ):
        groups[name] = [template.format(row=row, next=row % reals + 1) for row in range(1, reals + 1)]
    for name, function in (("Round", "ROUND"), ("RoundUp", "ROUNDUP"), ("RoundDown", "ROUNDDOWN"), ("Trunc", "TRUNC")):
        groups[name] = [f"={function}(Rounding!A{row},Rounding!B{row})" for row in range(1, roundings + 1)]
    groups["Int"] = [f"=INT(Wholes!A{row})" for row in range(1, wholes + 1)]
    groups["TruncWhole"] = [f"=TRUNC(Wholes!A{row})" for row in range(1, wholes + 1)]
    for name, function in (
        ("Mod", "MOD"), ("Quotient", "QUOTIENT"), ("Ceiling", "CEILING"), ("Floor", "FLOOR"),
        ("MRound", "MROUND"), ("CeilingMath", "CEILING.MATH"), ("FloorMath", "FLOOR.MATH"),
    ):
        groups[name] = [f"={function}(Divisions!A{row},Divisions!B{row})" for row in range(1, divisions + 1)]
    for name in ("HOUR", "MINUTE", "SECOND"):
        groups[name.title()] = [f"={name}(Times!A{row})" for row in range(1, times + 1)]
    groups["Days360"] = [f"=DAYS360(DatePairs!A{row},DatePairs!B{row})" for row in range(1, dated + 1)]
    groups["Days360Europe"] = [f"=DAYS360(DatePairs!A{row},DatePairs!B{row},TRUE)" for row in range(1, dated + 1)]
    for basis in range(5):
        groups[f"YearFrac{basis}"] = [f"=YEARFRAC(DatePairs!A{row},DatePairs!B{row},{basis})" for row in range(1, dated + 1)]
    groups["CountIf"] = [f"=COUNTIF(Crit!A1:A{rows},{test})" for test in CRITERIA]
    groups["SumIf"] = [f"=SUMIF(Crit!A1:A{rows},{test},Crit!B1:B{rows})" for test in CRITERIA]
    groups.update(EXTRA_GROUPS)
    groups.update(FAMILIES)
    groups.update(_fifth())
    return groups  # fmt: skip


#: Formulas written into the file by the library rather than typed into
#: Excel, to see what Excel makes of a literal with more digits than it
#: keeps when it reads one from a file. Not ``1E-320`` or ``1.8E+308``: a
#: formula with a number Excel cannot hold makes it refuse the whole file,
#: though a cell's value that small opens.
FILE_LITERALS = [
    "1.23456789012345678", "2.9999999999999996", "0.30000000000000004", "1234567890123456789",
    "1234567890123456", "1234567890123459", "0.1+0.2",
]  # fmt: skip

_BUILD = r"""
Public Function Build(ByVal Source As String, ByVal Target As String, ByVal Spec As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim groups As Variant
    Dim parts As Variant
    Dim g As Long
    Dim i As Long
    Dim out As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Source)
    Application.Calculation = xlCalculationManual
    wb.Worksheets("About").Range("A1").Value = Application.Version & " " & Application.Build
    wb.Worksheets("About").Range("A2").Value = Format(Date, "yyyy-mm-dd")
    groups = Split(Spec, ChrW(&HE002))
    For g = 0 To UBound(groups)
        parts = Split(groups(g), ChrW(&HE001))
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = parts(0)
        For i = 1 To UBound(parts)
            On Error Resume Next
            ws.Cells(i, 1).Formula = parts(i)
            If Err.Number <> 0 Then out = out & parts(0) & "!" & CStr(i) & " " & parts(i) & vbLf
            Err.Clear
            On Error GoTo 0
        Next i
    Next g
    Application.CalculateFull
    wb.SaveAs Filename:=Target, FileFormat:=51
    wb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Build = out
End Function
"""


def build_inputs(target: Path) -> None:
    """The input sheets, written by this library into Excel's empty
    workbook, so every number is exactly the double meant."""
    book = Workbook.open(FIXTURES / "empty.xlsx")
    book.rename_sheet(book.sheets[0].name, "Data")
    data = book["Data"]
    for row in range(1, 11):
        data[f"A{row}"] = row
    data["B1"] = "abc"
    data["B2"] = "10"
    data["B3"] = True
    data["B5"].formula = "NA()"
    data["B6"] = 2.5
    data["B7"] = -3
    data["B8"] = "Apple"
    data["B9"] = "apple"
    data["B10"] = dt.date(2020, 1, 15)
    for row, (key, value) in enumerate(zip("xyzxw", (10, 20, 30, 40, 50), strict=True), start=1):
        data[f"C{row}"] = key
        data[f"D{row}"] = value
    columns = (("Name", "pen", "book", "ink"), ("Qty", 3, 1, 5), ("Price", 1.5, 12, 4.25))
    for letter, values in zip("FGH", columns, strict=True):
        for row, value in enumerate(values, start=1):
            data[f"{letter}{row}"] = value
    book.add_sheet("About")
    numbers = book.add_sheet("Numbers")
    for row, number in enumerate(NUMBERS, start=1):
        numbers[f"A{row}"] = number
    pairs = book.add_sheet("Pairs")
    for row, (first, second) in enumerate(PAIRS, start=1):
        pairs[f"A{row}"] = first
        pairs[f"B{row}"] = second
    literals = book.add_sheet("FileLiterals")
    for row, formula in enumerate(FILE_LITERALS, start=1):
        literals[f"A{row}"].formula = formula
    for name, rows in (
        ("Powers", POWERS),
        ("Rounding", ROUNDINGS),
        ("Divisions", DIVISIONS),
        ("DatePairs", DATE_PAIRS),
    ):
        sheet = book.add_sheet(name)
        for row, (first, second) in enumerate(rows, start=1):
            sheet[f"A{row}"] = first
            sheet[f"B{row}"] = second
    for name, values in (("Reals", REALS), ("Wholes", WHOLES), ("Times", TIMES), ("Ties", TIES)):
        sheet = book.add_sheet(name)
        for row, value in enumerate(values, start=1):
            sheet[f"A{row}"] = value
    criteria = book.add_sheet("Crit")
    for row, value in enumerate(CRITERIA_CELLS, start=1):
        if isinstance(value, str) and value.startswith("="):
            criteria[f"A{row}"].formula = value[1:]
        elif value is not None:
            assert isinstance(value, (str, float, bool, dt.date))
            criteria[f"A{row}"] = value
        criteria[f"B{row}"] = float(row)
    _statistics_inputs(book)
    _database_inputs(book)
    _samples_inputs(book)
    book.save(target)


def _statistics_inputs(book: Workbook) -> None:
    """``Stats``: x in A, y in B, a growing series in C, a small table in
    C:D for CHISQ.TEST, numbers and a nested SUBTOTAL in E, an error in F;
    rows 5 and 6 hidden for SUBTOTAL and AGGREGATE to skip."""
    chance = random.Random(2718)
    sheet = book.add_sheet("Stats")
    for row in range(1, 21):
        x = float(row) + round(chance.uniform(-0.4, 0.4), 3)
        sheet[f"A{row}"] = x
        sheet[f"B{row}"] = round(2.5 * x + 3 + chance.uniform(-2, 2), 3)
    for row in range(1, 11):
        sheet[f"C{row}"] = round(1.3**row * (1 + chance.uniform(-0.05, 0.05)), 4)
    for row, values in enumerate(((20.0, 30.0), (40.0, 50.0), (60.0, 70.0)), start=1):
        sheet[f"D{row}"] = values[0]
        sheet[f"E{row + 10}"] = values[1]
    sheet["E1"] = 5.0
    sheet["E2"] = 7.0
    sheet["E3"].formula = "SUBTOTAL(9,E1:E2)"
    sheet["E4"] = 11.0
    sheet["F1"] = 1.0
    sheet["F2"].formula = "1/0"
    sheet["F3"] = 3.0
    sheet["F4"] = "x"
    sheet["F5"] = 5.0
    sheet.set_row_hidden(5, True)
    sheet.set_row_hidden(6, True)


def _database_inputs(book: Workbook) -> None:
    """``Db``: Microsoft's orchard database in A1:E11, and criteria ranges
    beside it: G1:H3 (Apple over 10 tall, or any Pear), J1:J2 (one Pear
    of 12), L1:M2 (a formula criterion), O1:O2 (a wildcard)."""
    sheet = book.add_sheet("Db")
    rows = [
        ("Tree", "Height", "Age", "Yield", "Profit"), ("Apple", 18.0, 20.0, 14.0, 105.0),
        ("Pear", 12.0, 12.0, 10.0, 96.0), ("Cherry", 13.0, 14.0, 9.0, 105.0), ("Apple", 14.0, 15.0, 10.0, 75.0),
        ("Pear", 9.0, 8.0, 8.0, 76.8), ("Apple", 8.0, 9.0, 6.0, 45.0), ("Cherry", 10.0, 10.0, 7.0, 60.0),
        ("Apple", 11.0, 13.0, 9.0, 66.0), ("Pear", 7.0, 6.0, 5.0, 40.0), ("Plum", 15.0, 11.0, 12.0, 90.0),
    ]  # fmt: skip
    for row, values in enumerate(rows, start=1):
        for column, value in zip("ABCDE", values, strict=True):
            sheet[f"{column}{row}"] = value
    sheet["G1"] = "Tree"
    sheet["H1"] = "Height"
    sheet["G2"] = "Apple"
    sheet["H2"] = ">10"
    sheet["G3"] = "Pear"
    sheet["J1"] = "Height"
    sheet["J2"] = "=12"
    sheet["L1"] = "Rich"
    sheet["L2"].formula = "E2>80"
    sheet["M1"] = "Tree"
    sheet["M2"] = "<>Plum"
    sheet["O1"] = "Tree"
    sheet["O2"] = "P*"


def _samples_inputs(book: Workbook) -> None:
    """``Samples``: ten columns of twelve positive numbers of different
    sizes under the headings V1 to V10, and in L1:L2 a criterion every row
    meets, for the fifth batch's second look at the spread functions."""
    chance = random.Random(314159)
    sheet = book.add_sheet("Samples")
    columns = "ABCDEFGHIJ"
    for index, column in enumerate(columns):
        sheet[f"{column}1"] = f"V{index + 1}"
        scale = 10.0 ** chance.randint(-2, 4)
        for row in range(2, 14):
            sheet[f"{column}{row}"] = round(chance.uniform(0.05, 9.95) * scale, 6)
    sheet["L1"] = "V1"
    sheet["L2"] = ">0"


def main() -> int:
    try:
        from pyvbaharness import ExcelSession
        from pyvbaharness.session import HarnessConfig
    except ImportError:
        print('pyvbaharness is not installed; pip install -e ".[dev]" --group live', file=sys.stderr)
        return 2
    groups = {**GROUPS, **_generated(), **_more()}
    spec = RS.join(FS.join([name, *formulas]) for name, formulas in groups.items())
    wait = float(os.environ.get("LIVE_EXCEL_LOCK_WAIT", "0"))
    with tempfile.TemporaryDirectory() as scratch:
        source = Path(scratch) / "inputs.xlsx"
        build_inputs(source)
        with ExcelSession(HarnessConfig(lock_wait_s=wait)) as excel:
            excel.new_workbook()
            result = excel.run_vba(_BUILD, proc="Build", args=(str(source), str(FIXTURE), spec), timeout=900)
            if result.outcome != "passed":
                raise SystemExit(f"Excel refused the build: {result!r}")
    refused = str(result.value or "").strip()
    if refused:
        print("formulas Excel would not take:")
        print(refused)
    count = sum(len(formulas) for formulas in groups.values())
    print(f"wrote {FIXTURE.name}: {count} formulas in {len(groups)} groups")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
