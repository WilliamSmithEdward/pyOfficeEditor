"""The formula engine's parts: the parser, Excel's numeric rules, and the
calculation of a whole workbook.

What any single formula gives is held to Excel in
``test_excel_formula_corpus.py``; this file covers what the corpus cannot
show one formula at a time: the tree a formula parses to, a workbook's
formulas reading one another, and what calculating writes back.
"""

from __future__ import annotations

import datetime as dt
import json
import math
import os
from decimal import Decimal
from pathlib import Path

import pytest

from pyofficeeditor.excel import (
    CellError,
    CellValue,
    FilterColumn,
    FormulaSyntaxError,
    Workbook,
    Worksheet,
    column_letter,
    criteria,
)
from pyofficeeditor.excel._calc import parse, special
from pyofficeeditor.excel._calc.dates import parse_date_time
from pyofficeeditor.excel._calc.functions.distributions import normal_inverse
from pyofficeeditor.excel._calc.nodes import (
    AreaReference,
    ArrayLiteral,
    AxisReference,
    Binary,
    Call,
    CellReference,
    ErrorLiteral,
    Invoke,
    Missing,
    NameReference,
    Number,
    Paren,
    Postfix,
    Prefix,
    StructuredReference,
    Text,
    Unary,
    function_key,
    render,
)
from pyofficeeditor.excel._calc.numbers import compare_numbers, literal_value, near_zero, number_text
from pyofficeeditor.excel._calc.values import text_to_number
from pyofficeeditor.excel._reference import AxisRef, CellRef
from pyofficeeditor.exceptions import UnsupportedFormulaError

TODAY = dt.date(2026, 9, 22)
DATE_TEXTS = Path(__file__).parent / "fixtures" / "excel" / "date_texts.json"

# ----------------------------------------------------------------------
# Parsing
# ----------------------------------------------------------------------


def test_negation_binds_tighter_than_power() -> None:
    assert parse("-2^2") == Binary("^", Unary("-", Number(2.0, "2")), Number(2.0, "2"))


def test_power_groups_from_the_left() -> None:
    tree = parse("2^3^2")
    assert isinstance(tree, Binary) and isinstance(tree.left, Binary)
    assert render(tree) == "2^3^2"


def test_percent_binds_tighter_than_power() -> None:
    assert parse("2^50%") == Binary("^", Number(2.0, "2"), Postfix("%", Number(50.0, "50")))


def test_a_leading_equals_sign_is_dropped() -> None:
    assert parse("=1+2") == parse("1+2")


def test_a_function_name_that_looks_like_a_cell() -> None:
    call = parse("LOG10(100)")
    assert isinstance(call, Call) and call.name == "LOG10"
    assert parse("LOG10") == CellReference(CellRef.parse("LOG10"))


def test_whole_rows_and_columns() -> None:
    assert parse("1:3") == AxisReference(AxisRef(True, 1, 3))
    assert parse("$A:$C") == AxisReference(AxisRef(False, 1, 3, True, True))


def test_sheet_qualifiers() -> None:
    assert parse("Data!A1") == CellReference(CellRef(1, 1), Prefix("Data"))
    assert parse("'Q1 Data'!A1:B2") == AreaReference(CellRef(1, 1), CellRef(2, 2), Prefix("Q1 Data"))
    assert parse("Jan:Mar!A1") == CellReference(CellRef(1, 1), Prefix("Jan", "Mar"))
    assert parse("[1]Data!A1") == CellReference(CellRef(1, 1), Prefix("Data", None, "1"))
    assert parse("'It''s'!A1") == CellReference(CellRef(1, 1), Prefix("It's"))


def test_a_range_between_two_qualified_cells_on_one_sheet_is_an_area() -> None:
    assert parse("Data!A1:Data!B2") == AreaReference(CellRef(1, 1), CellRef(2, 2), Prefix("Data"))


def test_a_space_between_references_is_the_intersection() -> None:
    tree = parse("A1:C3 B2:D4")
    assert isinstance(tree, Binary) and tree.op == " "


def test_a_space_elsewhere_is_whitespace() -> None:
    assert parse("SUM( A1 , B1 )") == parse("SUM(A1,B1)")
    assert parse("A1 + B1") == parse("A1+B1")


def test_a_comma_is_a_union_only_in_parentheses_of_its_own() -> None:
    call = parse("SUM((A1,B1),C1)")
    assert isinstance(call, Call) and len(call.args) == 2
    union = call.args[0]
    assert isinstance(union, Paren) and isinstance(union.inner, Binary) and union.inner.op == ","


def test_an_argument_left_empty_is_missing() -> None:
    call = parse("IF(A1,,2)")
    assert isinstance(call, Call) and call.args[1] == Missing()


def test_array_constants() -> None:
    assert parse('{1,-2;"a",TRUE}') == ArrayLiteral(((1.0, -2.0), ("a", True)))
    assert parse("{#N/A}") == ArrayLiteral(((CellError("#N/A"),),))


def test_errors_and_the_spill_operator() -> None:
    assert parse("#DIV/0!") == ErrorLiteral("#DIV/0!")
    assert parse("A1#") == Postfix("#", CellReference(CellRef(1, 1)))


def test_structured_references() -> None:
    assert parse("Table1[Qty]") == StructuredReference("Table1", (), "Qty", "Qty")
    assert parse("Table1[[#Headers],[Qty]:[Price]]") == StructuredReference("Table1", ("#Headers",), "Qty", "Price")
    assert parse("Table1[[#This Row],[Qty]]") == parse("Table1[@Qty]")
    assert parse("Table1[#All]") == StructuredReference("Table1", ("#All",))
    assert parse("Table1[Price'[USD']]") == StructuredReference("Table1", (), "Price[USD]", "Price[USD]")


def test_names_and_functions_behind_the_file_prefixes() -> None:
    call = parse("_xlfn._xlws.SORT(A1:A3)")
    assert isinstance(call, Call) and call.function == "SORT"
    assert function_key("_xlfn.STDEV.S") == "STDEV.S"
    assert parse("_xlpm.x") == NameReference("_xlpm.x")


def test_strings_keep_doubled_quotes() -> None:
    assert parse('"say ""hi"""') == Text('say "hi"')


def test_a_lambda_called_where_it_is_made() -> None:
    tree = parse("_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(4)")
    assert isinstance(tree, Invoke)
    assert isinstance(tree.target, Call) and tree.target.function == "LAMBDA"
    assert tree.args == (Number(4.0, "4"),)
    # With a space the parenthesis is an intersection, not a call.
    assert isinstance(parse("SUM(A1:A3) (A2:B2)"), Binary)


@pytest.mark.parametrize(
    "formula",
    [
        "3-",
        "SUM(1,2",
        "(1+2",
        '"open',
        "{1,2;3}",
        "1+*2",
        "",
        "#NOTANERROR",
        "Table1[#Nothing]",
    ],
)
def test_text_excel_would_refuse(formula: str) -> None:
    with pytest.raises(FormulaSyntaxError):
        parse(formula)


@pytest.mark.parametrize(
    "formula",
    [
        "-2^2",
        "(1+2)*3",
        "2^-1",
        "5%%",
        "SUM(Data!A1:A10,'Q1 Data'!B2)",
        "IF(A1,,2)",
        "A1:C3 B2:D4",
        "SUM((A1,B1))",
        '{1,2;"a",TRUE}',
        "Table1[[#Headers],[Qty]:[Price]]",
        "-(1+2)%",
        "INDEX(A1:C3,2,2):C3",
        "_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(4)",
        "(_xlfn.LAMBDA(_xlpm.x,_xlpm.x))(1)(2)",
    ],
)
def test_rendering_parses_back_to_the_same_tree(formula: str) -> None:
    tree = parse(formula)
    assert parse(render(tree)) == tree


# ----------------------------------------------------------------------
# Excel's numbers
# ----------------------------------------------------------------------


@pytest.mark.parametrize(
    ("value", "text"),
    [
        (0.1 + 0.2, "0.3"),
        (1 / 3, "0.333333333333333"),
        (1e20, "1E+20"),
        (1.5e19, "15000000000000000000"),
        (1e-18, "0.000000000000000001"),
        (1e-19, "1E-19"),
        (1.2345678901234567e-5, "1.23456789012346E-05"),
        (-2.5, "-2.5"),
        (1234567890123456.0, "1234567890123460"),
        (1.7976931348623157e308, "1.7976931348623E+308"),
    ],
)
def test_a_number_as_text(value: float, text: str) -> None:
    assert number_text(value) == text


def test_numbers_are_equal_to_fifteen_digits() -> None:
    assert compare_numbers(0.1 + 0.2, 0.3) == 0
    below_by_four = 1 - 4 * 2**-53
    below_by_five = 1 - 5 * 2**-53
    assert compare_numbers(1.0, below_by_four) == 0
    assert compare_numbers(1.0, below_by_five) == 1


def test_a_final_sum_near_zero_is_zero() -> None:
    assert near_zero(0.1 + 0.2, -0.3, 0.1 + 0.2 - 0.3) == 0.0
    assert near_zero(1.0, -(1 + 8 * 2**-52), -8 * 2**-52) != 0.0


def test_a_literal_keeps_fifteen_digits_cut_off() -> None:
    assert literal_value("2.9999999999999996") == 2.99999999999999
    assert literal_value("1234567890123456789") == 1.23456789012345e18
    assert literal_value("1E-308") == 0.0


@pytest.mark.parametrize(
    ("text", "number"),
    [
        ("1,000", 1000.0),
        ("$1,000.50", 1000.5),
        ("(5)", -5.0),
        ("50%", 0.5),
        ("1 1/2", 1.5),
        ("1/2/2020", 43832.0),
        ("12:30", 0.5208333333333334),
        ("Jan 15, 2020", 43845.0),
        ("1/15", 46037.0),
        # Measured: a group after the first may be longer than three digits,
        # the first may be any length, and a fraction may follow them.
        ("1,0000", 10000.0),
        ("12,34567,890", 1234567890.0),
        ("0010,000", 10000.0),
        ("-1,000 1/2", -1000.5),
        ("1 32767/1", 32768.0),
    ],
)
def test_text_that_reads_as_a_number(text: str, number: float) -> None:
    assert text_to_number(text, TODAY) == number


@pytest.mark.parametrize(
    "text",
    [
        *("", "TRUE", "1,00", "(-5)", "5-", "12:30PM", "1e308", "\t1"),
        # Measured: a first group of zeros, a currency sign with a percent
        # sign, a fraction after a decimal point, and a fraction's part past
        # 32767 are not numbers.
        *("0,123", "1,000,00", "$5%", "%$5", "($5%)", "1.5 1/2", "1 1/32768", "1 32768/1"),
    ],
)
def test_text_that_is_not_a_number(text: str) -> None:
    assert text_to_number(text, TODAY) is None


def test_text_reads_as_a_date_or_a_time_as_excel_reads_it() -> None:
    # Strings measured in Excel by scripts/measure_date_texts.py, each with
    # VALUE of it and whether DATEVALUE and TIMEVALUE read it, which they do
    # as VALUE's days and the rest.
    record = json.loads(DATE_TEXTS.read_text(encoding="utf-8"))
    measured = dt.date.fromisoformat(record["measured"])
    wrong: list[str] = []
    for key, epoch_1904 in (("cases", False), ("cases_1904", True)):
        for text, value, read in record[key]:
            found = parse_date_time(text, measured, epoch_1904=epoch_1904)
            as_measured = found == value if read else found is None
            if text_to_number(text, measured, epoch_1904=epoch_1904) != value or not as_measured:
                wrong.append(text)
    assert len(record["cases"]) > 18000
    assert wrong == []


@pytest.mark.parametrize(
    ("formula", "value"),
    [
        # Measured: a number that is no day of the month is a year.
        ('DATEVALUE("1/2020")', 43831.0),
        ('DATEVALUE("2/29")', 47150.0),
        ('DATEVALUE("Jan/99")', 36161.0),
        # A name, a day and a year take a comma.
        ('ISERROR(DATEVALUE("Jan 1 99"))', True),
        # One separator may trail a time, a comma with a space after it.
        ('TIMEVALUE("12:30, ")', 0.5208333333333334),
        ('ISERROR(TIMEVALUE("12:30,"))', True),
        # Minutes and seconds, and a field past its range.
        ('VALUE("93:22.15")', 0.06483969907407407),
        ('VALUE("25:00")', 1.0416666666666667),
        ('ISERROR(VALUE("24:60"))', True),
        # A month's name as a field is the negative of its number, unsigned.
        ('VALUE("Jan:5")', 178956970.6284722),
        # Numbers after a date and a time are dropped.
        ('VALUE("1/1/2020 12:00 4")', 43831.5),
        ('VALUE("12:00 PM1/1/2020")', 43831.5),
    ],
)
def test_a_formula_reads_a_date_or_a_time_from_text(book: Workbook, formula: str, value: CellValue) -> None:
    assert book["Data"].evaluate(formula, today=dt.date(2026, 9, 26)) == value


# ----------------------------------------------------------------------
# Calculating a workbook
# ----------------------------------------------------------------------


@pytest.fixture
def book(live_empty_xlsx: Path) -> Workbook:
    workbook = Workbook.open(live_empty_xlsx)
    workbook.rename_sheet(workbook.sheets[0].name, "Data")
    return workbook


def test_calculate_writes_each_result_into_its_cell(book: Workbook) -> None:
    sheet = book["Data"]
    sheet["A1"] = 2
    sheet["A2"].formula = "A1*3"
    sheet["A3"].formula = 'A2&" apples"'
    sheet["A4"].formula = "A2>5"
    sheet["A5"].formula = "1/0"
    sheet["A6"].formula = 'MID("x",5,1)'
    report = book.calculate(today=TODAY)
    assert report.complete and report.calculated == 5
    reopened = Workbook.from_bytes(book.to_bytes())["Data"]
    assert reopened["A2"].value == 6
    assert reopened["A3"].value == "6 apples"
    assert reopened["A4"].value is True
    assert reopened["A5"].value == CellError("#DIV/0!")
    assert reopened["A6"].value == ""
    assert reopened["A2"].formula == "A1*3"


def test_a_long_chain_calculates_without_deep_recursion(book: Workbook) -> None:
    sheet = book["Data"]
    sheet["A1"] = 1
    count = 3000
    # Written bottom up, so the first formula calculated needs every other.
    for row in range(count, 1, -1):
        sheet[f"A{row}"].formula = f"A{row - 1}+1"
    report = book.calculate()
    assert report.complete
    assert sheet[f"A{count}"].value == count


def test_a_sum_over_a_column_of_formulas(book: Workbook) -> None:
    sheet = book["Data"]
    for row in range(1, 201):
        sheet[f"A{row}"] = row
        sheet[f"B{row}"].formula = f"A{row}*2"
    sheet["C1"].formula = "SUM(B:B)"
    book.calculate()
    assert sheet["C1"].value == 2 * sum(range(1, 201))


def test_a_circular_reference_keeps_its_cached_values(book: Workbook) -> None:
    sheet = book["Data"]
    sheet["A1"].formula = "B1+1"
    sheet["B1"].formula = "A1+1"
    sheet["C1"].formula = "5"
    report = book.calculate()
    assert sorted(report.circular) == ["Data!A1", "Data!B1"]
    assert sheet["C1"].value == 5
    assert book.values_changed


def test_a_cell_read_beside_a_circular_reference_is_not_on_it(book: Workbook) -> None:
    """A1 reads B1 and C1 at once, and C1 reads A1 back: the loop is A1 and
    C1, and B1, asked for beside C1, is calculated as usual."""
    sheet = book["Data"]
    sheet["A1"].formula = "SUM(B1:C1)"
    sheet["B1"].formula = "5*2"
    sheet["C1"].formula = "A1"
    report = book.calculate()
    assert sorted(report.circular) == ["Data!A1", "Data!C1"]
    assert sheet["B1"].value == 10


def test_an_unknown_function_is_a_name_error(book: Workbook) -> None:
    sheet = book["Data"]
    sheet["A1"].formula = "NOSUCHFUNCTION(1)"
    assert book.calculate().complete
    assert sheet["A1"].value == CellError("#NAME?")


def test_a_function_the_engine_lacks_keeps_the_cached_value(book: Workbook) -> None:
    sheet = book["Data"]
    sheet["A1"].formula = "CUBEVALUE(1)"
    sheet["A2"].formula = "A1+1"
    report = book.calculate()
    assert report.unsupported == {"Data!A1": "the function CUBEVALUE"}
    assert report.dependent == ["Data!A2"]
    assert sheet["A1"].value is None
    assert book.values_changed


def test_an_array_formula_fills_its_block(book: Workbook) -> None:
    sheet = book["Data"]
    for row, value in enumerate((1, 2, 3), start=1):
        sheet[f"A{row}"] = value
    element = sheet.cell_element(CellRef(1, 2))
    formula = element.child("f")
    assert formula is None
    sheet["B1"].formula = "A1:A3*10"
    written = sheet.cell_element(CellRef(1, 2)).child("f")
    assert written is not None
    written.set("t", "array")
    written.set("ref", "B1:B4")
    book.calculate()
    assert [sheet[f"B{row}"].value for row in range(1, 5)] == [10, 20, 30, CellError("#N/A")]


def test_an_ordinary_formula_takes_the_range_cell_in_its_row(book: Workbook) -> None:
    sheet = book["Data"]
    for row, value in enumerate((1, 2, 3), start=1):
        sheet[f"A{row}"] = value
    sheet["B2"].formula = "A1:A3*10"
    sheet["B5"].formula = "A1:A3*10"
    book.calculate()
    assert sheet["B2"].value == 20
    assert sheet["B5"].value == CellError("#VALUE!")


def test_defined_names_and_other_sheets(book: Workbook) -> None:
    data = book["Data"]
    other = book.add_sheet("Other")
    data["A1"] = 4
    other["A1"] = 6
    book.add_defined_name("Rate", "0.5")
    book.add_defined_name("Both", "Data!$A$1+Other!$A$1")
    data["B1"].formula = "Both*Rate"
    data["B2"].formula = "SUM(Data:Other!A1)"
    book.calculate()
    assert data["B1"].value == 5
    assert data["B2"].value == 10


def test_a_table_is_read_by_its_structured_references(book: Workbook) -> None:
    sheet = book["Data"]
    for column, header in zip("ABC", ("Item", "Qty", "Price"), strict=True):
        sheet[f"{column}1"] = header
    for row, (item, quantity, price) in enumerate((("pen", 3, 1.5), ("ink", 5, 4.0)), start=2):
        sheet[f"A{row}"] = item
        sheet[f"B{row}"] = quantity
        sheet[f"C{row}"] = price
    sheet.add_table("Stock", "A1:C3")
    sheet["E1"].formula = "SUMPRODUCT(Stock[Qty],Stock[Price])"
    sheet["E2"].formula = "ROWS(Stock[#All])"
    sheet["E3"].formula = "INDEX(Stock[[#Headers],[Price]],1)"
    book.calculate()
    assert sheet["E1"].value == 24.5
    assert sheet["E2"].value == 3
    assert sheet["E3"].value == "Price"


def test_a_table_name_alone_is_no_reference_in_a_file(book: Workbook) -> None:
    """Measured by the live gate: Excel writes ``Stock[]`` for a table's
    data, and reads ``Stock`` alone as a defined name there is none of."""
    sheet = book["Data"]
    sheet["A1"] = "Item"
    sheet["A2"] = "pen"
    sheet.add_table("Stock", "A1:A2")
    sheet["C1"].formula = "ROWS(Stock[])"
    sheet["C2"].formula = "ROWS(Stock)"
    book.calculate()
    assert sheet["C1"].value == 1
    assert sheet["C2"].value == CellError("#NAME?")


def test_a_literal_written_here_keeps_fifteen_digits(book: Workbook) -> None:
    """Excel cuts a literal to fifteen digits when it reads a file, so the
    engine does too: INT(2.9999999999999996) is INT(2.99999999999999)."""
    sheet = book["Data"]
    sheet["A1"].formula = "INT(2.9999999999999996)"
    book.calculate()
    assert sheet["A1"].value == 2


def test_references_made_at_run_time(book: Workbook) -> None:
    sheet = book["Data"]
    for row in range(1, 6):
        sheet[f"A{row}"] = row
    sheet["B1"].formula = 'SUM(INDIRECT("A2:A"&4))'
    sheet["B2"].formula = "SUM(OFFSET(A1,1,0,3,1))"
    sheet["B3"].formula = "SUM(INDEX(A1:A5,2):INDEX(A1:A5,4))"
    book.calculate()
    assert [sheet[f"B{row}"].value for row in range(1, 4)] == [9, 9, 9]


def test_evaluate_calculates_what_a_formula_reads(book: Workbook) -> None:
    sheet = book["Data"]
    sheet["A1"] = 3
    sheet["A2"].formula = "A1^2"
    assert sheet.evaluate("A2+1") == 10
    assert sheet.evaluate("ROW()", at="C7") == 7
    assert sheet.evaluate('"x"&A1') == "x3"
    assert sheet["A2"].value is None


def test_evaluate_refuses_what_it_cannot_calculate(book: Workbook) -> None:
    sheet = book["Data"]
    with pytest.raises(UnsupportedFormulaError):
        sheet.evaluate("CUBEVALUE(1)")
    with pytest.raises(FormulaSyntaxError):
        sheet.evaluate("1+")


def test_today_and_now_come_from_the_caller(book: Workbook) -> None:
    sheet = book["Data"]
    moment = dt.datetime(2024, 2, 29, 18, 0)
    assert sheet.evaluate("TODAY()", now=moment) == 45351
    now = sheet.evaluate("NOW()", now=moment)
    assert isinstance(now, float) and math.isclose(now, 45351.75)
    assert sheet.evaluate('DATEVALUE("3/1")', today=dt.date(2024, 1, 1)) == 45352


def test_a_complete_calculation_trusts_the_cache_again(book: Workbook) -> None:
    sheet = book["Data"]
    sheet["A1"] = 1
    sheet["A2"].formula = "A1+1"
    assert book.values_changed
    assert book.calculate().complete
    assert not book.values_changed


def test_let_and_lambda(book: Workbook) -> None:
    sheet = book["Data"]
    sheet["A1"].formula = "_xlfn.LET(_xlpm.x,2,_xlpm.y,_xlpm.x+1,_xlpm.x*_xlpm.y)"
    sheet["A2"].formula = "_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(4)"
    sheet["A3"].formula = "_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.n,_xlpm.n+1),_xlpm.f(_xlpm.f(1)))"
    sheet["A4"].formula = "_xlfn.LAMBDA(_xlpm.x,_xlpm.x)"
    sheet["A5"].formula = (
        "_xlfn.LAMBDA(_xlpm.a,_xlpm.b,IF(_xlfn.ISOMITTED(_xlpm.b),_xlpm.a,_xlpm.a+_xlpm.b))(5,)"
    )
    sheet["A6"].formula = "_xlfn.REDUCE(1,{1,2,3,4},_xlfn.LAMBDA(_xlpm.p,_xlpm.v,_xlpm.p*_xlpm.v))"
    # Measured: an argument left out of the call is #VALUE!; left empty,
    # as in (5,) above, it is omitted.
    sheet["A7"].formula = "_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a)(5)"
    book.calculate()
    assert [sheet[f"A{row}"].value for row in range(1, 8)] == [6, 8, 3, CellError("#CALC!"), 5, 24, CellError("#VALUE!")]


def test_a_lambda_in_a_defined_name_can_call_itself(book: Workbook) -> None:
    book.add_defined_name(
        "Factorial", "_xlfn.LAMBDA(_xlpm.n,IF(_xlpm.n<=1,1,_xlpm.n*Factorial(_xlpm.n-1)))"
    )
    sheet = book["Data"]
    sheet["A1"].formula = "Factorial(6)"
    assert book.calculate().complete
    assert sheet["A1"].value == 720


def test_the_spill_operator_reads_a_dynamic_array_block(book: Workbook) -> None:
    sheet = book["Data"]
    sheet["B1"].formula = "_xlfn.SEQUENCE(3)"
    written = sheet.cell_element(CellRef(1, 2))
    formula = written.child("f")
    assert formula is not None
    formula.set("t", "array")
    formula.set("ref", "B1:B3")
    # Cell metadata is how a file tells a spill from Ctrl+Shift+Enter.
    written.set("cm", "1")
    sheet["C1"] = 7
    sheet["D1"].formula = "SUM(B1#)"
    sheet["D2"].formula = "SUM(_xlfn.ANCHORARRAY(B1))"
    sheet["D3"].formula = "SUM(C1#)"
    book.calculate()
    assert [sheet[f"D{row}"].value for row in range(1, 4)] == [6, 6, CellError("#REF!")]


def test_subtotal_leaves_out_filtered_rows_and_on_request_hidden_ones(book: Workbook) -> None:
    sheet = book["Data"]
    sheet["A1"] = "N"
    for row, value in enumerate((1, 2, 3, 4), start=2):
        sheet[f"A{row}"] = value
    # The filter hides row 4; row 5, below its range, is hidden by hand.
    sheet.set_auto_filter("A1:A4", [FilterColumn(0, criteria("<3"))])
    assert sheet.row_hidden(4)
    sheet.set_row_hidden(5, True)
    sheet["C1"].formula = "SUBTOTAL(9,A2:A5)"
    sheet["C2"].formula = "SUBTOTAL(109,A2:A5)"
    sheet["C3"].formula = "_xlfn.AGGREGATE(9,5,A2:A5)"
    sheet["C4"].formula = "_xlfn.AGGREGATE(9,4,A2:A5)"
    sheet["C5"].formula = "SUBTOTAL(9,C1:C2)"
    book.calculate()
    assert [sheet[f"C{row}"].value for row in range(1, 6)] == [7, 3, 3, 10, 0]


def test_database_criteria(book: Workbook) -> None:
    sheet = book["Data"]
    rows = [("Name", "Score"), ("Pear", 5), ("Pearmain", 7), ("Apple", 9)]
    for row, (name, score) in enumerate(rows, start=1):
        sheet[f"A{row}"] = name
        sheet[f"B{row}"] = score
    sheet["D1"] = "Name"
    sheet["D2"] = "Pear"
    sheet["E1"] = "Name"
    sheet["E2"] = "=Pear"
    sheet["F1"] = "High"
    sheet["F2"].formula = "B2>6"
    sheet["H1"].formula = 'DSUM(A1:B4,"Score",D1:D2)'
    sheet["H2"].formula = 'DSUM(A1:B4,"Score",E1:E2)'
    sheet["H3"].formula = "DCOUNT(A1:B4,,F1:F2)"
    book.calculate()
    # Text with no operator matches as a prefix; a formula under a label
    # that names no column is asked of every record.
    assert [sheet[f"H{row}"].value for row in range(1, 4)] == [12, 5, 2]


def test_cell_names_the_workbook_it_is_calculated_as(book: Workbook, live_empty_xlsx: Path) -> None:
    other = book.add_sheet("Other")
    other["B5"] = 1
    sheet = book["Data"]
    sheet["A1"].formula = 'CELL("address",Other!B5)'
    sheet["A2"].formula = 'CELL("address",Data!B5)'
    sheet["A3"].formula = 'CELL("filename",A1)'
    sheet["A4"].formula = 'HYPERLINK("https://example.com","site")'
    book.calculate()
    name = live_empty_xlsx.name
    assert sheet["A1"].value == f"[{name}]Other!$B$5"
    assert sheet["A2"].value == "$B$5"
    assert sheet["A3"].value == f"{live_empty_xlsx.parent}{os.sep}[{name}]Data"
    assert sheet["A4"].value == "site"


def test_cell_asking_about_formatting_keeps_the_cached_value(book: Workbook) -> None:
    sheet = book["Data"]
    sheet["A1"].formula = 'CELL("width",B1)'
    report = book.calculate()
    assert report.unsupported == {"Data!A1": 'CELL("width"), which reads formatting'}


def _flows(book: Workbook, flows: list[float]) -> str:
    sheet = book["Data"]
    for column, flow in enumerate(flows, start=1):
        sheet.cell(1, column).value = flow
    return f"A1:{sheet.cell(1, len(flows)).a1}"


def test_irr_stops_short_of_the_root_where_excel_does(book: Workbook) -> None:
    # 2.9e-10 of the result from the exact root: Excel's secant stops there.
    flows = _flows(book, [-23518.90552495122, 5897.322375982933, 8074.228465462295, 3418.6542983744653, 6167.091710440405])
    assert book["Data"].evaluate(f"IRR({flows},0.3)") == 0.0006747819308212666


def test_irr_starts_again_from_a_tenth_when_its_guess_fails(book: Workbook) -> None:
    # From a guess of 1 the secant overshoots and cycles for 200 steps.
    flows = _flows(
        book,
        [
            -24518.42305940892, 1937.6740514260678, 1767.627855347036, 1621.8383666307395, 3423.8749530313867,
            4005.481458699994, 3303.67042464501, 4648.740185700573, 2305.081428575526, 7245.145640629922,
            5949.791802306586,
        ],
    )
    assert book["Data"].evaluate(f"IRR({flows},1)") == 0.06297805183414074


def _dated(book: Workbook, flows: list[float], days: list[float]) -> str:
    sheet = book["Data"]
    for column, (flow, day) in enumerate(zip(flows, days, strict=True), start=1):
        sheet.cell(1, column).value = flow
        sheet.cell(2, column).value = day
    last = column_letter(len(flows))
    return f"A1:{last}1,A2:{last}2"


def test_xirr_halves_a_bracket_rather_than_stepping_to_the_root(book: Workbook) -> None:
    # A guess 1e-11 above the root: bisection from [0, 2g] still halves
    # down to g (1 - 2^-24) before its tolerance is met.
    series = _dated(book, [-1000.0, 300.0, 400.0, 500.0], [43000.0, 43365.0, 43730.0, 44095.0])
    # In a cell: typed into the formula, the guess would keep 15 digits.
    book["Data"]["A3"] = 0.08896339470224628
    assert book["Data"].evaluate(f"XIRR({series},A3)") == 0.08896338939961473


def test_xirr_edges(book: Workbook) -> None:
    sheet = book["Data"]
    year = [43000.0, 43365.0]
    # A guess of 0 stands for 0.00001.
    assert sheet.evaluate(f"XIRR({_dated(book, [-1000.0, 1200.0], year)},0)") == 0.19999999511718752
    # The root exactly at the end of a bracket, where XNPV is exactly 0.
    assert sheet.evaluate(f"XIRR({_dated(book, [1000.0, -1200.0], year)},0.1)") == 0.19999999403953553
    # Above a negative guess only [g, 0] is searched; below 1 only down to -1.
    assert sheet.evaluate(f"XIRR({_dated(book, [-1000.0, 1250.0], year)},-0.1)") == CellError("#NUM!")
    assert sheet.evaluate(f"XIRR({_dated(book, [-1000.0, 800.0], year)},1)") == CellError("#NUM!")


def test_irr_needs_a_residual_under_its_tolerance(book: Workbook) -> None:
    # Scaled by 2^40, only an exact zero passes. No rate reaches one here;
    # the second series reaches one only after its guess stalls.
    assert book["Data"].evaluate(f"IRR({_flows(book, [-1069510457926.0002, 1099511627776.0])})") == CellError("#NUM!")
    assert book["Data"].evaluate(f"IRR({_flows(book, [-977338919173.8009, 1099511627776.0])},0.05)") == 0.12500546760736642


def _bond(book: Workbook, values: list[float]) -> str:
    """The arguments in row 1, where no formula literal cuts them to 15
    digits."""
    sheet = book["Data"]
    for column, value in enumerate(values, start=1):
        sheet.cell(1, column).value = value
    return ",".join(f"{column_letter(column)}1" for column in range(1, len(values) + 1))


def test_duration_times_coupons_as_price_does_and_the_redemption_its_own_way(book: Workbook) -> None:
    sheet = book["Data"]
    # A coupon's time is index + DSC/E; the redemption's is DSC/E + N - 1,
    # which rounds another way.
    args = _bond(book, [44468.0, 44647.0, 0.035429952036797184, 0.16067767704236355, 4.0, 2.0])
    assert sheet.evaluate(f"DURATION({args})") == 0.492182006415991
    # Settled on a coupon date: each coupon weighs its time times its
    # present value, not its time times the coupon over the power.
    args = _bond(book, [49383.0, 50844.0, 0.13249040076812385, 0.030546357927438164, 2.0, 1.0])
    assert sheet.evaluate(f"DURATION({args})") == 3.3581234390639882


def test_mduration_divides_by_the_growth(book: Workbook) -> None:
    args = _bond(book, [41018.0, 41334.0, 0.10018741153059099, 0.11925232538933145, 4.0, 0.0])
    assert book["Data"].evaluate(f"MDURATION({args})") == 0.806202009349328


def test_price_takes_accrued_interest_from_the_rate(book: Workbook) -> None:
    sheet = book["Data"]
    # A/E * rate * 100 / frequency with coupons to come, not coupon * A/E.
    args = _bond(book, [45098.0, 45403.0, 0.13431405895368478, 0.18167305658967808, 100.0, 2.0, 2.0])
    assert sheet.evaluate(f"PRICE({args})") == 96.42924034785027
    args = _bond(book, [40129.0, 48923.0, 0.015269366925158684, 0.08153059843154511, 98.78907894617276, 2.0, 0.0])
    assert sheet.evaluate(f"PRICE({args})") == 30.415592647768403
    # With one coupon left it is coupon * (A/E).
    args = _bond(book, [44139.0, 44198.0, 0.0858656433341273, 0.015639045918274874, 98.14162125781019, 1.0, 3.0])
    assert sheet.evaluate(f"PRICE({args})") == 99.24148797292862


def test_yield_stops_up_to_1e_10_short_of_the_root(book: Workbook) -> None:
    sheet = book["Data"]
    at_price = "YIELD(A1,B1,C1,PRICE(A1,B1,C1,D1,E1,F1,G1),E1,F1,G1)"
    # One Newton step from the start lands 1e-10 below the root, 0.15; the
    # next step would be under 1e-10, so that is the answer.
    _bond(book, [40000.0, 41533.0, 0.141, 0.15, 100.0, 2.0, 4.0])
    assert sheet.evaluate(at_price) == 0.14999999990009072
    # On a zero-coupon bond the start is one Halley step from 0, a cubic in
    # the root away from it, so near 0 it is the answer itself.
    _bond(book, [40000.0, 47300.0, 0.0, 0.00010000000000000026, 100.0, 2.0, 0.0])
    assert sheet.evaluate(at_price) == 9.999996673483998e-05
    # The start's years count 30/360 as YEARFRAC does: from 28 February,
    # the month's last day, to 31 May is 91 days, not 90.
    args = _bond(book, [44985.0, 47999.0, 0.045, 91.5, 100.0, 1.0, 0.0])
    assert sheet.evaluate(f"YIELD({args})") == 0.058221415821190996


def test_yield_with_one_coupon_left(book: Workbook) -> None:
    sheet = book["Data"]
    # Actual/360: the period is the calendar's 182 days, not 360 / 2.
    args = _bond(book, [36505.0, 36596.0, 0.0, 91.36144230982995, 100.0, 2.0, 2.0])
    assert sheet.evaluate(f"YIELD({args})") == 0.3782145934550592
    # 30/360 from 28 April to 28 May is 30 days to redemption, not the 32
    # left of a period that began on 28 February.
    args = _bond(book, [38835.0, 38865.0, 0.0, 105.93784568694602, 92.60490729347165, 4.0, 0.0])
    assert sheet.evaluate(f"YIELD({args})") == -1.5102748190150097


def test_oddlprice_steps_quasi_coupon_dates_from_the_last_interest_date(book: Workbook) -> None:
    sheet = book["Data"]
    last = "DATE(2021,8,31)"
    # Each date a quarter after the one before: 30 November, 28 February,
    # then 28 May. From a month end, that last covers any maturity in May.
    assert sheet.evaluate(f"ODDLPRICE(DATE(2021,9,10),DATE(2022,5,30),{last},0.1,0.1,100,4,1)") == 100.0338820483653
    assert sheet.evaluate(f"ODDLPRICE(DATE(2021,9,10),DATE(2022,5,27),{last},0.1,0.1,100,4,1)") == 99.98155503258074
    assert sheet.evaluate("ODDLPRICE(DATE(2011,10,26),DATE(2012,7,17),DATE(2010,12,31),0.05,0.06,100,2,1)") == 99.13490142042163


def test_oddlprice_on_basis_0_counts_month_ends_as_the_30th_in_dc(book: Workbook) -> None:
    sheet = book["Data"]
    # 13 October to 31 March is 167 days, where YEARFRAC's rules count 168;
    # the quasi period from 28 February is 178 days long, not 180.
    assert sheet.evaluate("ODDLPRICE(DATE(2032,11,19),DATE(2033,3,31),DATE(2032,10,13),0.1,0,100,2,0)") == 103.63888888888889
    assert sheet.evaluate("ODDLPRICE(DATE(2018,4,2),DATE(2018,5,14),DATE(2018,2,28),0.1,0.1,100,2,0)") == 99.9895189314301


def test_oddlyield(book: Workbook) -> None:
    sheet = book["Data"]
    assert sheet.evaluate("ODDLYIELD(DATE(2021,2,3),DATE(2021,7,23),DATE(2021,1,1),0.05,98,100,2,0)") == 0.0938122427300492
    # 30/360 counts no days from 30 July to 31 July: the yield is 0.
    assert sheet.evaluate("ODDLYIELD(DATE(2021,7,30),DATE(2021,7,31),DATE(2021,6,15),0.05,98,100,2,0)") == 0


def test_oddfprice_with_a_short_or_long_first_period(book: Workbook) -> None:
    sheet = book["Data"]
    short = "DATE(2021,2,3),DATE(2023,6,30),DATE(2021,1,1),DATE(2021,6,30)"
    assert sheet.evaluate(f"ODDFPRICE({short},0.05,0.06,100,2,0)") == 97.78308404154141
    assert sheet.evaluate(f"ODDFYIELD({short},0.05,98,100,2,0)") == 0.059007229164258115
    # The coupon period comes from the first coupon's schedule, 30 November
    # to 28 February, not the maturity's 29 November.
    assert sheet.evaluate("ODDFPRICE(DATE(2019,2,5),DATE(2023,11,29),DATE(2018,12,1),DATE(2019,2,28),0.1,0,100,4,1)") == 148.13888888888889
    # A long first period to a month-end first coupon discounts one period
    # more, unless settlement shares its month with a quasi-coupon date.
    long = "DATE(2027,4,30),DATE(2021,10,29),DATE(2023,4,30),0.1,0.1,100,1,0"
    assert sheet.evaluate(f"ODDFPRICE(DATE(2021,11,3),{long})") == 90.60114237553395
    assert sheet.evaluate(f"ODDFPRICE(DATE(2022,4,12),{long})") == 99.54578216866932
    mid_month = "DATE(2022,4,12),DATE(2027,6,15),DATE(2021,10,29),DATE(2023,6,15),0.1"
    assert sheet.evaluate(f"ODDFPRICE({mid_month},0.1,100,1,0)") == 99.43079730491951
    assert sheet.evaluate(f"ODDFYIELD({mid_month},90,100,1,0)") == 0.12551573318812853


def test_oddfprice_refusals(book: Workbook) -> None:
    sheet = book["Data"]
    # An odd first period exactly one period long.
    odd = "DATE(2021,2,1),DATE(2023,6,30),DATE(2020,12,31),DATE(2021,6,30)"
    assert sheet.evaluate(f"ODDFPRICE({odd},0.05,0.06,100,2,0)") == CellError("#NUM!")
    # A first coupon that is not one of the maturity's coupon dates.
    off = "DATE(2021,2,3),DATE(2023,6,30),DATE(2021,1,1),DATE(2021,6,29)"
    assert sheet.evaluate(f"ODDFPRICE({off},0.05,0.06,100,2,0)") == CellError("#NUM!")


def test_pmt_divides_by_1_less_the_discount_and_multiplies_by_the_rate(book: Workbook) -> None:
    sheet = book["Data"]
    # w = 1 - (1 + r)^-n, and -((pv + fv)/w - fv) r: the exact payment,
    # rounded once, is a unit or two away in each of these.
    assert sheet.evaluate("PMT(0.01,12,1000,0,0)") == -88.84878867834168
    assert sheet.evaluate("PMT(-0.02,60,1000,0,0)") == -8.471904731198892
    assert sheet.evaluate("PMT(1E-20,12,1000)") == -83.33333333333334
    # At the start of each period r becomes 1/(1/r + 1).
    assert sheet.evaluate("PMT(0.01,360,0,5000,1)") == -1.4164651943319049
    assert sheet.evaluate("PMT(0.3,30,250000,5000,1)") == -57714.77666967648


def test_pmt_errors(book: Workbook) -> None:
    sheet = book["Data"]
    assert sheet.evaluate("PMT(0.05,0,1000)") == CellError("#NUM!")
    assert sheet.evaluate("PMT(-1,12,1000)") == CellError("#NUM!")
    # (1 + r)^-n overflows.
    assert sheet.evaluate("PMT(-0.05,1000000,1000)") == CellError("#NUM!")
    assert sheet.evaluate("PMT(0,7,0.1,0.2,1)") == -0.042857142857142864


def test_ipmt_and_ppmt_form_the_balance_and_the_discount_as_pmt_does(book: Workbook) -> None:
    sheet = book["Data"]
    # The balance going into the period, (pv + fv)/w_n w_m - fv with m
    # payments left, times the rate: the exact interest, rounded once, is
    # 0.9990000000000009 and -8.562690475450857.
    assert sheet.evaluate("IPMT(-0.999,2,10,1000)") == 0.9990000000000073
    assert sheet.evaluate("IPMT(0.01,2.5,10,1000)") == -8.562690475450859
    # The principal is (pv + fv)/w_n (1 + r)^-m times the rate, not the
    # payment less the interest.
    assert sheet.evaluate("PPMT(0.01,1,10,1000)") == -95.58207655117134
    assert sheet.evaluate("PPMT(1E-20,2,12,1000)") == -83.33333333333334
    # Paid at the start of each period, the first payment is all principal.
    assert sheet.evaluate("IPMT(0.01,1,10,1000,0,1)") == 0
    assert sheet.evaluate("PPMT(0.01,1,10,1000,0,1)") == -104.5367094566053


def test_ipmt_takes_a_period_up_to_the_term_plus_one(book: Workbook) -> None:
    sheet = book["Data"]
    assert sheet.evaluate("IPMT(0.01,10.5,10,1000)") == -0.5239837631578298
    assert sheet.evaluate("IPMT(0.01,11,10,1000)") == CellError("#NUM!")
    assert sheet.evaluate("PPMT(0.01,0.99,10,1000)") == CellError("#NUM!")
    # A rate of -1 or below is refused, as PMT refuses it.
    assert sheet.evaluate("IPMT(-1,2,10,1000)") == CellError("#NUM!")
    assert sheet.evaluate("PPMT(-1.5,2,10,1000)") == CellError("#NUM!")


def test_cumipmt_and_cumprinc_in_closed_form(book: Workbook) -> None:
    sheet = book["Data"]
    # The principal repaid is pv/w_n (1 + r)^-(n - end + 1) w_count, times
    # 1 + r when payments fall at the end of each period, and the interest
    # is the payments less it: the exact sum of the periods, rounded once,
    # is -55.82076551171361, -1000 and -6.5e-17.
    assert sheet.evaluate("CUMIPMT(0.01,10,1000,1,10,0)") == -55.82076551171349
    assert sheet.evaluate("CUMPRINC(5,10,1000,1,10,0)") == -1000.0000000000001
    assert sheet.evaluate("CUMIPMT(1E-20,12,1000,1,12,0)") == -1.1368683772161603e-13
    # At the start of each period a first payment in the span is all
    # principal, added on its own.
    assert sheet.evaluate("CUMPRINC(0.01,10,1000,1,2,1)") == -200.11878600777663
    assert sheet.evaluate("CUMIPMT(0.01,10,1000,1,2,1)") == -8.95463290543394


def test_cumipmt_rounds_a_fractional_start_up_and_end_down(book: Workbook) -> None:
    sheet = book["Data"]
    # From period 2 and to period 9.
    assert sheet.evaluate("CUMIPMT(0.01,10,1000,1.1,10,0)") == -45.820765511713375
    assert sheet.evaluate("CUMIPMT(0.01,10,1000,2,9.9,0)") == -44.77539841714736
    assert sheet.evaluate("CUMPRINC(0.01,10,1000,1.4,1.4,1)") == 0
    # The checks read the periods as given.
    assert sheet.evaluate("CUMIPMT(0.01,10,1000,0.9,10,0)") == CellError("#NUM!")
    assert sheet.evaluate("CUMIPMT(0.01,10.5,1000,1,10.7,0)") == CellError("#NUM!")
    assert sheet.evaluate("CUMIPMT(0.01,10,1000,2.5,2.4,0)") == CellError("#NUM!")


def test_below_1_the_hyperbolic_functions_take_e_to_the_x_less_1_by_kahans_trick(book: Workbook) -> None:
    sheet = book["Data"]
    # a = e^x - 1 and b = e^-x - 1, each (u - 1) x / ln u with u = EXP(x):
    # SINH is (a - b)/2, TANH (a - b)/(a + b + 2).
    assert sheet.evaluate("SINH(0.7)") == 0.7585837018395336
    assert sheet.evaluate("SINH(-0.01)") == -0.010000166667500001
    assert sheet.evaluate("TANH(0.001)") == 0.0009999996666668002
    assert sheet.evaluate("TANH(0.4)") == 0.3799489622552249


def test_weibull_takes_its_distribution_as_1_less_e_to_the_minus_t_below_1(book: Workbook) -> None:
    sheet = book["Data"]
    # -(e^-t - 1) by Kahan's trick, not 1 - EXP(-t), up to t = 1.
    assert sheet.evaluate("WEIBULL.DIST(0.1,1,1,TRUE)") == 0.09516258196404045
    assert sheet.evaluate("WEIBULL.DIST(0.85,1,1,TRUE)") == 0.5725850680512734


# ----------------------------------------------------------------------
# GROUPBY and PIVOTBY
# ----------------------------------------------------------------------

#: Sheet In: orders with a blank region, a number for one, one region
#: spelled in lower case, a blank product, and a blank and a text quantity.
ORDERS: list[tuple[CellValue, ...]] = [
    ("Region", "Product", "Year", "Qty", "Price", "Flag"),
    ("East", "Pen", 2023, 3, 1.5, True), ("West", "Book", 2024, 1, 12.0, False),
    ("East", "Ink", 2024, 5, 4.25, True), ("North", "Pen", 2023, 2, 1.5, True),
    ("West", "Pen", 2023, 4, 1.25, True), ("east", "Book", 2023, 2, 11.5, False),
    ("South", "Ink", 2024, 7, 4.0, True), ("North", "Book", 2024, 1, 13.0, True),
    ("West", "Ink", 2023, 6, 4.5, False), ("East", "Pen", 2024, 8, 1.75, True),
    (None, "Pen", 2024, 2, 1.5, True), ("South", None, 2023, 3, 4.0, True),
    ("North", "Ink", 2023, None, 4.25, True), ("East", "Book", 2024, 2, 12.5, False),
    ("West", "Book", 2023, "x", 12.0, True), (7, "Pen", 2024, 1, 1.5, True),
    ("South", "Pen", 2024, 9, 1.25, False), ("North", "Pen", 2024, 3, 1.5, True),
]  # fmt: skip

#: Sheet Pv: groups whose key order, total order and leaves' orders all
#: differ. a: x = 4 + 6, y = 1; b: x = 3, y = 4; c: x = 2, y = 20; d: x = 9,
#: y = 6.
LEVELS: list[tuple[CellValue, ...]] = [
    ("K1", "K2", "T", "V"), ("c", "y", "t", 20), ("a", "x", "t", 4), ("d", "y", "t", 6), ("b", "x", "t", 3),
    ("a", "y", "t", 1), ("c", "x", "t", 2), ("d", "x", "t", 9), ("b", "y", "t", 4), ("a", "x", "t", 6),
]  # fmt: skip


def _grouped(book: Workbook) -> None:
    """The sheets the grouping tests read: In, Pv, and Mw, where each leaf
    of A/B/C by p/q has one row and value columns worth 1, 100, 10000 and
    1000000 times the powers of two, so that every sum says what it added."""
    for name, rows in (("In", ORDERS), ("Pv", LEVELS)):
        sheet = book.add_sheet(name)
        for row, values in enumerate(rows, start=1):
            for column, value in enumerate(values, start=1):
                if value is not None:
                    sheet[f"{column_letter(column)}{row}"] = value
    sheet = book.add_sheet("Mw")
    for row, (first, second) in enumerate([("A", "p"), ("A", "q"), ("B", "p"), ("B", "q"), ("C", "p"), ("C", "q")], start=2):
        sheet[f"A{row}"] = "r"
        sheet[f"B{row}"] = first
        sheet[f"C{row}"] = second
        for offset, column in enumerate("DEFG"):
            sheet[f"{column}{row}"] = 2 ** (row - 2) * 100**offset


def _text(sheet: Worksheet, formula: str) -> object:
    """The formula's result as ARRAYTOTEXT(...,1) writes it, which is how
    Excel's results were read."""
    return sheet.evaluate(f"ARRAYTOTEXT({formula},1)")


def test_groupby_matches_keys_as_sort_does_and_spells_each_by_its_first_leaf(book: Workbook) -> None:
    _grouped(book)
    sheet = book["Data"]
    # "east" and "East" are one group, spelled from the first row of its
    # first leaf: Book, found in "east" 2023; with products descending,
    # Pen, found first in "East". A blank key sorts last either way.
    assert _text(sheet, "GROUPBY(In!A2:B19,In!D2:D19,_xleta.SUM,0,2)") == (
        '{7,"Pen",1;7,"",1;"east","Book",4;"east","Ink",5;"east","Pen",11;"east","",20;"North","Book",1;'
        '"North","Ink",0;"North","Pen",5;"North","",6;"South","Ink",7;"South","Pen",9;"South",,3;"South","",19;'
        '"West","Book",1;"West","Ink",6;"West","Pen",4;"West","",11;,"Pen",2;,"",2;"Grand Total","",59}'
    )
    assert _text(sheet, "GROUPBY(In!A2:B19,In!D2:D19,_xleta.SUM,0,1,{1,-2})") == (
        '{7,"Pen",1;"East","Pen",11;"East","Ink",5;"East","Book",4;"North","Pen",5;"North","Ink",0;'
        '"North","Book",1;"South","Pen",9;"South","Ink",7;"South",,3;"West","Pen",4;"West","Ink",6;'
        '"West","Book",1;,"Pen",2;"Total","",59}'
    )


def test_groupby_sorts_every_level_by_its_own_groups_result(book: Workbook) -> None:
    _grouped(book)
    assert _text(book["Data"], "GROUPBY(Pv!A2:B10,Pv!D2:D10,_xleta.SUM,0,1,3)") == (
        '{"b","x",3;"b","y",4;"a","y",1;"a","x",10;"d","y",6;"d","x",9;"c","x",2;"c","y",20;"Total","",55}'
    )


def test_groupby_takes_headers_when_the_values_start_with_text_above_numbers(book: Workbook) -> None:
    sheet = book["Data"]
    for row, (key, number, text) in enumerate([("k", "v", "v"), ("a", None, 1), ("b", 1, "z"), ("a", 2, 2), ("b", 3, 3)], start=1):
        sheet[f"A{row}"] = key
        if number is not None:
            sheet[f"B{row}"] = number
        sheet[f"C{row}"] = text
    assert _text(sheet, "GROUPBY(A1:A5,B1:B5,_xleta.SUM)") == '{"a",2;"b",4;"Total",6}'
    # Text below the header is #VALUE!; text all the way down is no header.
    assert _text(sheet, "GROUPBY(A1:A5,C1:C5,_xleta.SUM)") == CellError("#VALUE!")
    assert _text(sheet, 'GROUPBY({1;2;1},{"a";"b";"c"},_xleta.SUM)') == '{1,0;2,0;"Total",0}'


def test_groupby_names_the_value_column_beside_each_stacked_function(book: Workbook) -> None:
    _grouped(book)
    assert _text(book["Data"], "GROUPBY(In!A1:A19,In!D1:E19,VSTACK(_xleta.SUM,_xleta.MAX),3)") == (
        '{"Region","","","";7,"SUM","Qty",1;7,"MAX","Price",1.5;"East","SUM","Qty",20;"East","MAX","Price",12.5;'
        '"North","SUM","Qty",6;"North","MAX","Price",13;"South","SUM","Qty",19;"South","MAX","Price",4;'
        '"West","SUM","Qty",11;"West","MAX","Price",12;,"SUM","Qty",2;,"MAX","Price",1.5;"Total","SUM","Qty",59;'
        '"Total","MAX","Price",13}'
    )


def test_groupby_keeps_an_error_in_its_cell_but_refuses_an_array(book: Workbook) -> None:
    _grouped(book)
    sheet = book["Data"]
    assert _text(sheet, 'GROUPBY(In!A2:A19,In!D2:D19,LAMBDA(x,IF(ROWS(x)>3,VALUE("a"),1)))') == (
        '{7,1;"East",#VALUE!;"North",#VALUE!;"South",1;"West",#VALUE!;,1;"Total",#VALUE!}'
    )
    # A group reaches the function as an array, even a group of one row,
    # and a result that is an array refuses the whole summary.
    assert _text(sheet, "GROUPBY(In!A2:A19,In!D2:D19,LAMBDA(x,x))") == CellError("#VALUE!")
    # COUNTBLANK takes only a range: given the group's array it gives an
    # array, one #VALUE! an item, and an array result refuses it all.
    assert _text(sheet, "GROUPBY(In!A2:A19,In!D2:D19,LAMBDA(x,COUNTBLANK(x)))") == CellError("#VALUE!")


def test_pivotby_spreads_a_field_across_the_top_with_subtotals_first(book: Workbook) -> None:
    _grouped(book)
    sheet = book["Data"]
    assert _text(sheet, "PIVOTBY(In!A2:A19,In!C2:C19,In!E2:E19,_xleta.SUM)") == (
        '{"",2023,2024,"Total";7,"",1.5,1.5;"East",13,18.5,31.5;"North",5.75,14.5,20.25;"South",4,5.25,9.25;'
        '"West",17.75,12,29.75;,"",1.5,1.5;"Total",40.5,53.25,93.75}'
    )
    assert _text(sheet, "PIVOTBY(In!A2:A19,In!B2:C19,In!E2:E19,_xleta.SUM,0,1,,-2)") == (
        '{"","Grand Total","Book","Book","Book","Ink","Ink","Ink","Pen","Pen","Pen",,;'
        '"","","",2023,2024,"",2023,2024,"",2023,2024,"",2023;7,1.5,"","","","","","",1.5,"",1.5,"","";'
        '"East",31.5,24,11.5,12.5,4.25,"",4.25,3.25,1.5,1.75,"","";"North",20.25,13,"",13,4.25,4.25,"",3,1.5,1.5,"","";'
        '"South",9.25,"","","",4,"",4,1.25,"",1.25,4,4;"West",29.75,24,12,12,4.5,4.5,"",1.25,1.25,"","","";'
        ',1.5,"","","","","","",1.5,"",1.5,"","";"Total",93.75,61,23.5,37.5,17,8.75,8.25,11.75,4.25,7.5,4,4}'
    )


def test_pivotby_shows_the_column_fields_names_as_one_text(book: Workbook) -> None:
    _grouped(book)
    assert _text(book["Data"], "PIVOTBY(In!A1:B19,In!B1:C19,In!E1:E19,_xleta.SUM,3)") == (
        '{"","","Product, Year","","","","","","","";"","","Book","Book","Ink","Ink","Pen","Pen",,"Total";'
        '"","",2023,2024,2023,2024,2023,2024,2023,"";"Region","Product","Price","Price","Price","Price","Price",'
        '"Price","Price","Price";7,"Pen","","","","","",1.5,"",1.5;"east","Book",11.5,12.5,"","","","","",24;'
        '"east","Ink","","","",4.25,"","","",4.25;"east","Pen","","","","",1.5,1.75,"",3.25;'
        '"North","Book","",13,"","","","","",13;"North","Ink","","",4.25,"","","","",4.25;'
        '"North","Pen","","","","",1.5,1.5,"",3;"South","Ink","","","",4,"","","",4;'
        '"South","Pen","","","","","",1.25,"",1.25;"South",,"","","","","","",4,4;"West","Book",12,12,"","","","","",24;'
        '"West","Ink","","",4.5,"","","","",4.5;"West","Pen","","","","",1.25,"","",1.25;'
        ',"Pen","","","","","",1.5,"",1.5;"Total","",23.5,37.5,8.75,8.25,4.25,7.5,4,93.75}'
    )


def test_pivotby_relative_to_the_parent_row(book: Workbook) -> None:
    _grouped(book)
    # ROWS(y) counts the rows of the cell's parent row, the grand total's
    # for a subtotal, its own for the grand total.
    formula = "PIVOTBY(In!A2:B19,In!B2:C19,In!E2:E19,LAMBDA(x,y,ROWS(y)),0,2,,2,,,4)"
    assert _text(book["Data"], formula) == (
        '{"","","Book","Book","Book","Ink","Ink","Ink","Pen","Pen","Pen",,,"Grand Total";'
        '"","",2023,2024,"",2023,2024,"",2023,2024,"",2023,"","";7,"Pen","","","","","","","",1,1,"","",1;'
        '7,"","","","","","","","",5,8,"","",18;"east","Book",1,1,2,"","","","","","","","",5;'
        '"east","Ink","","","","",1,1,"","","","","",5;"east","Pen","","","","","","",1,1,2,"","",5;'
        '"east","",2,3,5,"",2,4,3,5,8,"","",18;"North","Book","",1,1,"","","","","","","","",4;'
        '"North","Ink","","","",1,"",1,"","","","","",4;"North","Pen","","","","","","",1,1,2,"","",4;'
        '"North","","",3,5,2,"",4,3,5,8,"","",18;"South","Ink","","","","",1,1,"","","","","",3;'
        '"South","Pen","","","","","","","",1,1,"","",3;"South",,"","","","","","","","","",1,1,3;'
        '"South","","","","","",2,4,"",5,8,1,1,18;"West","Book",1,1,2,"","","","","","","","",4;'
        '"West","Ink","","","",1,"",1,"","","","","",4;"West","Pen","","","","","","",1,"",1,"","",4;'
        '"West","",2,3,5,2,"",4,3,"",8,"","",18;,"Pen","","","","","","","",1,1,"","",1;'
        ',"","","","","","","","",5,8,"","",18;"Grand Total","",2,3,5,2,2,4,3,5,8,1,1,18}'
    )


def test_pivotby_sorts_by_a_value_only_where_there_is_a_total_to_sort_by(book: Workbook) -> None:
    _grouped(book)
    sheet = book["Data"]
    # Across the top the leaves sort by their totals, but a, b, c and d,
    # having no columns of their own, keep key order until subtotals show.
    assert _text(sheet, "PIVOTBY(Pv!C2:C10,Pv!A2:B10,Pv!D2:D10,_xleta.SUM,0,1,,1,3)") == (
        '{"","a","a","b","b","c","c","d","d","Total";"","y","x","x","y","x","y","y","x","";'
        '"t",1,10,3,4,2,20,6,9,55;"Total",1,10,3,4,2,20,6,9,55}'
    )
    assert _text(sheet, "PIVOTBY(Pv!C2:C10,Pv!A2:B10,Pv!D2:D10,_xleta.SUM,0,1,,2,3)") == (
        '{"","b","b","b","a","a","a","d","d","d","c","c","c","Grand Total";"","x","y","","y","x","","y","x","","x","y","","";'
        '"t",3,4,7,1,10,11,6,9,15,2,20,22,55;"Total",3,4,7,1,10,11,6,9,15,2,20,22,55}'
    )
    # Down the side every level sorts, but only beside a total column.
    assert _text(sheet, "PIVOTBY(Pv!A2:B10,Pv!C2:C10,Pv!D2:D10,_xleta.SUM,0,1,-3)") == (
        '{"","","t","Total";"c","y",20,20;"c","x",2,2;"d","x",9,9;"d","y",6,6;"a","x",10,10;"a","y",1,1;'
        '"b","y",4,4;"b","x",3,3;"Total","",55,55}'
    )
    assert _text(sheet, "PIVOTBY(Pv!A2:A10,Pv!B2:B10,Pv!D2:D10,_xleta.SUM,0,1,2,0)") == (
        '{"","x","y";"a",10,1;"b",3,4;"c",2,20;"d",9,6;"Total",24,31}'
    )


def test_pivotby_misplaces_one_functions_totals_over_several_value_columns(book: Workbook) -> None:
    _grouped(book)
    sheet = book["Data"]
    # Each total goes one column along a node, the grand total first, and
    # the leaves are written over them: of the grand total (63 and 6300)
    # only 6300 survives, in B's subtotal.
    assert _text(sheet, "PIVOTBY(Mw!A2:A7,Mw!B2:C7,Mw!D2:E7,_xleta.SUM,0,0,,2)") == (
        '{"","A","A","A","A","A","A","B","B","B","B","B","B","C","C","C","C","C","C","Grand Total","Grand Total";'
        '"","p","p","q","q","","","p","p","q","q","","","p","p","q","q","","","","";'
        '"r",1,100,2,200,"","",4,400,8,800,6300,"",16,1600,32,3200,"","","",""}'
    )
    # Totals first move the columns as they were placed.
    assert _text(sheet, "PIVOTBY(Mw!A2:A7,Mw!B2:C7,Mw!D2:G7,_xleta.SUM,0,0,,-2)") == (
        '{"","Grand Total","Grand Total","Grand Total","Grand Total","A","A","A","A","A","A","A","A","A","A","A","A",'
        '"B","B","B","B","B","B","B","B","B","B","B","B","C","C","C","C","C","C","C","C","C","C","C","C";'
        '"","","","","","","","","","p","p","p","p","q","q","q","q","","","","","p","p","p","p","q","q","q","q",'
        '"","","","","p","p","p","p","q","q","q","q";"r","","","","","",63,6300,630000,1,100,10000,1000000,'
        '2,200,20000,2000000,"","","","",4,400,40000,4000000,8,800,80000,8000000,"","","","",'
        '16,1600,160000,16000000,32,3200,320000,32000000}'
    )
    assert _text(sheet, "PIVOTBY(Mw!A2:A7,Mw!A2:A7,Mw!D2:E7,_xleta.SUM)") == (
        '{"","r","r","Total","Total";"r",63,6300,6300,"";"Total",63,6300,6300,""}'
    )
    # Two functions over the two columns total as they should.
    assert _text(sheet, "PIVOTBY(Mw!A2:A7,Mw!A2:A7,Mw!D2:E7,HSTACK(_xleta.SUM,_xleta.SUM),0,0)") == (
        '{"","r","r","Total","Total";"","SUM","SUM","SUM","SUM";"r",63,6300,63,6300}'
    )


def test_pivotby_names_functions_and_their_value_columns(book: Workbook) -> None:
    _grouped(book)
    sheet = book["Data"]
    assert _text(sheet, "PIVOTBY(In!A1:A19,In!C1:C19,In!E1:E19,HSTACK(_xleta.SUM,_xleta.MAX),3)") == (
        '{"","Year","","","","","";"",2023,2023,2024,2024,"Total","Total";"","SUM","MAX","SUM","MAX","SUM","MAX";'
        '"Region","Price","Price","Price","Price","Price","Price";7,"","",1.5,1.5,1.5,1.5;'
        '"East",13,11.5,18.5,12.5,31.5,12.5;"North",5.75,4.25,14.5,13,20.25,13;"South",4,4,5.25,4,9.25,4;'
        '"West",17.75,12,12,12,29.75,12;,"","",1.5,1.5,1.5,1.5;"Total",40.5,12,53.25,13,93.75,13}'
    )
    # Stacked over two value columns, each row names its value column and
    # the row fields' names move up beside the last row of keys.
    formula = "PIVOTBY(In!A1:A19,In!C1:C19,In!D1:E19,VSTACK(_xleta.SUM,_xleta.MAX),3)"
    assert _text(sheet, formula) == (
        '{"","","","Year","","";"Region","","",2023,2024,"Total";7,"SUM","Qty","",1,1;7,"MAX","Price","",1.5,1.5;'
        '"East","SUM","Qty",5,15,20;"East","MAX","Price",11.5,12.5,12.5;"North","SUM","Qty",2,4,6;'
        '"North","MAX","Price",4.25,13,13;"South","SUM","Qty",3,16,19;"South","MAX","Price",4,4,4;'
        '"West","SUM","Qty",10,1,11;"West","MAX","Price",12,12,12;,"SUM","Qty","",2,2;,"MAX","Price","",1.5,1.5;'
        '"Total","SUM","Qty",20,39,59;"Total","MAX","Price",12,13,13}'
    )


def test_a_lambdas_arguments_are_counted_and_may_be_optional(book: Workbook) -> None:
    sheet = book["Data"]
    optional = "_xlfn.LAMBDA(_xlpm.x,_xlop.y,IF(_xlfn.ISOMITTED(_xlpm.y),\"om\",_xlpm.y))"
    assert sheet.evaluate(f"{optional}(1)") == "om"
    assert sheet.evaluate(f"{optional}(1,5)") == 5
    assert sheet.evaluate(f"{optional}(1,)") == "om"
    assert sheet.evaluate(f"{optional}(1,2,3)") == CellError("#VALUE!")
    assert sheet.evaluate("LAMBDA(x,x+1)()") == CellError("#VALUE!")
    assert sheet.evaluate("REDUCE(0,{1,2},_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlop.c,_xlpm.a+_xlpm.b))") == 3
    # PERCENTOF needs two arguments where BYROW gives one.
    assert sheet.evaluate("BYROW({1,2;3,4},_xleta.PERCENTOF)") == CellError("#VALUE!")


def test_counta_counts_every_item_of_an_array_but_only_filled_cells(book: Workbook) -> None:
    sheet = book["Data"]
    sheet["C1"] = "a"
    sheet["C3"] = "c"
    assert sheet.evaluate("COUNTA(C1:C3)") == 2
    assert sheet.evaluate("COUNTA(VSTACK(C1:C3))") == 3
    # A blank cell for ignore_empty is FALSE; left out, it is TRUE.
    assert sheet.evaluate('TEXTJOIN("|",,C1:C3)') == "a|c"
    assert sheet.evaluate('TEXTJOIN("|",C2,C1:C3)') == "a||c"
    assert sheet.evaluate('TEXTJOIN("|",C5:C6,C1:C3)') == CellError("#VALUE!")


def test_a_parameter_only_a_range_will_do_takes_an_array_one_item_at_a_time(book: Workbook) -> None:
    _grouped(book)
    sheet = book["Data"]
    assert _text(sheet, "LET(q,SEQUENCE(2,2),COUNTIF(q,1))") == "{#VALUE!,#VALUE!;#VALUE!,#VALUE!}"
    assert _text(sheet, "LET(a,SEQUENCE(3),COUNTIF(a,{1,2}))") == "{#VALUE!,#VALUE!;#VALUE!,#VALUE!;#VALUE!,#VALUE!}"
    assert sheet.evaluate("LET(s,5,COUNTIF(s,1))") == CellError("#VALUE!")
    assert sheet.evaluate("COUNTBLANK((In!D2:D4,In!E2:E4))") == CellError("#VALUE!")


def test_take_and_drop_of_a_range_are_ranges(book: Workbook) -> None:
    _grouped(book)
    sheet = book["Data"]
    assert _text(sheet, "ROW(TAKE(In!D2:D10,3))") == "{2;3;4}"
    assert sheet.evaluate("COUNTIF(DROP(In!D2:E10,1,-1),1)") == 2
    assert sheet.evaluate("ISREF(TAKE(In!D2:D10,20))") is True
    assert sheet.evaluate("ISREF(TAKE({1,2},1))") is False
    assert sheet.evaluate("TAKE((In!D2:D4,In!E2:E4),1)") == CellError("#VALUE!")


def test_trimrange_trims_an_array_of_its_blank_edges(book: Workbook) -> None:
    sheet = book["Data"]
    assert _text(sheet, "TRIMRANGE(VSTACK(Z1,1,Z2))") == "{1}"
    assert _text(sheet, "TRIMRANGE(HSTACK(Z1,1,Z2),,2)") == "{,1}"
    assert sheet.evaluate("TRIMRANGE(VSTACK(Z1,Z2))") == CellError("#VALUE!")


def test_transpose_in_a_formula_before_dynamic_arrays_cuts_its_range_to_the_row(book: Workbook) -> None:
    sheet = book["Data"]
    for row in (1, 2, 3):
        sheet[f"A{row}"] = row
    # Measured with the formulas entered as Excel before dynamic arrays
    # took them: the range is cut to the formula's row, as a single value.
    sheet["C1"].formula = "TRANSPOSE(A1:A2)"
    sheet["C2"].formula = "TRANSPOSE(A1:A3)"
    sheet["C5"].formula = "SUM(TRANSPOSE(A1:A3))"
    sheet["C6"].formula = "INDEX(TRANSPOSE(A1:A3),1,2)"
    book.calculate()
    assert [sheet[f"C{row}"].value for row in (1, 2, 5, 6)] == [1, 2, CellError("#VALUE!"), CellError("#VALUE!")]


def test_bahttext_spells_an_amount_as_excel_does(book: Workbook) -> None:
    sheet = book["Data"]
    # A 1 in the ones is "et" after any other digit, those before a
    # million included; a million's multiple is spelled on its own.
    assert sheet.evaluate("BAHTTEXT(21)") == "ยี่สิบเอ็ดบาทถ้วน"
    assert sheet.evaluate("BAHTTEXT(1000001)") == "หนึ่งล้านเอ็ดบาทถ้วน"
    assert sheet.evaluate("BAHTTEXT(1000001000000)") == "หนึ่งล้านเอ็ดล้านบาทถ้วน"
    assert sheet.evaluate("BAHTTEXT(1E+100)") == "หนึ่งหมื่น" + "ล้าน" * 16 + "บาทถ้วน"
    # Satang are rounded as ROUND rounds, and stand alone below a baht.
    assert sheet.evaluate("BAHTTEXT(1.005)") == "หนึ่งบาทหนึ่งสตางค์"
    assert sheet.evaluate("BAHTTEXT(0.995)") == "หนึ่งบาทถ้วน"
    assert sheet.evaluate("BAHTTEXT(0.21)") == "ยี่สิบเอ็ดสตางค์"
    # A negative amount keeps its minus even when it rounds to nothing.
    assert sheet.evaluate("BAHTTEXT(-0.004)") == "ลบศูนย์บาทถ้วน"
    assert sheet.evaluate("BAHTTEXT(Z99)") == CellError("#VALUE!")


def test_sheet_and_sheets_given_what_is_not_a_range(book: Workbook) -> None:
    sheet = book["Data"]
    assert sheet.evaluate('SHEET("Data")') == 1
    assert sheet.evaluate("SHEET(1)") == CellError("#N/A")
    assert sheet.evaluate('SHEET({"Data"})') == CellError("#N/A")
    assert sheet.evaluate('SHEETS("Data")') == CellError("#N/A")
    assert sheet.evaluate("SHEET(1/0)") == CellError("#DIV/0!")


# ----------------------------------------------------------------------
# Special functions
# ----------------------------------------------------------------------


def _near(value: Decimal) -> float:
    return float(value)


def test_the_error_function() -> None:
    assert _near(special.erf(Decimal(0))) == 0.0
    assert _near(special.erf(Decimal(1))) == 0.8427007929497149
    assert _near(special.erf(Decimal(-1))) == -0.8427007929497149
    assert _near(special.erfc(Decimal(1))) == 0.15729920705028513
    # Far in the tail the complement keeps its digits.
    assert _near(special.erfc(Decimal(5))) == 1.537459794428035e-12
    assert _near(special.normal_cdf(Decimal(0))) == 0.5


def test_gamma_and_its_logarithm() -> None:
    assert _near(special.gamma(Decimal(5))) == 24.0
    # The double nearest the square root of pi, one above the square root
    # of the double nearest pi.
    assert _near(special.gamma(Decimal("0.5"))) == 1.772453850905516
    assert _near(special.lgamma(Decimal(4))) == math.log(6)
    with pytest.raises(ValueError, match="pole"):
        special.gamma(Decimal(-2))


def test_the_incomplete_gamma_and_beta_functions() -> None:
    # P(1, x) is 1 - exp(-x); I_x(2, 3) at 1/2 is 11/16 exactly.
    assert _near(special.gamma_lower(Decimal(1), Decimal(2))) == -math.expm1(-2)
    assert _near(special.gamma_upper(Decimal(1), Decimal(2))) == math.exp(-2)
    assert _near(special.beta_lower(Decimal(2), Decimal(3), Decimal("0.5"))) == 0.6875
    assert _near(special.beta_upper(Decimal(2), Decimal(3), Decimal("0.5"))) == 0.3125


def test_an_inverse_to_the_last_digit() -> None:
    # The quantile every statistics course quotes, 1.959963984540054, is
    # that of 0.975 itself; the double nearest 0.975 is a little below it.
    assert normal_inverse(0.975) == 1.9599639845400538
