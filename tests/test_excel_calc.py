"""The formula engine's parts: the parser, Excel's numeric rules, and the
calculation of a whole workbook.

What any single formula gives is held to Excel in
``test_excel_formula_corpus.py``; this file covers what the corpus cannot
show one formula at a time: the tree a formula parses to, a workbook's
formulas reading one another, and what calculating writes back.
"""

from __future__ import annotations

import datetime as dt
import math
import os
from decimal import Decimal
from pathlib import Path

import pytest

from pyofficeeditor.excel import CellError, FilterColumn, FormulaSyntaxError, Workbook, column_letter, criteria
from pyofficeeditor.excel._calc import parse, special
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
    ],
)
def test_text_that_reads_as_a_number(text: str, number: float) -> None:
    assert text_to_number(text, TODAY) == number


@pytest.mark.parametrize("text", ["", "TRUE", "1,00", "(-5)", "5-", "12:30PM", "1e308", "\t1"])
def test_text_that_is_not_a_number(text: str) -> None:
    assert text_to_number(text, TODAY) is None


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
        "_xlfn.LAMBDA(_xlpm.a,_xlpm.b,IF(_xlfn.ISOMITTED(_xlpm.b),_xlpm.a,_xlpm.a+_xlpm.b))(5)"
    )
    sheet["A6"].formula = "_xlfn.REDUCE(1,{1,2,3,4},_xlfn.LAMBDA(_xlpm.p,_xlpm.v,_xlpm.p*_xlpm.v))"
    book.calculate()
    assert [sheet[f"A{row}"].value for row in range(1, 7)] == [6, 8, 3, CellError("#CALC!"), 5, 24]


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
