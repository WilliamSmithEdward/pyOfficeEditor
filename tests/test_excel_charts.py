"""Charts, and the references in them that follow the cells they read.

``charts.xlsx`` was built by Excel with charts on its data sheet, on another
sheet and on a chart sheet. ``charts_answers.json`` records each chart's
series formulas and title as Excel's object model reported them, and every
reference in each chart part as Excel wrote it: in the file as built, and
in the file Excel saved after each of nine edits. Each edit here is held to
the file Excel wrote after the same edit. The object model cannot be: it
reports a deleted reference by its old address until the file is reopened,
and then refuses to report the series at all.
"""

from __future__ import annotations

import json
import shutil
from collections.abc import Callable
from pathlib import Path

import pytest

from pyofficeeditor.excel import Chart, ChartSheet, Workbook

Answers = dict[str, dict[str, object]]

#: Each edit Excel made, made the same way here.
EDITS: dict[str, Callable[[Workbook], object]] = {
    "insert rows 3:4": lambda book: book["Data"].insert_rows(3, 2),
    "insert row 1": lambda book: book["Data"].insert_rows(1, 1),
    "delete row 4": lambda book: book["Data"].delete_rows(4, 1),
    "delete rows 2:6": lambda book: book["Data"].delete_rows(2, 5),
    "delete row 1": lambda book: book["Data"].delete_rows(1, 1),
    "insert column B": lambda book: book["Data"].insert_columns(2, 1),
    "delete column C": lambda book: book["Data"].delete_columns(3, 1),
    "delete column A": lambda book: book["Data"].delete_columns(1, 1),
    "rename Data": lambda book: book.rename_sheet("Data", "Q1 Data"),
}


@pytest.fixture(scope="module")
def answers(live_charts_answers: Path) -> Answers:
    return json.loads(live_charts_answers.read_text(encoding="utf-8"))


@pytest.fixture
def book(tmp_path: Path, live_charts_xlsx: Path) -> Workbook:
    target = tmp_path / "charts.xlsx"
    shutil.copy(live_charts_xlsx, target)
    return Workbook.open(target)


def described(chart: Chart) -> dict[str, object]:
    """A chart as the builder's VBA reports it: the series formulas, and
    the title as ``ChartTitle.Formula`` gives it, a linked one with its
    ``=``."""
    title = chart.title or "" if chart.title_reference is None else "=" + chart.title_reference
    return {"series": [series.formula for series in chart.series], "title": title}


def charts(book: Workbook) -> dict[str, Chart]:
    """Every chart in the workbook, keyed as the answers key them."""
    found: dict[str, Chart] = {}
    for sheet in book.sheets:
        for chart in sheet.charts:
            found[f"{sheet.name}/{chart.name}"] = chart
    for chart_sheet in book.chart_sheets:
        chart = chart_sheet.chart
        assert chart is not None
        found[chart_sheet.name] = chart
    return found


def references(book: Workbook, state: str) -> Answers:
    return {f"{state}|{key}": {"references": list(chart.references)} for key, chart in charts(book).items()}


def expected(answers: Answers, state: str) -> Answers:
    return {key: value for key, value in answers.items() if key.startswith(f"{state}|")}


class TestReadingExcelsOwn:
    def test_every_chart_as_excel_reported_it(self, book: Workbook, answers: Answers) -> None:
        reported = {f"reported base|{key}": described(chart) for key, chart in charts(book).items()}
        assert reported == expected(answers, "reported base")

    def test_every_reference_as_excel_wrote_it(self, book: Workbook, answers: Answers) -> None:
        assert references(book, "base") == expected(answers, "base")

    def test_what_kind_each_is(self, book: Workbook) -> None:
        kinds = {chart.name: chart.kinds for sheet in book.sheets for chart in sheet.charts}
        assert kinds == {"Columns": ("column",), "Line": ("line",), "Pie": ("pie",), "Scatter": ("scatter",)}
        bars = book.chart_sheet("Bars").chart
        assert bars is not None and bars.kinds == ("bar",)

    def test_a_series_keeps_the_name_it_last_read(self, book: Workbook) -> None:
        (columns, _) = book["Data"].charts
        assert [series.name for series in columns.series] == ["Sales", "Costs"]


class TestEditsMoveChartsAsExcelDoes:
    @pytest.mark.parametrize("edit", list(EDITS))
    def test_every_reference_after(self, book: Workbook, answers: Answers, edit: str) -> None:
        EDITS[edit](book)
        book.save()
        assert book.path is not None
        assert references(Workbook.open(book.path), edit) == expected(answers, edit)


class TestChartSheets:
    def test_it_is_a_tab_but_not_a_worksheet(self, book: Workbook) -> None:
        assert book.sheet_names == ["Data", "Bars", "Report"]
        assert [sheet.name for sheet in book.sheets] == ["Data", "Report"]
        assert [sheet.name for sheet in book.chart_sheets] == ["Bars"]
        assert "Bars" not in book
        assert len(book) == 2
        assert book[1].name == "Report", "counted among the worksheets"
        with pytest.raises(KeyError, match="chart sheet"):
            book["Bars"]

    def test_the_active_sheet_can_be_one(self, book: Workbook) -> None:
        book.active = "Bars"
        assert isinstance(book.active, ChartSheet)
        book.active = "Data"
        assert book.active.name == "Data"

    def test_it_can_be_renamed(self, book: Workbook) -> None:
        sheet = book.rename_chart_sheet("Bars", "Totals")
        assert (sheet.name, book.sheet_names) == ("Totals", ["Data", "Totals", "Report"])
        book.save()
        assert book.path is not None
        assert Workbook.open(book.path).chart_sheet("Totals").chart is not None

    def test_it_can_be_moved(self, book: Workbook) -> None:
        book.move_sheet("Bars", 2)
        assert book.sheet_names == ["Data", "Report", "Bars"]

    def test_removing_it_takes_its_drawing_and_chart(self, book: Workbook) -> None:
        """Nothing else used either, and Excel writes neither out once the
        sheet is gone."""
        chart = book.chart_sheet("Bars").chart
        assert chart is not None
        package = book.package
        drawings = [r.target_part for r in package.relationships(book.chart_sheet("Bars").part_name)]
        assert drawings
        book.remove_sheet("Bars")
        assert book.sheet_names == ["Data", "Report"]
        for part in [*drawings, chart.part_name]:
            assert not package.has_part(part), part
        assert [chart.name for chart in book["Report"].charts] == ["Pie", "Scatter"]


class TestRemovingASheetWithCharts:
    def test_its_charts_go_and_the_rest_stay(self, book: Workbook) -> None:
        report = [chart.part_name for chart in book["Report"].charts]
        data = [chart.part_name for chart in book["Data"].charts]
        book.remove_sheet("Report")
        package = book.package
        assert not any(package.has_part(part) for part in report)
        assert all(package.has_part(part) for part in data)
        book.save()
        assert book.path is not None
        again = Workbook.open(book.path)
        assert [chart.name for chart in again["Data"].charts] == ["Columns", "Line"]
