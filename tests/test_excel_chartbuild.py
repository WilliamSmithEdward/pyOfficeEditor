"""Charts written as Excel's Insert Chart writes them.

``chartkinds.xlsx`` was built by Excel with ``Shapes.AddChart2`` and its
default style: one chart of each kind from ``Data!A1:C6``, and a column
chart with a title typed in. ``chartkinds_answers.json`` records each
chart's series formulas as Excel reported them. A chart added here from the
same cells is held to Excel's part byte for byte, its cache of the cells'
values included, bar the ids Excel draws at random.
"""

from __future__ import annotations

import json
import re
import shutil
import zipfile
from collections.abc import Callable
from pathlib import Path

import pytest

from pyofficeeditor.excel import CellRef, Workbook, Worksheet, _chartbuild
from pyofficeeditor.excel._chartbuild import CHART_KINDS, ChartKind

#: The colour of a chart's n-th series.
accent: Callable[[int], str] = getattr(_chartbuild, "_accent")

#: The chart Excel made for each name, in the order it made them.
MADE: list[tuple[str, ChartKind, str | None]] = [
    ("Chart 1", "column", None),
    ("Chart 2", "bar", None),
    ("Chart 3", "line", None),
    ("Chart 4", "lineMarkers", None),
    ("Chart 5", "pie", None),
    ("Chart 6", "doughnut", None),
    ("Chart 7", "scatter", None),
    ("Chart 8", "area", None),
    ("Chart 9", "column", "Sales by month"),
]


def normal(text: str) -> str:
    """A chart part with the ids Excel draws at random replaced: its axes',
    and the tail every series' id shares."""
    text = re.sub(r'<c:(axId|crossAx) val="\d+"/>', r'<c:\1 val="N"/>', text)
    return re.sub(r'val="\{([0-9A-F]{8})-[0-9A-F-]+\}"', r'val="{\1-X}"', text)


@pytest.fixture(scope="module")
def answers(live_chart_kinds_answers: Path) -> dict[str, dict[str, object]]:
    return json.loads(live_chart_kinds_answers.read_text(encoding="utf-8"))


@pytest.fixture
def data(tmp_path: Path, live_empty_xlsx: Path) -> Worksheet:
    """Excel's empty workbook with the cells ``chartkinds.xlsx`` charts."""
    target = tmp_path / "charts.xlsx"
    shutil.copy(live_empty_xlsx, target)
    book = Workbook.open(target)
    sheet = book.rename_sheet(book.sheet_names[0], "Data")
    for column, text in enumerate(["Month", "Sales", "Costs"], start=1):
        sheet.set_value(CellRef(1, column), text)
    for row in range(2, 7):
        sheet.set_value(CellRef(row, 1), f"M{row - 1}")
        sheet.set_value(CellRef(row, 2), row * 10)
        sheet.set_value(CellRef(row, 3), row * 4)
    return sheet


def saved(sheet: Worksheet) -> Path:
    sheet.workbook.save()
    path = sheet.workbook.path
    assert path is not None
    return path


class TestAsExcelWritesThem:
    def test_every_kind(self, data: Worksheet, live_chart_kinds_xlsx: Path) -> None:
        for index, (_, kind, title) in enumerate(MADE):
            data.add_chart(kind, "A1:C6", left=250, top=20 + 230 * index, title=title)
        with zipfile.ZipFile(saved(data)) as ours, zipfile.ZipFile(live_chart_kinds_xlsx) as theirs:
            for number, (name, kind, _) in enumerate(MADE, start=1):
                part = f"xl/charts/chart{number}.xml"
                assert normal(ours.read(part).decode("utf-8")) == normal(theirs.read(part).decode("utf-8")), (
                    name,
                    kind,
                )

    def test_each_reads_back_as_excel_reported_it(
        self, data: Worksheet, answers: dict[str, dict[str, object]]
    ) -> None:
        for index, (name, kind, title) in enumerate(MADE):
            chart = data.add_chart(kind, "A1:C6", left=250, top=20 + 230 * index, title=title)
            assert chart.name == name
            assert [series.formula for series in chart.series] == answers[name]["series"]
            assert chart.title == title

    def test_the_kind_it_reads_as(self, data: Worksheet) -> None:
        kinds = [data.add_chart(kind, "A1:C6", left=250, top=20).kinds for kind in CHART_KINDS]
        assert kinds == [
            ("column",), ("bar",), ("line",), ("line",), ("pie",), ("doughnut",), ("scatter",), ("area",)
        ]


class TestTheBlock:
    """Series run down a block's columns only when it is taller than it is
    wide, as Excel ran them for six blocks, measured; the tall A1:C6 of
    ``chartkinds.xlsx`` runs them down its columns."""

    @pytest.mark.parametrize(
        ("block", "series"),
        [
            ("A1:C2", ["=SERIES(Data!$A$2,Data!$B$1:$C$1,Data!$B$2:$C$2,1)"]),
            (
                "A1:C3",
                [
                    "=SERIES(Data!$A$2,Data!$B$1:$C$1,Data!$B$2:$C$2,1)",
                    "=SERIES(Data!$A$3,Data!$B$1:$C$1,Data!$B$3:$C$3,2)",
                ],
            ),
            ("A1:B2", ["=SERIES(Data!$A$2,Data!$B$1,Data!$B$2,1)"]),
        ],
    )
    def test_a_block_no_taller_than_it_is_wide_runs_them_along_its_rows(
        self, data: Worksheet, block: str, series: list[str]
    ) -> None:
        chart = data.add_chart("column", block, left=0, top=0)
        assert [one.formula for one in chart.series] == series

    def test_the_way_the_series_run_can_be_chosen(self, data: Worksheet) -> None:
        chart = data.add_chart("column", "A1:C6", left=0, top=0, series_in="rows")
        assert chart.series[0].formula == "=SERIES(Data!$A$2,Data!$B$1:$C$1,Data!$B$2:$C$2,1)"
        assert len(chart.series) == 5

    def test_a_block_on_another_sheet(self, data: Worksheet) -> None:
        report = data.workbook.add_sheet("Q1 Report")
        chart = report.add_chart("pie", "Data!A1:B6", left=0, top=0)
        assert chart.series[0].formula == "=SERIES(Data!$B$1,Data!$A$2:$A$6,Data!$B$2:$B$6,1)"
        assert [c.name for c in report.charts] == ["Chart 1"]
        assert data.charts == []

    def test_numbers_as_categories_are_kept_as_numbers(self, data: Worksheet) -> None:
        for row in range(2, 7):
            data.set_value(CellRef(row, 1), 2020 + row)
        data.add_chart("line", "A1:B6", left=0, top=0)
        with zipfile.ZipFile(saved(data)) as package:
            part = package.read("xl/charts/chart1.xml").decode("utf-8")
        assert "<c:cat><c:numRef><c:f>Data!$A$2:$A$6</c:f><c:numCache>" in part
        assert '<c:pt idx="0"><c:v>2022</c:v></c:pt>' in part

    def test_the_chart_moves_with_the_cells_it_reads(self, data: Worksheet) -> None:
        data.add_chart("column", "A1:C6", left=250, top=20)
        data.insert_rows(1, 2)
        (chart,) = data.charts
        assert chart.series[0].formula == "=SERIES(Data!$B$3,Data!$A$4:$A$8,Data!$B$4:$B$8,1)"

    def test_it_survives_a_save(self, data: Worksheet) -> None:
        data.add_chart("scatter", "A1:C6", left=250, top=20, name="Spread")
        (chart,) = Workbook.open(saved(data))["Data"].charts
        assert (chart.name, chart.kinds, len(chart.series)) == ("Spread", ("scatter",), 2)


class TestRefusals:
    def test_a_kind_it_does_not_add(self, data: Worksheet) -> None:
        with pytest.raises(ValueError, match="not a kind"):
            data.add_chart("radar", "A1:C6", left=0, top=0)  # type: ignore[arg-type]

    @pytest.mark.parametrize("block", ["A1:A6", "A1:C1", "B2"])
    def test_a_block_without_names_categories_and_values(self, data: Worksheet, block: str) -> None:
        with pytest.raises(ValueError, match="needs a row of series names"):
            data.add_chart("column", block, left=0, top=0)

    def test_no_size(self, data: Worksheet) -> None:
        with pytest.raises(ValueError, match="size above nothing"):
            data.add_chart("column", "A1:C6", left=0, top=0, width=0)

    def test_a_name_the_sheet_has(self, data: Worksheet) -> None:
        data.add_chart("column", "A1:C6", left=0, top=0)
        with pytest.raises(ValueError, match="already has a shape named 'Chart 1'"):
            data.add_chart("line", "A1:C6", left=0, top=0, name="Chart 1")


class TestColours:
    """As Excel coloured twenty series of one chart, measured."""

    def test_the_first_six_series_take_the_six_accents(self) -> None:
        assert [accent(index) for index in range(6)] == [f'<a:schemeClr val="accent{n}"/>' for n in range(1, 7)]

    def test_each_round_after_takes_the_palettes_next_variation(self) -> None:
        assert accent(6) == '<a:schemeClr val="accent1"><a:lumMod val="60000"/></a:schemeClr>'
        assert accent(11) == '<a:schemeClr val="accent6"><a:lumMod val="60000"/></a:schemeClr>'
        assert accent(12) == '<a:schemeClr val="accent1"><a:lumMod val="80000"/><a:lumOff val="20000"/></a:schemeClr>'
        assert accent(18) == '<a:schemeClr val="accent1"><a:lumMod val="80000"/></a:schemeClr>'
        assert accent(19) == '<a:schemeClr val="accent2"><a:lumMod val="80000"/></a:schemeClr>'
