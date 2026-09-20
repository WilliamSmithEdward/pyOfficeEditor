"""Conditional formatting rules, and the ranges they cover.

The XML quoted here is Excel's own, captured by driving real Excel through
pyvbaharness and reading the part back. The compatibility formulas in
particular are copied rather than reconstructed, because Excel's are the
only ones that work: see the module docstring of ``_conditional``.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyofficeeditor._xml import XmlDocument
from pyofficeeditor.excel import Workbook, Worksheet
from pyofficeeditor.excel._conditional import (
    ColorScale,
    ConditionalFormatting,
    DataBar,
    IconSet,
    average,
    bar,
    begins_with,
    cell_is,
    compatibility_formula,
    contains_text,
    duplicates,
    during,
    ends_with,
    expression,
    gradient,
    icons,
    is_blank,
    is_error,
    not_contains_text,
    parse_sqref,
    top,
    uniques,
)
from pyofficeeditor.excel._dxf import Dxf
from pyofficeeditor.excel._reference import CellRef

#: Excel's own output, one per rule family.
EXCEL_BLOCKS: dict[str, bytes] = {
    "cellIs": (
        b'<conditionalFormatting sqref="A1:A10"><cfRule type="cellIs" dxfId="7" '
        b'priority="1" stopIfTrue="1" operator="greaterThan"><formula>50</formula>'
        b"</cfRule></conditionalFormatting>"
    ),
    "between": (
        b'<conditionalFormatting sqref="B1:B10"><cfRule type="cellIs" dxfId="6" '
        b'priority="2" stopIfTrue="1" operator="between"><formula>20</formula>'
        b"<formula>50</formula></cfRule></conditionalFormatting>"
    ),
    "colorScale": (
        b'<conditionalFormatting sqref="C1:C10"><cfRule type="colorScale" priority="3">'
        b'<colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/>'
        b'<cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/>'
        b'<color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>'
    ),
    "iconSet": (
        b'<conditionalFormatting sqref="G1:G10"><cfRule type="iconSet" priority="10">'
        b'<iconSet iconSet="3Symbols2" showValue="0" reverse="1">'
        b'<cfvo type="percent" val="0"/><cfvo type="percent" val="33"/>'
        b'<cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>'
    ),
    "top10": (
        b'<conditionalFormatting sqref="B1:B10"><cfRule type="top10" dxfId="8" '
        b'priority="2" percent="1" bottom="1" rank="20"/></conditionalFormatting>'
    ),
    "aboveAverage": (
        b'<conditionalFormatting sqref="E1:E10"><cfRule type="aboveAverage" dxfId="3" '
        b'priority="7" stopIfTrue="1" aboveAverage="0"/><cfRule type="aboveAverage" '
        b'dxfId="2" priority="8" stopIfTrue="1" aboveAverage="0" stdDev="2"/>'
        b"</conditionalFormatting>"
    ),
    "containsText": (
        b'<conditionalFormatting sqref="G1:G10"><cfRule type="containsText" dxfId="4" '
        b'priority="7" stopIfTrue="1" operator="containsText" text="text 1">'
        b'<formula>NOT(ISERROR(SEARCH("text 1",G1)))</formula></cfRule>'
        b"</conditionalFormatting>"
    ),
    "notContainsText": (
        b'<conditionalFormatting sqref="C1:C10"><cfRule type="notContainsText" dxfId="5" '
        b'priority="5" stopIfTrue="1" operator="notContains" text="zz">'
        b'<formula>ISERROR(SEARCH("zz",C1))</formula></cfRule></conditionalFormatting>'
    ),
    "timePeriod": (
        b'<conditionalFormatting sqref="A3"><cfRule type="timePeriod" dxfId="14" '
        b'priority="3" stopIfTrue="1" timePeriod="last7Days">'
        b"<formula>AND(TODAY()-FLOOR(A3,1)&lt;=6,FLOOR(A3,1)&lt;=TODAY())</formula>"
        b"</cfRule></conditionalFormatting>"
    ),
    "multiArea": (
        b'<conditionalFormatting sqref="K3:K5 M8:M9"><cfRule type="containsText" '
        b'dxfId="0" priority="17" stopIfTrue="1" operator="containsText" text="qq">'
        b'<formula>NOT(ISERROR(SEARCH("qq",K3)))</formula></cfRule>'
        b"</conditionalFormatting>"
    ),
    #: The one that carries an x14 twin, linked by GUID.
    "dataBar": (
        b'<conditionalFormatting sqref="D1:D10"><cfRule type="dataBar" priority="4">'
        b'<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>'
        b'<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" '
        b'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
        b"<x14:id>{50FAB905-7736-4DB6-B664-8F181147ABB9}</x14:id></ext></extLst>"
        b"</cfRule></conditionalFormatting>"
    ),
}


def parse(raw: bytes) -> ConditionalFormatting:
    return ConditionalFormatting.read(XmlDocument.parse(raw).root)


class TestRoundTrip:
    @pytest.mark.parametrize("name", sorted(EXCEL_BLOCKS))
    def test_excel_output_comes_back_byte_for_byte(self, name: str) -> None:
        raw = EXCEL_BLOCKS[name]
        assert parse(raw).write().to_xml() == raw.decode()

    @pytest.mark.parametrize("name", sorted(EXCEL_BLOCKS))
    def test_parsing_is_stable(self, name: str) -> None:
        once = parse(EXCEL_BLOCKS[name])
        assert parse(once.write().to_xml().encode()) == once

    def test_a_data_bars_extension_survives(self) -> None:
        """The ``x14:id`` names the twin rule in the sheet's extLst. Dropping
        it leaves that twin pointing at a rule with no id."""
        rule = parse(EXCEL_BLOCKS["dataBar"]).rules[0]
        assert rule.extension is not None
        assert "50FAB905-7736-4DB6-B664-8F181147ABB9" in rule.extension

    def test_attribute_order_follows_the_schema(self) -> None:
        """``CT_CfRule`` declares aboveAverage before stdDev, and Excel
        writes them that way."""
        rendered = parse(EXCEL_BLOCKS["aboveAverage"]).write().to_xml()
        assert rendered.index("aboveAverage=") < rendered.index("stdDev=")


class TestReading:
    def test_a_comparison(self) -> None:
        rule = parse(EXCEL_BLOCKS["cellIs"]).rules[0]
        assert rule.kind == "cellIs"
        assert rule.operator == "greaterThan"
        assert rule.formulas == ("50",)
        assert rule.dxf_id == 7
        assert rule.stop_if_true is True

    def test_two_operands(self) -> None:
        assert parse(EXCEL_BLOCKS["between"]).rules[0].formulas == ("20", "50")

    def test_below_average_is_an_attribute_set_to_zero(self) -> None:
        """``aboveAverage="0"`` is what makes it a below-average rule. A
        reader treating the attribute as a flag that is only ever true would
        report the opposite rule."""
        rules = parse(EXCEL_BLOCKS["aboveAverage"]).rules
        assert rules[0].above_average is False
        assert rules[1].std_dev == 2

    def test_top10_carries_its_modifiers(self) -> None:
        rule = parse(EXCEL_BLOCKS["top10"]).rules[0]
        assert (rule.rank, rule.percent, rule.bottom) == (20, True, True)

    def test_a_colour_scale(self) -> None:
        scale = parse(EXCEL_BLOCKS["colorScale"]).rules[0].color_scale
        assert scale is not None
        assert [value.kind for value in scale.values] == ["min", "percentile", "max"]
        assert len(scale.colors) == 3

    def test_an_icon_set(self) -> None:
        icon = parse(EXCEL_BLOCKS["iconSet"]).rules[0].icon_set
        assert icon is not None
        assert icon.icons == "3Symbols2"
        assert icon.show_value is False
        assert icon.reverse is True

    def test_two_rules_in_one_block(self) -> None:
        assert len(parse(EXCEL_BLOCKS["aboveAverage"]).rules) == 2


class TestSqref:
    def test_one_range(self) -> None:
        assert [b.a1 for b in parse_sqref("A1:A10")] == ["A1:A10"]

    def test_several_are_space_separated(self) -> None:
        assert [b.a1 for b in parse_sqref("K3:K5 M8:M9")] == ["K3:K5", "M8:M9"]

    def test_a_single_cell(self) -> None:
        assert [b.a1 for b in parse_sqref("A3")] == ["A3"]

    def test_it_round_trips(self) -> None:
        assert parse(EXCEL_BLOCKS["multiArea"]).sqref == "K3:K5 M8:M9"

    def test_the_anchor_is_the_first_areas_corner(self) -> None:
        """Not the bounding box's. ``K3:K5 M8:M9`` anchors at K3, and the
        formula Excel wrote says so."""
        assert parse(EXCEL_BLOCKS["multiArea"]).anchor.a1 == "K3"


class TestCompatibilityFormula:
    """Excel's templates, verbatim. A rule written without one is present,
    valid, counted by ``FormatConditions.Count``, and never fires."""

    @pytest.mark.parametrize(
        ("kind", "text", "expected"),
        [
            ("containsText", "text 1", 'NOT(ISERROR(SEARCH("text 1",G1)))'),
            ("notContainsText", "zz", 'ISERROR(SEARCH("zz",G1))'),
            ("beginsWith", "word", 'LEFT(G1,LEN("word"))="word"'),
            ("endsWith", "3", 'RIGHT(G1,LEN("3"))="3"'),
            ("containsBlanks", None, "LEN(TRIM(G1))=0"),
            ("notContainsBlanks", None, "LEN(TRIM(G1))>0"),
            ("containsErrors", None, "ISERROR(G1)"),
            ("notContainsErrors", None, "NOT(ISERROR(G1))"),
        ],
    )
    def test_templates(self, kind: str, text: str | None, expected: str) -> None:
        anchor = CellRef.parse("G1")
        assert compatibility_formula(kind, anchor, text=text, time_period=None) == expected

    @pytest.mark.parametrize(
        ("period", "expected"),
        [
            ("today", "FLOOR(A1,1)=TODAY()"),
            ("yesterday", "FLOOR(A1,1)=TODAY()-1"),
            ("tomorrow", "FLOOR(A1,1)=TODAY()+1"),
            ("last7Days", "AND(TODAY()-FLOOR(A1,1)<=6,FLOOR(A1,1)<=TODAY())"),
            ("thisMonth", "AND(MONTH(A1)=MONTH(TODAY()),YEAR(A1)=YEAR(TODAY()))"),
        ],
    )
    def test_time_periods(self, period: str, expected: str) -> None:
        anchor = CellRef.parse("A1")
        assert compatibility_formula("timePeriod", anchor, text=None, time_period=period) == expected

    def test_every_period_has_one(self) -> None:
        anchor = CellRef.parse("A1")
        for period in (
            "today",
            "yesterday",
            "tomorrow",
            "last7Days",
            "thisWeek",
            "lastWeek",
            "nextWeek",
            "thisMonth",
            "lastMonth",
            "nextMonth",
        ):
            assert compatibility_formula("timePeriod", anchor, text=None, time_period=period)

    def test_a_quote_is_doubled_not_xml_escaped(self) -> None:
        """Two escapings, in two places. The ``text`` attribute is XML and
        gets ``&quot;``; the formula is a string literal and doubles it."""
        formula = compatibility_formula(
            "containsText", CellRef.parse("H2"), text='a"b', time_period=None
        )
        assert formula == 'NOT(ISERROR(SEARCH("a""b",H2)))'

    def test_a_comparison_needs_none(self) -> None:
        anchor = CellRef.parse("A1")
        assert compatibility_formula("cellIs", anchor, text=None, time_period=None) is None
        assert compatibility_formula("colorScale", anchor, text=None, time_period=None) is None


class TestAnchoring:
    def test_the_formula_follows_the_range(self) -> None:
        rule = contains_text("zz")
        assert rule.anchored_at(CellRef.parse("I7")).formulas == (
            'NOT(ISERROR(SEARCH("zz",I7)))',
        )

    def test_a_condition_formula_is_left_alone(self) -> None:
        """For cellIs and expression the formula *is* the rule. Rebuilding
        it would throw the condition away."""
        rule = cell_is("greaterThan", 50)
        assert rule.anchored_at(CellRef.parse("Z99")).formulas == ("50",)

    def test_an_expression_keeps_its_own(self) -> None:
        rule = expression("=$A1>5")
        assert rule.anchored_at(CellRef.parse("D4")).formulas == ("$A1>5",)


class TestConstructors:
    def test_cell_is_quotes_a_string(self) -> None:
        """A bare word in a ``<formula>`` reads as a defined name, not text."""
        assert cell_is("equal", "North").formulas == ('"North"',)

    def test_cell_is_leaves_a_number_bare(self) -> None:
        assert cell_is("greaterThan", 50).formulas == ("50",)

    def test_cell_is_between_needs_two(self) -> None:
        assert cell_is("between", 20, 50).formulas == ("20", "50")
        with pytest.raises(ValueError, match="needs two values"):
            cell_is("between", 20)

    def test_a_formula_operand_passes_through(self) -> None:
        assert cell_is("greaterThan", "=AVERAGE($B$1:$B$9)").formulas == (
            "AVERAGE($B$1:$B$9)",
        )

    def test_the_text_operator_differs_from_the_type(self) -> None:
        """``notContainsText`` carries ``operator="notContains"``."""
        rule = not_contains_text("zz")
        assert (rule.kind, rule.operator) == ("notContainsText", "notContains")

    def test_text_rules(self) -> None:
        assert contains_text("a").kind == "containsText"
        assert begins_with("a").kind == "beginsWith"
        assert ends_with("a").kind == "endsWith"

    def test_blanks_and_errors(self) -> None:
        assert is_blank().kind == "containsBlanks"
        assert is_blank(blank=False).kind == "notContainsBlanks"
        assert is_error().kind == "containsErrors"
        assert is_error(error=False).kind == "notContainsErrors"

    def test_duplicates_and_uniques(self) -> None:
        assert duplicates().kind == "duplicateValues"
        assert uniques().kind == "uniqueValues"

    def test_top_and_bottom(self) -> None:
        assert top(3).rank == 3
        assert top(20, percent=True, bottom=True).percent is True
        with pytest.raises(ValueError, match="at least 1"):
            top(0)

    def test_above_average_writes_nothing_and_below_writes_zero(self) -> None:
        """Absent means above, which is the default Excel relies on."""
        assert average(above=True).above_average is None
        assert average(above=False).above_average is False

    def test_standard_deviations(self) -> None:
        assert average(above=False, standard_deviations=2).std_dev == 2

    def test_during_rejects_an_invented_period(self) -> None:
        with pytest.raises(ValueError, match="not a time period"):
            during("lastFortnight")  # type: ignore[arg-type]

    def test_the_visual_rules_get_excels_defaults(self) -> None:
        assert gradient().color_scale == ColorScale.three()
        assert bar().data_bar == DataBar()
        assert icons().icon_set == IconSet.three()


class TestOnAWorksheet:
    @pytest.fixture()
    def sheet(self, live_sample_xlsx: Path) -> Worksheet:
        return Workbook.open(live_sample_xlsx)["Data"]

    def test_a_clean_sheet_has_none(self, sheet: Worksheet) -> None:
        assert sheet.conditional_formats == []

    def test_adding_one(self, sheet: Worksheet) -> None:
        sheet.add_conditional_format("B2:B5", cell_is("greaterThan", 100), dxf=Dxf.of(fill="FFC7CE"))
        blocks = sheet.conditional_formats
        assert len(blocks) == 1
        assert blocks[0].sqref == "B2:B5"
        assert blocks[0].rules[0].operator == "greaterThan"

    def test_the_dxf_lands_in_the_workbook(self, sheet: Worksheet) -> None:
        rule = sheet.add_conditional_format(
            "B2:B5", cell_is("greaterThan", 100), dxf=Dxf.of(fill="FFC7CE")
        )
        styles = sheet.workbook.styles
        assert styles is not None
        assert rule.dxf_id is not None
        assert styles.dxf(rule.dxf_id) == Dxf.of(fill="FFC7CE")

    def test_two_rules_sharing_a_format_share_the_entry(self, sheet: Worksheet) -> None:
        pink = Dxf.of(fill="FFC7CE")
        first = sheet.add_conditional_format("B2:B5", cell_is("greaterThan", 100), dxf=pink)
        second = sheet.add_conditional_format("C2:C5", cell_is("lessThan", 5), dxf=pink)
        assert first.dxf_id == second.dxf_id
        styles = sheet.workbook.styles
        assert styles is not None
        assert styles.dxf_count() == 1

    def test_priority_increases(self, sheet: Worksheet) -> None:
        """A rule added later loses to one added earlier, which is what
        Excel's own New Rule does."""
        first = sheet.add_conditional_format("B2:B5", cell_is("greaterThan", 100))
        second = sheet.add_conditional_format("C2:C5", cell_is("lessThan", 5))
        assert (first.priority, second.priority) == (1, 2)

    def test_an_explicit_priority_wins(self, sheet: Worksheet) -> None:
        rule = sheet.add_conditional_format("B2:B5", cell_is("greaterThan", 1), priority=7)
        assert rule.priority == 7

    def test_the_formula_is_built_for_where_it_lands(self, sheet: Worksheet) -> None:
        rule = sheet.add_conditional_format("G4:G9", contains_text("North"))
        assert rule.formulas == ('NOT(ISERROR(SEARCH("North",G4)))',)

    def test_several_ranges_as_one_block(self, sheet: Worksheet) -> None:
        sheet.add_conditional_format(["B2:B5", "D2:D5"], cell_is("greaterThan", 1))
        block = sheet.conditional_formats[0]
        assert block.sqref == "B2:B5 D2:D5"
        assert block.anchor.a1 == "B2"

    def test_a_space_separated_string_is_several_ranges(self, sheet: Worksheet) -> None:
        sheet.add_conditional_format("B2:B5 D2:D5", cell_is("greaterThan", 1))
        assert sheet.conditional_formats[0].sqref == "B2:B5 D2:D5"

    def test_rules_at_a_cell_come_back_in_priority_order(self, sheet: Worksheet) -> None:
        sheet.add_conditional_format("B1:B9", cell_is("greaterThan", 1), priority=5)
        sheet.add_conditional_format("B1:B9", cell_is("lessThan", 9), priority=2)
        sheet.add_conditional_format("Z1:Z9", cell_is("equal", 0), priority=1)
        found = sheet.conditional_rules_at("B2")
        assert [rule.priority for rule in found] == [2, 5]

    def test_no_rules_at_an_untouched_cell(self, sheet: Worksheet) -> None:
        sheet.add_conditional_format("B2:B5", cell_is("greaterThan", 1))
        assert sheet.conditional_rules_at("Z99") == []

    def test_the_block_lands_in_schema_order(self, sheet: Worksheet) -> None:
        """After ``mergeCells`` and before ``hyperlinks``, or Excel repairs
        the sheet."""
        sheet.add_conditional_format("B2:B5", cell_is("greaterThan", 1))
        rendered = sheet.document.to_bytes().decode()
        assert rendered.index("<mergeCells") < rendered.index("<conditionalFormatting")
        assert rendered.index("<conditionalFormatting") < rendered.index("<pageMargins")

    def test_an_empty_range_is_refused(self, sheet: Worksheet) -> None:
        with pytest.raises(ValueError, match="at least one range"):
            sheet.add_conditional_format("", cell_is("greaterThan", 1))

    def test_it_survives_a_save(self, sheet: Worksheet) -> None:
        sheet.add_conditional_format("B2:B5", contains_text("x"), dxf=Dxf.of(fill="FFC7CE"))
        book = sheet.workbook
        reopened = Workbook.from_bytes(book.to_bytes())["Data"]
        block = reopened.conditional_formats[0]
        assert block.sqref == "B2:B5"
        assert block.rules[0].formulas == ('NOT(ISERROR(SEARCH("x",B2)))',)
        styles = reopened.workbook.styles
        assert styles is not None
        assert styles.dxf(block.rules[0].dxf_id or 0) == Dxf.of(fill="FFC7CE")


class TestClearing:
    @pytest.fixture()
    def sheet(self, live_sample_xlsx: Path) -> Worksheet:
        sheet = Workbook.open(live_sample_xlsx)["Data"]
        sheet.add_conditional_format("B2:B5", cell_is("greaterThan", 1))
        sheet.add_conditional_format("D2:D9", cell_is("lessThan", 1))
        return sheet

    def test_everything(self, sheet: Worksheet) -> None:
        assert sheet.clear_conditional_formats() == 2
        assert sheet.conditional_formats == []

    def test_only_what_falls_inside(self, sheet: Worksheet) -> None:
        assert sheet.clear_conditional_formats("B1:B9") == 1
        assert [b.sqref for b in sheet.conditional_formats] == ["D2:D9"]

    def test_a_block_that_overhangs_is_left_alone(self, sheet: Worksheet) -> None:
        """Narrowing it would change what the cells outside the range show,
        which is not what "clear this range" asks for."""
        assert sheet.clear_conditional_formats("D2:D5") == 0
        assert len(sheet.conditional_formats) == 2

    def test_clearing_nothing_reports_zero(self, sheet: Worksheet) -> None:
        assert sheet.clear_conditional_formats("Z1:Z9") == 0

    def test_the_element_goes_rather_than_emptying(self, sheet: Worksheet) -> None:
        sheet.clear_conditional_formats()
        assert b"conditionalFormatting" not in sheet.document.to_bytes()
