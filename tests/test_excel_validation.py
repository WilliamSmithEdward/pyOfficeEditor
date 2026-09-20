"""Data validation.

The XML quoted here is Excel's own, captured by driving real Excel and
reading the part back. The inverted ``showDropDown`` in particular is
measured rather than inferred: the attribute name says the opposite of what
it does.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import pytest

from pyofficeeditor._xml import XmlDocument
from pyofficeeditor.excel import DataValidation, Workbook, Worksheet

#: Excel's own output, one per variant.
EXCEL: dict[str, bytes] = {
    "literalList": (
        b'<dataValidation type="list" allowBlank="1" showInputMessage="1" '
        b'showErrorMessage="1" sqref="C2:C9" xr:uid="{1DA609D7-5467-413C-A134-AD43110C3470}">'
        b'<formula1>"red,green,blue"</formula1></dataValidation>'
    ),
    #: The dropdown turned OFF. showDropDown="1" is what does that.
    "noDropdown": (
        b'<dataValidation type="list" allowBlank="1" showDropDown="1" '
        b'showInputMessage="1" showErrorMessage="1" sqref="D2:D9" '
        b'xr:uid="{A29037C7-C628-400A-A16F-4D0B6084D244}">'
        b"<formula1>$A$1:$A$2</formula1></dataValidation>"
    ),
    "messages": (
        b'<dataValidation type="whole" allowBlank="1" showInputMessage="1" '
        b'showErrorMessage="1" errorTitle="No" error="That is not one to ten." '
        b'promptTitle="Heads up" prompt="One to ten." sqref="E2:E9" '
        b'xr:uid="{5441153C-506D-41A5-A1ED-440A2B27F720}">'
        b"<formula1>1</formula1><formula2>10</formula2></dataValidation>"
    ),
    "warning": (
        b'<dataValidation type="whole" errorStyle="warning" operator="equal" '
        b'allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="F2:F9" '
        b'xr:uid="{267E381E-EA54-4C2C-A0A7-A37DBE9A7F47}">'
        b"<formula1>5</formula1></dataValidation>"
    ),
    "date": (
        b'<dataValidation type="date" operator="greaterThan" allowBlank="1" '
        b'showInputMessage="1" showErrorMessage="1" sqref="I2:I9" '
        b'xr:uid="{F5670D7F-8B82-4108-86E9-549C25DA37D8}">'
        b"<formula1>46023</formula1></dataValidation>"
    ),
    "custom": (
        b'<dataValidation type="custom" allowBlank="1" showInputMessage="1" '
        b'showErrorMessage="1" sqref="L2:L9" '
        b'xr:uid="{39E62B06-AD83-4A81-B9D7-EDB2C324368E}">'
        b"<formula1>ISNUMBER(L2)</formula1></dataValidation>"
    ),
    "multiArea": (
        b'<dataValidation type="custom" allowBlank="1" showInputMessage="1" '
        b'showErrorMessage="1" sqref="I2:I9 K2:K4" '
        b'xr:uid="{0C915CD0-3322-475E-A0AD-B9D6CE7F0146}">'
        b"<formula1>ISNUMBER($A2)</formula1></dataValidation>"
    ),
}


def parse(raw: bytes) -> DataValidation:
    return DataValidation.read(XmlDocument.parse(raw).root)


class TestRoundTrip:
    @pytest.mark.parametrize("name", sorted(EXCEL))
    def test_excel_output_comes_back_byte_for_byte(self, name: str) -> None:
        assert parse(EXCEL[name]).write().to_xml() == EXCEL[name].decode()

    @pytest.mark.parametrize("name", sorted(EXCEL))
    def test_parsing_is_stable(self, name: str) -> None:
        once = parse(EXCEL[name])
        assert parse(once.write().to_xml().encode()) == once


class TestTheDropdownIsInverted:
    """``showDropDown="1"`` hides the dropdown. Measured by setting
    ``Validation.InCellDropdown`` both ways and reading the part back: the
    list that has a dropdown writes no attribute at all."""

    def test_the_attribute_absent_means_a_dropdown(self) -> None:
        assert parse(EXCEL["literalList"]).hide_dropdown is False

    def test_the_attribute_present_means_no_dropdown(self) -> None:
        assert parse(EXCEL["noDropdown"]).hide_dropdown is True

    def test_writing_it(self) -> None:
        shown = DataValidation.any_of(["a", "b"])
        hidden = DataValidation.any_of(["a", "b"], hide_dropdown=True)
        assert "showDropDown" not in shown.write().to_xml()
        assert 'showDropDown="1"' in hidden.write().to_xml()


class TestLists:
    def test_a_typed_in_list_is_a_quoted_string(self) -> None:
        """Unquoted, Excel reads the items as cell references."""
        assert DataValidation.any_of(["red", "green"]).formula1 == '"red,green"'

    def test_a_range_list_is_bare(self) -> None:
        assert DataValidation.any_of("=$A$1:$A$2").formula1 == "$A$1:$A$2"

    def test_reading_the_items_back(self) -> None:
        assert parse(EXCEL["literalList"]).items == ("red", "green", "blue")

    def test_a_range_list_has_no_items_to_read(self) -> None:
        """They are in cells, and reading them is the caller's business."""
        assert parse(EXCEL["noDropdown"]).items is None

    def test_a_comma_in_an_item_is_refused(self) -> None:
        """The list is one comma-separated string, so Excel would read it
        as two items. Silently splitting the caller's value is worse."""
        with pytest.raises(ValueError, match="contains a comma"):
            DataValidation.any_of(["a,b", "c"])

    def test_a_quote_in_an_item_is_refused(self) -> None:
        with pytest.raises(ValueError, match="contains a quote"):
            DataValidation.any_of(['say "hi"'])


class TestDefaultsAreAbsent:
    def test_between_writes_no_operator(self) -> None:
        rendered = DataValidation.whole_number(1, 10).write().to_xml()
        assert "operator" not in rendered

    def test_another_operator_is_written(self) -> None:
        rendered = DataValidation.whole_number(5, operator="equal").write().to_xml()
        assert 'operator="equal"' in rendered

    def test_stop_writes_no_error_style(self) -> None:
        rule = DataValidation.whole_number(1, 10).with_error("no", style="stop")
        assert "errorStyle" not in rule.write().to_xml()

    def test_warning_is_written(self) -> None:
        rule = DataValidation.whole_number(1, 10).with_error("no", style="warning")
        assert 'errorStyle="warning"' in rule.write().to_xml()

    def test_a_missing_operator_reads_as_none_not_a_guess(self) -> None:
        assert parse(EXCEL["messages"]).operator is None


class TestBounds:
    def test_two_bounds_give_between(self) -> None:
        rule = DataValidation.whole_number(1, 10)
        assert (rule.operator, rule.formula1, rule.formula2) == ("between", "1", "10")

    def test_one_bound_gives_at_least(self) -> None:
        rule = DataValidation.whole_number(5)
        assert (rule.operator, rule.formula1) == ("greaterThanOrEqual", "5")

    def test_between_needs_two(self) -> None:
        with pytest.raises(ValueError, match="needs two bounds"):
            DataValidation.whole_number(1, operator="between")

    def test_no_bound_at_all(self) -> None:
        with pytest.raises(ValueError, match="at least one bound"):
            DataValidation.whole_number()

    def test_a_date_becomes_its_serial(self) -> None:
        """Which is what the file holds, and what Excel shows back as a
        date through its own object model."""
        assert DataValidation.date(dt.date(2026, 1, 1)).formula1 == "46023"

    def test_a_time_becomes_a_fraction_of_a_day(self) -> None:
        assert DataValidation.time(dt.time(9, 0)).formula1 == "0.375"

    def test_a_decimal_keeps_its_point(self) -> None:
        assert DataValidation.decimal(0.5, 9.5).formula1 == "0.5"

    def test_a_reference_as_a_bound(self) -> None:
        assert DataValidation.whole_number("=$B$5").formula1 == "$B$5"

    def test_a_custom_formula_loses_its_equals(self) -> None:
        assert DataValidation.custom("=ISNUMBER(A1)").formula1 == "ISNUMBER(A1)"


class TestMessages:
    def test_an_error(self) -> None:
        rule = DataValidation.whole_number(1, 10).with_error("nope", title="No")
        assert rule.error_message == "nope"
        assert rule.error_title == "No"
        assert rule.show_error is True

    def test_a_prompt(self) -> None:
        rule = DataValidation.whole_number(1, 10).with_prompt("one to ten", title="Hi")
        assert rule.prompt_message == "one to ten"
        assert rule.prompt_title == "Hi"

    def test_reading_them(self) -> None:
        rule = parse(EXCEL["messages"])
        assert rule.error_title == "No"
        assert rule.error_message == "That is not one to ten."
        assert rule.prompt_title == "Heads up"


class TestOnAWorksheet:
    @pytest.fixture()
    def sheet(self, live_sample_xlsx: Path) -> Worksheet:
        return Workbook.open(live_sample_xlsx)["Data"]

    def test_a_clean_sheet_has_none(self, sheet: Worksheet) -> None:
        assert sheet.data_validations == []

    def test_adding_one(self, sheet: Worksheet) -> None:
        sheet.add_data_validation("B2:B9", DataValidation.any_of(["a", "b"]))
        found = sheet.data_validations
        assert len(found) == 1
        assert found[0].sqref == "B2:B9"
        assert found[0].items == ("a", "b")

    def test_finding_the_one_on_a_cell(self, sheet: Worksheet) -> None:
        sheet.add_data_validation("B2:B9", DataValidation.whole_number(1, 10))
        assert sheet.data_validation_at("B5") is not None
        assert sheet.data_validation_at("Z99") is None

    def test_several_ranges_as_one_entry(self, sheet: Worksheet) -> None:
        sheet.add_data_validation(["B2:B5", "D2:D5"], DataValidation.whole_number(1, 10))
        assert sheet.data_validations[0].sqref == "B2:B5 D2:D5"

    def test_adding_over_an_existing_one_replaces_it(self, sheet: Worksheet) -> None:
        """A cell carries one validation. Leaving two behind makes which one
        applies depend on document order."""
        sheet.add_data_validation("B2:B9", DataValidation.whole_number(1, 10))
        sheet.add_data_validation("B2:B9", DataValidation.any_of(["a", "b"]))
        found = sheet.data_validations
        assert len(found) == 1
        assert found[0].kind == "list"

    def test_a_partial_overlap_drops_the_overlapping_area(self, sheet: Worksheet) -> None:
        sheet.add_data_validation(["B2:B9", "D2:D9"], DataValidation.whole_number(1, 10))
        sheet.add_data_validation("B5:B6", DataValidation.any_of(["a"]))
        kinds = {v.kind: v.sqref for v in sheet.data_validations}
        assert kinds["whole"] == "D2:D9", "the overlapping area went whole"
        assert kinds["list"] == "B5:B6"

    def test_a_validation_with_no_type_is_refused(self, sheet: Worksheet) -> None:
        with pytest.raises(ValueError, match="no type"):
            sheet.add_data_validation("B2:B9", DataValidation())

    def test_an_empty_range_is_refused(self, sheet: Worksheet) -> None:
        with pytest.raises(ValueError, match="at least one range"):
            sheet.add_data_validation("", DataValidation.any_of(["a"]))

    def test_the_count_follows(self, sheet: Worksheet) -> None:
        sheet.add_data_validation("B2:B9", DataValidation.any_of(["a"]))
        sheet.add_data_validation("C2:C9", DataValidation.any_of(["b"]))
        assert '<dataValidations count="2">' in sheet.document.to_bytes().decode()

    def test_the_block_lands_in_schema_order(self, sheet: Worksheet) -> None:
        sheet.add_data_validation("B2:B9", DataValidation.any_of(["a"]))
        rendered = sheet.document.to_bytes().decode()
        assert rendered.index("<mergeCells") < rendered.index("<dataValidations")
        assert rendered.index("<dataValidations") < rendered.index("<pageMargins")

    def test_it_survives_a_save(self, sheet: Worksheet) -> None:
        sheet.add_data_validation(
            "B2:B9", DataValidation.any_of(["a", "b"], hide_dropdown=True)
        )
        reopened = Workbook.from_bytes(sheet.workbook.to_bytes())["Data"]
        found = reopened.data_validations[0]
        assert found.items == ("a", "b")
        assert found.hide_dropdown is True


class TestClearing:
    @pytest.fixture()
    def sheet(self, live_sample_xlsx: Path) -> Worksheet:
        sheet = Workbook.open(live_sample_xlsx)["Data"]
        sheet.add_data_validation("B2:B9", DataValidation.whole_number(1, 10))
        sheet.add_data_validation("D2:D9", DataValidation.any_of(["a"]))
        return sheet

    def test_everything(self, sheet: Worksheet) -> None:
        assert sheet.clear_data_validations() == 2
        assert sheet.data_validations == []

    def test_the_wrapper_goes_with_the_last_entry(self, sheet: Worksheet) -> None:
        """An empty ``<dataValidations count="0"/>`` is not something Excel
        accepts."""
        sheet.clear_data_validations()
        assert b"dataValidations" not in sheet.document.to_bytes()

    def test_only_what_falls_inside(self, sheet: Worksheet) -> None:
        assert sheet.clear_data_validations("B1:B20") == 1
        assert [v.sqref for v in sheet.data_validations] == ["D2:D9"]

    def test_one_that_overhangs_is_left_alone(self, sheet: Worksheet) -> None:
        assert sheet.clear_data_validations("D2:D5") == 0
        assert len(sheet.data_validations) == 2


class TestShifting:
    """The shifting already existed before the model did, so these check the
    two stay in step."""

    @pytest.fixture()
    def sheet(self, live_sample_xlsx: Path) -> Worksheet:
        return Workbook.open(live_sample_xlsx)["Data"]

    def test_an_insertion_moves_the_range(self, sheet: Worksheet) -> None:
        sheet.add_data_validation("B2:B9", DataValidation.whole_number(1, 10))
        sheet.insert_rows(3, 2)
        assert sheet.data_validations[0].sqref == "B2:B11"

    def test_an_insertion_moves_the_bounds(self, sheet: Worksheet) -> None:
        """A bound is a formula and carries references of its own."""
        sheet.add_data_validation("B2:B9", DataValidation.whole_number("=$B$5"))
        sheet.insert_rows(3, 2)
        assert sheet.data_validations[0].formula1 == "$B$7"

    def test_a_deletion_shrinks_it(self, sheet: Worksheet) -> None:
        sheet.add_data_validation("B2:B9", DataValidation.whole_number(1, 10))
        sheet.delete_rows(3, 2)
        assert sheet.data_validations[0].sqref == "B2:B7"

    def test_a_deleted_range_takes_the_validation(self, sheet: Worksheet) -> None:
        sheet.add_data_validation("B3:B4", DataValidation.whole_number(1, 10))
        sheet.delete_rows(3, 2)
        assert sheet.data_validations == []
