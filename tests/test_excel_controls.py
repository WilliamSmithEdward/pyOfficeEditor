"""What a Forms-toolbar control is wired to.

``controls.xlsm`` was built by Excel with one of every control that has
something to say, wired to cells and ranges, and ``controls_answers.json``
beside it is what Excel's own object model then answered for each. The
reader is held to those answers rather than to a reading of the markup.

The measuring is the point. The current value is spelled three different
ways and none of them is a plain ``val``:

- a check box and an option button write ``checked="Checked"`` or
  ``checked="Mixed"``, and write nothing at all when they are off
- a drop down and a list box write ``sel``, the 1-based selection, and
  leave ``val`` at 0
- a spinner and a scroll bar write ``val``

So reading ``val`` alone answers 0 for a ticked box and 0 for a drop down
with the second item chosen. Both look right. Both are wrong. And the off
state is not 0 either: Excel answers -4146.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyofficeeditor import XmlDocument
from pyofficeeditor.excel import (
    XL_MIXED,
    XL_OFF,
    XL_ON,
    FormControl,
    Workbook,
    Worksheet,
)
from pyofficeeditor.excel._shapes import read_control


@pytest.fixture(scope="module")
def sheet(live_controls_xlsm: Path) -> Worksheet:
    return Workbook.open(live_controls_xlsm)["Controls"]


@pytest.fixture(scope="module")
def answers(live_controls_answers: Path) -> dict[str, dict[str, object]]:
    return json.loads(live_controls_answers.read_text(encoding="utf-8"))


def control_of(sheet: Worksheet, name: str) -> FormControl:
    found = sheet.shape(name).control
    assert found is not None, f"{name} should carry control details"
    return found


class TestWhatTheSheetHolds:
    def test_every_control_is_found(self, sheet: Worksheet) -> None:
        named = {shape.name for shape in sheet.shapes}
        assert named == {
            "Go",
            "Tick",
            "Untick",
            "Part",
            "Pick",
            "Many",
            "First",
            "Second",
            "Step",
            "Slide",
            "Set",
            "Caption",
            "Plain",
        }

    def test_an_ordinary_shape_carries_no_control(self, sheet: Worksheet) -> None:
        """The sheet has an AutoShape on it too, so the reader has to tell
        the two apart rather than treating every shape as a control."""
        plain = sheet.shape("Plain")
        assert plain.control is None
        assert plain.kind == "shape"

    def test_a_control_is_a_form_control(self, sheet: Worksheet) -> None:
        assert sheet.shape("Tick").kind == "formControl"
        assert sheet.shape("Tick").mso_type == 8


class TestAgainstExcelsOwnAnswers:
    """Every field, against what Excel reported, with one documented
    divergence."""

    def test_the_linked_cell(
        self, sheet: Worksheet, answers: dict[str, dict[str, object]]
    ) -> None:
        for name, expected in answers.items():
            if name in ("Plain", "Second"):
                continue
            assert control_of(sheet, name).linked_cell == expected["linked_cell"], name

    def test_the_list_range(
        self, sheet: Worksheet, answers: dict[str, dict[str, object]]
    ) -> None:
        for name, expected in answers.items():
            if name == "Plain":
                continue
            assert control_of(sheet, name).list_range == expected["list_range"], name

    def test_the_value(
        self, sheet: Worksheet, answers: dict[str, dict[str, object]]
    ) -> None:
        """Excel answers nothing for a kind with no value, which is
        recorded as an empty string; those carry 0 here."""
        for name, expected in answers.items():
            if name == "Plain":
                continue
            wanted = expected["value"]
            assert control_of(sheet, name).value == (wanted if wanted != "" else 0), name

    def test_an_option_button_group_stores_its_link_once(
        self, sheet: Worksheet, answers: dict[str, dict[str, object]]
    ) -> None:
        """The one place the reader and the object model disagree, on
        purpose.

        Excel answers ``$D$10`` for both option buttons. Only the first
        one's part holds it; the second's carries no ``fmlaLink`` at all,
        and inventing the group's cell for it would mean guessing what
        makes a group.
        """
        assert answers["Second"]["linked_cell"] == "$D$10"

        first = control_of(sheet, "First")
        second = control_of(sheet, "Second")
        assert first.linked_cell == "$D$10"
        assert first.first_button is True
        assert second.linked_cell == ""
        assert second.first_button is False


class TestTheValueIsNotOneAttribute:
    """The three spellings, each pinned to the kind that uses it."""

    def test_a_ticked_box_says_checked_not_val(self, sheet: Worksheet) -> None:
        tick = control_of(sheet, "Tick")
        assert tick.kind == "CheckBox"
        assert tick.value == XL_ON
        assert tick.checked is True

    def test_an_unticked_box_writes_nothing_and_is_not_zero(
        self, sheet: Worksheet
    ) -> None:
        """Absence is the off state, and off is -4146 rather than 0."""
        untick = control_of(sheet, "Untick")
        assert untick.value == XL_OFF == -4146
        assert untick.checked is False

    def test_the_third_state(self, sheet: Worksheet) -> None:
        part = control_of(sheet, "Part")
        assert part.value == XL_MIXED == 2
        assert part.checked is None, "mixed is neither ticked nor unticked"

    def test_a_drop_down_uses_sel_and_leaves_val_at_zero(
        self, sheet: Worksheet
    ) -> None:
        """The case that catches a reader looking at ``val``: the file says
        ``sel="2" val="0"`` and the answer is 2."""
        pick = control_of(sheet, "Pick")
        assert pick.kind == "Drop"
        assert pick.value == 2
        assert pick.list_range == "$A$1:$A$3"

    def test_a_spinner_uses_val(self, sheet: Worksheet) -> None:
        assert control_of(sheet, "Step").value == 7

    def test_a_scroll_bar_uses_val(self, sheet: Worksheet) -> None:
        assert control_of(sheet, "Slide").value == 25

    def test_a_kind_with_no_value_reads_zero(self, sheet: Worksheet) -> None:
        for name in ("Go", "Set", "Caption"):
            assert control_of(sheet, name).value == 0, name
            assert control_of(sheet, name).checked is None, name


class TestTheKind:
    def test_excels_own_spelling_is_kept(self, sheet: Worksheet) -> None:
        """Not translated: what is read is what can be written back."""
        assert {name: control_of(sheet, name).kind for name in
                ("Go", "Tick", "Pick", "Many", "First", "Step", "Slide", "Set", "Caption")} == {
            "Go": "Button",
            "Tick": "CheckBox",
            "Pick": "Drop",
            "Many": "List",
            "First": "Radio",
            "Step": "Spin",
            "Slide": "Scroll",
            "Set": "GBox",
            "Caption": "Label",
        }


class TestWhereItComesFrom:
    def test_the_part_and_relationship_are_carried(self, sheet: Worksheet) -> None:
        """So an edit can land back in the part it was read from."""
        tick = control_of(sheet, "Tick")
        assert tick.part_name.startswith("xl/ctrlProps/")
        assert tick.relationship.startswith("rId")

    def test_each_control_has_its_own_part(self, sheet: Worksheet) -> None:
        parts = [
            control_of(sheet, shape.name).part_name
            for shape in sheet.shapes
            if shape.control is not None
        ]
        assert len(parts) == len(set(parts)) == 12

    def test_the_macro_still_comes_from_the_sheet(self, sheet: Worksheet) -> None:
        """A control's macro is on the sheet's own ``<control>``, not in
        the control part and not on the drawing shape.

        The bracketed number indexes the workbook holding the procedure
        and is not fixed across files: this fixture says ``[0]`` where
        ``shapes.xlsm`` says ``[1]``. Only the name after it is stable.
        """
        macro = sheet.shape("Go").macro
        assert macro.endswith("!Clicked")
        assert macro.startswith("[")
        assert control_of(sheet, "Go").linked_cell == ""


class TestReadingAPartDirectly:
    """The reader on its own, without a workbook around it."""

    def test_an_unknown_kind_reads_no_value(self) -> None:
        root = XmlDocument.parse(b'<formControlPr objectType="Nonesuch" val="9"/>').root
        assert read_control(root).value == 0

    def test_a_missing_checked_attribute_is_off(self) -> None:
        root = XmlDocument.parse(b'<formControlPr objectType="CheckBox"/>').root
        assert read_control(root).value == XL_OFF

    def test_an_explicit_unchecked_is_also_off(self) -> None:
        """Excel omits the attribute, but the schema allows spelling it."""
        root = XmlDocument.parse(
            b'<formControlPr objectType="CheckBox" checked="Unchecked"/>'
        ).root
        assert read_control(root).value == XL_OFF

    def test_a_junk_value_does_not_raise(self) -> None:
        root = XmlDocument.parse(
            b'<formControlPr objectType="Spin" val="not a number"/>'
        ).root
        assert read_control(root).value == 0
