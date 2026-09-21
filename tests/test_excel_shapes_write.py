"""Adding, removing and rewiring shapes.

Every structural rule here was measured by writing a workbook and asking
Excel to open it, because Excel refuses a package it disagrees with rather
than repairing it, and a refused file is indistinguishable from a good one
until you try. Four of those rules cost a refusal each to find:

- a control's anchor writes ``<from><xdr:col>``: the wrapper loses the
  prefix and its children keep it
- a connector is ``prst="line"``, not the ``rect`` every other shape gets
- the VML spells a tick box ``Checkbox`` where its own part spells it
  ``CheckBox``
- a sheet that has never held a control declares neither ``xdr`` nor
  ``x14``, and the control's markup needs both

The live gate in ``test_excel_live_gate.py`` is what proves Excel still
accepts these. These tests pin the bytes so a regression names itself
without needing Office.
"""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

import pytest

from pyofficeeditor.excel import FormControl, Workbook
from pyofficeeditor.excel._shapes import (
    CONTROL_KINDS,
    FIRST_CONTROL_ID,
    RANGE_DEFAULTS,
    XL_MIXED,
    XL_OFF,
    XL_ON,
    Shape,
    SheetGrid,
    check_control_kind,
    control_properties,
    control_range,
    control_vml,
    corner_markup,
    new_anchor,
    qualified_macro,
    set_vml_macro,
    vml_id,
)


@pytest.fixture
def book(tmp_path: Path, live_controls_xlsm: Path) -> Workbook:
    """A copy of the Excel-authored controls workbook, per test."""
    target = tmp_path / "controls.xlsm"
    shutil.copy(live_controls_xlsm, target)
    return Workbook.open(target)


@pytest.fixture
def bare(tmp_path: Path, live_sample_xlsx: Path) -> Workbook:
    """A workbook whose sheet has no drawing, no VML and no controls."""
    target = tmp_path / "sample.xlsx"
    shutil.copy(live_sample_xlsx, target)
    return Workbook.open(target)


def parts_of(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def dangling(path: Path) -> list[str]:
    """Relationships pointing at a part the package does not hold.

    The failure that stays invisible until Excel opens the file, and then
    costs the whole workbook rather than the one shape.
    """
    missing: list[str] = []
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        for entry in names:
            if not entry.endswith(".rels"):
                continue
            base = entry.rsplit("_rels/", 1)[0]
            body = archive.read(entry).decode("utf-8")
            for target in re.findall(r'Target="([^"]+)"', body):
                if target.startswith(("http", "mailto", "file", "ftp")):
                    continue
                resolved = f"{base}{target}"
                while "/../" in resolved:
                    head, _, tail = resolved.partition("/../")
                    resolved = f"{head.rsplit('/', 1)[0]}/{tail}"
                if resolved not in names:
                    missing.append(f"{entry} -> {target}")
    return missing


class TestAddingADrawingShape:
    def test_it_comes_back(self, book: Workbook) -> None:
        sheet = book["Controls"]
        made = sheet.add_shape("Box", left=300, top=20, width=120, height=50)
        assert made.name == "Box"
        assert sheet.shape("Box").kind == "shape"

    def test_the_box_is_kept(self, book: Workbook) -> None:
        sheet = book["Controls"]
        sheet.add_shape("Box", left=300, top=20, width=120, height=50, text="Hi")
        found = sheet.shape("Box")
        assert (found.left, found.top) == (300.0, 20.0)
        assert (found.width, found.height) == (120.0, 50.0)
        assert found.text == "Hi"

    def test_a_line_is_not_a_rectangle(self, book: Workbook) -> None:
        """``prst="rect"`` on a connector makes Excel refuse the file."""
        sheet = book["Controls"]
        sheet.add_shape("Edge", left=10, top=10, width=90, height=40, kind="line")
        assert sheet.shape("Edge").geometry == "line"

    def test_a_named_geometry_wins(self, book: Workbook) -> None:
        sheet = book["Controls"]
        sheet.add_shape(
            "Round", left=10, top=10, width=90, height=40, geometry="roundRect"
        )
        assert sheet.shape("Round").geometry == "roundRect"

    def test_a_text_box_is_a_flag_not_an_element(self, book: Workbook) -> None:
        sheet = book["Controls"]
        sheet.add_shape("Note", left=10, top=10, width=90, height=40, kind="textBox")
        assert sheet.shape("Note").kind == "textBox"

    def test_two_lines_of_text(self, book: Workbook) -> None:
        sheet = book["Controls"]
        sheet.add_shape("Two", left=10, top=10, width=90, height=40, text="a\nb")
        assert sheet.shape("Two").text == "a\nb"

    def test_a_duplicate_name_is_refused(self, book: Workbook) -> None:
        sheet = book["Controls"]
        with pytest.raises(ValueError, match="already has a shape named"):
            sheet.add_shape("Plain", left=10, top=10, width=90, height=40)

    def test_a_blank_name_is_refused(self, book: Workbook) -> None:
        sheet = book["Controls"]
        with pytest.raises(ValueError, match="needs a name"):
            sheet.add_shape("  ", left=10, top=10, width=90, height=40)

    def test_a_control_is_refused_here(self, book: Workbook) -> None:
        """It needs three more parts than this writes."""
        sheet = book["Controls"]
        with pytest.raises(ValueError, match="add_form_control"):
            sheet.add_shape(
                "No", left=10, top=10, width=90, height=40, kind="formControl"
            )

    def test_ids_do_not_collide(self, book: Workbook) -> None:
        sheet = book["Controls"]
        first = sheet.add_shape("A", left=10, top=10, width=50, height=20)
        second = sheet.add_shape("B", left=10, top=40, width=50, height=20)
        assert first.shape_id != second.shape_id

    def test_a_drawing_shape_keeps_out_of_control_numbering(
        self, book: Workbook
    ) -> None:
        """Excel numbers controls from 1025 and drawing shapes from 2."""
        sheet = book["Controls"]
        made = sheet.add_shape("A", left=10, top=10, width=50, height=20)
        assert made.shape_id < FIRST_CONTROL_ID


class TestASheetWithNoDrawing:
    def test_the_part_is_made(self, bare: Workbook, tmp_path: Path) -> None:
        sheet = bare["Data"]
        sheet.add_shape("Box", left=100, top=20, width=120, height=50)
        bare.save()
        parts = parts_of(tmp_path / "sample.xlsx")
        assert "xl/drawings/drawing1.xml" in parts

    def test_it_is_declared_and_related(self, bare: Workbook, tmp_path: Path) -> None:
        sheet = bare["Data"]
        sheet.add_shape("Box", left=100, top=20, width=120, height=50)
        bare.save()
        parts = parts_of(tmp_path / "sample.xlsx")
        types = parts["[Content_Types].xml"].decode()
        assert "/xl/drawings/drawing1.xml" in types
        rels = parts["xl/worksheets/_rels/sheet1.xml.rels"].decode()
        assert "drawings/drawing1.xml" in rels
        assert "<drawing" in parts["xl/worksheets/sheet1.xml"].decode()

    def test_nothing_dangles(self, bare: Workbook, tmp_path: Path) -> None:
        bare["Data"].add_shape("Box", left=100, top=20, width=120, height=50)
        bare.save()
        assert dangling(tmp_path / "sample.xlsx") == []


class TestAddingAFormControl:
    def test_all_four_parts_are_written(self, bare: Workbook, tmp_path: Path) -> None:
        """Leave one out and the control is invisible, inert, or the file
        does not open."""
        bare["Data"].add_form_control(
            "Go", left=100, top=20, width=90, height=30, text="Press", macro="Run"
        )
        bare.save()
        parts = parts_of(tmp_path / "sample.xlsx")

        assert "xl/drawings/drawing1.xml" in parts, "the anchor"
        assert "xl/drawings/vmlDrawing1.vml" in parts, "what Excel draws it from"
        assert "xl/ctrlProps/ctrlProp1.xml" in parts, "what it is wired to"
        sheet = parts["xl/worksheets/sheet1.xml"].decode()
        assert "<control " in sheet, "the sheet's own record"

    def test_the_prefixes_it_uses_are_declared(
        self, bare: Workbook, tmp_path: Path
    ) -> None:
        """A sheet that never held a control declares neither, and the
        markup uses both. Either missing and Excel refuses the file."""
        bare["Data"].add_form_control("Go", left=100, top=20, width=90, height=30)
        bare.save()
        root = parts_of(tmp_path / "sample.xlsx")["xl/worksheets/sheet1.xml"].decode()
        head = root[root.find("<worksheet") : root.find(">", root.find("<worksheet"))]
        assert "xmlns:xdr=" in head
        assert "xmlns:x14=" in head

    def test_the_vml_content_type_is_a_default(
        self, bare: Workbook, tmp_path: Path
    ) -> None:
        """By extension, not an override, which is how Excel writes it."""
        bare["Data"].add_form_control("Go", left=100, top=20, width=90, height=30)
        bare.save()
        types = parts_of(tmp_path / "sample.xlsx")["[Content_Types].xml"].decode()
        assert '<Default Extension="vml"' in types

    def test_defaults_come_before_overrides(
        self, bare: Workbook, tmp_path: Path
    ) -> None:
        """Excel refuses a package where a Default follows an Override."""
        bare["Data"].add_form_control("Go", left=100, top=20, width=90, height=30)
        bare.save()
        types = parts_of(tmp_path / "sample.xlsx")["[Content_Types].xml"].decode()
        assert types.rfind("<Default") < types.find("<Override")

    def test_the_control_reads_back(self, bare: Workbook) -> None:
        sheet = bare["Data"]
        sheet.add_form_control(
            "Tick", kind="CheckBox", left=100, top=20, width=110, height=20,
            linked_cell="$H$1", value=XL_ON,
        )
        found = sheet.shape("Tick")
        assert found.kind == "formControl"
        assert found.control is not None
        assert found.control.kind == "CheckBox"
        assert found.control.linked_cell == "$H$1"
        assert found.control.checked is True

    def test_controls_are_numbered_from_1025(self, bare: Workbook) -> None:
        sheet = bare["Data"]
        made = sheet.add_form_control("Go", left=100, top=20, width=90, height=30)
        assert made.shape_id == FIRST_CONTROL_ID

    def test_a_second_control_gets_the_next_id(self, book: Workbook) -> None:
        sheet = book["Controls"]
        highest = max(
            shape.shape_id for shape in sheet.shapes if shape.shape_id >= FIRST_CONTROL_ID
        )
        made = sheet.add_form_control("New", left=10, top=10, width=90, height=30)
        assert made.shape_id == highest + 1

    def test_an_unknown_kind_is_refused(self, bare: Workbook) -> None:
        with pytest.raises(ValueError, match="not a form control"):
            bare["Data"].add_form_control(
                "Odd", kind="Nonesuch", left=10, top=10, width=90, height=30
            )

    def test_a_refused_kind_leaves_nothing_behind(
        self, bare: Workbook, tmp_path: Path
    ) -> None:
        """Building a control writes four things, and the VML is the one
        that knows the kind is wrong. Noticing it there would leave a
        control part, a relationship and an anchor for a control that was
        never made, so the kind is checked before anything is written."""
        before = parts_of(tmp_path / "sample.xlsx")
        with pytest.raises(ValueError):
            bare["Data"].add_form_control(
                "Odd", kind="Nonesuch", left=10, top=10, width=90, height=30
            )
        bare.save()
        after = parts_of(tmp_path / "sample.xlsx")
        assert set(after) - set(before) == set(), "a part was left behind"
        assert dangling(tmp_path / "sample.xlsx") == []

    def test_nothing_dangles(self, bare: Workbook, tmp_path: Path) -> None:
        sheet = bare["Data"]
        sheet.add_form_control("Go", left=100, top=20, width=90, height=30)
        sheet.add_form_control(
            "Tick", kind="CheckBox", left=100, top=60, width=110, height=20
        )
        bare.save()
        assert dangling(tmp_path / "sample.xlsx") == []


class TestRemoving:
    def test_a_drawing_shape_goes(self, book: Workbook) -> None:
        sheet = book["Controls"]
        assert sheet.shape("Plain") is not None
        sheet.remove_shape("Plain")
        assert "Plain" not in {shape.name for shape in sheet.shapes}

    def test_the_others_stay(self, book: Workbook) -> None:
        sheet = book["Controls"]
        before = [shape.name for shape in sheet.shapes]
        sheet.remove_shape("Plain")
        after = [shape.name for shape in sheet.shapes]
        assert after == [name for name in before if name != "Plain"]

    def test_a_control_takes_its_part_with_it(
        self, book: Workbook, tmp_path: Path
    ) -> None:
        sheet = book["Controls"]
        control = sheet.shape("Tick").control
        assert control is not None
        part = control.part_name
        sheet.remove_shape("Tick")
        book.save()
        assert part not in parts_of(tmp_path / "controls.xlsm")

    def test_and_its_relationship(self, book: Workbook, tmp_path: Path) -> None:
        sheet = book["Controls"]
        control = sheet.shape("Tick").control
        assert control is not None
        sheet.remove_shape("Tick")
        book.save()
        rels = parts_of(tmp_path / "controls.xlsm")[
            "xl/worksheets/_rels/sheet1.xml.rels"
        ].decode()
        assert f'Id="{control.relationship}"' not in rels

    def test_and_its_vml_shape(self, book: Workbook, tmp_path: Path) -> None:
        sheet = book["Controls"]
        shape_id = sheet.shape("Tick").shape_id
        sheet.remove_shape("Tick")
        book.save()
        vml = parts_of(tmp_path / "controls.xlsm")[
            "xl/drawings/vmlDrawing1.vml"
        ].decode()
        assert f"_x0000_s{shape_id}" not in vml

    def test_and_its_record_on_the_sheet(self, book: Workbook, tmp_path: Path) -> None:
        sheet = book["Controls"]
        shape_id = sheet.shape("Tick").shape_id
        sheet.remove_shape("Tick")
        book.save()
        text = parts_of(tmp_path / "controls.xlsm")["xl/worksheets/sheet1.xml"].decode()
        assert f'shapeId="{shape_id}"' not in text

    def test_nothing_dangles(self, book: Workbook, tmp_path: Path) -> None:
        sheet = book["Controls"]
        sheet.remove_shape("Tick")
        sheet.remove_shape("Plain")
        book.save()
        assert dangling(tmp_path / "controls.xlsm") == []

    def test_an_unknown_name_says_what_is_there(self, book: Workbook) -> None:
        with pytest.raises(KeyError, match="no shape named"):
            book["Controls"].remove_shape("Nonesuch")


class TestSettingAMacro:
    def test_on_a_drawing_shape(self, book: Workbook) -> None:
        sheet = book["Controls"]
        sheet.set_shape_macro("Plain", "Clicked")
        assert sheet.shape("Plain").macro == "Clicked"

    def test_on_a_control_it_keeps_the_prefix(self, book: Workbook) -> None:
        """The bracketed number indexes the workbook holding the
        procedure, and what decides it is not modelled here, so an
        existing one is kept rather than replaced."""
        sheet = book["Controls"]
        before = sheet.shape("Go").macro
        prefix = before.partition("!")[0]
        sheet.set_shape_macro("Go", "Renamed")
        assert sheet.shape("Go").macro == f"{prefix}!Renamed"

    def test_a_qualified_name_goes_through_untouched(self, book: Workbook) -> None:
        sheet = book["Controls"]
        sheet.set_shape_macro("Go", "[3]!Elsewhere")
        assert sheet.shape("Go").macro == "[3]!Elsewhere"

    def test_both_copies_move(self, book: Workbook, tmp_path: Path) -> None:
        """Excel reads the sheet's and ignores the VML's, measured. They
        are still written together: a stale name left in the package is a
        trap for whatever reads it next."""
        sheet = book["Controls"]
        shape_id = sheet.shape("Go").shape_id
        sheet.set_shape_macro("Go", "Renamed")
        book.save()
        parts = parts_of(tmp_path / "controls.xlsm")
        assert "Renamed" in parts["xl/worksheets/sheet1.xml"].decode()
        vml = parts["xl/drawings/vmlDrawing1.vml"].decode()
        block = vml[vml.find(f"_x0000_s{shape_id}") :]
        assert "Renamed" in block[: block.find("</v:shape>")]

    def test_clearing_one(self, book: Workbook) -> None:
        sheet = book["Controls"]
        sheet.set_shape_macro("Go", "")
        assert sheet.shape("Go").macro == ""

    def test_an_unknown_name_is_refused(self, book: Workbook) -> None:
        with pytest.raises(KeyError, match="no shape named"):
            book["Controls"].set_shape_macro("Nonesuch", "Clicked")


class TestByteFidelity:
    """Adding a shape changes the parts it has to and no others.

    ``xl/workbook.xml`` is in every expected set below because any save
    sets ``fullCalcOnLoad`` on it, so Excel recalculates rather than
    trusting the cached results still in the bytes. That is the library's
    behaviour everywhere and has nothing to do with shapes.
    """

    CALC = "xl/workbook.xml"

    def test_only_the_drawing_moves(self, book: Workbook, tmp_path: Path) -> None:
        before = parts_of(tmp_path / "controls.xlsm")
        book["Controls"].add_shape("Box", left=300, top=20, width=120, height=50)
        book.save()
        after = parts_of(tmp_path / "controls.xlsm")
        changed = {name for name in before if before[name] != after.get(name)}
        assert changed == {"xl/drawings/drawing1.xml", self.CALC}

    def test_a_macro_on_a_control_leaves_the_drawing_alone(
        self, book: Workbook, tmp_path: Path
    ) -> None:
        before = parts_of(tmp_path / "controls.xlsm")
        book["Controls"].set_shape_macro("Go", "Renamed")
        book.save()
        after = parts_of(tmp_path / "controls.xlsm")
        changed = {name for name in before if before[name] != after.get(name)}
        assert changed == {
            "xl/worksheets/sheet1.xml",
            "xl/drawings/vmlDrawing1.vml",
            self.CALC,
        }

    def test_a_macro_on_a_drawing_shape_leaves_the_sheet_alone(
        self, book: Workbook, tmp_path: Path
    ) -> None:
        """It lives on the shape's own element, not on the sheet."""
        before = parts_of(tmp_path / "controls.xlsm")
        book["Controls"].set_shape_macro("Plain", "Clicked")
        book.save()
        after = parts_of(tmp_path / "controls.xlsm")
        changed = {name for name in before if before[name] != after.get(name)}
        assert changed == {"xl/drawings/drawing1.xml", self.CALC}


class TestTheMarkupItself:
    """The four rules each of which cost a refusal to find."""

    def test_a_control_anchor_keeps_the_prefix_on_its_children(self) -> None:
        markup = corner_markup("from", SheetGrid(), 100.0, 50.0, wrapper="")
        assert markup.startswith("<from>")
        assert "<xdr:col>" in markup
        assert "<col>" not in markup

    def test_a_drawing_anchor_prefixes_both(self) -> None:
        markup = corner_markup("from", SheetGrid(), 100.0, 50.0)
        assert markup.startswith("<xdr:from>")
        assert "<xdr:col>" in markup

    def test_the_vml_spells_a_tick_box_differently(self) -> None:
        """``CheckBox`` in its own part and ``Checkbox`` in the VML."""
        shape = Shape(
            name="T", kind="formControl", shape_id=1025,
            control=FormControl(kind="CheckBox"),
        )
        assert 'objectType="CheckBox"' in control_properties(shape.control or FormControl())
        assert 'ObjectType="Checkbox"' in control_vml(shape, SheetGrid())

    def test_a_connector_gets_the_line_preset(self) -> None:
        shape = Shape(name="E", kind="line", width=90, height=40, shape_id=2)
        markup = new_anchor(shape, SheetGrid())
        assert 'prst="line"' in markup
        assert "<xdr:cxnSp" in markup
        assert "txBody" not in markup, "a connector holds no text"

    def test_an_autoshape_defaults_to_a_rectangle(self) -> None:
        shape = Shape(name="B", kind="shape", width=90, height=40, shape_id=2)
        assert 'prst="rect"' in new_anchor(shape, SheetGrid())


class TestTheValueGoesBackTheWayItCame:
    """The same three spellings the reader has to undo."""

    def test_a_ticked_box_writes_checked(self) -> None:
        part = control_properties(FormControl(kind="CheckBox", value=XL_ON))
        assert 'checked="Checked"' in part
        assert "val=" not in part

    def test_the_third_state(self) -> None:
        part = control_properties(FormControl(kind="CheckBox", value=XL_MIXED))
        assert 'checked="Mixed"' in part

    def test_an_unticked_box_writes_nothing(self) -> None:
        """Which is the absence the reader turns back into xlOff."""
        part = control_properties(FormControl(kind="CheckBox", value=XL_OFF))
        assert "checked=" not in part

    def test_a_list_writes_sel(self) -> None:
        part = control_properties(FormControl(kind="Drop", value=2))
        assert 'sel="2"' in part

    def test_a_spinner_writes_val(self) -> None:
        part = control_properties(FormControl(kind="Spin", value=7))
        assert 'val="7"' in part


class TestBounds:
    def test_a_spinner_without_a_maximum_is_not_pinned_at_zero(self) -> None:
        """Excel's own default is 30000. Writing 0 gives a control that
        sits at 0 whatever value it holds, in a file that is valid."""
        part = control_properties(FormControl(kind="Spin", value=7))
        assert 'max="30000"' in part

    def test_a_scroll_bar_defaults_to_a_hundred(self) -> None:
        assert 'max="100"' in control_properties(FormControl(kind="Scroll"))

    def test_a_given_maximum_wins(self) -> None:
        part = control_properties(FormControl(kind="Spin", maximum=50))
        assert 'max="50"' in part

    def test_the_defaults_are_per_kind(self) -> None:
        assert control_range(FormControl(kind="Spin")) == RANGE_DEFAULTS["Spin"]
        assert control_range(FormControl(kind="Scroll")) == RANGE_DEFAULTS["Scroll"]

    def test_a_kind_with_no_range_gets_nothing(self) -> None:
        assert "max=" not in control_properties(FormControl(kind="Button"))


class TestEveryKindSurvivesARoundTrip:
    """A control read from a file and written back keeps what it was.

    The reader takes ``kind`` from the part's ``objectType`` and the
    writer puts it back, so both sides speak one vocabulary: Excel's.
    Keying the writer on a second set of names is how a Radio silently
    becomes a Button, in a valid file, with nothing raised. pyOpenVBA
    loses four of nine kinds that way, which is
    WilliamSmithEdward/pyOpenVBA#25; these keep this library honest about
    the same thing.
    """

    KINDS = (
        "Button", "CheckBox", "Drop", "List", "Radio",
        "Spin", "Scroll", "GBox", "Label",
    )

    def test_the_kind_comes_back(self, bare: Workbook) -> None:
        sheet = bare["Data"]
        for index, kind in enumerate(self.KINDS):
            sheet.add_form_control(
                f"C{index}", kind=kind, left=300, top=20 + index * 40,
                width=110, height=30,
            )
        for index, kind in enumerate(self.KINDS):
            control = sheet.shape(f"C{index}").control
            assert control is not None
            assert control.kind == kind, f"C{index} went in as {kind}"

    def test_the_writer_and_the_reader_share_a_vocabulary(self) -> None:
        """Not a round trip through a file: the tables themselves.

        ``CONTROL_KINDS`` is what the reader reports, and the writer has
        to accept every one of them.
        """
        for kind in CONTROL_KINDS:
            assert check_control_kind(kind) == kind

    def test_the_wiring_comes_back_too(self, bare: Workbook) -> None:
        sheet = bare["Data"]
        sheet.add_form_control(
            "Pick", kind="Drop", left=300, top=20, width=110, height=20,
            linked_cell="$H$1", list_range="$A$1:$A$3", value=2,
        )
        control = sheet.shape("Pick").control
        assert control is not None
        assert control.linked_cell == "$H$1"
        assert control.list_range == "$A$1:$A$3"
        assert control.value == 2

    def test_an_option_button_keeps_its_state(self, bare: Workbook) -> None:
        """The kind pyOpenVBA reads as 0 in both states."""
        sheet = bare["Data"]
        sheet.add_form_control(
            "On", kind="Radio", left=300, top=20, width=110, height=20, value=XL_ON
        )
        sheet.add_form_control(
            "Off", kind="Radio", left=300, top=60, width=110, height=20, value=XL_OFF
        )
        on = sheet.shape("On").control
        off = sheet.shape("Off").control
        assert on is not None and off is not None
        assert on.value == XL_ON
        assert on.checked is True
        assert off.value == XL_OFF
        assert off.checked is False


class TestHelpers:
    def test_a_bare_name_is_qualified(self) -> None:
        assert qualified_macro("Clicked") == "[0]!Clicked"

    def test_an_existing_prefix_is_kept(self) -> None:
        assert qualified_macro("Clicked", "[2]!Old") == "[2]!Clicked"

    def test_a_qualified_name_is_left_alone(self) -> None:
        assert qualified_macro("[7]!There", "[2]!Old") == "[7]!There"

    def test_clearing_stays_clear(self) -> None:
        assert qualified_macro("", "[2]!Old") == ""

    def test_a_vml_id_drops_what_it_cannot_hold(self) -> None:
        assert vml_id("Press Me") == "PressMe"

    def test_a_vml_id_cannot_start_with_a_digit(self) -> None:
        assert vml_id("2nd") == "_2nd"

    def test_setting_a_macro_on_absent_vml_changes_nothing(self) -> None:
        assert set_vml_macro("<xml></xml>", 1025, "Run") == "<xml></xml>"
