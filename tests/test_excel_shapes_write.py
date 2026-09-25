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

from pyofficeeditor._xml import Element, XmlDocument
from pyofficeeditor.excel import FormControl, Workbook, Worksheet
from pyofficeeditor.excel._schema import WORKSHEET_CHILD_ORDER, insert_in_schema_order
from pyofficeeditor.excel._shapes import (
    CONTROL_KINDS,
    FIRST_CONTROL_ID,
    RANGE_DEFAULTS,
    XL_MIXED,
    XL_OFF,
    XL_ON,
    Shape,
    SheetGrid,
    VmlControl,
    anchor_holding,
    check_control_kind,
    control_properties,
    control_range,
    control_vml,
    corner_markup,
    css_points,
    find_shape_element,
    new_anchor,
    qualified_macro,
    set_vml_macro,
    vml_controls,
    vml_id,
)

#: The relationship Excel names an ActiveX control's part by.
RT_ACTIVEX = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control"


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


@pytest.fixture
def default_named(tmp_path: Path, live_refused_xlsx: Path) -> Workbook:
    """Excel's workbook with a button still called "Button 1" beside a
    note. Excel writes that button's VML as ``id="_x0000_s1025"`` with no
    ``o:spid``: the id is what finds it."""
    target = tmp_path / "refused.xlsx"
    shutil.copy(live_refused_xlsx, target)
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


def unreached(path: Path) -> list[str]:
    """Parts no relationship points at: left behind by a removal that took
    the relationship and not the part, or the part and not what it used."""
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        reached: set[str] = set()
        for entry in names:
            if not entry.endswith(".rels"):
                continue
            base = entry.rsplit("_rels/", 1)[0]
            for target in re.findall(r'Target="([^"]+)"', archive.read(entry).decode("utf-8")):
                resolved = target.lstrip("/") if target.startswith("/") else f"{base}{target}"
                while "/../" in resolved:
                    head, _, tail = resolved.partition("/../")
                    resolved = f"{head.rsplit('/', 1)[0]}/{tail}"
                reached.add(resolved)
    return sorted(
        name for name in names
        if name not in reached and not name.endswith(".rels") and name != "[Content_Types].xml"
    )


def drawing_of(sheet: Worksheet) -> tuple[str, XmlDocument]:
    """A sheet's drawing part, for a test that rearranges it by hand."""
    found: tuple[str, XmlDocument] = getattr(sheet, "_drawing_part")()
    return found


def detached(element: Element) -> Element:
    """An element taken out of wherever it is, to be put somewhere else."""
    parent = element.parent
    assert parent is not None
    parent.remove(element)
    return element


def group_into_pair(book: Workbook, sheet: Worksheet, name: str) -> None:
    """Move a shape into ``shapes.xlsm``'s group ``Pair``: its own element
    goes inside the group and its anchor goes."""
    part, drawing = drawing_of(sheet)
    holder = anchor_holding(drawing.root, name)
    group = find_shape_element(drawing.root, "Pair")
    body = find_shape_element(drawing.root, name)
    assert holder is not None and group is not None and body is not None
    group.append(detached(body))
    drawing.root.remove(holder)
    book.package.write(part, drawing.to_bytes())


def retype_as_activex(book: Workbook, sheet: Worksheet) -> None:
    """Point a sheet's first control record at an ActiveX part, as Excel's
    record for an ActiveX control does."""
    control = next(sheet.document.root.descendants("control"))
    relationships = book.package.xml(f"{sheet.part_name.rpartition('/')[0]}/_rels/{sheet.part_name.rpartition('/')[2]}.rels")
    entry = next(
        node for node in relationships.root.children_named("Relationship")
        if node.get("Id") == control.get("r:id")
    )
    entry.set("Type", RT_ACTIVEX)


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
    def test_activex_is_read_without_treating_its_part_as_a_form_control(
        self, tmp_path: Path, live_shapes_xlsm: Path
    ) -> None:
        target = tmp_path / "shapes.xlsm"
        shutil.copy(live_shapes_xlsm, target)
        book = Workbook.open(target)
        sheet = book["Shapes"]
        retype_as_activex(book, sheet)
        assert sheet.shape("Go").kind == "activeX"
        with pytest.raises(ValueError, match="ActiveX control removal"):
            sheet.remove_shape("Go")

        # With no drawing twin its record still names and places it.
        part, drawing = drawing_of(sheet)
        holder = anchor_holding(drawing.root, "Go")
        assert holder is not None
        drawing.root.remove(holder)
        book.package.write(part, drawing.to_bytes())
        orphan = sheet.shape("Go")
        assert (orphan.kind, orphan.cells) == ("activeX", "B19:C21")
        assert [shape.name for shape in sheet.shapes].count("Go") == 1

    def test_an_activex_fallback_record_does_not_hide_the_full_one(
        self, tmp_path: Path, live_shapes_xlsm: Path
    ) -> None:
        # Excel writes an ActiveX record twice: in full in an mc:Choice and
        # bare in the mc:Fallback, which comes after it.
        target = tmp_path / "shapes.xlsm"
        shutil.copy(live_shapes_xlsm, target)
        book = Workbook.open(target)
        sheet = book["Shapes"]
        retype_as_activex(book, sheet)
        record = next(sheet.document.root.descendants("control"))
        properties = record.child("controlPr")
        assert properties is not None
        properties.set("altText", "Runs the report")
        choice = record.parent
        assert choice is not None and choice.parent is not None
        fallback = XmlDocument.parse(
            f'<mc:Fallback><control shapeId="{record.get("shapeId")}" r:id="{record.get("r:id")}"'
            f' name="Go"/></mc:Fallback>'.encode()
        ).root
        choice.parent.append(fallback)
        go = sheet.shape("Go")
        assert (go.alt_text, go.macro) == ("Runs the report", "[1]!Clicked")

    def test_an_activex_control_in_a_group_stops_the_group_going(
        self, tmp_path: Path, live_shapes_xlsm: Path
    ) -> None:
        target = tmp_path / "shapes.xlsm"
        shutil.copy(live_shapes_xlsm, target)
        book = Workbook.open(target)
        sheet = book["Shapes"]
        retype_as_activex(book, sheet)
        group_into_pair(book, sheet, "Go")
        # Listed once, inside its group.
        assert [shape.name for shape in sheet.shapes].count("Go") == 0
        with pytest.raises(ValueError, match="ActiveX"):
            sheet.remove_shape("Pair")
        assert "Pair" in {shape.name for shape in sheet.shapes}

    def test_a_chart_takes_its_part_and_relationship(
        self, bare: Workbook, tmp_path: Path
    ) -> None:
        sheet = bare["Data"]
        chart = sheet.add_chart("column", "A1:C6", left=250, top=20)
        part = chart.part_name
        sheet.remove_shape(chart.name)
        bare.save()
        saved = tmp_path / "sample.xlsx"
        assert part not in parts_of(saved)
        assert dangling(saved) == []
        assert sheet.charts == []

    def test_an_excel_chart_takes_its_style_and_colour_parts(
        self, tmp_path: Path, live_chart_kinds_xlsx: Path
    ) -> None:
        target = tmp_path / "chartkinds.xlsx"
        shutil.copy(live_chart_kinds_xlsx, target)
        book = Workbook.open(target)
        sheet = book["Data"]
        chart = sheet.charts[0]
        owned = [one.target_part for one in book.package.relationships(chart.part_name)]
        assert owned, "Excel gives every chart a style part and a colour part"
        sheet.remove_shape(chart.name)
        book.save()
        parts = parts_of(target)
        types = parts["[Content_Types].xml"].decode()
        for part in (chart.part_name, *owned):
            assert part not in parts
            assert f'PartName="/{part}"' not in types
        assert unreached(target) == unreached(live_chart_kinds_xlsx)

    def test_a_part_already_gone_does_not_stop_the_removal(
        self, tmp_path: Path, live_shapes_xlsm: Path
    ) -> None:
        target = tmp_path / "shapes.xlsm"
        shutil.copy(live_shapes_xlsm, target)
        book = Workbook.open(target)
        sheet = book["Shapes"]
        chart = sheet.add_chart("column", "A1:C6", left=10, top=10)
        group_into_pair(book, sheet, chart.name)
        group_into_pair(book, sheet, "Go")
        book.package.remove_part(chart.part_name)
        sheet.remove_shape("Pair")
        assert "Pair" not in {shape.name for shape in sheet.shapes}
        # The control in the group went too, record and all.
        assert 'shapeId="1025"' not in sheet.document.to_text()

    def test_a_hyperlink_goes_with_its_shape(self, book: Workbook, tmp_path: Path) -> None:
        sheet = book["Controls"]
        part, drawing = drawing_of(sheet)
        body = find_shape_element(drawing.root, "Plain")
        assert body is not None
        naming = next(body.descendants("cNvPr"))
        link = book.package.relationships(part).add(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            "https://example.com/",
            external=True,
        )
        naming.append(XmlDocument.parse(f'<a:hlinkClick r:id="{link.id}"/>'.encode()).root)
        sheet.remove_shape("Plain")
        book.save()
        # The link was the drawing's only relationship, so its .rels goes too.
        rels = parts_of(tmp_path / "controls.xlsm").get("xl/drawings/_rels/drawing1.xml.rels", b"")
        assert b"example.com" not in rels

    def test_a_group_takes_its_child_picture_part(
        self, tmp_path: Path, live_shapes_xlsm: Path
    ) -> None:
        target = tmp_path / "shapes.xlsm"
        shutil.copy(live_shapes_xlsm, target)
        book = Workbook.open(target)
        sheet = book["Shapes"]
        gif = (
            b"GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff!\xf9\x04\x01\x00\x00\x00\x00"
            b",\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;"
        )
        picture = sheet.add_picture("Grouped picture", gif, left=10, top=10)
        image = picture.image
        group_into_pair(book, sheet, picture.name)
        assert any(child.image == image for child in sheet.shape("Pair").children)

        sheet.remove_shape("Pair")
        book.save()
        assert image not in parts_of(target)
        assert dangling(target) == dangling(live_shapes_xlsm)

    def test_a_group_takes_its_child_form_control_parts(
        self, tmp_path: Path, live_shapes_xlsm: Path
    ) -> None:
        target = tmp_path / "shapes.xlsm"
        shutil.copy(live_shapes_xlsm, target)
        book = Workbook.open(target)
        sheet = book["Shapes"]
        control = sheet.shape("Go")
        assert control.control is not None
        control_part = control.control.part_name
        group_into_pair(book, sheet, "Go")
        member = next(child for child in sheet.shape("Pair").children if child.name == "Go")
        # Still a control in a group: the drawing twin's hidden="1" is
        # Excel's, and the VML says it shows.
        assert (member.kind, member.hidden, member.macro) == ("formControl", False, "[1]!Clicked")

        sheet.remove_shape("Pair")
        book.save()
        parts = parts_of(target)
        assert control_part not in parts
        assert f'shapeId="{control.shape_id}"' not in parts["xl/worksheets/sheet1.xml"].decode()
        assert f"_x0000_s{control.shape_id}" not in parts["xl/drawings/vmlDrawing1.vml"].decode()

    def test_a_group_takes_its_child_chart_part(
        self, tmp_path: Path, live_shapes_xlsm: Path
    ) -> None:
        target = tmp_path / "shapes.xlsm"
        shutil.copy(live_shapes_xlsm, target)
        book = Workbook.open(target)
        sheet = book["Shapes"]
        chart = sheet.add_chart("column", "A1:C6", left=10, top=10)
        group_into_pair(book, sheet, chart.name)

        sheet.remove_shape("Pair")
        book.save()
        assert chart.part_name not in parts_of(target)
        assert dangling(target) == dangling(live_shapes_xlsm)

    def test_a_group_member_is_not_reached_by_its_name(
        self, tmp_path: Path, live_shapes_xlsm: Path
    ) -> None:
        # The sheet lists groups, not members, so a member's name is neither
        # something to remove nor something a new shape may take.
        target = tmp_path / "shapes.xlsm"
        shutil.copy(live_shapes_xlsm, target)
        sheet = Workbook.open(target)["Shapes"]
        with pytest.raises(KeyError):
            sheet.remove_shape("GroupB")
        with pytest.raises(ValueError, match="inside a group"):
            sheet.add_shape("GroupB", left=10, top=10, width=40, height=20)
        with pytest.raises(ValueError, match="inside a group"):
            sheet.update_shape("Box", new_name="GroupA")
        assert [child.name for child in sheet.shape("Pair").children] == ["GroupA", "GroupB"]

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

    def test_the_last_control_takes_the_wrapper_around_the_records(
        self, book: Workbook, tmp_path: Path
    ) -> None:
        # Excel refuses a sheet that keeps an empty mc:AlternateContent; the
        # live gate measures it.
        sheet = book["Controls"]
        for name in [shape.name for shape in sheet.shapes if shape.kind == "formControl"]:
            sheet.remove_shape(name)
        book.save()
        text = parts_of(tmp_path / "controls.xlsm")["xl/worksheets/sheet1.xml"].decode()
        assert "controls>" not in text
        assert "AlternateContent" not in text

    def test_an_unknown_name_says_what_is_there(self, book: Workbook) -> None:
        with pytest.raises(KeyError, match="no shape named"):
            book["Controls"].remove_shape("Nonesuch")


class TestUpdating:
    def test_move_resize_rename_and_text_keep_the_shape(
        self, book: Workbook, tmp_path: Path
    ) -> None:
        sheet = book["Controls"]
        before = sheet.shape("Plain")
        after = sheet.update_shape(
            "Plain", new_name="Moved", left=100, top=50, width=120,
            height=45, text="New text", alt_text="Useful description", hidden=True,
        )
        assert (after.name, after.left, after.top, after.width, after.height) == (
            "Moved", 100, 50, 120, 45,
        )
        assert (after.text, after.alt_text, after.hidden) == (
            "New text", "Useful description", True,
        )
        assert after.shape_id == before.shape_id
        assert all(shape.name != "Plain" for shape in sheet.shapes)
        book.save()
        reopened = Workbook.open(tmp_path / "controls.xlsm")["Controls"].shape("Moved")
        assert (reopened.left, reopened.top, reopened.width, reopened.height) == (
            100, 50, 120, 45,
        )
        assert (reopened.text, reopened.alt_text, reopened.hidden) == (
            "New text", "Useful description", True,
        )

    def test_control_updates_all_its_parts(self, book: Workbook, tmp_path: Path) -> None:
        sheet = book["Controls"]
        before = sheet.shape("Go")
        after = sheet.update_shape(
            "Go", new_name="Moved", left=110, top=35, width=125, height=40,
            text="Now run", alt_text="Run the macro", hidden=True,
            linked_cell="$D$6", list_range="$H$1:$H$9",
        )
        assert after.shape_id == before.shape_id
        assert (after.name, after.left, after.top, after.width, after.height) == (
            "Moved", 110, 35, 125, 40,
        )
        assert (after.text, after.alt_text, after.hidden) == (
            "Now run", "Run the macro", True,
        )
        assert after.control is not None
        assert (after.control.linked_cell, after.control.list_range) == (
            "$D$6", "$H$1:$H$9",
        )
        book.save()
        reopened = Workbook.open(tmp_path / "controls.xlsm")["Controls"].shape("Moved")
        assert (reopened.left, reopened.top, reopened.width, reopened.height) == (
            110, 35, 125, 40,
        )
        assert reopened.control is not None
        assert reopened.control.linked_cell == "$D$6"
        assert (reopened.text, reopened.alt_text, reopened.hidden) == (
            "Now run", "Run the macro", True,
        )
        vml = parts_of(tmp_path / "controls.xlsm")["xl/drawings/vmlDrawing1.vml"].decode()
        assert "z-index:1;visibility:hidden;mso-wrap-style:tight" in vml
        assert dangling(tmp_path / "controls.xlsm") == []

    def test_control_wiring_and_visibility_can_be_cleared(self, book: Workbook) -> None:
        sheet = book["Controls"]
        sheet.update_shape("Go", hidden=True, linked_cell="$D$6", list_range="$H$1:$H$9")
        cleared = sheet.update_shape("Go", hidden=False, linked_cell="", list_range="")
        assert cleared.hidden is False
        assert cleared.control is not None
        assert (cleared.control.linked_cell, cleared.control.list_range) == ("", "")

    def test_a_caption_keeps_how_it_is_drawn(self, book: Workbook, tmp_path: Path) -> None:
        sheet = book["Controls"]
        sheet.update_shape("Go", text="Run it")
        book.save()
        parts = parts_of(tmp_path / "controls.xlsm")
        vml = parts["xl/drawings/vmlDrawing1.vml"].decode()
        assert re.search(r'<font face="Aptos Narrow" size="220"\s+color="#000000">Run it</font>', vml)
        drawing = parts["xl/drawings/drawing1.xml"].decode()
        run = re.search(r'<a:t>Run it</a:t>', drawing)
        assert run is not None
        body = drawing[drawing.rfind("<xdr:txBody>", 0, run.start()) : run.start()]
        assert 'anchor="ctr"' in body and 'algn="ctr"' in body and 'sz="1100"' in body

    def test_a_control_with_no_caption_refuses_text_and_changes_nothing(
        self, book: Workbook
    ) -> None:
        sheet = book["Controls"]
        with pytest.raises(ValueError, match="shows no text"):
            sheet.update_shape("Pick", new_name="Chooser", text="x")
        assert "Pick" in {shape.name for shape in sheet.shapes}
        assert "Chooser" not in {shape.name for shape in sheet.shapes}

    def test_a_control_made_here_takes_new_text(self, bare: Workbook) -> None:
        sheet = bare["Data"]
        sheet.add_form_control("Run", left=10, top=10, width=80, height=24, text="Old")
        assert sheet.shape("Run").text == "Old"
        assert sheet.update_shape("Run", new_name="Go", text="New").text == "New"

    def test_a_line_can_be_renamed(self, book: Workbook) -> None:
        sheet = book["Controls"]
        sheet.add_shape("Rule", kind="line", left=10, top=300, width=120, height=0)
        assert sheet.update_shape("Rule", new_name="Divider", alt_text="a rule").name == "Divider"

    def test_a_side_that_is_not_a_number_is_refused(self, book: Workbook) -> None:
        sheet = book["Controls"]
        for bad in (float("nan"), float("inf"), -1.0):
            with pytest.raises(ValueError, match="points"):
                sheet.update_shape("Go", left=bad)
            with pytest.raises(ValueError, match="points"):
                sheet.update_shape("Plain", width=bad)
        assert sheet.shape("Go").cells == "E2:G4"

    def test_a_side_left_out_keeps_where_the_anchor_is(self, book: Workbook) -> None:
        # Rows put in above move the anchor, and a width change must not
        # move the shape back to where it was.
        sheet = book["Controls"]
        before = sheet.shape("Plain").cells
        sheet.insert_rows(1, 5)
        moved = sheet.shape("Plain").cells
        assert moved != before
        after = sheet.update_shape("Plain", width=sheet.shape("Plain").width + 10)
        assert after.cells.split(":")[0] == moved.split(":")[0]

    def test_a_chart_moves_by_its_anchor_alone(self, bare: Workbook, tmp_path: Path) -> None:
        sheet = bare["Data"]
        chart = sheet.add_chart("column", "A1:C6", left=250, top=20)
        sheet.update_shape(chart.name, left=300, top=40)
        bare.save()
        drawing = parts_of(tmp_path / "sample.xlsx")["xl/drawings/drawing1.xml"].decode()
        assert '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' in drawing
        assert sheet.shape(chart.name).left == pytest.approx(300, abs=0.01)

    def test_new_text_drops_a_link_to_a_cell(self, book: Workbook) -> None:
        sheet = book["Controls"]
        _, drawing = drawing_of(sheet)
        body = find_shape_element(drawing.root, "Plain")
        assert body is not None
        body.set("textlink", "$A$1")
        sheet.update_shape("Plain", text="Fresh words")
        assert body.get("textlink") == ""

    def test_a_new_link_goes_in_its_slot(self, book: Workbook, tmp_path: Path) -> None:
        sheet = book["Controls"]
        sheet.update_shape("Part", linked_cell="$D$10")
        book.save()
        vml = parts_of(tmp_path / "controls.xlsm")["xl/drawings/vmlDrawing1.vml"].decode()
        block = re.search(r'o:spid="_x0000_s1028".*?</v:shape>', vml, re.DOTALL)
        assert block is not None
        assert block.group(0).index("<x:FmlaLink>") < block.group(0).index("<x:NoThreeD/>")

    def test_both_copies_of_an_alternate_content_shape_change(
        self, book: Workbook
    ) -> None:
        # A shape newer Excel draws, with a picture copy for older Excel in
        # the mc:Fallback: both carry the name, and a move reaches both.
        sheet = book["Controls"]
        _, drawing = drawing_of(sheet)
        holder = anchor_holding(drawing.root, "Plain")
        assert holder is not None
        markup = holder.to_xml()
        wrapper = XmlDocument.parse(
            (
                "<mc:AlternateContent><mc:Choice Requires=\"a14\">"
                f"{markup}</mc:Choice><mc:Fallback>{markup}</mc:Fallback></mc:AlternateContent>"
            ).encode()
        ).root
        drawing.root.insert_before(holder, wrapper)
        drawing.root.remove(holder)
        sheet.update_shape("Plain", new_name="Both", alt_text="twice", left=200, top=30)
        copies = [(node.get("name"), node.get("descr")) for node in wrapper.descendants("cNvPr")]
        assert copies == [("Both", "twice"), ("Both", "twice")]
        corners = [node.to_xml() for node in wrapper.descendants("from")]
        assert len(corners) == 2 and corners[0] == corners[1]

    def test_lengths_in_the_vml_never_use_an_exponent(self) -> None:
        assert css_points(1_500_000) == "1500000pt"
        assert css_points(12.75) == "12.75pt"
        assert css_points(0) == "0pt"


class TestReadingTheVml:
    def test_hidden_is_read_from_the_style_by_the_writers_rule(self) -> None:
        vml = (
            '<v:shape id="A" o:spid="_x0000_s1025" style=\'position:absolute;Visibility:Hidden;'
            "mso-wrap-style:tight'><v:textbox><div><font face=\"X\">One<br>\r\n Two &amp; three"
            '</font></div></v:textbox></v:shape>'
            '<v:shape id="_x0000_s1026" style=\'visibility:hidden\'></v:shape>'
        )
        facts = vml_controls(vml)
        assert facts[1025].hidden is True
        assert facts[1025].caption == "One\nTwo & three"
        # With no o:spid, the id is the spid, as Excel writes a shape still
        # called by its default name.
        assert facts[1026] == VmlControl(hidden=True, caption=None)

    def test_an_spid_wins_over_an_id_that_looks_like_one(self) -> None:
        vml = '<v:shape id="_x0000_s1030" o:spid="_x0000_s1025"></v:shape>'
        assert set(vml_controls(vml)) == {1025}

    def test_a_self_closed_shape_does_not_swallow_the_next(self) -> None:
        vml = (
            '<v:shape id="_x0000_s1025"/>'
            "<v:shape id=\"_x0000_s1026\" style='visibility:hidden'></v:shape>"
        )
        assert vml_controls(vml) == {1026: VmlControl(hidden=True, caption=None)}


class TestADefaultNamedControl:
    """Excel writes the VML of a control still called by its default name
    as ``<v:shape id="_x0000_s1025">``, and gives it an ``o:spid`` only once
    it is renamed. Every edit has to find it the first way too."""

    def test_its_caption_and_visibility_are_read(self, default_named: Workbook) -> None:
        shape = default_named["R"].shape("Button 1")
        assert (shape.kind, shape.text, shape.hidden) == ("formControl", "Button 1", False)

    def test_it_takes_alt_text(self, default_named: Workbook) -> None:
        sheet = default_named["R"]
        assert sheet.update_shape("Button 1", alt_text="described").alt_text == "described"

    def test_it_can_be_hidden(self, default_named: Workbook, tmp_path: Path) -> None:
        default_named["R"].update_shape("Button 1", hidden=True)
        default_named.save()
        assert Workbook.open(tmp_path / "refused.xlsx")["R"].shape("Button 1").hidden is True

    def test_it_moves(self, default_named: Workbook, tmp_path: Path) -> None:
        default_named["R"].update_shape("Button 1", left=20)
        default_named.save()
        vml = parts_of(tmp_path / "refused.xlsx")["xl/drawings/vmlDrawing1.vml"].decode()
        head = vml[vml.find('id="_x0000_s1025"') :]
        assert "margin-left:20pt" in head[: head.find(">")]

    def test_renaming_it_gives_its_vml_an_spid(self, default_named: Workbook, tmp_path: Path) -> None:
        """The id stops being the spid once it holds the name, so the spid
        moves to ``o:spid``, where Excel keeps a renamed control's."""
        default_named["R"].update_shape("Button 1", new_name="Renamed")
        default_named.save()
        vml = parts_of(tmp_path / "refused.xlsx")["xl/drawings/vmlDrawing1.vml"].decode()
        assert 'id="Renamed" o:spid="_x0000_s1025"' in vml
        again = Workbook.open(tmp_path / "refused.xlsx")["R"]
        assert again.update_shape("Renamed", hidden=True).hidden is True

    def test_its_macro_reaches_the_vml(self, default_named: Workbook, tmp_path: Path) -> None:
        default_named["R"].set_shape_macro("Button 1", "Clicked")
        default_named.save()
        vml = parts_of(tmp_path / "refused.xlsx")["xl/drawings/vmlDrawing1.vml"].decode()
        block = vml[vml.find('id="_x0000_s1025"') :]
        assert "<x:FmlaMacro>[0]!Clicked</x:FmlaMacro>" in block[: block.find("</v:shape>")]

    def test_removing_it_takes_its_vml_shape_and_leaves_the_note(
        self, default_named: Workbook, tmp_path: Path
    ) -> None:
        default_named["R"].remove_shape("Button 1")
        default_named.save()
        vml = parts_of(tmp_path / "refused.xlsx")["xl/drawings/vmlDrawing1.vml"].decode()
        assert 'id="_x0000_s1025"' not in vml
        assert 'id="_x0000_s1026"' in vml


#: A header picture's VML part as Excel writes it, less the shape type.
HEADER_VML = (
    '<xml xmlns:v="urn:schemas-microsoft-com:vml"\r\n'
    ' xmlns:o="urn:schemas-microsoft-com:office:office"\r\n'
    ' xmlns:x="urn:schemas-microsoft-com:office:excel">\r\n'
    ' <o:shapelayout v:ext="edit">\r\n  <o:idmap v:ext="edit" data="1"/>\r\n'
    ' </o:shapelayout><v:shape id="LH" o:spid="_x0000_s1025" type="#_x0000_t75"\r\n'
    "  style='position:absolute;margin-left:0;margin-top:0;width:30pt;height:15pt;\r\n"
    "  z-index:1'>\r\n"
    '  <v:imagedata o:relid="rId1" o:title="logo"/>\r\n'
    '  <o:lock v:ext="edit" rotation="t"/>\r\n'
    " </v:shape></xml>"
)


def with_header_picture(book: Workbook, sheet: Worksheet) -> str:
    """Give a sheet a header picture's VML part, related first, as Excel
    relates it, and named by ``<legacyDrawingHF>``. Returns the part."""
    package = book.package
    part = book.free_part_name("xl/drawings/vmlDrawing{n}.vml")
    package.content_types.set_default("vml", "application/vnd.openxmlformats-officedocument.vmlDrawing")
    package.write(part, HEADER_VML.encode("utf-8"))
    relationship = package.relationships(sheet.part_name).add_part(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing", part
    )
    insert_in_schema_order(
        sheet.document.root,
        Element.create("legacyDrawingHF", {"r:id": relationship.id}),
        WORKSHEET_CHILD_ORDER,
    )
    return part


class TestAHeaderPicture:
    """A header or footer picture is VML related by the same type as the
    sheet's notes and controls, in a part of its own. Measured: Excel lists
    it first and numbers the picture from 1025, as it numbers the notes."""

    def test_a_note_and_a_control_go_in_the_sheets_own_vml(self, bare: Workbook) -> None:
        sheet = bare["Data"]
        header = with_header_picture(bare, sheet)
        sheet.add_form_control("Go", left=100, top=60, width=80, height=24, text="Go")
        sheet.set_comment("D4", "a note")
        assert bare.package.read(header).decode("utf-8") == HEADER_VML
        drawing = getattr(sheet, "_legacy_vml_parts")()
        assert header not in drawing
        assert len(drawing) == 1
        vml = bare.package.read(drawing[0]).decode("utf-8")
        assert 'ObjectType="Button"' in vml and 'ObjectType="Note"' in vml
        assert sheet.comment("D4") is not None

    def test_an_edit_to_a_control_leaves_the_header_alone(self, bare: Workbook) -> None:
        sheet = bare["Data"]
        header = with_header_picture(bare, sheet)
        shape_id = sheet.add_form_control("Go", left=100, top=60, width=80, height=24, text="Go").shape_id
        # Excel's own numbering: the picture shares its number with a note
        # or a control on the sheet.
        clashing = HEADER_VML.replace("_x0000_s1025", f"_x0000_s{shape_id}")
        bare.package.write(header, clashing.encode("utf-8"))

        sheet.update_shape("Go", new_name="Moved", left=120, text="Now", hidden=True)
        assert sheet.shape("Moved").hidden is True
        sheet.set_shape_macro("Moved", "Clicked")
        sheet.remove_shape("Moved")
        assert bare.package.read(header).decode("utf-8") == clashing


def with_ole_object(book: Workbook, sheet: Worksheet, name: str, *, hidden: bool = False) -> int:
    """Embed an OLE object the way Excel does, measured: a hidden drawing
    twin, an ``<oleObject>`` record in the sheet's ``<oleObjects>``, a VML
    shape that draws it, and the embedded part. Returns its shape id."""
    package = book.package
    shape_id: int = getattr(sheet, "_next_control_id")()
    getattr(sheet, "_declare_control_namespaces")()
    _, drawing = drawing_of(sheet)
    corners = (
        "<xdr:from><xdr:col>2</xdr:col><xdr:colOff>47625</xdr:colOff><xdr:row>2</xdr:row>"
        "<xdr:rowOff>123825</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col>"
        "<xdr:colOff>590550</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>66675</xdr:rowOff></xdr:to>"
    )
    twin = (
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">'
        f'<xdr:twoCellAnchor editAs="oneCell">{corners}<xdr:sp macro="" textlink=""><xdr:nvSpPr>'
        f'<xdr:cNvPr id="{shape_id}" name="{name}" hidden="1"><a:extLst>'
        '<a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}">'
        f'<a14:compatExt spid="_x0000_s{shape_id}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/>'
        '</xdr:nvSpPr><xdr:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>'
        "<xdr:clientData/></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent>"
    )
    drawing.root.append(XmlDocument.parse(twin.encode("utf-8")).root)

    embedded = book.free_part_name("xl/embeddings/oleObject{n}.bin")
    package.write(embedded, b"\xd0\xcf\x11\xe0", content_type="application/vnd.openxmlformats-officedocument.oleObject")
    relationship = package.relationships(sheet.part_name).add_part(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject", embedded
    )
    record = (
        '<oleObjects><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        f'<mc:Choice Requires="x14"><oleObject progId="Packager Shell Object" shapeId="{shape_id}" '
        f'r:id="{relationship.id}"><objectPr defaultSize="0"><anchor moveWithCells="1">'
        f'{corners.replace("xdr:from", "from").replace("xdr:to>", "to>")}</anchor></objectPr></oleObject>'
        f'</mc:Choice><mc:Fallback><oleObject progId="Packager Shell Object" shapeId="{shape_id}" '
        f'r:id="{relationship.id}"/></mc:Fallback></mc:AlternateContent></oleObjects>'
    )
    insert_in_schema_order(
        sheet.document.root, XmlDocument.parse(record.encode("utf-8")).root, WORKSHEET_CHILD_ORDER
    )

    part: str = getattr(sheet, "_vml_part")()
    vml = package.read(part).decode("utf-8")
    shape = (
        f'<v:shape id="{name}" o:spid="_x0000_s{shape_id}" type="#_x0000_t75"\r\n'
        "  style='position:absolute;margin-left:99.75pt;margin-top:39.75pt;width:42.75pt;\r\n"
        f"  height:40.5pt;z-index:1{';visibility:hidden' if hidden else ''}' filled=\"t\"\r\n"
        '  o:insetmode="auto">\r\n  <v:imagedata o:relid="rId1" o:title=""/>\r\n'
        '  <x:ClientData ObjectType="Pict">\r\n   <x:SizeWithCells/>\r\n'
        "   <x:Anchor>\r\n    2, 5, 2, 13, 2, 62, 5, 7</x:Anchor>\r\n   <x:CF>Pict</x:CF>\r\n"
        "   <x:AutoPict/>\r\n  </x:ClientData>\r\n </v:shape>"
    )
    package.write(part, vml.replace("</xml>", f"{shape}</xml>").encode("utf-8"))
    return shape_id


class TestAnOleObject:
    """Excel draws an embedded object from its VML and keeps a hidden twin
    of it in the drawing, as it does a control. Measured: ``Shape.Type`` is
    7, and the twin says hidden whether the object shows or not."""

    def test_it_reads_as_one_and_shown(self, bare: Workbook) -> None:
        sheet = bare["Data"]
        with_ole_object(bare, sheet, "Attached")
        shape = sheet.shape("Attached")
        assert (shape.kind, shape.mso_type, shape.hidden, shape.cells) == ("oleObject", 7, False, "C3:C6")

    def test_a_hidden_one_reads_hidden(self, bare: Workbook) -> None:
        sheet = bare["Data"]
        with_ole_object(bare, sheet, "Tucked", hidden=True)
        assert sheet.shape("Tucked").hidden is True

    def test_update_shape_leaves_it_alone(self, bare: Workbook) -> None:
        sheet = bare["Data"]
        with_ole_object(bare, sheet, "Attached")
        with pytest.raises(ValueError, match="OLE object"):
            sheet.update_shape("Attached", new_name="Renamed", hidden=True)
        assert sheet.shape("Attached").hidden is False

    def test_remove_shape_leaves_it_alone(self, bare: Workbook) -> None:
        """Taking only its twin leaves Excel drawing it from the record and
        the VML, measured, so nothing goes."""
        sheet = bare["Data"]
        with_ole_object(bare, sheet, "Attached")
        with pytest.raises(ValueError, match="OLE object removal"):
            sheet.remove_shape("Attached")
        assert sheet.shape("Attached").kind == "oleObject"


class TestSettingAMacro:
    def test_on_a_drawing_shape(self, book: Workbook) -> None:
        sheet = book["Controls"]
        sheet.set_shape_macro("Plain", "Clicked")
        assert sheet.shape("Plain").macro == "Clicked"

    def test_an_activex_control_has_none_to_set(self, book: Workbook) -> None:
        sheet = book["Controls"]
        retype_as_activex(book, sheet)
        with pytest.raises(ValueError, match="event procedure"):
            sheet.set_shape_macro("Go", "Other")
        assert sheet.shape("Go").macro == "[0]!Clicked"

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
