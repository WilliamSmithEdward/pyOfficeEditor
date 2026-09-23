"""The OPC package layer.

Two things matter here. The path arithmetic, because a relationship target
is relative to the folder of the part that declares it and getting that
wrong points Office at a part that is not there. And the flush discipline:
a part nobody modified must never be re-serialized, so a no-op save stays a
no-op.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from pyofficeeditor._xml import Element
from pyofficeeditor._zip import ZipArchive
from pyofficeeditor.exceptions import PackageError
from pyofficeeditor.opc import (
    CONTENT_TYPES_PART,
    RT_OFFICE_DOCUMENT,
    OpcPackage,
    normalize_part_name,
    relative_target,
    rels_part_for,
    resolve_target,
)

RT_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
RT_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"


def sheet_part_by_name(package: OpcPackage, name: str) -> str:
    """Resolve a worksheet the way Office does.

    Not by indexing the worksheet relationships: their order is the rels
    part's, which is not sheet order. In the Excel-authored sample fixture
    ``rId2 -> worksheets/sheet2.xml`` is listed *before*
    ``rId1 -> worksheets/sheet1.xml``, so ``by_type(...)[0]`` is the second
    sheet. The order and the names live in ``<sheets>`` inside
    ``xl/workbook.xml``, and each entry's ``r:id`` names the relationship.
    """
    workbook_part = package.main_document_part()
    workbook = package.xml(workbook_part)
    for entry in workbook.root.require("sheets").children_named("sheet"):
        if entry.get("name") == name:
            relationship_id = entry.get("r:id") or entry.get("id")
            assert relationship_id is not None, f"<sheet name={name!r}> has no r:id"
            return package.relationships(workbook_part).by_id(relationship_id).target_part
    raise AssertionError(f"no sheet named {name!r}")


@pytest.fixture()
def package(minimal_xlsm_bytes: bytes) -> OpcPackage:
    return OpcPackage.from_bytes(minimal_xlsm_bytes)


class TestPathArithmetic:
    @pytest.mark.parametrize(
        ("given", "expected"),
        [
            ("/xl/workbook.xml", "xl/workbook.xml"),
            ("xl/workbook.xml", "xl/workbook.xml"),
            ("xl/./a/../workbook.xml", "xl/workbook.xml"),
            ("xl\\workbook.xml", "xl/workbook.xml"),
        ],
    )
    def test_normalize(self, given: str, expected: str) -> None:
        assert normalize_part_name(given) == expected

    def test_a_name_climbing_above_the_root_is_refused(self) -> None:
        with pytest.raises(PackageError, match="above the package root"):
            normalize_part_name("../outside.xml")

    @pytest.mark.parametrize(
        ("part", "expected"),
        [
            ("", "_rels/.rels"),
            ("xl/workbook.xml", "xl/_rels/workbook.xml.rels"),
            ("doc.xml", "_rels/doc.xml.rels"),
            ("a/b/c.xml", "a/b/_rels/c.xml.rels"),
        ],
    )
    def test_rels_part_for(self, part: str, expected: str) -> None:
        assert rels_part_for(part) == expected

    @pytest.mark.parametrize(
        ("source", "target", "expected"),
        [
            ("xl/workbook.xml", "worksheets/sheet1.xml", "xl/worksheets/sheet1.xml"),
            ("xl/workbook.xml", "/docProps/core.xml", "docProps/core.xml"),
            ("xl/worksheets/sheet1.xml", "../media/i.png", "xl/media/i.png"),
            ("", "xl/workbook.xml", "xl/workbook.xml"),
            ("xl/workbook.xml", "./styles.xml", "xl/styles.xml"),
        ],
    )
    def test_resolve_target(self, source: str, target: str, expected: str) -> None:
        assert resolve_target(source, target) == expected

    @pytest.mark.parametrize(
        ("source", "target"),
        [
            ("xl/workbook.xml", "xl/worksheets/sheet1.xml"),
            ("xl/worksheets/sheet1.xml", "xl/media/image1.png"),
            ("", "xl/workbook.xml"),
            ("xl/workbook.xml", "docProps/core.xml"),
            ("a/b/c.xml", "a/b/d.xml"),
            ("a/b/c.xml", "x/y/z.xml"),
        ],
    )
    def test_relative_target_is_the_inverse_of_resolve(self, source: str, target: str) -> None:
        assert resolve_target(source, relative_target(source, target)) == target


class TestOpening:
    def test_a_zip_without_content_types_is_not_a_package(self) -> None:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("hello.txt", b"hi")
        with pytest.raises(PackageError, match=CONTENT_TYPES_PART.replace("[", r"\[")):
            OpcPackage.from_bytes(buffer.getvalue())

    def test_open_reports_a_missing_file(self, tmp_path: Path) -> None:
        with pytest.raises(PackageError, match="no such file"):
            OpcPackage.open(tmp_path / "absent.xlsx")

    def test_open_records_the_path(self, tmp_path: Path, minimal_xlsm_bytes: bytes) -> None:
        target = tmp_path / "book.xlsm"
        target.write_bytes(minimal_xlsm_bytes)
        with OpcPackage.open(target) as opened:
            assert opened.path == target

    def test_part_names_are_the_archives(self, package: OpcPackage, minimal_xlsm_bytes: bytes) -> None:
        assert package.part_names() == ZipArchive.from_bytes(minimal_xlsm_bytes).names()

    def test_has_part_normalizes(self, package: OpcPackage) -> None:
        assert package.has_part("/xl/workbook.xml")
        assert not package.has_part("xl/absent.xml")

    def test_reading_an_absent_part_raises(self, package: OpcPackage) -> None:
        with pytest.raises(PackageError, match="no part"):
            package.read("xl/absent.xml")


class TestNoOpSave:
    def test_every_authored_package_saves_unchanged(
        self, excel_authored_packages: list[tuple[str, bytes]], openpyxl_xlsx_bytes: bytes
    ) -> None:
        for name, data in [*excel_authored_packages, ("openpyxl.xlsx", openpyxl_xlsx_bytes)]:
            assert OpcPackage.from_bytes(data).to_bytes() == data, name

    def test_merely_reading_parts_does_not_dirty_them(self, package: OpcPackage, minimal_xlsm_bytes: bytes) -> None:
        for name in package.part_names():
            if name.endswith((".xml", ".rels")):
                package.xml(name)
        assert package.to_bytes() == minimal_xlsm_bytes

    def test_touching_content_types_does_not_dirty_it(self, package: OpcPackage, minimal_xlsm_bytes: bytes) -> None:
        package.content_types.of("xl/workbook.xml")
        assert package.to_bytes() == minimal_xlsm_bytes

    def test_saving_to_disk_writes_the_same_bytes(
        self, tmp_path: Path, minimal_xlsm_bytes: bytes
    ) -> None:
        target = tmp_path / "book.xlsm"
        target.write_bytes(minimal_xlsm_bytes)
        with OpcPackage.open(target) as opened:
            assert opened.save() == target
        assert target.read_bytes() == minimal_xlsm_bytes

    def test_a_package_from_bytes_needs_a_path_to_save(self, package: OpcPackage) -> None:
        with pytest.raises(PackageError, match="pass a path"):
            package.save()


class TestNavigation:
    def test_the_main_document_is_found_by_relationship(self, package: OpcPackage) -> None:
        assert package.main_document_part() == "xl/workbook.xml"

    def test_root_relationships(self, package: OpcPackage) -> None:
        relationships = package.relationships()
        assert len(relationships) == 3
        assert relationships.by_id("rId1").target == "xl/workbook.xml"
        assert relationships.one(RT_OFFICE_DOCUMENT).target_part == "xl/workbook.xml"

    def test_a_missing_relationship_id_is_named(self, package: OpcPackage) -> None:
        with pytest.raises(PackageError, match="rId99"):
            package.relationships().by_id("rId99")

    def test_one_refuses_when_there_is_not_exactly_one(self, package: OpcPackage) -> None:
        with pytest.raises(PackageError, match=r"declares 0 .*hyperlink"):
            package.relationships().one(RT_HYPERLINK)

    def test_the_worksheet_relationship_resolves_to_a_real_part(self, package: OpcPackage) -> None:
        sheets = package.relationships("xl/workbook.xml").by_type(RT_WORKSHEET)
        assert len(sheets) == 1
        assert package.has_part(sheets[0].target_part)

    def test_next_id_skips_what_is_used(self, package: OpcPackage) -> None:
        assert package.relationships().next_id() == "rId4"

    def test_the_wider_package_navigates_too(self, powerquery_xlsx_bytes: bytes) -> None:
        wide = OpcPackage.from_bytes(powerquery_xlsx_bytes)
        main = wide.main_document_part()
        assert main == "xl/workbook.xml"
        for relationship in wide.relationships(main):
            if not relationship.is_external:
                assert wide.has_part(relationship.target_part), relationship.target


class TestContentTypes:
    def test_resolution_by_override_then_by_extension(self, package: OpcPackage) -> None:
        types = package.content_types
        assert types.of("xl/worksheets/sheet1.xml") == CT_WORKSHEET
        assert types.of("_rels/.rels") == (
            "application/vnd.openxmlformats-package.relationships+xml"
        )
        assert types.of("xl/unknown.bin") is None

    def test_setting_and_removing_an_override(self, package: OpcPackage) -> None:
        package.content_types.set_override("xl/new.xml", "application/x-test+xml")
        assert package.content_types.of("xl/new.xml") == "application/x-test+xml"
        reopened = OpcPackage.from_bytes(package.to_bytes())
        assert reopened.content_types.of("xl/new.xml") == "application/x-test+xml"
        assert reopened.content_types.remove_override("xl/new.xml") is True
        assert reopened.content_types.remove_override("xl/new.xml") is False

    def test_setting_an_existing_override_replaces_it(self, package: OpcPackage) -> None:
        package.content_types.set_override("xl/worksheets/sheet1.xml", "application/x-other+xml")
        assert (
            OpcPackage.from_bytes(package.to_bytes()).content_types.of("xl/worksheets/sheet1.xml")
            == "application/x-other+xml"
        )

    def test_a_new_default_lands_before_the_overrides(self, package: OpcPackage) -> None:
        """Office writes every Default first; a Default after an Override is
        legal but is not the shape Excel produces."""
        package.content_types.set_default("png", "image/png")
        text = package.read(CONTENT_TYPES_PART).decode()
        assert text.index('Extension="png"') < text.index("<Override")
        assert OpcPackage.from_bytes(package.to_bytes()).content_types.of("xl/media/i.png") == "image/png"


class TestEditing:
    def test_the_cached_tree_is_shared(self, package: OpcPackage) -> None:
        assert package.xml("xl/workbook.xml") is package.xml("xl/workbook.xml")

    def test_an_edit_persists_and_leaves_the_rest_stored(
        self, minimal_xlsm_bytes: bytes
    ) -> None:
        package = OpcPackage.from_bytes(minimal_xlsm_bytes)
        part = package.relationships("xl/workbook.xml").by_type(RT_WORKSHEET)[0].target_part
        package.xml(part).root.require("dimension").set("ref", "A1:B9")
        saved = package.to_bytes()

        reopened = OpcPackage.from_bytes(saved)
        assert reopened.xml(part).root.require("dimension").get("ref") == "A1:B9"
        before = ZipArchive.from_bytes(minimal_xlsm_bytes)
        after = ZipArchive.from_bytes(saved)
        for name in after.names():
            if name != part:
                assert after.member(name).stored == before.member(name).stored, name

    def test_read_reflects_a_pending_edit(self, package: OpcPackage) -> None:
        package.xml("xl/worksheets/sheet1.xml").root.require("dimension").set("ref", "A1:Z99")
        assert b"A1:Z99" in package.read("xl/worksheets/sheet1.xml")

    def test_writing_a_part_invalidates_its_cached_tree(self, package: OpcPackage) -> None:
        package.xml("xl/worksheets/sheet1.xml")
        package.write("xl/worksheets/sheet1.xml", b"<worksheet><dimension ref=\"A1:A1\"/></worksheet>")
        assert package.xml("xl/worksheets/sheet1.xml").root.require("dimension").get("ref") == "A1:A1"

    def test_a_new_part_without_a_content_type_is_refused(self, package: OpcPackage) -> None:
        with pytest.raises(PackageError, match="no content type"):
            package.write("xl/custom.bin", b"\x00\x01")

    def test_a_new_part_with_a_content_type_is_written(self, package: OpcPackage) -> None:
        package.write("xl/custom.xml", b"<c/>", content_type="application/x-custom+xml")
        reopened = OpcPackage.from_bytes(package.to_bytes())
        assert reopened.read("xl/custom.xml") == b"<c/>"
        assert reopened.content_types.of("xl/custom.xml") == "application/x-custom+xml"

    def test_a_new_xml_part_needs_no_override_when_the_default_covers_it(self, package: OpcPackage) -> None:
        """``Default Extension="xml"`` resolves to ``application/xml``, which
        is a content type, so the write is allowed without an override."""
        package.write("xl/plain.xml", b"<p/>")
        assert OpcPackage.from_bytes(package.to_bytes()).read("xl/plain.xml") == b"<p/>"

    def test_removing_a_part_removes_its_override(self, package: OpcPackage) -> None:
        package.remove_part("docProps/app.xml")
        reopened = OpcPackage.from_bytes(package.to_bytes())
        assert not reopened.has_part("docProps/app.xml")
        assert "docProps/app.xml" not in reopened.content_types.overrides

    def test_removing_an_absent_part_raises(self, package: OpcPackage) -> None:
        package.remove_part("docProps/app.xml")
        with pytest.raises(PackageError, match="no part"):
            package.remove_part("docProps/app.xml")

    def test_removing_a_part_removes_its_rels(self, package: OpcPackage) -> None:
        assert package.has_part("xl/_rels/workbook.xml.rels")
        package.remove_part("xl/workbook.xml")
        assert not package.has_part("xl/_rels/workbook.xml.rels")


class TestRelationshipWriting:
    def test_an_unused_rels_part_is_never_written(self, package: OpcPackage, minimal_xlsm_bytes: bytes) -> None:
        relationships = package.relationships("xl/worksheets/sheet1.xml")
        assert len(relationships) == 0
        assert package.to_bytes() == minimal_xlsm_bytes
        assert not package.has_part("xl/worksheets/_rels/sheet1.xml.rels")

    def test_a_used_rels_part_is_created(self, package: OpcPackage) -> None:
        package.relationships("xl/worksheets/sheet1.xml").add(
            RT_HYPERLINK, "https://example.invalid/", external=True
        )
        reopened = OpcPackage.from_bytes(package.to_bytes())
        assert reopened.has_part("xl/worksheets/_rels/sheet1.xml.rels")
        link = next(iter(reopened.relationships("xl/worksheets/sheet1.xml")))
        assert link.target == "https://example.invalid/"
        assert link.is_external

    def test_a_new_parts_relationships_survive_writing_it_again(self, package: OpcPackage) -> None:
        """Writing a part dropped its relationships' wrapper, and the next
        lookup made an empty ``.rels`` over the one not yet saved: every
        picture added to a new drawing lost its image but the last."""
        package.write("xl/drawings/drawing9.xml", b"<wsDr/>", content_type="application/xml")
        package.relationships("xl/drawings/drawing9.xml").add_part(RT_WORKSHEET, "xl/media/image1.png")
        package.write("xl/drawings/drawing9.xml", b"<wsDr/>", content_type="application/xml")
        package.relationships("xl/drawings/drawing9.xml").add_part(RT_WORKSHEET, "xl/media/image2.png")
        targets = [r.target_part for r in package.relationships("xl/drawings/drawing9.xml")]
        assert targets == ["xl/media/image1.png", "xl/media/image2.png"]

    def test_an_external_target_refuses_to_resolve_to_a_part(self, package: OpcPackage) -> None:
        link = package.relationships("xl/worksheets/sheet1.xml").add(
            RT_HYPERLINK, "https://example.invalid/", external=True
        )
        with pytest.raises(PackageError, match="is external"):
            _ = link.target_part

    def test_add_part_writes_a_folder_relative_target(self, package: OpcPackage) -> None:
        added = package.relationships("xl/workbook.xml").add_part(RT_WORKSHEET, "xl/worksheets/sheet2.xml")
        assert added.target == "worksheets/sheet2.xml"
        assert added.target_part == "xl/worksheets/sheet2.xml"

    def test_a_duplicate_relationship_id_is_refused(self, package: OpcPackage) -> None:
        with pytest.raises(PackageError, match="already has a relationship"):
            package.relationships().add(RT_HYPERLINK, "x", relationship_id="rId1")

    def test_removing_a_relationship(self, package: OpcPackage) -> None:
        package.relationships("xl/workbook.xml").remove("rId1")
        reopened = OpcPackage.from_bytes(package.to_bytes())
        assert all(r.id != "rId1" for r in reopened.relationships("xl/workbook.xml"))

    def test_a_relationship_target_can_be_repointed(self, package: OpcPackage) -> None:
        relationships = package.relationships("xl/workbook.xml")
        sheet = relationships.by_type(RT_WORKSHEET)[0]
        sheet.target = "worksheets/renamed.xml"
        reopened = OpcPackage.from_bytes(package.to_bytes())
        assert reopened.relationships("xl/workbook.xml").by_type(RT_WORKSHEET)[0].target_part == (
            "xl/worksheets/renamed.xml"
        )


class TestRicherExcelAuthoredPackage:
    """Gates against the fixture real Excel builds on demand.

    It carries what the committed fixtures do not: a sharedStrings part,
    formula cells with cached values, dates, booleans, an error cell and a
    merged range. These skip until scripts/build_excel_fixtures.py has run.
    """

    def test_it_round_trips_unchanged(self, live_sample_xlsx: Path) -> None:
        data = live_sample_xlsx.read_bytes()
        assert OpcPackage.from_bytes(data).to_bytes() == data

    def test_it_has_the_parts_the_surface_work_needs(self, live_sample_xlsx: Path) -> None:
        package = OpcPackage.open(live_sample_xlsx)
        names = package.part_names()
        assert any("sharedStrings" in n for n in names), names
        assert sum(1 for n in names if n.startswith("xl/worksheets/sheet")) == 2, names

    def test_every_internal_relationship_points_at_a_real_part(self, live_sample_xlsx: Path) -> None:
        package = OpcPackage.open(live_sample_xlsx)
        for source in ["", *package.part_names()]:
            for relationship in package.relationships(source):
                if not relationship.is_external:
                    assert package.has_part(relationship.target_part), (
                        f"{source or 'root'} -> {relationship.target}"
                    )

    def test_every_part_has_a_content_type(self, live_sample_xlsx: Path) -> None:
        """A part Office cannot resolve a content type for is a part Office
        ignores."""
        package = OpcPackage.open(live_sample_xlsx)
        for name in package.part_names():
            if name == CONTENT_TYPES_PART:
                continue
            assert package.content_types.of(name) is not None, name

    def test_relationship_order_is_not_sheet_order(self, live_sample_xlsx: Path) -> None:
        """The other fact the cell layer has to be built around.

        Excel wrote ``rId2 -> worksheets/sheet2.xml`` before
        ``rId1 -> worksheets/sheet1.xml`` in the workbook's rels part, so
        indexing the worksheet relationships picks the wrong sheet. Sheet
        order and sheet names live in ``<sheets>`` inside the workbook part,
        and each entry points at its relationship by ``r:id``.
        """
        package = OpcPackage.open(live_sample_xlsx)
        workbook_part = package.main_document_part()
        by_relationship = [
            r.target_part for r in package.relationships(workbook_part).by_type(RT_WORKSHEET)
        ]
        by_declaration = [
            sheet_part_by_name(package, entry.get("name") or "")
            for entry in package.xml(workbook_part).root.require("sheets").children_named("sheet")
        ]
        assert by_declaration == ["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"]
        assert by_relationship != by_declaration, (
            "the fixture exists to prove these disagree; if Excel ever stops "
            "reordering them, the helper is still the correct navigation"
        )
        assert sorted(by_relationship) == sorted(by_declaration), "same parts, different order"

    def test_a_range_assigned_formula_is_stored_as_a_shared_formula(
        self, live_sample_xlsx: Path
    ) -> None:
        """The fact the cell layer has to be built around.

        The fixture assigns ``=B2*C2`` to D2:D5 in one statement. Excel
        stores the text once, on the first cell, and leaves the rest
        pointing at it by index:

            D2: <f t="shared" ref="D2:D5" si="0">B2*C2</f><v>510</v>
            D3: <f t="shared" si="0"/><v>1445</v>

        So three of those four cells carry a cached value and no formula
        text at all. A cell-level reader that returns ``<f>``'s text would
        report an empty formula for D3, D4 and D5, which is wrong rather
        than merely incomplete: the formula must be translated from the
        master cell's, shifting its relative references.
        """
        package = OpcPackage.open(live_sample_xlsx)
        part = sheet_part_by_name(package, "Data")
        cells = {
            c.get("r"): c
            for c in package.xml(part).root.require("sheetData").descendants("c")
        }

        master = cells["D2"].require("f")
        assert master.get("t") == "shared"
        assert master.get("ref") == "D2:D5"
        assert master.text == "B2*C2"

        for reference, cached in (("D3", "1445"), ("D4", "522.5"), ("D5", "787.5")):
            follower = cells[reference].require("f")
            assert follower.get("t") == "shared"
            assert follower.get("si") == master.get("si")
            assert follower.text == "", f"{reference} carries no formula text of its own"
            assert cells[reference].require("v").text == cached

        # A formula entered into a single cell is stored plainly.
        assert cells["D6"].require("f").text == "SUM(D2:D5)"
        assert cells["D6"].require("f").get("t") is None

    def test_the_cell_types_the_surface_layer_must_handle(self, live_sample_xlsx: Path) -> None:
        """Excel's cell type attribute, as it actually appears.

        A string is an index into sharedStrings, a boolean is 1 or 0, an
        error carries its text, and a date is a plain serial number whose
        only clue is the style index. None of these can be read by looking
        at ``<v>`` alone.
        """
        package = OpcPackage.open(live_sample_xlsx)
        part = sheet_part_by_name(package, "Data")
        cells = {
            c.get("r"): c
            for c in package.xml(part).root.require("sheetData").descendants("c")
        }

        assert cells["A2"].get("t") == "s", "a string is an index, not text"
        assert cells["A2"].require("v").text.isdigit()
        assert cells["B2"].get("t") is None, "a number carries no type at all"
        assert cells["E2"].get("t") == "b" and cells["E2"].require("v").text == "1"
        assert cells["E3"].require("v").text == "0"
        assert cells["B8"].get("t") == "e" and cells["B8"].require("v").text == "#DIV/0!"
        assert cells["F2"].get("t") is None, "a date is a number"
        assert cells["F2"].get("s") is not None, "only the style says it is a date"
        assert cells["F2"].require("v").text == "46037"


class TestInteroperability:
    def test_the_standard_library_reads_what_we_write(self, package: OpcPackage) -> None:
        package.xml("xl/worksheets/sheet1.xml").root.require("dimension").set("ref", "A1:D4")
        reference = zipfile.ZipFile(io.BytesIO(package.to_bytes()))
        assert reference.testzip() is None
        assert b"A1:D4" in reference.read("xl/worksheets/sheet1.xml")

    def test_openpyxl_reads_a_package_we_edited(self, openpyxl_xlsx_bytes: bytes, tmp_path: Path) -> None:
        openpyxl = pytest.importorskip("openpyxl")
        package = OpcPackage.from_bytes(openpyxl_xlsx_bytes)
        part = sheet_part_by_name(package, "Data")
        # Add a row through the XML layer, then let openpyxl read it back.
        sheet_data = package.xml(part).root.require("sheetData")
        row = Element.create("row", {"r": "9"})
        cell = Element.create("c", {"r": "B9"})
        value = Element.create("v")
        value.set_text("777")
        cell.append(value)
        row.append(cell)
        sheet_data.append(row)

        target = tmp_path / "edited.xlsx"
        target.write_bytes(package.to_bytes())
        workbook = openpyxl.load_workbook(target)
        assert workbook["Data"]["B9"].value == 777
        assert workbook["Data"]["A2"].value == "North", "the original content survived"
