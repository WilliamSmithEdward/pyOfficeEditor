"""Notes on cells, which the file and Excel's object model call comments.

``comments.xlsx`` was built by Excel with notes of every shape, and
``comments_answers.json`` beside it records what Excel's object model said
of each. Reading is held to that. Writing is held to the markup Excel wrote
when measured: the text in the comments part in Excel's own run, a box in
the VML part placed as Excel places one, and ``<legacyDrawing>`` where the
schema puts it. The live gate has Excel open what is written here.
"""

from __future__ import annotations

import datetime as dt
import json
import re
import shutil
from pathlib import Path

import pytest

from pyofficeeditor.excel import Comment, Reply, ThreadedComment, Workbook, Worksheet
from pyofficeeditor.excel._comments import (
    CT_PERSONS,
    CT_THREADED_COMMENTS,
    RT_PERSONS,
    RT_THREADED_COMMENTS,
    format_moment,
    next_id,
    note_anchor,
    parse_moment,
)
from pyofficeeditor.excel._reference import CellRef

#: What Excel wrote for a thread with two replies and a resolved one,
#: measured, with the person Excel took from the signed-in account swapped
#: for one no account stands behind.
EXCELS_THREADS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"'
    ' xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<threadedComment ref="B2" dT="2026-09-23T02:58:08.55" personId="{0FB1E72A-1D4A-4F0E-9AAA-118B4066800E}"'
    ' id="{B2151283-40E3-491D-9AF4-5D80C1AFE97A}" done="1"><text>resolved one</text></threadedComment>'
    '<threadedComment ref="C3" dT="2026-09-23T02:59:22.45" personId="{0FB1E72A-1D4A-4F0E-9AAA-118B4066800E}"'
    ' id="{02CFA75F-6C4A-4847-A343-D8EF98CA9187}"><text>first line\r\nsecond line</text></threadedComment>'
    '<threadedComment ref="C3" dT="2026-09-23T02:59:22.46" personId="{0FB1E72A-1D4A-4F0E-9AAA-118B4066800E}"'
    ' id="{00324A76-90C2-4B36-A188-F7D680B5082B}" parentId="{02CFA75F-6C4A-4847-A343-D8EF98CA9187}">'
    "<text>reply one</text></threadedComment>"
    '<threadedComment ref="C3" dT="2026-09-23T02:59:22.46" personId="{0FB1E72A-1D4A-4F0E-9AAA-118B4066800E}"'
    ' id="{00324A76-90C2-4B37-A188-F7D680B5082B}" parentId="{02CFA75F-6C4A-4847-A343-D8EF98CA9187}">'
    "<text>reply two</text></threadedComment></ThreadedComments>"
)
EXCELS_PERSONS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"'
    ' xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<person displayName="Ada" id="{0FB1E72A-1D4A-4F0E-9AAA-118B4066800E}" userId="Ada" providerId="None"/>'
    "</personList>"
)
LATER = dt.datetime(2026, 9, 22, 18, 30, tzinfo=dt.timezone.utc)


@pytest.fixture(scope="module")
def answers(live_comments_answers: Path) -> dict[str, dict[str, object]]:
    return json.loads(live_comments_answers.read_text(encoding="utf-8"))


@pytest.fixture
def excels(tmp_path: Path, live_comments_xlsx: Path) -> Workbook:
    target = tmp_path / "comments.xlsx"
    shutil.copy(live_comments_xlsx, target)
    return Workbook.open(target)


@pytest.fixture
def data(tmp_path: Path, live_sample_xlsx: Path) -> Worksheet:
    target = tmp_path / "sample.xlsx"
    shutil.copy(live_sample_xlsx, target)
    return Workbook.open(target)["Data"]


def part_text(sheet: Worksheet, suffix: str) -> str:
    """The text of the one part of a kind the package holds."""
    package = sheet.workbook.package
    names = [name for name in package.part_names() if name.endswith(suffix)]
    assert len(names) == 1, names
    return package.read(names[0]).decode("utf-8")


def vml_of(sheet: Worksheet) -> str:
    package = sheet.workbook.package
    for relationship in package.relationships(sheet.part_name).by_type(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
    ):
        return package.read(relationship.target_part).decode("utf-8")
    raise AssertionError("the sheet has no VML part")


class TestReadingExcelsOwn:
    def test_every_note_as_excel_reported_it(
        self, excels: Workbook, answers: dict[str, dict[str, object]]
    ) -> None:
        seen = {
            f"{sheet.name}!{comment.ref}": comment
            for sheet in excels
            for comment in sheet.comments
        }
        assert set(seen) == set(answers)
        for key, said in answers.items():
            comment = seen[key]
            assert (comment.text, comment.author, comment.visible) == (
                said["text"], said["author"], said["visible"]
            ), key

    def test_a_note_in_two_runs_reads_as_one_text(self, excels: Workbook) -> None:
        """Excel split it where the bold stops."""
        assert excels["Notes"].comment("D10") == Comment("D10", "bold start", "William")

    def test_one_cell(self, excels: Workbook) -> None:
        sheet = excels["Notes"]
        assert sheet["E5"].comment == Comment("E5", "two\nlines", "William", visible=True)
        assert sheet["B2"].comment is None
        assert sheet.comment("$C$3") == sheet.comment("C3")


class TestWhereExcelPutsTheBox:
    """Measured on Excel 16 at 100% scaling: 144 by 79 pixels, 15 right of
    the cell and 10 above it, or 2 below the top edge on the first row."""

    COLUMNS = [64] * 30
    ROWS = [20] * 30

    @pytest.mark.parametrize(
        ("cell", "anchor", "left", "top"),
        [
            ("A1", (1, 15, 0, 2, 3, 31, 4, 1), 59.25, 1.5),
            ("B1", (2, 15, 0, 2, 4, 31, 4, 1), 107.25, 1.5),
            ("A5", (1, 15, 3, 10, 3, 31, 7, 9), 59.25, 52.5),
            ("C3", (3, 15, 1, 10, 5, 31, 5, 9), 155.25, 22.5),
            ("E5", (5, 15, 3, 10, 7, 31, 7, 9), 251.25, 52.5),
        ],
    )
    def test_on_the_default_grid(
        self, cell: str, anchor: tuple[int, ...], left: float, top: float
    ) -> None:
        assert note_anchor(CellRef.parse(cell), self.COLUMNS, self.ROWS) == (anchor, left, top)

    def test_beside_resized_columns_and_rows(self) -> None:
        """Column D at 20.71 characters is 145 pixels and E at 3.71 is 26;
        row 2 at 30 points is 40 pixels and row 4 at 6 is 8."""
        columns = [64, 64, 64, 145, 26] + [64] * 20
        rows = [20, 40, 20, 8] + [20] * 20
        assert note_anchor(CellRef.parse("C3"), columns, rows) == ((3, 15, 1, 30, 4, 14, 6, 1), 155.25, 37.5)


class TestWriting:
    def test_a_first_note_makes_its_parts(self, data: Worksheet) -> None:
        data.set_comment("C3", "plain note", author="Ada")
        assert data.comments == [Comment("C3", "plain note", "Ada")]
        comments = part_text(data, "comments1.xml")
        assert '<comment ref="C3" authorId="0" shapeId="0">' in comments
        assert '<rFont val="Tahoma"/>' in comments, "Excel's own note run"
        vml = vml_of(data)
        assert 'ObjectType="Note"' in vml
        assert "<x:Row>2</x:Row>" in vml and "<x:Column>2</x:Column>" in vml
        assert data.document.root.child("legacyDrawing") is not None
        types = data.workbook.package.read("[Content_Types].xml").decode("utf-8")
        assert "spreadsheetml.comments+xml" in types
        assert 'Extension="vml"' in types

    def test_notes_are_kept_in_cell_order(self, data: Worksheet) -> None:
        """As Excel keeps them, whatever order they were added in."""
        for ref in ("E5", "A1", "C3", "B5"):
            data.set_comment(ref, ref)
        assert [comment.ref for comment in data.comments] == ["A1", "C3", "B5", "E5"]

    def test_authors_are_listed_once(self, data: Worksheet) -> None:
        data.set_comment("C3", "one", author="Ada")
        data.set_comment("C4", "two", author="Bo")
        data.set_comment("C5", "three", author="Ada")
        assert [comment.author for comment in data.comments] == ["Ada", "Bo", "Ada"]
        assert part_text(data, "comments1.xml").count("<author>Ada</author>") == 1

    def test_replacing_keeps_the_box(self, data: Worksheet) -> None:
        data.set_comment("C3", "first")
        before = vml_of(data)
        data.set_comment("C3", "second", visible=True)
        assert data.comment("C3") == Comment("C3", "second", visible=True)
        after = vml_of(data)
        assert after.count("<v:shape ") == before.count("<v:shape ") == 1
        assert "<x:Visible/>" in after and "visibility:visible" in after
        data.set_comment("C3", "third")
        assert "<x:Visible/>" not in vml_of(data)

    def test_text_keeps_its_lines_and_its_spaces(self, data: Worksheet) -> None:
        data.set_comment("C3", "two\nlines")
        data.set_comment("C4", "  padded  ")
        assert data.comment("C3") == Comment("C3", "two\nlines")
        assert data.comment("C4") == Comment("C4", "  padded  ")

    def test_through_a_cell(self, data: Worksheet) -> None:
        data["B2"].comment = "set here"
        assert data["B2"].comment == Comment("B2", "set here")
        data["B2"].comment = None
        assert data["B2"].comment is None

    def test_it_survives_a_save(self, data: Worksheet, tmp_path: Path) -> None:
        data.set_comment("C3", "plain", author="Ada")
        data.set_comment("E5", "shown", visible=True)
        data.workbook.save()
        reopened = Workbook.open(tmp_path / "sample.xlsx")["Data"]
        assert reopened.comments == [Comment("C3", "plain", "Ada"), Comment("E5", "shown", visible=True)]


class TestRemoving:
    def test_the_last_note_takes_its_parts_with_it(self, data: Worksheet) -> None:
        package = data.workbook.package
        # Any edit drops the calculation chain, which Excel rebuilds.
        data["A40"].value = 1
        parts_before = set(package.part_names())
        data.set_comment("C3", "gone soon")
        assert data.remove_comment("C3") is True
        assert set(package.part_names()) == parts_before
        assert data.document.root.child("legacyDrawing") is None
        assert data.comments == []

    def test_one_of_several(self, data: Worksheet) -> None:
        data.set_comment("C3", "stays")
        data.set_comment("C4", "goes")
        data.remove_comment("C4")
        assert data.comments == [Comment("C3", "stays")]
        assert vml_of(data).count("<v:shape ") == 1

    def test_nothing_to_remove(self, data: Worksheet) -> None:
        assert data.remove_comment("C3") is False
        data.set_comment("C3", "here")
        assert data.remove_comment("D9") is False

    def test_excels_own(self, excels: Workbook) -> None:
        sheet = excels["Notes"]
        assert sheet.remove_comment("D10") is True
        assert [comment.ref for comment in sheet.comments] == ["A1", "C3", "E5", "B8"]
        assert "<x:Row>9</x:Row>" not in vml_of(sheet)


class TestBesideFormControls:
    """Measured: a sheet's notes and form controls share its VML part and
    one sequence of shape ids, from the block the part claims."""

    def test_a_note_and_a_button_take_consecutive_ids(self, data: Worksheet) -> None:
        data.set_comment("B2", "first")
        data.add_form_control("Press", left=150, top=40, width=60, height=20, text="Click")
        data.set_comment("D6", "after the button")
        ids = [int(number) for number in re.findall(r"_x0000_s(\d+)", vml_of(data))]
        assert sorted(ids) == [1025, 1026, 1027]
        assert data.shape("Press").shape_id == 1026

    def test_on_excels_sheet_with_a_button(self, excels: Workbook) -> None:
        """Excel's second sheet claims block 2, where the button is 2049
        and the note 2050."""
        sheet = excels["Both"]
        sheet.set_comment("D4", "another")
        sheet.add_form_control("Again", left=200, top=100, width=60, height=20, text="Again")
        assert "_x0000_s2051" in vml_of(sheet)
        assert sheet.shape("Again").shape_id == 2052

    def test_removing_the_note_keeps_the_button(self, data: Worksheet) -> None:
        data.add_form_control("Press", left=150, top=40, width=60, height=20, text="Click")
        data.set_comment("B2", "brief")
        data.remove_comment("B2")
        assert data.document.root.child("legacyDrawing") is not None
        assert 'ObjectType="Button"' in vml_of(data)
        assert 'ObjectType="Note"' not in vml_of(data)

    def test_a_second_sheet_claims_a_block_of_its_own(self, data: Worksheet) -> None:
        data.set_comment("C3", "first sheet")
        notes = data.workbook["Notes"]
        notes.set_comment("B2", "second sheet")
        assert 'data="2"' in vml_of(notes)
        assert "_x0000_s2049" in vml_of(notes)


class TestWhereLegacyDrawingGoes:
    def test_before_a_tables_parts(self, tmp_path: Path, live_structures_xlsx: Path) -> None:
        """``CT_Worksheet`` puts it after ``drawing`` and before ``tableParts``.
        Missing from the order, it was appended after them, and Excel
        refused every sheet with a table that got a note or a control."""
        target = tmp_path / "structures.xlsx"
        shutil.copy(live_structures_xlsx, target)
        sheet = Workbook.open(target)["Tabled"]
        sheet.set_comment("E10", "beside a table")
        sheet.add_form_control("Press", left=300, top=150, width=60, height=20)
        names = [child.name for child in sheet.document.root.elements()]
        assert names.index("drawing") < names.index("legacyDrawing") < names.index("tableParts")


class TestRowsAndColumnsMoving:
    def test_notes_follow_their_cells(self, data: Worksheet) -> None:
        data.set_comment("C3", "moves")
        data.insert_rows(2, 2)
        data.insert_columns(1)
        assert data.comments == [Comment("D5", "moves")]
        vml = vml_of(data)
        assert "<x:Row>4</x:Row>" in vml and "<x:Column>3</x:Column>" in vml

    def test_a_deleted_cells_note_goes_box_and_all(self, data: Worksheet) -> None:
        data.set_comment("C3", "deleted with its row")
        data.set_comment("C5", "stays")
        data.delete_rows(3)
        assert data.comments == [Comment("C4", "stays")]
        vml = vml_of(data)
        assert vml.count('ObjectType="Note"') == 1
        assert "<x:Row>3</x:Row>" in vml


def with_excels_threads(sheet: Worksheet) -> None:
    """Give a sheet the threads part and the person list Excel wrote."""
    package = sheet.workbook.package
    package.write("xl/threadedComments/threadedComment1.xml", EXCELS_THREADS.encode(), content_type=CT_THREADED_COMMENTS)
    package.relationships(sheet.part_name).add_part(RT_THREADED_COMMENTS, "xl/threadedComments/threadedComment1.xml")
    package.write("xl/persons/person.xml", EXCELS_PERSONS.encode(), content_type=CT_PERSONS)
    package.relationships(sheet.workbook.workbook_part).add_part(RT_PERSONS, "xl/persons/person.xml")


class TestReadingThreads:
    def test_excels_markup(self, data: Worksheet) -> None:
        with_excels_threads(data)
        at = dt.timezone.utc
        assert data.threaded_comments == [
            ThreadedComment("B2", "resolved one", "Ada", dt.datetime(2026, 9, 23, 2, 58, 8, 550000, at), resolved=True),
            ThreadedComment(
                "C3", "first line\nsecond line", "Ada", dt.datetime(2026, 9, 23, 2, 59, 22, 450000, at),
                replies=(
                    Reply("reply one", "Ada", dt.datetime(2026, 9, 23, 2, 59, 22, 460000, at)),
                    Reply("reply two", "Ada", dt.datetime(2026, 9, 23, 2, 59, 22, 460000, at)),
                ),
            ),
        ]
        assert data.threaded_comment("C3") is not None
        assert data.threaded_comment("D4") is None

    def test_a_sheet_with_none(self, data: Worksheet) -> None:
        assert data.threaded_comments == []


class TestWritingThreads:
    def test_a_first_thread_makes_its_parts(self, data: Worksheet) -> None:
        thread = data.add_threaded_comment("C3", "first line\nsecond line", author="Ada", when=LATER)
        assert thread == ThreadedComment("C3", "first line\nsecond line", "Ada", LATER)
        assert data.threaded_comments == [thread]
        threads = part_text(data, "threadedComment1.xml")
        assert 'dT="2026-09-22T18:30:00.00"' in threads
        assert "<text>first line\r\nsecond line</text>" in threads, "CRLF, as Excel stores it"
        persons = part_text(data, "person.xml")
        assert 'displayName="Ada"' in persons and 'providerId="None"' in persons

    def test_the_placeholder_note_beside_it(self, data: Worksheet) -> None:
        """Excel's own words, the thread's text, then each reply's, with
        the thread's id as its author and its ``xr:uid``."""
        data.add_threaded_comment("C3", "start", author="Ada", when=LATER)
        data.add_threaded_reply("C3", "reply one", author="Bo", when=LATER)
        comments = part_text(data, "comments1.xml")
        identifier = re.search(r'<author>tc=(\{[^}]+\})</author>', comments)
        assert identifier is not None
        assert f'xr:uid="{identifier.group(1)}"' in comments
        assert (
            "Learn more: https://go.microsoft.com/fwlink/?linkid=870924\r\n\r\n"
            "Comment:\r\n    start\r\nReply:\r\n    reply one</t>"
        ) in comments
        assert 'ObjectType="Note"' in vml_of(data), "and a box, as any note has"
        assert data.comments == [], "the placeholder is not a note of its own"

    def test_replies_keep_their_order(self, data: Worksheet) -> None:
        """Measured: Excel lists replies written at one moment in the order
        of their ids, so each takes the id after the last."""
        data.add_threaded_comment("C3", "start", author="Ada", when=LATER)
        data.add_threaded_reply("C3", "one", author="Bo", when=LATER)
        data.add_threaded_reply("C3", "two", author="Ada", when=LATER)
        thread = data.threaded_comment("C3")
        assert thread is not None
        assert [reply.text for reply in thread.replies] == ["one", "two"]
        ids = re.findall(r'id="(\{[^}]+\})" parentId', part_text(data, "threadedComment1.xml"))
        assert ids == sorted(ids) and len(ids) == 2

    def test_threads_are_kept_in_cell_order(self, data: Worksheet) -> None:
        for ref in ("E5", "A1", "C3"):
            data.add_threaded_comment(ref, ref, author="Ada", when=LATER)
        assert [thread.ref for thread in data.threaded_comments] == ["A1", "C3", "E5"]

    def test_the_person_list_is_shared(self, data: Worksheet) -> None:
        data.add_threaded_comment("C3", "one", author="Ada", when=LATER)
        data.add_threaded_reply("C3", "two", author="Ada", when=LATER)
        data.workbook["Notes"].add_threaded_comment("A1", "three", author="Ada", when=LATER)
        assert part_text(data, "person.xml").count("<person ") == 1

    def test_resolving(self, data: Worksheet) -> None:
        data.add_threaded_comment("C3", "start", author="Ada", when=LATER)
        assert data.resolve_threaded_comment("C3").resolved is True
        assert 'done="1"' in part_text(data, "threadedComment1.xml")
        assert data.resolve_threaded_comment("C3", False).resolved is False

    def test_a_time_without_a_zone_is_utc(self, data: Worksheet) -> None:
        data.add_threaded_comment("C3", "start", author="Ada", when=dt.datetime(2026, 1, 2, 3, 4, 5))
        assert 'dT="2026-01-02T03:04:05.00"' in part_text(data, "threadedComment1.xml")

    def test_it_survives_a_save(self, data: Worksheet, tmp_path: Path) -> None:
        data.add_threaded_comment("C3", "start", author="Ada", when=LATER)
        data.add_threaded_reply("C3", "reply", author="Bo", when=LATER)
        data.workbook.save()
        reopened = Workbook.open(tmp_path / "sample.xlsx")["Data"]
        assert reopened.threaded_comments == data.threaded_comments


class TestThreadRefusals:
    def test_a_note_on_a_threads_cell(self, data: Worksheet) -> None:
        data.add_threaded_comment("C3", "start", author="Ada", when=LATER)
        with pytest.raises(ValueError, match="threaded comment"):
            data.set_comment("C3", "a note")

    def test_a_second_thread_or_a_thread_on_a_note(self, data: Worksheet) -> None:
        data.add_threaded_comment("C3", "start", author="Ada", when=LATER)
        data.set_comment("D4", "a note")
        for ref in ("C3", "D4"):
            with pytest.raises(ValueError, match="already has a comment"):
                data.add_threaded_comment(ref, "again", author="Ada")

    def test_replying_or_resolving_with_no_thread(self, data: Worksheet) -> None:
        with pytest.raises(ValueError, match="no thread"):
            data.add_threaded_reply("C3", "hello", author="Ada")
        with pytest.raises(ValueError, match="no thread"):
            data.resolve_threaded_comment("C3")


class TestRemovingThreads:
    def test_a_thread_goes_with_its_replies_placeholder_and_box(self, data: Worksheet) -> None:
        package = data.workbook.package
        data["A40"].value = 1
        parts_before = set(package.part_names())
        data.add_threaded_comment("C3", "start", author="Ada", when=LATER)
        data.add_threaded_reply("C3", "reply", author="Bo", when=LATER)
        assert data.remove_comment("C3") is True
        assert data.threaded_comments == []
        # The person list stays, as Excel keeps it.
        assert set(package.part_names()) - parts_before == {"xl/persons/person.xml"}

    def test_one_of_two(self, data: Worksheet) -> None:
        data.add_threaded_comment("C3", "goes", author="Ada", when=LATER)
        data.add_threaded_comment("D4", "stays", author="Ada", when=LATER)
        data.remove_comment("C3")
        assert [thread.ref for thread in data.threaded_comments] == ["D4"]
        assert vml_of(data).count('ObjectType="Note"') == 1


class TestThreadsMoving:
    def test_a_thread_follows_its_cell(self, data: Worksheet) -> None:
        data.add_threaded_comment("C3", "start", author="Ada", when=LATER)
        data.add_threaded_reply("C3", "reply", author="Bo", when=LATER)
        data.insert_rows(2)
        thread = data.threaded_comment("C4")
        assert thread is not None and [reply.text for reply in thread.replies] == ["reply"]
        assert data.comment("C4") is None, "its placeholder moved with it, and is still not a note"
        assert "tc=" in part_text(data, "comments1.xml")

    def test_a_deleted_cells_thread_goes(self, data: Worksheet) -> None:
        data.add_threaded_comment("C3", "goes", author="Ada", when=LATER)
        data.add_threaded_comment("C5", "stays", author="Ada", when=LATER)
        data.delete_rows(3)
        assert [thread.ref for thread in data.threaded_comments] == ["C4"]


class TestMoments:
    def test_excels_form(self) -> None:
        moment = parse_moment("2026-09-23T02:59:22.45")
        assert moment == dt.datetime(2026, 9, 23, 2, 59, 22, 450000, dt.timezone.utc)
        assert moment is not None and format_moment(moment) == "2026-09-23T02:59:22.45"

    def test_another_zone_is_turned_to_utc(self) -> None:
        pacific = dt.timezone(dt.timedelta(hours=-7))
        assert format_moment(dt.datetime(2026, 9, 22, 19, 59, 22, tzinfo=pacific)) == "2026-09-23T02:59:22.00"

    def test_nonsense_reads_as_no_time(self) -> None:
        assert parse_moment("yesterday") is None
        assert parse_moment(None) is None

    def test_the_next_id_sorts_after(self) -> None:
        assert next_id("{00324A76-90C2-4B36-A188-F7D680B5082B}") == "{00324A76-90C2-4B36-A188-F7D680B5082C}"
        assert next_id("not an id").startswith("{")
