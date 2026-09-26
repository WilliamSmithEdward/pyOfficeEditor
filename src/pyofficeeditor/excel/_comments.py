"""Cell comments: notes, the older kind, and threaded comments.

A note is two parts that have to agree, and a third the sheet points at:

- the comments part, which holds the text and who wrote it::

      <comments><authors><author>Ada</author></authors><commentList>
        <comment ref="C3" authorId="0" shapeId="0"><text><r><rPr>...</rPr>
          <t>plain note</t></r></text></comment></commentList></comments>

- the VML part, which holds the yellow box Excel draws: where it sits,
  whether it shows, and which cell it belongs to, stored again as a
  zero-based ``<x:Row>`` and ``<x:Column>``
- ``<legacyDrawing>`` on the sheet, naming the VML part, which a sheet with
  form controls already has: the notes then share its VML part

All of the markup here is what Excel wrote, measured: the Tahoma 9 run
every note's text is written in, ``shapeId="0"`` on every comment, a line
break stored as CRLF, and a box 144 by 79 pixels placed 15 pixels right of
the cell and 10 above it. The text of a note is read from its runs;
formatting inside a note is not modelled, and a note written here carries
Excel's own.

A threaded comment, Excel's newer kind, is a conversation on a cell: a
first comment and its replies, each with an author and a time, in a part
of its own, and the authors in a person list the whole workbook shares.
Excel also writes a placeholder note beside each thread, so a version that
cannot show threads has something to show, and draws it a box like any
note's. All four are written here as Excel writes them, and a thread's
placeholder is kept up to date with its replies.
"""

from __future__ import annotations

import datetime as dt
import re
import uuid
from dataclasses import dataclass

from pyofficeeditor._xml import Element, XmlDocument, escape_attribute, escape_text
from pyofficeeditor.excel._reference import CellRef
from pyofficeeditor.excel._sharedstrings import needs_space_preserved
from pyofficeeditor.excel._xstring import decode, encode_attribute, encode_text

CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
RT_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
CT_THREADED_COMMENTS = "application/vnd.ms-excel.threadedcomments+xml"
RT_THREADED_COMMENTS = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
CT_PERSONS = "application/vnd.ms-excel.person+xml"
RT_PERSONS = "http://schemas.microsoft.com/office/2017/10/relationships/person"
NS_THREADED_COMMENTS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"

#: A comments part with nothing in it yet, as Excel starts one, with the
#: revision namespace a thread's placeholder names itself in.
EMPTY_COMMENTS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr"'
    ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">'
    "<authors/><commentList/></comments>"
)

#: The same for a sheet's threads and the workbook's people.
EMPTY_THREADED_COMMENTS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    f'<ThreadedComments xmlns="{NS_THREADED_COMMENTS}"'
    ' xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'
)
EMPTY_PERSONS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    f'<personList xmlns="{NS_THREADED_COMMENTS}"'
    ' xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'
)

#: What Excel writes at the head of the note it keeps beside a thread.
#: The thread's text follows, then each reply's under ``Reply:``.
PLACEHOLDER_HEAD = (
    "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, "
    "any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: "
    "https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    "
)

#: The run Excel writes every note's text in.
_NOTE_RUN = '<rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr>'

#: The size Excel gives a new note, in pixels: 108 by 59.25 points.
NOTE_WIDTH_PIXELS = 144
NOTE_HEIGHT_PIXELS = 79

#: The shape type a note's box is drawn from, as Excel writes it.
NOTE_SHAPE_TYPE = (
    '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"\r\n'
    '  path="m,l,21600r21600,l21600,xe">\r\n'
    '  <v:stroke joinstyle="miter"/>\r\n'
    '  <v:path gradientshapeok="t" o:connecttype="rect"/>\r\n'
    " </v:shapetype>"
)

#: One shape in a VML part, whatever it draws.
_VML_SHAPES = re.compile(r"<v:shape\b.*?</v:shape>", re.DOTALL)


@dataclass(frozen=True)
class Comment:
    """A note on a cell: its text, who wrote it, and whether it shows
    without the pointer over the cell."""

    ref: str
    text: str
    author: str = ""
    visible: bool = False


def read_comments(root: Element) -> list[Comment]:
    """The notes a comments part holds, in the order it holds them, each
    shown as hidden: whether one shows is the VML's to say."""
    listed = root.child("authors")
    authors = [] if listed is None else [decode(author.text) for author in listed.children_named("author")]
    container = root.child("commentList")
    if container is None:
        return []
    found: list[Comment] = []
    for element in container.children_named("comment"):
        ref = element.get("ref")
        if ref is None:
            continue
        try:
            index = int(element.get("authorId") or "0")
        except ValueError:
            index = -1
        author = authors[index] if 0 <= index < len(authors) else ""
        found.append(Comment(ref=ref, text=comment_text(element), author=author))
    return found


def comment_text(element: Element) -> str:
    """A note's text: the runs joined, or the one ``<t>`` a note without
    runs carries.

    Measured: Excel writes a line break inside a note as a raw CRLF, and
    its object model reports it as one line feed, which is how the parser
    here reads it too. A carriage return of its own is ``_x000D_``, as in
    a cell's text.
    """
    text = element.child("text")
    if text is None:
        return ""
    pieces: list[str] = []
    for child in text.elements():
        local = child.name.rpartition(":")[2]
        if local == "t":
            pieces.append(decode(child.text))
        elif local == "r":
            run_text = child.child("t")
            if run_text is not None:
                pieces.append(decode(run_text.text))
    return "".join(pieces)


def write_comment(
    root: Element, ref: str, text: str, author: str, *, run: bool = True, uid: str | None = None
) -> None:
    """Put a note's text on a cell, replacing any it had, in Excel's run.

    A new note goes where Excel puts one: the part keeps its notes in cell
    order, row by row, whatever order they were added in. A thread's
    placeholder is written as Excel writes one: plain text with no run,
    named after its thread in ``xr:uid`` where the part declares ``xr``.
    """
    authors = _made(root, "authors")
    names = [decode(entry.text) for entry in authors.children_named("author")]
    if author in names:
        author_id = names.index(author)
    else:
        entry = Element.create("author")
        entry.set_text(encode_text(author))
        authors.append(entry)
        author_id = len(names)

    space = ' xml:space="preserve"' if needs_space_preserved(text) else ""
    body = f"<t{space}>{escape_text(encode_text(text))}</t>"
    if run:
        body = f"<r>{_NOTE_RUN}{body}</r>"
    # The parser here reads a prefix as part of the name, so the fragment
    # needs no declaration of its own; the part's root has one.
    named = f' xr:uid="{uid}"' if uid is not None and root.has("xmlns:xr") else ""
    markup = f'<comment ref="{ref}" authorId="{author_id}" shapeId="0"{named}><text>{body}</text></comment>'
    element = XmlDocument.parse(markup.encode("utf-8")).root
    container = _made(root, "commentList")
    existing = find_comment(root, ref)
    if existing is not None:
        container.insert_before(existing, element)
        container.remove(existing)
        return
    wanted = _position(ref)
    for other in container.children_named("comment"):
        if _position(other.get("ref") or "") > wanted:
            container.insert_before(other, element)
            return
    container.append(element)


def _position(ref: str) -> tuple[int, int]:
    """Where a cell sorts among a part's notes: by row, then column."""
    try:
        cell = CellRef.parse(ref)
    except ValueError:
        return (0, 0)
    return (cell.row, cell.column)


# ---------------------------------------------------------------------------
# Threaded comments
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Reply:
    """One reply in a thread."""

    text: str
    author: str = ""
    #: When it was written, in UTC, as Excel stores it.
    when: dt.datetime | None = None


@dataclass(frozen=True)
class ThreadedComment:
    """A conversation on a cell: the first comment, its replies in order,
    and whether it was marked resolved."""

    ref: str
    text: str
    author: str = ""
    when: dt.datetime | None = None
    replies: tuple[Reply, ...] = ()
    resolved: bool = False


def read_persons(root: Element) -> dict[str, str]:
    """The workbook's people, by id, as the names Excel shows."""
    return {
        person.get("id") or "": decode(person.get("displayName") or "")
        for person in root.children_named("person")
    }


def read_threads(root: Element, persons: dict[str, str]) -> list[ThreadedComment]:
    """The threads a sheet's part holds, each with its replies, in the
    order the part holds the first comments."""
    threads: dict[str, ThreadedComment] = {}
    for element in root.children_named("threadedComment"):
        identifier = element.get("id") or ""
        text_element = element.child("text")
        text = decode(text_element.text) if text_element is not None else ""
        author = persons.get(element.get("personId") or "", "")
        when = parse_moment(element.get("dT"))
        parent = element.get("parentId")
        if parent is not None and parent in threads:
            thread = threads[parent]
            threads[parent] = ThreadedComment(
                ref=thread.ref, text=thread.text, author=thread.author, when=thread.when,
                replies=(*thread.replies, Reply(text, author, when)), resolved=thread.resolved,
            )
            continue
        threads[identifier] = ThreadedComment(
            ref=element.get("ref") or "", text=text, author=author, when=when,
            resolved=(element.get("done") or "") in ("1", "true"),
        )
    return list(threads.values())


def thread_ids(root: Element) -> dict[str, str]:
    """Each thread's id, by the cell it is on."""
    return {
        element.get("ref") or "": element.get("id") or ""
        for element in root.children_named("threadedComment")
        if element.get("parentId") is None
    }


def write_thread(
    root: Element,
    ref: str,
    text: str,
    person: str,
    when: dt.datetime,
    *,
    parent: str | None = None,
) -> str:
    """Add a first comment, or a reply to one, as Excel writes it; its id.

    A first comment goes in cell order, as Excel keeps them, and a reply
    after the last entry of its thread. A reply's id is the one after the
    thread's last reply, as Excel numbers them: measured, Excel lists
    replies written at the same moment in the order of their ids, and a
    random id put a later reply first.
    """
    identifier = new_id()
    if parent is not None:
        replies = [
            entry.get("id") or ""
            for entry in root.children_named("threadedComment")
            if entry.get("parentId") == parent
        ]
        if replies:
            identifier = next_id(replies[-1])
    attributes = f'ref="{ref}" dT="{format_moment(when)}" personId="{person}" id="{identifier}"'
    if parent is not None:
        attributes += f' parentId="{parent}"'
    markup = f"<threadedComment {attributes}><text>{escape_text(encode_text(text))}</text></threadedComment>"
    element = XmlDocument.parse(markup.encode("utf-8")).root
    entries = list(root.children_named("threadedComment"))
    if parent is not None:
        members = [entry for entry in entries if parent in (entry.get("id"), entry.get("parentId"))]
        if members and members[-1] is not entries[-1]:
            following = entries[entries.index(members[-1]) + 1]
            root.insert_before(following, element)
        else:
            root.append(element)
        return identifier
    wanted = _position(ref)
    for entry in entries:
        if entry.get("parentId") is None and _position(entry.get("ref") or "") > wanted:
            root.insert_before(entry, element)
            return identifier
    root.append(element)
    return identifier


def set_resolved(root: Element, ref: str, resolved: bool) -> bool:
    """Mark a cell's thread resolved or open again; whether it has one."""
    for element in root.children_named("threadedComment"):
        if element.get("ref") == ref and element.get("parentId") is None:
            if resolved:
                element.set("done", "1")
            else:
                element.unset("done")
            return True
    return False


def delete_thread(root: Element, ref: str) -> bool:
    """Take a cell's thread out, replies and all; whether it had one."""
    found = [element for element in root.children_named("threadedComment") if element.get("ref") == ref]
    for element in found:
        root.remove(element)
    return bool(found)


def person_id(root: Element, name: str) -> str:
    """The id of the person Excel shows as ``name``, added if the list has
    none: a person no account stands behind, which Excel writes with
    ``providerId="None"`` and the name again as ``userId``."""
    for person in root.children_named("person"):
        if decode(person.get("displayName") or "") == name:
            return person.get("id") or ""
    identifier = new_id()
    quoted = escape_attribute(encode_attribute(name))
    markup = f'<person displayName="{quoted}" id="{identifier}" userId="{quoted}" providerId="None"/>'
    root.append(XmlDocument.parse(markup.encode("utf-8")).root)
    return identifier


def placeholder_text(thread: ThreadedComment) -> str:
    """The note Excel keeps beside a thread, for a version that cannot
    show one: a notice, the thread's text, then each reply's."""
    text = PLACEHOLDER_HEAD + thread.text
    for reply in thread.replies:
        text += "\nReply:\n    " + reply.text
    return text


def new_id() -> str:
    """An id in the form Excel gives a thread, a reply and a person."""
    return "{" + str(uuid.uuid4()).upper() + "}"


def next_id(previous: str) -> str:
    """The id after ``previous``, which sorts after it: Excel's own ids for
    successive replies count up by one."""
    try:
        number = uuid.UUID(previous.strip("{}")).int + 1
    except ValueError:
        return new_id()
    if number >= 1 << 128:
        return new_id()
    return "{" + str(uuid.UUID(int=number)).upper() + "}"


_MOMENT = re.compile(r"(\d{4})-(\d\d)-(\d\d)T(\d\d):(\d\d):(\d\d)(?:\.(\d+))?")


def parse_moment(raw: str | None) -> dt.datetime | None:
    """A thread's ``dT``, which Excel writes in UTC to the hundredth of a
    second: ``2026-09-23T02:59:22.45``."""
    if raw is None:
        return None
    found = _MOMENT.match(raw)
    if found is None:
        return None
    year, month, day, hour, minute, second = (int(part) for part in found.groups()[:6])
    fraction = found.group(7) or "0"
    microseconds = int((fraction + "000000")[:6])
    try:
        return dt.datetime(year, month, day, hour, minute, second, microseconds, tzinfo=dt.timezone.utc)
    except ValueError:
        return None


def format_moment(moment: dt.datetime) -> str:
    """A moment as Excel writes ``dT``: UTC, to the hundredth of a second.
    A moment with no time zone is taken as UTC already."""
    if moment.tzinfo is not None:
        moment = moment.astimezone(dt.timezone.utc)
    return moment.strftime("%Y-%m-%dT%H:%M:%S") + f".{moment.microsecond // 10000:02d}"


def find_comment(root: Element, ref: str) -> Element | None:
    container = root.child("commentList")
    if container is None:
        return None
    for element in container.children_named("comment"):
        if element.get("ref") == ref:
            return element
    return None


def delete_comment(root: Element, ref: str) -> bool:
    """Take a cell's note out of the part; whether there was one."""
    element = find_comment(root, ref)
    if element is None:
        return False
    parent = element.parent
    if parent is not None:
        parent.remove(element)
    return True


def is_empty(root: Element) -> bool:
    container = root.child("commentList")
    return container is None or next(container.children_named("comment"), None) is None


def _made(root: Element, name: str) -> Element:
    """A child the part has to have, made where Excel puts it if missing."""
    found = root.child(name)
    if found is None:
        found = Element.create(name)
        if name == "authors":
            root.insert(0, found)
        else:
            root.append(found)
    return found


# ---------------------------------------------------------------------------
# The box Excel draws
# ---------------------------------------------------------------------------


def note_anchor(
    cell: CellRef, column_pixels: list[int], row_pixels: list[int]
) -> tuple[tuple[int, int, int, int, int, int, int, int], float, float]:
    """Where Excel puts a new note's box, as the VML anchor's eight numbers
    and the box's top left corner in points.

    ``column_pixels`` and ``row_pixels`` are the widths and heights of the
    sheet's columns and rows from the first, far enough to hold the box.
    Measured: the box starts 15 pixels right of the cell's right edge and
    10 above its top, 2 below the sheet's top edge on the first row, and
    the anchor names each corner by the zero-based column and row it falls
    in and how many pixels into it.
    """
    left = sum(column_pixels[: cell.column]) + 15
    top = max(sum(row_pixels[: cell.row - 1]) - 10, 2)
    first_column, first_x = _cell_at(left, column_pixels)
    first_row, first_y = _cell_at(top, row_pixels)
    last_column, last_x = _cell_at(left + NOTE_WIDTH_PIXELS, column_pixels)
    last_row, last_y = _cell_at(top + NOTE_HEIGHT_PIXELS, row_pixels)
    anchor = (first_column, first_x, first_row, first_y, last_column, last_x, last_row, last_y)
    return anchor, left * 0.75, top * 0.75


def _cell_at(position: int, sizes: list[int]) -> tuple[int, int]:
    """Which column or row a pixel falls in, and how far into it."""
    start = 0
    for index, size in enumerate(sizes):
        if position < start + size:
            return index, position - start
        start += size
    return len(sizes), position - start


def note_vml(
    shape_id: int,
    cell: CellRef,
    anchor: tuple[int, int, int, int, int, int, int, int],
    left: float,
    top: float,
    *,
    visible: bool,
    z_index: int,
) -> str:
    """A note's box, as Excel writes one."""
    corners = ", ".join(str(number) for number in anchor)
    shown = "\r\n   <x:Visible/>" if visible else ""
    return (
        f'<v:shape id="_x0000_s{shape_id}" type="#_x0000_t202" style=\'position:absolute;\r\n'
        f"  margin-left:{left:g}pt;margin-top:{top:g}pt;width:108pt;height:59.25pt;"
        f"z-index:{z_index};\r\n"
        f"  visibility:{'visible' if visible else 'hidden'}' fillcolor=\"infoBackground [80]\""
        ' strokecolor="none [81]"\r\n'
        '  o:insetmode="auto">\r\n'
        '  <v:fill color2="infoBackground [80]"/>\r\n'
        '  <v:shadow color="none [81]" obscured="t"/>\r\n'
        '  <v:path o:connecttype="none"/>\r\n'
        "  <v:textbox style='mso-direction-alt:auto'>\r\n"
        "   <div style='text-align:left'></div>\r\n"
        "  </v:textbox>\r\n"
        '  <x:ClientData ObjectType="Note">\r\n'
        "   <x:MoveWithCells/>\r\n"
        "   <x:SizeWithCells/>\r\n"
        "   <x:Anchor>\r\n"
        f"    {corners}</x:Anchor>\r\n"
        "   <x:AutoFill>False</x:AutoFill>\r\n"
        f"   <x:Row>{cell.row - 1}</x:Row>\r\n"
        f"   <x:Column>{cell.column - 1}</x:Column>{shown}\r\n"
        "  </x:ClientData>\r\n"
        " </v:shape>"
    )


def note_shapes(vml: str) -> dict[CellRef, str]:
    """Every note's box in a VML part, by the cell it belongs to."""
    found: dict[CellRef, str] = {}
    for match in _VML_SHAPES.finditer(vml):
        shape = match.group(0)
        if 'ObjectType="Note"' not in shape:
            continue
        row = re.search(r"<x:Row>\s*(\d+)\s*</x:Row>", shape)
        column = re.search(r"<x:Column>\s*(\d+)\s*</x:Column>", shape)
        if row is None or column is None:
            continue
        found[CellRef(int(row.group(1)) + 1, int(column.group(1)) + 1)] = shape
    return found


#: A note box's anchor: eight numbers, the third and seventh its top and
#: bottom rows, zero-based.
_ANCHOR = re.compile(r"(<x:Anchor>)(.*?)(</x:Anchor>)", re.DOTALL)
_LAST_ANCHOR_ROW = 1048575


def move_notes(vml: str, steps: dict[CellRef, int]) -> str:
    """A VML part with each note on a cell of ``steps`` moved that many rows
    down, or up when negative: the cell it belongs to, and its box with it,
    as a sort moves a note with its cell. The part is rebuilt in one pass,
    so notes trading places cannot overwrite each other."""
    pieces: list[str] = []
    last = 0
    for match in _VML_SHAPES.finditer(vml):
        shape = match.group(0)
        if 'ObjectType="Note"' not in shape:
            continue
        row = re.search(r"<x:Row>\s*(\d+)\s*</x:Row>", shape)
        column = re.search(r"<x:Column>\s*(\d+)\s*</x:Column>", shape)
        if row is None or column is None:
            continue
        step = steps.get(CellRef(int(row.group(1)) + 1, int(column.group(1)) + 1), 0)
        if not step:
            continue
        moved = shape[: row.start(1)] + str(int(row.group(1)) + step) + shape[row.end(1) :]
        moved = _ANCHOR.sub(
            lambda found, step=step: found.group(1) + _box_moved(found.group(2), step) + found.group(3), moved, 1
        )
        pieces += [vml[last : match.start()], moved]
        last = match.end()
    if not pieces:
        return vml
    return "".join(pieces) + vml[last:]


def _box_moved(anchor: str, step: int) -> str:
    """An anchor's text with its two rows moved by ``step``, its spacing
    kept."""
    slot = iter(range(8))

    def move(number: re.Match[str]) -> str:
        if next(slot, None) not in (2, 6):
            return number.group(0)
        return str(min(max(int(number.group(0)) + step, 0), _LAST_ANCHOR_ROW))

    return re.sub(r"-?\d+", move, anchor)


def shows(shape: str) -> bool:
    """Whether a note's box shows without the pointer over its cell."""
    return re.search(r"<x:Visible\s*/>", shape) is not None


def shown_as(shape: str, visible: bool) -> str:
    """A note's box, shown or hidden. Excel says it twice: in the style,
    and with ``<x:Visible/>`` after the cell the note belongs to."""
    shape = re.sub(
        r"visibility:(?:hidden|visible)", f"visibility:{'visible' if visible else 'hidden'}", shape, count=1
    )
    flag = re.search(r"\s*<x:Visible\s*/>", shape)
    if visible and flag is None:
        shape = re.sub(r"(</x:Column>)", "\\1\r\n   <x:Visible/>", shape, count=1)
    elif not visible and flag is not None:
        shape = shape.replace(flag.group(0), "", 1)
    return shape


def with_note_shape_type(vml: str) -> str:
    """A VML part that can draw a note: Excel puts the shape type in front
    of the first note it adds, after anything already there."""
    if 'id="_x0000_t202"' in vml:
        return vml
    at = vml.rfind("</xml>")
    return vml if at < 0 else vml[:at] + NOTE_SHAPE_TYPE + vml[at:]


def shape_count(vml: str) -> int:
    """How many shapes a VML part draws, which is where the next one's
    ``z-index`` starts."""
    return len(_VML_SHAPES.findall(vml))


__all__ = [
    "CT_COMMENTS",
    "CT_PERSONS",
    "CT_THREADED_COMMENTS",
    "EMPTY_COMMENTS",
    "EMPTY_PERSONS",
    "EMPTY_THREADED_COMMENTS",
    "NOTE_HEIGHT_PIXELS",
    "NOTE_SHAPE_TYPE",
    "NOTE_WIDTH_PIXELS",
    "NS_THREADED_COMMENTS",
    "PLACEHOLDER_HEAD",
    "RT_COMMENTS",
    "RT_PERSONS",
    "RT_THREADED_COMMENTS",
    "Comment",
    "Reply",
    "ThreadedComment",
    "comment_text",
    "delete_comment",
    "delete_thread",
    "find_comment",
    "format_moment",
    "is_empty",
    "move_notes",
    "new_id",
    "next_id",
    "note_anchor",
    "note_shapes",
    "note_vml",
    "parse_moment",
    "person_id",
    "placeholder_text",
    "read_comments",
    "read_persons",
    "read_threads",
    "set_resolved",
    "shape_count",
    "shown_as",
    "shows",
    "thread_ids",
    "with_note_shape_type",
    "write_comment",
    "write_thread",
]
