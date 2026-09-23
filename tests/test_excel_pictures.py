"""Pictures on a sheet.

``pictures.xlsx`` was built by Excel's ``Shapes.AddPicture``, and
``pictures_answers.json`` beside it records what Excel's object model said
of each picture. Reading is held to that; writing to the markup Excel
wrote when measured, and the live gate has Excel open what is written.
"""

from __future__ import annotations

import json
import shutil
import struct
import zlib
from pathlib import Path

import pytest

from pyofficeeditor.excel import Workbook, Worksheet
from pyofficeeditor.excel._pictures import image_info
from pyofficeeditor.excel._shapes import emu

#: The one-pixel GIF every web page used to carry.
GIF = (
    b"GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff!\xf9\x04\x01\x00\x00\x00\x00"
    b",\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;"
)


def png(width: int, height: int, dpi: int | None = None) -> bytes:
    """A plain red PNG, with a ``pHYs`` chunk when ``dpi`` is given."""

    def chunk(kind: bytes, body: bytes) -> bytes:
        crc = zlib.crc32(kind + body) & 0xFFFFFFFF
        return struct.pack(">I", len(body)) + kind + body + struct.pack(">I", crc)

    rows = b"".join(b"\x00" + b"\xff\x00\x00" * width for _ in range(height))
    body = chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
    if dpi is not None:
        per_metre = round(dpi / 0.0254)
        body += chunk(b"pHYs", struct.pack(">IIB", per_metre, per_metre, 1))
    body += chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b"")
    return b"\x89PNG\r\n\x1a\n" + body


def jpeg(width: int, height: int) -> bytes:
    """Just enough of a JPEG to be sized: its markers, and no picture."""
    frame = struct.pack(">BHHB", 8, height, width, 1) + b"\x01\x11\x00"
    return b"\xff\xd8" + b"\xff\xe0" + struct.pack(">H", 16) + b"JFIF\x00\x01\x01\x01\x00\x48\x00\x48\x00\x00" + (
        b"\xff\xc0" + struct.pack(">H", len(frame) + 2) + frame
    ) + b"\xff\xd9"


@pytest.fixture(scope="module")
def answers(live_pictures_answers: Path) -> dict[str, dict[str, object]]:
    return json.loads(live_pictures_answers.read_text(encoding="utf-8"))


@pytest.fixture
def excels(tmp_path: Path, live_pictures_xlsx: Path) -> Worksheet:
    target = tmp_path / "pictures.xlsx"
    shutil.copy(live_pictures_xlsx, target)
    return Workbook.open(target)["Pictures"]


@pytest.fixture
def blank(tmp_path: Path, live_sample_xlsx: Path) -> Worksheet:
    target = tmp_path / "sample.xlsx"
    shutil.copy(live_sample_xlsx, target)
    return Workbook.open(target)["Data"]


def drawing_text(sheet: Worksheet) -> str:
    package = sheet.workbook.package
    names = [name for name in package.part_names() if name.startswith("xl/drawings/drawing")]
    assert len(names) == 1, names
    return package.read(names[0]).decode("utf-8")


def media(sheet: Worksheet) -> list[str]:
    return sorted(name for name in sheet.workbook.package.part_names() if name.startswith("xl/media/"))


class TestReadingExcelsOwn:
    def test_every_picture_where_excel_put_it(
        self, excels: Worksheet, answers: dict[str, dict[str, object]]
    ) -> None:
        shapes = {shape.name: shape for shape in excels.shapes}
        assert set(shapes) == set(answers)
        for name, said in answers.items():
            shape = shapes[name]
            width, height = said["width"], said["height"]
            assert isinstance(width, float) and isinstance(height, float)
            assert shape.kind == "picture" and shape.mso_type == said["type"] == 13, name
            assert (round(shape.width, 3), round(shape.height, 3)) == (round(width, 3), round(height, 3)), name

    def test_one_image_shown_twice_is_stored_once(self, excels: Worksheet) -> None:
        shapes = {shape.name: shape for shape in excels.shapes}
        assert shapes["Natural"].image == shapes["Stretched"].image
        assert len(media(excels)) == 3

    def test_the_bytes_behind_a_picture(self, excels: Worksheet) -> None:
        assert excels.picture_data("Dot") == GIF
        assert excels.picture_data("Natural").startswith(b"\x89PNG")


class TestSizing:
    """What Excel made of each image's own size, measured: the extents in
    EMU it wrote into the drawing for each."""

    @pytest.mark.parametrize(
        ("data", "extent"),
        [
            (png(120, 80, 96), (1142857, 761905)),
            (png(120, 80, 144), (762039, 508026)),
            (png(120, 80), (1143000, 762000)),
            (GIF, (9525, 9525)),
            (jpeg(3840, 2400), (36576000, 22860000)),
        ],
    )
    def test_the_size_excel_gives(self, data: bytes, extent: tuple[int, int]) -> None:
        width, height = image_info(data).size
        assert (emu(width), emu(height)) == extent

    def test_a_jpegs_own_density_is_ignored(self) -> None:
        """One claiming 72 to the inch came in at 96 all the same."""
        assert image_info(jpeg(3840, 2400)).size == (2880.0, 1800.0)

    def test_another_format_is_refused(self) -> None:
        with pytest.raises(ValueError, match="PNG, a JPEG or a GIF"):
            image_info(b"BM" + b"\x00" * 60)


class TestWriting:
    def test_a_picture_and_its_parts(self, blank: Worksheet) -> None:
        shape = blank.add_picture("Logo", png(120, 80, 144), left=100, top=50, description="our logo")
        assert shape.kind == "picture" and shape.image == "xl/media/image1.png"
        assert (round(shape.width, 4), round(shape.height, 4)) == (60.0031, 40.002)
        drawing = drawing_text(blank)
        assert '<xdr:twoCellAnchor editAs="oneCell">' in drawing
        assert 'descr="our logo"' in drawing
        assert '<a:picLocks noChangeAspect="1"/>' in drawing
        types = blank.workbook.package.read("[Content_Types].xml").decode("utf-8")
        assert '<Default Extension="png" ContentType="image/png"/>' in types
        assert blank.picture_data("Logo") == png(120, 80, 144)

    def test_given_a_size_it_no_longer_locks_its_aspect(self, blank: Worksheet) -> None:
        blank.add_picture("Stretched", png(120, 80), left=10, top=10, width=60, height=30)
        assert "<a:picLocks/>" in drawing_text(blank)

    def test_one_side_given_keeps_the_proportions(self, blank: Worksheet) -> None:
        shape = blank.add_picture("Wide", png(120, 80), left=10, top=10, width=45)
        assert (shape.width, shape.height) == (45, 30)
        taller = blank.add_picture("Tall", png(120, 80), left=10, top=100, height=60)
        assert (taller.width, taller.height) == (90, 60)

    def test_an_image_shown_twice_is_stored_once(self, blank: Worksheet) -> None:
        blank.add_picture("One", png(20, 20), left=10, top=10)
        blank.add_picture("Two", png(20, 20), left=100, top=10)
        blank.add_picture("Other", GIF, left=200, top=10)
        assert media(blank) == ["xl/media/image1.png", "xl/media/image2.gif"]
        assert drawing_text(blank).count('r:embed="') == 3
        relationships = blank.workbook.package.relationships("xl/drawings/drawing1.xml")
        assert len(list(relationships)) == 2

    def test_a_jpeg_is_stored_as_excel_stores_one(self, blank: Worksheet) -> None:
        blank.add_picture("Photo", jpeg(40, 20), left=10, top=10)
        assert media(blank) == ["xl/media/image1.jpeg"]
        assert 'cstate="print"' in drawing_text(blank)

    def test_from_a_file(self, blank: Worksheet, tmp_path: Path) -> None:
        source = tmp_path / "dot.gif"
        source.write_bytes(GIF)
        blank.add_picture("Dot", source, left=10, top=10)
        assert blank.picture_data("Dot") == GIF

    def test_it_survives_a_save(self, blank: Worksheet, tmp_path: Path) -> None:
        blank.add_picture("Logo", png(30, 20), left=100, top=50)
        blank.workbook.save()
        reopened = Workbook.open(tmp_path / "sample.xlsx")["Data"]
        assert reopened.picture_data("Logo") == png(30, 20)

    def test_refusals(self, blank: Worksheet) -> None:
        blank.add_picture("Logo", png(30, 20), left=100, top=50)
        with pytest.raises(ValueError, match="already has a shape"):
            blank.add_picture("Logo", png(30, 20), left=100, top=50)
        with pytest.raises(ValueError, match="PNG, a JPEG or a GIF"):
            blank.add_picture("Bitmap", b"BM" + b"\x00" * 60, left=0, top=0)
        blank.add_shape("Box", left=0, top=0, width=10, height=10)
        with pytest.raises(ValueError, match="not a picture"):
            blank.picture_data("Box")


class TestRemoving:
    def test_the_last_picture_of_an_image_takes_it_with_it(self, blank: Worksheet) -> None:
        blank.add_picture("Gone", png(10, 10), left=10, top=10)
        blank.remove_shape("Gone")
        assert media(blank) == []
        assert len(list(blank.workbook.package.relationships("xl/drawings/drawing1.xml"))) == 0

    def test_an_image_another_picture_shows_stays(self, excels: Worksheet) -> None:
        excels.remove_shape("Stretched")
        assert len(media(excels)) == 3
        assert excels.picture_data("Natural").startswith(b"\x89PNG")
        excels.remove_shape("Natural")
        assert len(media(excels)) == 2
