"""Pictures on a sheet: the image, its part, and the anchor Excel draws it by.

A picture is three things, as Excel writes one:

- the image itself, in a media part of its own, ``xl/media/image1.png``,
  with a content type by extension; Excel writes a JPEG as ``.jpeg``, and
  stores an image once however many pictures show it
- an image relationship from the sheet's drawing part to that media part
- an ``<xdr:pic>`` in the drawing, in a two-cell anchor that moves with its
  cells and keeps its size, whose ``<a:blip>`` names the relationship

Its size, when none is asked for, is the one Excel gives it, measured:
the image's pixels at 96 to the inch, except that a PNG saying how many
pixels it has to the metre is taken at its word. A JPEG's own density is
ignored: one claiming 72 to the inch came in at the same size as one
claiming nothing.

Only the formats Excel stores as they are, PNG, JPEG and GIF, are written;
anything else is refused by name rather than stored as something Excel
would convert first.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass

from pyofficeeditor._xml import escape_attribute
from pyofficeeditor.excel._shapes import SheetGrid, corner_markup, emu

RT_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

#: How many pixels to the inch a size in pixels means, when the image says
#: nothing, or says something Excel ignores.
DEFAULT_DPI = 96.0


@dataclass(frozen=True)
class ImageInfo:
    """What a picture's size is worked out from."""

    extension: str
    content_type: str
    width: int
    height: int
    dpi_x: float = DEFAULT_DPI
    dpi_y: float = DEFAULT_DPI

    @property
    def size(self) -> tuple[float, float]:
        """The size Excel gives the picture, in points."""
        return self.width * 72 / self.dpi_x, self.height * 72 / self.dpi_y


def image_info(data: bytes) -> ImageInfo:
    """The format and size of an image, read from its own header."""
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return _png(data)
    if data.startswith(b"\xff\xd8"):
        return _jpeg(data)
    if data[:6] in (b"GIF87a", b"GIF89a") and len(data) >= 10:
        width, height = struct.unpack("<HH", data[6:10])
        return ImageInfo("gif", "image/gif", width, height)
    raise ValueError(
        "the image is not a PNG, a JPEG or a GIF, the formats Excel stores as they are; "
        "convert it to PNG first."
    )


def _png(data: bytes) -> ImageInfo:
    if len(data) < 24 or data[12:16] != b"IHDR":
        raise ValueError("the PNG has no header to size it by.")
    width, height = struct.unpack(">II", data[16:24])
    dpi_x = dpi_y = DEFAULT_DPI
    at = 8
    while at + 8 <= len(data):
        length = struct.unpack(">I", data[at : at + 4])[0]
        kind = data[at + 4 : at + 8]
        if kind == b"pHYs" and length >= 9:
            per_x, per_y, unit = struct.unpack(">IIB", data[at + 8 : at + 17])
            if unit == 1 and per_x and per_y:
                dpi_x, dpi_y = per_x * 0.0254, per_y * 0.0254
            break
        if kind in (b"IDAT", b"IEND"):
            break
        at += 12 + length
    return ImageInfo("png", "image/png", width, height, dpi_x, dpi_y)


#: The start-of-frame markers, which carry a JPEG's size; the others in
#: that range are tables and are skipped.
_FRAMES = frozenset(range(0xC0, 0xD0)) - {0xC4, 0xC8, 0xCC}


def _jpeg(data: bytes) -> ImageInfo:
    at = 2
    while at + 4 <= len(data):
        if data[at] != 0xFF:
            break
        marker = data[at + 1]
        if marker == 0xFF:
            at += 1
            continue
        length = struct.unpack(">H", data[at + 2 : at + 4])[0]
        if marker in _FRAMES and at + 9 <= len(data):
            height, width = struct.unpack(">HH", data[at + 5 : at + 9])
            return ImageInfo("jpeg", "image/jpeg", width, height)
        at += 2 + length
    raise ValueError("the JPEG has no frame header to size it by.")


def picture_anchor(
    *,
    shape_id: int,
    name: str,
    description: str,
    relationship: str,
    extension: str,
    left: float,
    top: float,
    width: float,
    height: float,
    grid: SheetGrid,
    keeps_aspect: bool,
) -> str:
    """A picture's anchor, as Excel writes one for ``Shapes.AddPicture``.

    ``keeps_aspect`` is Excel's ``noChangeAspect``, which it sets on a
    picture put in at its own size and leaves off one given a size; a JPEG's
    blip says it is compressed for print, and a GIF's carries no DPI note.
    """
    described = f' descr="{escape_attribute(description)}"' if description else ""
    locks = '<a:picLocks noChangeAspect="1"/>' if keeps_aspect else "<a:picLocks/>"
    state = ' cstate="print"' if extension == "jpeg" else ""
    local_dpi = (
        ""
        if extension == "gif"
        else '<a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">'
        '<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>'
        "</a:ext></a:extLst>"
    )
    blip = (
        '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        f' r:embed="{relationship}"{state}>{local_dpi}</a:blip>'
        if local_dpi
        else '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        f' r:embed="{relationship}"{state}/>'
    )
    return (
        '<xdr:twoCellAnchor editAs="oneCell">'
        f"{corner_markup('from', grid, left, top)}"
        f"{corner_markup('to', grid, left + width, top + height)}"
        "<xdr:pic><xdr:nvPicPr>"
        f'<xdr:cNvPr id="{shape_id}" name="{escape_attribute(name)}"{described}/>'
        f"<xdr:cNvPicPr>{locks}</xdr:cNvPicPr></xdr:nvPicPr>"
        f"<xdr:blipFill>{blip}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>"
        "<xdr:spPr><a:xfrm>"
        f'<a:off x="{emu(left)}" y="{emu(top)}"/><a:ext cx="{emu(width)}" cy="{emu(height)}"/>'
        '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
        "</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>"
    )


__all__ = [
    "DEFAULT_DPI",
    "RT_IMAGE",
    "ImageInfo",
    "image_info",
    "picture_anchor",
]
