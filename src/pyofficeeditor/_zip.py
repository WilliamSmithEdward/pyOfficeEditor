"""The ZIP container, read and written field for field.

An OOXML file is a ZIP archive, and this library edits parts of one while
leaving the rest alone.  That makes byte fidelity a correctness property
rather than a nicety: a no-op save must produce the input file, so that
version control shows an edit and nothing else, and a one-cell edit must
not re-deflate a hundred megabytes of untouched parts.

The standard library's :mod:`zipfile` cannot deliver either.  Its public
write API re-deflates every member, and ``ZipInfo.extra`` exposes only the
central directory's copy of the extra field.  Excel's copies differ: in a
freshly authored workbook, ``[Content_Types].xml`` carries a 520-byte extra
field in its *local* header and none in the central directory (measured on
``xlsm_file_with_no_vba_entered_yet.xlsm``).  Anything built on
``zipfile``'s writer silently drops those bytes.  The field is tag
``0xa220``, the Microsoft Open Packaging Growth Hint: a signature word, a
padding word, then reserved zero bytes Excel keeps so it can grow a part in
place.  It is preserved verbatim, on rewritten members too, because the
padding stays legal and the intent is Excel's to keep.

So this module reads and writes the container itself.  Reading captures
every field of both headers plus each member's stored bytes; writing emits
them back.  The gate is that :func:`ZipArchive.to_bytes` reproduces its
input exactly when nothing was changed, which holds across
Excel-authored ``.xlsm``, ``.xlsb`` and openpyxl-authored ``.xlsx``.

Not supported, and refused rather than guessed at: zip64 archives, and any
compression method other than stored (0) and deflate (8).  Neither appears
in Office output at the sizes this library targets, and a wrong guess would
corrupt a document.
"""

from __future__ import annotations

import struct
import zlib
from dataclasses import dataclass, field

from pyofficeeditor.exceptions import UnsupportedFormatError, ZipError

_LOCAL_SIG = b"PK\x03\x04"
_CENTRAL_SIG = b"PK\x01\x02"
_EOCD_SIG = b"PK\x05\x06"
_ZIP64_EOCD_LOCATOR_SIG = b"PK\x06\x07"
_DATA_DESCRIPTOR_SIG = b"PK\x07\x08"

_LOCAL_FIXED = 30
_CENTRAL_FIXED = 46
_EOCD_FIXED = 22

#: Sentinel the spec uses in a 32-bit size field to mean "read the zip64
#: extra field instead".
_ZIP64_SENTINEL = 0xFFFFFFFF

#: Flag bit 3: the sizes and CRC follow the data in a descriptor record.
_FLAG_DATA_DESCRIPTOR = 0x0008
#: Flag bit 11: the member name is UTF-8 rather than cp437.
_FLAG_UTF8_NAME = 0x0800
#: Flag bits 1 and 2 are an advisory hint about the deflate level used.
#: No reader needs them to decode, so a member this library re-deflates
#: clears them rather than making a false claim about its own bytes.
_FLAG_DEFLATE_LEVEL_HINT = 0x0006

STORED = 0
DEFLATED = 8

#: Deflate level for members this library writes.  Chosen to match what
#: Office readers expect to inflate, which is any valid deflate stream;
#: the level only trades size against time.
_DEFLATE_LEVEL = 9


@dataclass
class ZipMember:
    """One member of the archive, with both headers' fields kept apart.

    ``stored`` holds the bytes as they sit in the file, still compressed.
    An untouched member is written back from ``stored`` without ever being
    inflated.
    """

    name: str
    stored: bytes

    # Fields the local and central headers agree on.
    extract_version: int = 20
    flags: int = 0
    method: int = DEFLATED
    mod_time: int = 0
    mod_date: int = 0x21  # 1980-01-01, the epoch of the DOS date format.
    crc: int = 0
    compressed_size: int = 0
    file_size: int = 0

    # Fields the two headers may disagree on, so each is kept as found.
    local_extra: bytes = b""
    central_extra: bytes = b""

    # A member with a data descriptor zeroes these three in its local
    # header and states the truth only in the central directory and the
    # descriptor (measured: that is what CPython's zipfile writes to a
    # non-seekable stream).  ``None`` means "write the central value".
    local_crc: int | None = None
    local_compressed_size: int | None = None
    local_file_size: int | None = None

    # Central-directory-only fields.
    create_version: int = 20
    create_system: int = 0
    comment: bytes = b""
    disk_start: int = 0
    internal_attr: int = 0
    external_attr: int = 0

    #: The trailing data-descriptor record, when the member has one.  Kept
    #: so a member written back unchanged reproduces its bytes; dropped
    #: when the member is rewritten, since the sizes then go in the header.
    data_descriptor: bytes = b""

    #: Set while writing; not part of the member's identity.
    header_offset: int = field(default=0, compare=False)

    @property
    def raw_name(self) -> bytes:
        return self.name.encode("utf-8" if self.flags & _FLAG_UTF8_NAME else "cp437")

    def inflate(self) -> bytes:
        """The member's decompressed content."""
        if self.method == STORED:
            data = self.stored
        elif self.method == DEFLATED:
            try:
                data = zlib.decompress(self.stored, -15)
            except zlib.error as exc:
                raise ZipError(f"{self.name}: deflate stream is corrupt ({exc}).") from exc
        else:
            raise UnsupportedFormatError(
                f"{self.name}: compression method {self.method} is not supported "
                f"(only stored and deflate are)."
            )
        if len(data) != self.file_size:
            raise ZipError(
                f"{self.name}: inflated to {len(data)} bytes, the directory says {self.file_size}."
            )
        actual = zlib.crc32(data)
        if actual != self.crc:
            raise ZipError(f"{self.name}: CRC-32 is {actual:#010x}, the directory says {self.crc:#010x}.")
        return data

    def replace(self, data: bytes, *, method: int | None = None) -> None:
        """Put ``data`` in this member, re-deflating it.

        The growth-hint extra field and the timestamps are kept; the
        data-descriptor record is dropped because the new sizes are written
        into the header directly.
        """
        use = self.method if method is None else method
        if use == STORED:
            self.stored = data
        elif use == DEFLATED:
            compressor = zlib.compressobj(_DEFLATE_LEVEL, zlib.DEFLATED, -15)
            self.stored = compressor.compress(data) + compressor.flush()
        else:
            raise UnsupportedFormatError(f"cannot write compression method {use}.")
        self.method = use
        self.crc = zlib.crc32(data)
        self.file_size = len(data)
        self.compressed_size = len(self.stored)
        self.flags &= ~(_FLAG_DATA_DESCRIPTOR | _FLAG_DEFLATE_LEVEL_HINT)
        self.data_descriptor = b""
        # With no descriptor the local header states the sizes itself.
        self.local_crc = None
        self.local_compressed_size = None
        self.local_file_size = None


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise ZipError(message)


class ZipArchive:
    """A ZIP archive held in memory, addressable by member name.

    Member order is the archive's own and is preserved; a new member is
    appended.  Names are matched exactly, as OPC part names are
    case-sensitive on the wire even though OPC comparison rules are not.
    """

    def __init__(self, members: list[ZipMember] | None = None, comment: bytes = b"") -> None:
        self._members: list[ZipMember] = members if members is not None else []
        self.comment: bytes = comment

    # ------------------------------------------------------------------
    # Reading
    # ------------------------------------------------------------------

    @classmethod
    def from_bytes(cls, data: bytes) -> ZipArchive:
        """Parse an archive.  Raises :class:`ZipError` on anything malformed."""
        eocd = data.rfind(_EOCD_SIG)
        _require(eocd >= 0, "not a ZIP archive: no end-of-central-directory record.")
        _require(
            eocd + _EOCD_FIXED <= len(data),
            "truncated ZIP archive: the end-of-central-directory record runs past the file.",
        )
        if data.rfind(_ZIP64_EOCD_LOCATOR_SIG, 0, eocd) >= 0:
            raise UnsupportedFormatError(
                "zip64 archives are not supported; this package is larger than the ZIP "
                "format's 32-bit fields allow."
            )

        count, cd_size, cd_offset, comment_len = struct.unpack("<HIIH", data[eocd + 10 : eocd + 22])
        comment = data[eocd + _EOCD_FIXED : eocd + _EOCD_FIXED + comment_len]
        _require(
            cd_offset + cd_size <= len(data),
            "truncated ZIP archive: the central directory runs past the file.",
        )

        members: list[ZipMember] = []
        pos = cd_offset
        for index in range(count):
            _require(
                pos + _CENTRAL_FIXED <= cd_offset + cd_size,
                f"truncated central directory at entry {index}.",
            )
            _require(
                data[pos : pos + 4] == _CENTRAL_SIG,
                f"bad central directory signature at offset {pos}.",
            )
            (
                create_version,
                create_system,
                extract_version,
                flags,
                method,
                mod_time,
                mod_date,
                crc,
                compressed_size,
                file_size,
                name_len,
                extra_len,
                comment_len2,
                disk_start,
                internal_attr,
                external_attr,
                local_offset,
            ) = struct.unpack("<4xBBHHHHHIIIHHHHHII", data[pos : pos + _CENTRAL_FIXED])

            if _ZIP64_SENTINEL in (compressed_size, file_size, local_offset):
                raise UnsupportedFormatError(
                    "zip64 archives are not supported; a member's size or offset needs 64 bits."
                )

            base = pos + _CENTRAL_FIXED
            raw_name = data[base : base + name_len]
            central_extra = data[base + name_len : base + name_len + extra_len]
            member_comment = data[
                base + name_len + extra_len : base + name_len + extra_len + comment_len2
            ]
            pos = base + name_len + extra_len + comment_len2

            name = raw_name.decode("utf-8" if flags & _FLAG_UTF8_NAME else "cp437")

            _require(
                local_offset + _LOCAL_FIXED <= len(data),
                f"{name}: local header offset {local_offset} is past the end of the file.",
            )
            _require(
                data[local_offset : local_offset + 4] == _LOCAL_SIG,
                f"{name}: no local file header at offset {local_offset}.",
            )
            local_crc, local_compressed_size, local_file_size = struct.unpack(
                "<III", data[local_offset + 14 : local_offset + 26]
            )
            local_name_len, local_extra_len = struct.unpack(
                "<HH", data[local_offset + 26 : local_offset + 30]
            )
            extra_start = local_offset + _LOCAL_FIXED + local_name_len
            local_extra = data[extra_start : extra_start + local_extra_len]
            body = extra_start + local_extra_len
            _require(
                body + compressed_size <= len(data),
                f"{name}: member data runs past the end of the file.",
            )
            stored = data[body : body + compressed_size]

            descriptor = b""
            if flags & _FLAG_DATA_DESCRIPTOR:
                # The sizes we trust came from the central directory, so the
                # descriptor is only captured to be written back unchanged.
                # Its signature is optional, which is why the length is read
                # from the bytes rather than assumed.
                tail = body + compressed_size
                length = 16 if data[tail : tail + 4] == _DATA_DESCRIPTOR_SIG else 12
                _require(
                    tail + length <= len(data),
                    f"{name}: data descriptor runs past the end of the file.",
                )
                descriptor = data[tail : tail + length]

            members.append(
                ZipMember(
                    name=name,
                    stored=stored,
                    extract_version=extract_version,
                    flags=flags,
                    method=method,
                    mod_time=mod_time,
                    mod_date=mod_date,
                    crc=crc,
                    compressed_size=compressed_size,
                    file_size=file_size,
                    local_extra=local_extra,
                    central_extra=central_extra,
                    local_crc=None if local_crc == crc else local_crc,
                    local_compressed_size=(
                        None if local_compressed_size == compressed_size else local_compressed_size
                    ),
                    local_file_size=None if local_file_size == file_size else local_file_size,
                    create_version=create_version,
                    create_system=create_system,
                    comment=member_comment,
                    disk_start=disk_start,
                    internal_attr=internal_attr,
                    external_attr=external_attr,
                    data_descriptor=descriptor,
                )
            )

        return cls(members, comment)

    # ------------------------------------------------------------------
    # Writing
    # ------------------------------------------------------------------

    def to_bytes(self) -> bytes:
        """Serialize the archive.

        With no member changed this reproduces the bytes
        :meth:`from_bytes` was given.
        """
        out = bytearray()
        for member in self._members:
            member.header_offset = len(out)
            raw_name = member.raw_name
            out += struct.pack(
                "<4sHHHHHIIIHH",
                _LOCAL_SIG,
                member.extract_version,
                member.flags,
                member.method,
                member.mod_time,
                member.mod_date,
                member.crc if member.local_crc is None else member.local_crc,
                (
                    member.compressed_size
                    if member.local_compressed_size is None
                    else member.local_compressed_size
                ),
                member.file_size if member.local_file_size is None else member.local_file_size,
                len(raw_name),
                len(member.local_extra),
            )
            out += raw_name
            out += member.local_extra
            out += member.stored
            out += member.data_descriptor

        cd_offset = len(out)
        for member in self._members:
            raw_name = member.raw_name
            out += struct.pack(
                "<4sBBHHHHHIIIHHHHHII",
                _CENTRAL_SIG,
                member.create_version,
                member.create_system,
                member.extract_version,
                member.flags,
                member.method,
                member.mod_time,
                member.mod_date,
                member.crc,
                member.compressed_size,
                member.file_size,
                len(raw_name),
                len(member.central_extra),
                len(member.comment),
                member.disk_start,
                member.internal_attr,
                member.external_attr,
                member.header_offset,
            )
            out += raw_name
            out += member.central_extra
            out += member.comment
        cd_size = len(out) - cd_offset

        if cd_offset > _ZIP64_SENTINEL or len(self._members) > 0xFFFF:
            raise UnsupportedFormatError(
                "the package has outgrown the ZIP format's 32-bit fields; zip64 is not supported."
            )

        out += struct.pack(
            "<4sHHHHIIH",
            _EOCD_SIG,
            0,
            0,
            len(self._members),
            len(self._members),
            cd_size,
            cd_offset,
            len(self.comment),
        )
        out += self.comment
        return bytes(out)

    # ------------------------------------------------------------------
    # Member access
    # ------------------------------------------------------------------

    def names(self) -> list[str]:
        """Member names, in archive order."""
        return [m.name for m in self._members]

    def __contains__(self, name: str) -> bool:
        return any(m.name == name for m in self._members)

    def __len__(self) -> int:
        return len(self._members)

    def members(self) -> list[ZipMember]:
        """The members themselves, in archive order.  The list is a copy;
        the members are not."""
        return list(self._members)

    def member(self, name: str) -> ZipMember:
        for m in self._members:
            if m.name == name:
                return m
        raise KeyError(name)

    def read(self, name: str) -> bytes:
        """The decompressed content of one member."""
        try:
            member = self.member(name)
        except KeyError:
            raise ZipError(f"no member named {name!r} in the archive.") from None
        return member.inflate()

    def write(self, name: str, data: bytes) -> ZipMember:
        """Replace a member's content, or append a new member."""
        try:
            member = self.member(name)
        except KeyError:
            member = ZipMember(name=name, stored=b"", flags=_FLAG_UTF8_NAME if not name.isascii() else 0)
            self._members.append(member)
        member.replace(data)
        return member

    def remove(self, name: str) -> None:
        for index, m in enumerate(self._members):
            if m.name == name:
                del self._members[index]
                return
        raise ZipError(f"no member named {name!r} in the archive.")


__all__ = ["DEFLATED", "STORED", "ZipArchive", "ZipMember"]
