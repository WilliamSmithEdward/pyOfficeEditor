"""The ZIP container.

The gate is byte fidelity: an archive read and written back with nothing
changed must be the bytes that came in, and an edit to one member must not
disturb another member's stored bytes.  Everything else in the library
rests on that.
"""

from __future__ import annotations

import io
import zipfile

import pytest

from pyofficeeditor._zip import DEFLATED, STORED, ZipArchive, ZipMember
from pyofficeeditor.exceptions import UnsupportedFormatError, ZipError


class TestRoundTrip:
    def test_excel_authored_packages_are_byte_identical(
        self, excel_authored_packages: list[tuple[str, bytes]]
    ) -> None:
        for name, data in excel_authored_packages:
            assert ZipArchive.from_bytes(data).to_bytes() == data, name

    def test_openpyxl_authored_package_is_byte_identical(self, openpyxl_xlsx_bytes: bytes) -> None:
        assert ZipArchive.from_bytes(openpyxl_xlsx_bytes).to_bytes() == openpyxl_xlsx_bytes

    def test_data_descriptor_members_are_byte_identical(self, data_descriptor_zip_bytes: bytes) -> None:
        assert ZipArchive.from_bytes(data_descriptor_zip_bytes).to_bytes() == data_descriptor_zip_bytes

    def test_binary_members_are_not_assumed_to_be_text(self, binary_xlsb_bytes: bytes) -> None:
        """An .xlsb keeps its worksheets as .bin parts.  The container layer
        must not care."""
        archive = ZipArchive.from_bytes(binary_xlsb_bytes)
        assert archive.to_bytes() == binary_xlsb_bytes
        binary = [n for n in archive.names() if n.endswith(".bin")]
        assert len(binary) == 5, binary
        for name in binary:
            assert archive.read(name), name

    def test_member_order_and_content_match_the_standard_library(self, minimal_xlsm_bytes: bytes) -> None:
        archive = ZipArchive.from_bytes(minimal_xlsm_bytes)
        reference = zipfile.ZipFile(io.BytesIO(minimal_xlsm_bytes))
        assert archive.names() == reference.namelist()
        for name in archive.names():
            assert archive.read(name) == reference.read(name), name


class TestHeaderFidelity:
    def test_local_and_central_extra_fields_are_kept_apart(self, minimal_xlsm_bytes: bytes) -> None:
        """Excel populates the local extra field and leaves the central one
        empty.  A reader that keeps only one copy loses 520 bytes here."""
        member = ZipArchive.from_bytes(minimal_xlsm_bytes).member("[Content_Types].xml")
        assert len(member.local_extra) == 520
        assert member.central_extra == b""

    def test_the_growth_hint_is_recognisable(self, minimal_xlsm_bytes: bytes) -> None:
        member = ZipArchive.from_bytes(minimal_xlsm_bytes).member("[Content_Types].xml")
        tag = int.from_bytes(member.local_extra[:2], "little")
        assert tag == 0xA220, "the Microsoft Open Packaging Growth Hint"

    def test_a_descriptor_member_zeroes_its_local_sizes(self, data_descriptor_zip_bytes: bytes) -> None:
        member = ZipArchive.from_bytes(data_descriptor_zip_bytes).member("a.xml")
        assert member.local_crc == 0
        assert member.local_compressed_size == 0
        assert member.local_file_size == 0
        assert member.crc != 0
        assert len(member.data_descriptor) in (12, 16)

    def test_excel_version_words_survive(self, minimal_xlsm_bytes: bytes) -> None:
        member = ZipArchive.from_bytes(minimal_xlsm_bytes).member("xl/workbook.xml")
        assert member.create_version == 45
        assert member.extract_version == 20


class TestEditing:
    def test_editing_one_member_leaves_the_others_stored(self, minimal_xlsm_bytes: bytes) -> None:
        archive = ZipArchive.from_bytes(minimal_xlsm_bytes)
        before = {m.name: m.stored for m in archive.members()}
        archive.write("xl/worksheets/sheet1.xml", b"<x/>")
        for member in archive.members():
            if member.name != "xl/worksheets/sheet1.xml":
                assert member.stored == before[member.name], member.name

    def test_an_edit_survives_a_reparse(self, minimal_xlsm_bytes: bytes) -> None:
        archive = ZipArchive.from_bytes(minimal_xlsm_bytes)
        archive.write("xl/worksheets/sheet1.xml", b"<x/>")
        assert ZipArchive.from_bytes(archive.to_bytes()).read("xl/worksheets/sheet1.xml") == b"<x/>"

    def test_the_standard_library_can_read_what_we_write(self, minimal_xlsm_bytes: bytes) -> None:
        archive = ZipArchive.from_bytes(minimal_xlsm_bytes)
        archive.write("xl/worksheets/sheet1.xml", b"<changed/>")
        reference = zipfile.ZipFile(io.BytesIO(archive.to_bytes()))
        assert reference.testzip() is None
        assert reference.read("xl/worksheets/sheet1.xml") == b"<changed/>"

    def test_a_rewritten_member_keeps_its_growth_hint_and_timestamp(self, minimal_xlsm_bytes: bytes) -> None:
        archive = ZipArchive.from_bytes(minimal_xlsm_bytes)
        before = archive.member("[Content_Types].xml")
        extra, mod_time, mod_date = before.local_extra, before.mod_time, before.mod_date
        archive.write("[Content_Types].xml", b"<Types/>")
        after = archive.member("[Content_Types].xml")
        assert after.local_extra == extra
        assert (after.mod_time, after.mod_date) == (mod_time, mod_date)

    def test_a_rewritten_member_drops_its_descriptor(self, data_descriptor_zip_bytes: bytes) -> None:
        archive = ZipArchive.from_bytes(data_descriptor_zip_bytes)
        archive.write("a.xml", b"<new/>")
        member = archive.member("a.xml")
        assert member.data_descriptor == b""
        assert member.local_crc is None
        assert ZipArchive.from_bytes(archive.to_bytes()).read("a.xml") == b"<new/>"

    def test_adding_and_removing_members(self, minimal_xlsm_bytes: bytes) -> None:
        archive = ZipArchive.from_bytes(minimal_xlsm_bytes)
        count = len(archive)
        archive.write("xl/added.xml", b"<added/>")
        assert len(archive) == count + 1
        assert ZipArchive.from_bytes(archive.to_bytes()).read("xl/added.xml") == b"<added/>"
        archive.remove("xl/added.xml")
        assert "xl/added.xml" not in ZipArchive.from_bytes(archive.to_bytes())

    def test_removing_an_absent_member_raises(self, minimal_xlsm_bytes: bytes) -> None:
        archive = ZipArchive.from_bytes(minimal_xlsm_bytes)
        with pytest.raises(ZipError, match="no member named"):
            archive.remove("xl/nope.xml")

    def test_reading_an_absent_member_raises(self, minimal_xlsm_bytes: bytes) -> None:
        archive = ZipArchive.from_bytes(minimal_xlsm_bytes)
        with pytest.raises(ZipError, match="no member named"):
            archive.read("xl/nope.xml")

    @pytest.mark.parametrize("method", [STORED, DEFLATED])
    def test_both_supported_methods_round_trip(self, method: int) -> None:
        member = ZipMember(name="a.xml", stored=b"")
        member.replace(b"payload" * 50, method=method)
        archive = ZipArchive([member])
        assert ZipArchive.from_bytes(archive.to_bytes()).read("a.xml") == b"payload" * 50


class TestMalformedInput:
    def test_not_a_zip_at_all(self) -> None:
        with pytest.raises(ZipError, match="no end-of-central-directory"):
            ZipArchive.from_bytes(b"this is not a zip file")

    def test_empty_input(self) -> None:
        with pytest.raises(ZipError):
            ZipArchive.from_bytes(b"")

    def test_truncated_archive(self, minimal_xlsm_bytes: bytes) -> None:
        with pytest.raises(ZipError):
            ZipArchive.from_bytes(minimal_xlsm_bytes[:200])

    def test_central_directory_pointing_past_the_file(self, minimal_xlsm_bytes: bytes) -> None:
        data = bytearray(minimal_xlsm_bytes)
        end = data.rfind(b"PK\x05\x06")
        data[end + 16 : end + 20] = (len(data) * 4).to_bytes(4, "little")
        with pytest.raises(ZipError):
            ZipArchive.from_bytes(bytes(data))

    def test_local_header_signature_corrupted(self, minimal_xlsm_bytes: bytes) -> None:
        data = bytearray(minimal_xlsm_bytes)
        data[0:4] = b"XXXX"
        with pytest.raises(ZipError, match="no local file header"):
            ZipArchive.from_bytes(bytes(data))

    def test_a_corrupt_deflate_stream_is_reported(self, minimal_xlsm_bytes: bytes) -> None:
        archive = ZipArchive.from_bytes(minimal_xlsm_bytes)
        member = archive.member("xl/workbook.xml")
        member.stored = b"\x00" * len(member.stored)
        with pytest.raises(ZipError):
            archive.read("xl/workbook.xml")

    def test_a_crc_mismatch_is_reported(self) -> None:
        member = ZipMember(name="a.xml", stored=b"")
        member.replace(b"payload")
        member.crc ^= 0xFFFF
        with pytest.raises(ZipError, match="CRC-32"):
            member.inflate()

    def test_a_size_mismatch_is_reported(self) -> None:
        member = ZipMember(name="a.xml", stored=b"")
        member.replace(b"payload")
        member.file_size += 1
        with pytest.raises(ZipError, match="inflated to"):
            member.inflate()

    def test_an_unsupported_method_is_refused(self) -> None:
        member = ZipMember(name="a.xml", stored=b"x", method=14)  # LZMA
        with pytest.raises(UnsupportedFormatError, match="compression method 14"):
            member.inflate()
        with pytest.raises(UnsupportedFormatError):
            member.replace(b"x", method=14)

    def test_zip64_is_refused_rather_than_misread(self) -> None:
        """A zip64 archive is refused by name.  Silently reading the 32-bit
        fields would produce a plausible, wrong package."""
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", allowZip64=True) as archive:
            info = zipfile.ZipInfo("big.xml")
            info.file_size = 0
            archive.writestr(info, b"small")
        # Force the zip64 end-of-central-directory locator into the bytes.
        data = buffer.getvalue()
        end = data.rfind(b"PK\x05\x06")
        spiked = data[:end] + b"PK\x06\x07" + b"\x00" * 16 + data[end:]
        with pytest.raises(UnsupportedFormatError, match="zip64"):
            ZipArchive.from_bytes(spiked)
