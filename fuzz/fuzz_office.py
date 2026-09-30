"""Coverage-guided fuzzing of the readers that take untrusted Office files.

Each target feeds Atheris-generated input to one reader. A reader must either
succeed or raise a PyOfficeEditorError (FormulaSyntaxError for the formula
parser); anything else, a crash or a hang, is a finding. Where the reader
promises it, a successful read must also write back exactly the bytes it
was given:

  zip       ZipArchive.from_bytes(), every member read, and to_bytes()
            keeps every member: reading its output gives the same members,
            and writing that again gives the same bytes. (Byte-identity with
            the input holds only for the layout ZIP writers produce, which
            tests/test_zip.py checks on real archives; the fuzzer makes others.)
  xml       XmlDocument.parse(), and to_bytes() reproduces the input.
  workbook  Workbook.from_bytes(), every sheet's cells and merged ranges.
  formula   the formula parser, on the input as UTF-8 text.

    python fuzz/fuzz_office.py <target> [libFuzzer options] [corpus dirs]
    python fuzz/fuzz_office.py xml -max_total_time=60 tests/fuzz_corpus/xml

The .github/workflows/fuzz.yml workflow runs each target from
tests/fuzz_corpus/<target>. A finding becomes a seed there, which
tests/test_fuzz_corpus.py replays on every CI run.
"""

import sys

import atheris

with atheris.instrument_imports():
    from pyofficeeditor import PyOfficeEditorError
    from pyofficeeditor._xml import XmlDocument
    from pyofficeeditor._zip import ZipArchive
    from pyofficeeditor.excel import Workbook
    from pyofficeeditor.excel._calc.lexer import FormulaSyntaxError
    from pyofficeeditor.excel._calc.parser import parse as parse_formula


def fuzz_zip(data):
    try:
        archive = ZipArchive.from_bytes(data)
        for name in archive.names():
            archive.read(name)
    except PyOfficeEditorError:
        return
    written = archive.to_bytes()
    end = written.rfind(b"PK\x05\x06")
    if written[end - 20 : end - 16] == b"PK\x06\x07":
        # The format's own ambiguity: a member comment that ends in the zip64
        # locator's signature and 16 more bytes lands where the locator
        # would be, and every reader, Python's zipfile among them, takes it
        # for one. No ZIP writer puts those bytes there by accident.
        return
    again = ZipArchive.from_bytes(written)
    if [(m.name, m.stored) for m in again.members()] != [(m.name, m.stored) for m in archive.members()]:
        raise AssertionError("writing an archive changed its members")
    if again.to_bytes() != written:
        raise AssertionError("writing an archive twice gave different bytes")


def fuzz_xml(data):
    try:
        document = XmlDocument.parse(data)
    except PyOfficeEditorError:
        return
    if document.to_bytes() != data:
        raise AssertionError("an unchanged document does not write back its input")


def fuzz_workbook(data):
    try:
        book = Workbook.from_bytes(data)
        for sheet in book.sheets:
            list(sheet.cell_elements())
            _ = sheet.merged_ranges
    except PyOfficeEditorError:
        pass


def fuzz_formula(data):
    try:
        parse_formula(data.decode("utf-8", errors="replace"))
    except FormulaSyntaxError:
        pass


TARGETS = {"zip": fuzz_zip, "xml": fuzz_xml, "workbook": fuzz_workbook, "formula": fuzz_formula}


def main():
    if len(sys.argv) < 2 or sys.argv[1] not in TARGETS:
        sys.exit(f"usage: fuzz_office.py <{'|'.join(TARGETS)}> [libFuzzer options]")
    atheris.Setup([sys.argv[0], *sys.argv[2:]], TARGETS[sys.argv[1]])
    atheris.Fuzz()


if __name__ == "__main__":
    main()
