"""Fonts, fills, borders and alignment.

Formatting is shared: many cells point at one ``cellXfs`` entry, so editing
an entry repaints every cell using it. Everything here rests on that, and the
tests that matter most are the ones asserting what was *not* changed.
"""

from __future__ import annotations

import re
from dataclasses import replace
from pathlib import Path

import pytest

from pyofficeeditor._xml import XmlDocument
from pyofficeeditor.excel import (
    Alignment,
    Border,
    CellFormat,
    Color,
    Fill,
    Font,
    Protection,
    Side,
    Styles,
    Workbook,
)
from pyofficeeditor.excel._formats import RESERVED_FILL_COUNT


@pytest.fixture()
def book(live_sample_xlsx: Path) -> Workbook:
    return Workbook.open(live_sample_xlsx)


@pytest.fixture()
def styles(book: Workbook) -> Styles:
    found = book.styles
    assert found is not None
    return found


def roundtrip(component: Font | Fill | Border | Alignment | Protection) -> object:
    """Serialize a component, reparse it, and read it back."""
    raw = component.write().to_xml().encode()
    element = XmlDocument.parse(raw).root
    return type(component).read(element)  # type: ignore[arg-type]


class TestColor:
    def test_from_rgb_pads_alpha(self) -> None:
        assert Color.from_rgb("FF0000").rgb == "FFFF0000"
        assert Color.from_rgb("FFFF0000").rgb == "FFFF0000"
        assert Color.from_rgb("#00ff00").rgb == "FF00FF00"

    @pytest.mark.parametrize("value", ["", "xyz", "FF00", "GGGGGG", "FF0000FF00"])
    def test_a_bad_hex_value_is_refused(self, value: str) -> None:
        with pytest.raises(ValueError, match="hex color"):
            Color.from_rgb(value)

    @pytest.mark.parametrize(
        "source",
        [
            b'<color rgb="FFFF0000"/>',
            b'<color theme="1"/>',
            b'<color theme="4" tint="0.4"/>',
            b'<color indexed="64"/>',
            b'<color auto="1"/>',
        ],
    )
    def test_each_spelling_round_trips(self, source: bytes) -> None:
        """Theme, rgb, indexed and auto are kept apart rather than converted:
        resolving a theme index needs the theme part, so converting would
        make the round trip lossy."""
        colour = Color.read(XmlDocument.parse(source).root)
        assert colour is not None
        assert colour.write().to_xml().encode() == source

    def test_an_empty_color_is_recognised(self) -> None:
        assert Color().is_empty
        assert not Color(theme=1).is_empty
        assert not Color(auto=True).is_empty


class TestFont:
    def test_the_real_default_font_round_trips(self) -> None:
        source = (
            b'<font><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/>'
            b'<family val="2"/><scheme val="minor"/></font>'
        )
        font = Font.read(XmlDocument.parse(source).root)
        assert font.name == "Aptos Narrow"
        assert font.size == 11
        assert font.color == Color(theme=1)
        assert font.write().to_xml().encode() == source

    def test_the_real_bold_font_round_trips(self) -> None:
        source = (
            b'<font><b/><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/>'
            b'<family val="2"/><scheme val="minor"/></font>'
        )
        font = Font.read(XmlDocument.parse(source).root)
        assert font.bold is True
        assert font.write().to_xml().encode() == source

    def test_a_bare_flag_means_true(self) -> None:
        """``<b/>`` and ``<b val="1"/>`` both mean bold; ``val="0"`` does not."""
        assert Font.read(XmlDocument.parse(b"<font><b/></font>").root).bold is True
        assert Font.read(XmlDocument.parse(b'<font><b val="1"/></font>').root).bold is True
        assert Font.read(XmlDocument.parse(b'<font><b val="0"/></font>').root).bold is False

    def test_a_bare_underline_means_single(self) -> None:
        assert Font.read(XmlDocument.parse(b"<font><u/></font>").root).underline == "single"
        assert (
            Font.read(XmlDocument.parse(b'<font><u val="double"/></font>').root).underline
            == "double"
        )

    def test_single_underline_is_written_bare(self) -> None:
        """Which is how Excel writes it."""
        assert Font(underline="single").write().to_xml() == "<font><u/></font>"
        assert Font(underline="double").write().to_xml() == '<font><u val="double"/></font>'

    @pytest.mark.parametrize(
        "font",
        [
            Font(),
            Font(bold=True),
            Font(bold=True, italic=True, strike=True),
            Font(name="Calibri", size=12),
            Font(size=9.5),
            Font(color=Color.from_rgb("FF0000")),
            Font(underline="doubleAccounting"),
            Font(script="superscript"),
            Font(condense=True, extend=True, outline=True, shadow=True),
            Font(name="X", size=1, family=2, charset=1, scheme="major"),
        ],
    )
    def test_every_aspect_round_trips(self, font: Font) -> None:
        assert roundtrip(font) == font

    def test_children_are_written_in_excels_order(self) -> None:
        font = Font(name="X", size=11, bold=True, italic=True, underline="single", color=Color(theme=1))
        names = re.findall(r"<(\w+)", font.write().to_xml())[1:]
        assert names == ["b", "i", "u", "sz", "color", "name"]

    def test_a_fractional_size_keeps_its_decimal(self) -> None:
        assert '<sz val="9.5"/>' in Font(size=9.5).write().to_xml()
        assert '<sz val="11"/>' in Font(size=11.0).write().to_xml(), "no trailing .0"


class TestFill:
    def test_the_two_reserved_fills_round_trip(self) -> None:
        for source in (
            b'<fill><patternFill patternType="none"/></fill>',
            b'<fill><patternFill patternType="gray125"/></fill>',
        ):
            fill = Fill.read(XmlDocument.parse(source).root)
            assert fill.write().to_xml().encode() == source

    def test_a_solid_fill_puts_the_colour_in_fgcolor(self) -> None:
        """Excel stores a solid fill's colour in fgColor, not bgColor, which
        is the opposite of what the names suggest and the reliable way to
        produce a cell that looks unfilled."""
        raw = Fill.solid("FFFF00").write().to_xml()
        assert 'patternType="solid"' in raw
        assert '<fgColor rgb="FFFFFF00"/>' in raw
        assert "bgColor" not in raw

    def test_solid_accepts_a_colour_or_a_string(self) -> None:
        assert Fill.solid("FF0000") == Fill.solid(Color.from_rgb("FF0000"))

    @pytest.mark.parametrize(
        "fill",
        [
            Fill(),
            Fill.none(),
            Fill.solid("00FF00"),
            Fill(pattern="gray125"),
            Fill(pattern="darkGrid", foreground=Color(theme=2), background=Color(indexed=64)),
        ],
    )
    def test_round_trip(self, fill: Fill) -> None:
        assert roundtrip(fill) == fill

    def test_a_gradient_fill_is_kept_as_it_was(self) -> None:
        """Nothing here can build one, so discarding it would be worse than
        carrying it opaquely."""
        source = (
            b'<fill><gradientFill degree="90"><stop position="0">'
            b'<color rgb="FFFFFFFF"/></stop></gradientFill></fill>'
        )
        fill = Fill.read(XmlDocument.parse(source).root)
        assert fill.is_gradient
        assert fill.write().to_xml().encode() == source


class TestBorder:
    def test_the_real_empty_border_round_trips(self) -> None:
        source = b"<border><left/><right/><top/><bottom/><diagonal/></border>"
        border = Border.read(XmlDocument.parse(source).root)
        assert border.is_empty
        assert border.write().to_xml().encode() == source

    def test_all_sides(self) -> None:
        border = Border.all_sides("medium", "FF0000")
        assert border.left.style == "medium"
        assert border.left.color == Color.from_rgb("FF0000")
        assert border.top == border.bottom == border.left == border.right
        assert border.diagonal.is_empty

    def test_the_four_sides_are_always_written_in_order(self) -> None:
        """CT_Border is a sequence, and Excel writes all five even when
        empty."""
        names = re.findall(r"<(\w+)", Border().write().to_xml())[1:]
        assert names == ["left", "right", "top", "bottom", "diagonal"]

    def test_vertical_and_horizontal_are_written_only_when_used(self) -> None:
        assert "vertical" not in Border().write().to_xml()
        raw = Border(vertical=Side(style="thin")).write().to_xml()
        assert raw.index("vertical") > raw.index("diagonal")

    def test_start_and_end_are_read_as_left_and_right(self) -> None:
        """The modern spellings of the same edges."""
        border = Border.read(
            XmlDocument.parse(b'<border><start style="thin"/><end style="thick"/></border>').root
        )
        assert border.left.style == "thin"
        assert border.right.style == "thick"

    @pytest.mark.parametrize(
        "border",
        [
            Border(),
            Border.all_sides(),
            Border.all_sides("double", "0000FF"),
            Border(diagonal=Side(style="thin"), diagonal_up=True, diagonal_down=True),
            Border(vertical=Side(style="hair"), horizontal=Side(style="dotted")),
            Border(outline=True, left=Side(style="slantDashDot", color=Color(theme=3))),
        ],
    )
    def test_round_trip(self, border: Border) -> None:
        assert roundtrip(border) == border

    def test_an_unknown_style_is_dropped_rather_than_guessed(self) -> None:
        border = Border.read(XmlDocument.parse(b'<border><left style="wavy"/></border>').root)
        assert border.left.style is None


class TestAlignment:
    @pytest.mark.parametrize(
        "alignment",
        [
            Alignment(),
            Alignment(horizontal="center"),
            Alignment(vertical="top"),
            Alignment(wrap_text=True),
            Alignment(shrink_to_fit=True),
            Alignment(indent=2),
            Alignment(text_rotation=90),
            Alignment(text_rotation=255, justify_last_line=True, reading_order=2),
            Alignment(horizontal="distributed", vertical="justify", relative_indent=1),
        ],
    )
    def test_round_trip(self, alignment: Alignment) -> None:
        assert roundtrip(alignment) == alignment

    def test_an_empty_alignment_is_recognised(self) -> None:
        assert Alignment().is_empty
        assert not Alignment(wrap_text=True).is_empty

    def test_an_unknown_value_is_dropped_rather_than_guessed(self) -> None:
        parsed = Alignment.read(XmlDocument.parse(b'<alignment horizontal="sideways"/>').root)
        assert parsed.horizontal is None


class TestProtection:
    def test_the_default_is_locked_and_not_hidden(self) -> None:
        """Which is what an absent protection element means."""
        assert Protection.read(None) == Protection(locked=True, hidden=False)
        assert Protection().is_default

    @pytest.mark.parametrize(
        "protection",
        [Protection(), Protection(locked=False), Protection(hidden=True), Protection(False, True)],
    )
    def test_round_trip(self, protection: Protection) -> None:
        assert roundtrip(protection) == protection


class TestReadingRealFormats:
    def test_the_bold_header(self, book: Workbook) -> None:
        fmt = book["Data"]["A1"].format
        assert fmt.font.bold is True
        assert fmt.font.name == "Aptos Narrow"
        assert fmt.font.size == 11
        assert fmt.fill.pattern == "none"
        assert fmt.border.is_empty
        assert fmt.number_format == "General"

    def test_a_cell_with_no_style_resolves_through_entry_zero(self, book: Workbook) -> None:
        """A cell with no ``s`` is not unformatted: it uses ``cellXfs[0]``,
        which names the workbook's default font. Reporting an empty format
        would make "add bold" silently change the typeface."""
        font = book["Data"]["B2"].font
        assert font.name == "Aptos Narrow"
        assert font.size == 11
        assert font.bold is False

    def test_the_date_cell_keeps_its_number_format(self, book: Workbook) -> None:
        assert book["Data"]["F2"].format.number_format == r"yyyy\-mm\-dd"

    def test_a_style_index_past_the_table(self, styles: Styles) -> None:
        assert styles.cell_format(9999).font == Font()

    def test_the_component_tables_as_excel_wrote_them(self, styles: Styles) -> None:
        assert styles.font(0).name == "Aptos Narrow"
        assert styles.font(1).bold is True
        assert styles.fill(0).pattern == "none"
        assert styles.fill(1).pattern == "gray125"
        assert styles.border(0).is_empty

    def test_an_index_past_a_component_table(self, styles: Styles) -> None:
        assert styles.font(99) == Font()
        assert styles.fill(99) == Fill()
        assert styles.border(99) == Border()


class TestWritingFormats:
    def test_bolding_reuses_the_existing_bold_font(self, book: Workbook, styles: Styles) -> None:
        """B2 resolves through cellXfs[0], so adding bold produces exactly the
        font the header already uses."""
        before = styles.font_count
        book["Data"]["B2"].format = book["Data"]["B2"].format.with_font(bold=True)
        assert styles.font_count == before, "no new font was needed"
        assert book["Data"]["B2"].style_index == 1, "the header's entry was reused"

    def test_a_solid_fill(self, book: Workbook) -> None:
        book["Data"]["B3"].fill = "FFFF00"
        reopened = Workbook.from_bytes(book.to_bytes())
        fill = reopened["Data"]["B3"].fill
        assert fill.pattern == "solid"
        assert fill.foreground == Color.from_rgb("FFFF00")

    def test_a_fill_lands_past_the_reserved_indices(self, book: Workbook, styles: Styles) -> None:
        """Indices 0 and 1 belong to none and gray125."""
        book["Data"]["B3"].fill = "FFFF00"
        index = styles.ensure_fill(Fill.solid("FFFF00"))
        assert index >= RESERVED_FILL_COUNT

    def test_a_border(self, book: Workbook) -> None:
        book["Data"]["B4"].border = Border.all_sides("thin", "FF0000")
        reopened = Workbook.from_bytes(book.to_bytes())
        border = reopened["Data"]["B4"].border
        assert border.left.style == "thin"
        assert border.left.color == Color.from_rgb("FF0000")
        assert border.bottom.style == "thin"

    def test_alignment(self, book: Workbook) -> None:
        book["Data"]["B5"].alignment = Alignment(horizontal="center", wrap_text=True)
        reopened = Workbook.from_bytes(book.to_bytes())
        alignment = reopened["Data"]["B5"].alignment
        assert alignment.horizontal == "center"
        assert alignment.wrap_text is True

    def test_protection_round_trips(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet["B6"].format = replace(sheet["B6"].format, protection=Protection(locked=False))
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"]["B6"].format.protection == Protection(locked=False)

    def test_a_number_format_through_the_format_object(self, book: Workbook) -> None:
        book["Data"]["B7"].format = book["Data"]["B7"].format.with_number_format("0.00%")
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"]["B7"].number_format == "0.00%"

    def test_setting_one_aspect_preserves_the_others(self, book: Workbook) -> None:
        """The reason formats are whole values: a bold header that gains a
        fill has to stay bold."""
        sheet = book["Data"]
        sheet["A1"].fill = "FFFF00"
        reopened = Workbook.from_bytes(book.to_bytes())
        cell = reopened["Data"]["A1"]
        assert cell.font.bold is True, "still bold"
        assert cell.fill.foreground == Color.from_rgb("FFFF00")
        assert cell.value == "Region", "and still holding its value"

    def test_formatting_a_date_keeps_it_a_date(self, book: Workbook) -> None:
        import datetime as dt

        book["Data"]["F2"].font = replace(book["Data"]["F2"].font, bold=True)
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"]["F2"].value == dt.date(2026, 1, 15)
        assert reopened["Data"]["F2"].font.bold is True

    def test_formatting_does_not_disturb_other_cells(self, book: Workbook) -> None:
        """Entries are shared, so a naive implementation that edited one in
        place would repaint every cell using it."""
        sheet = book["Data"]
        assert sheet["A1"].font.bold and sheet["B1"].font.bold
        sheet["A1"].font = replace(sheet["A1"].font, italic=True)
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"]["A1"].font.italic is True
        assert reopened["Data"]["B1"].font.italic is False, "its neighbour is untouched"
        assert reopened["Data"]["B1"].font.bold is True

    def test_the_same_format_twice_adds_one_entry(self, book: Workbook, styles: Styles) -> None:
        book["Data"]["J1"].fill = "FFFF00"
        after_first = styles.cell_format_count
        book["Data"]["J2"].fill = "FFFF00"
        assert styles.cell_format_count == after_first

    def test_a_whole_column_of_one_format_adds_one_entry(self, book: Workbook, styles: Styles) -> None:
        before = styles.cell_format_count
        book["Data"].range("K1:K50").apply_fill("00FF00")
        assert styles.cell_format_count == before + 1
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"]["K50"].fill.foreground == Color.from_rgb("00FF00")

    def test_the_apply_flags_are_set(self, book: Workbook) -> None:
        """They say the entry overrides its named style for that aspect, and
        some readers honour them."""
        book["Data"]["B3"].fill = "FFFF00"
        text = book.package.read("xl/styles.xml").decode()
        assert 'applyFill="1"' in text

    def test_the_table_counts_are_maintained(self, book: Workbook) -> None:
        book["Data"]["B3"].fill = "FFFF00"
        book["Data"]["B4"].border = Border.all_sides()
        book["Data"]["B5"].font = Font(italic=True)
        text = book.package.read("xl/styles.xml").decode()
        for table, entry in (("fonts", "font"), ("fills", "fill"), ("borders", "border"), ("cellXfs", "xf")):
            declared = int(re.search(rf'<{table} count="(\d+)"', text).group(1))  # type: ignore[union-attr]
            actual = len(re.findall(rf"<{entry}[ />]", re.search(rf"<{table}.*?</{table}>", text, re.S).group(0)))  # type: ignore[union-attr]
            assert declared == actual, f"{table}: says {declared}, has {actual}"

    def test_reading_a_format_is_not_a_change(self, book: Workbook, live_sample_xlsx: Path) -> None:
        for sheet in book:
            for row in sheet.rows():
                for cell in row:
                    _ = cell.format, cell.font, cell.fill, cell.border, cell.alignment
        assert book.is_modified is False
        assert book.to_bytes() == live_sample_xlsx.read_bytes()


class TestRangeFormatting:
    def test_apply_font_keeps_each_cells_other_formatting(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet["C2"].fill = "FFFF00"
        sheet.range("C2:C5").apply_font(italic=True)
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"]["C2"].font.italic is True
        assert reopened["Data"]["C2"].fill.foreground == Color.from_rgb("FFFF00"), "kept its fill"
        assert reopened["Data"]["C3"].font.italic is True
        assert reopened["Data"]["C3"].fill.pattern in (None, "none"), "and C3 gained none"

    def test_apply_fill(self, book: Workbook) -> None:
        book["Data"].range("A1:C1").apply_fill("CCCCCC")
        reopened = Workbook.from_bytes(book.to_bytes())
        for reference in ("A1", "B1", "C1"):
            assert reopened["Data"][reference].fill.foreground == Color.from_rgb("CCCCCC")
            assert reopened["Data"][reference].font.bold is True, "the header stayed bold"

    def test_apply_border(self, book: Workbook) -> None:
        book["Data"].range("A1:B2").apply_border(Border.all_sides("medium"))
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"]["B2"].border.top.style == "medium"

    def test_apply_alignment(self, book: Workbook) -> None:
        book["Data"].range("A1:B1").apply_alignment(horizontal="center")
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"]["A1"].alignment.horizontal == "center"

    def test_apply_number_format(self, book: Workbook) -> None:
        book["Data"].range("C2:C5").apply_number_format("0.00")
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"]["C2"].number_format == "0.00"
        assert reopened["Data"]["C2"].value == 4.25, "the value is untouched"

    def test_set_format_replaces_wholesale(self, book: Workbook) -> None:
        book["Data"].range("A1:B1").set_format(CellFormat(font=Font(italic=True)))
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"]["A1"].font.italic is True
        assert reopened["Data"]["A1"].font.bold is False, "a whole replacement drops the bold"


class TestFunctionalUpdates:
    def test_with_font(self) -> None:
        base = CellFormat(font=Font(name="X", size=11, bold=True))
        assert base.with_font(italic=True).font == Font(name="X", size=11, bold=True, italic=True)
        assert base.font.italic is False, "the original is untouched"

    def test_with_fill_accepts_three_kinds(self) -> None:
        base = CellFormat()
        assert base.with_fill("FF0000").fill == Fill.solid("FF0000")
        assert base.with_fill(Color.from_rgb("FF0000")).fill == Fill.solid("FF0000")
        assert base.with_fill(Fill.none()).fill == Fill.none()

    def test_with_alignment(self) -> None:
        base = CellFormat(alignment=Alignment(horizontal="left"))
        updated = base.with_alignment(wrap_text=True)
        assert updated.alignment == Alignment(horizontal="left", wrap_text=True)

    def test_with_number_format(self) -> None:
        assert CellFormat().with_number_format("0.00").number_format == "0.00"

    def test_the_default_format(self) -> None:
        default = CellFormat()
        assert default.number_format == "General"
        assert default.font == Font()
        assert default.fill == Fill()
        assert default.border.is_empty
        assert default.alignment.is_empty
        assert default.protection.is_default


class TestBuildingStylesFromNothing:
    def test_tables_are_created_in_schema_order(self) -> None:
        """CT_Stylesheet is a sequence, so a table cannot just be appended."""
        document = XmlDocument.parse(b"<styleSheet><cellStyles/></styleSheet>")
        styles = Styles(document)
        styles.ensure_cell_format(CellFormat(font=Font(bold=True)))
        from pyofficeeditor.excel._schema import STYLESHEET_CHILD_ORDER

        names = [
            child.name for child in document.root.elements() if child.name in STYLESHEET_CHILD_ORDER
        ]
        positions = [STYLESHEET_CHILD_ORDER.index(n) for n in names]
        assert positions == sorted(positions), names

    def test_the_reserved_fills_are_created_when_the_table_is_empty(self) -> None:
        document = XmlDocument.parse(b"<styleSheet></styleSheet>")
        styles = Styles(document)
        index = styles.ensure_fill(Fill.solid("FF0000"))
        assert index == RESERVED_FILL_COUNT, "after none and gray125"
        assert styles.fill(0).pattern == "none"
        assert styles.fill(1).pattern == "gray125"

    def test_an_existing_fill_table_is_not_renumbered(self) -> None:
        """Inserting the reserved fills in front of existing entries would
        shift every fillId in the workbook and repaint every cell."""
        document = XmlDocument.parse(
            b'<styleSheet><fills count="1"><fill><patternFill patternType="solid">'
            b'<fgColor rgb="FFFF0000"/></patternFill></fill></fills></styleSheet>'
        )
        styles = Styles(document)
        assert styles.ensure_fill(Fill.solid("FF0000")) == 0, "the existing entry keeps its index"
        assert styles.fill_count == 1

    def test_a_format_built_from_nothing_resolves(self) -> None:
        document = XmlDocument.parse(b"<styleSheet></styleSheet>")
        styles = Styles(document)
        wanted = CellFormat(
            number_format="0.00",
            font=Font(bold=True, size=14),
            fill=Fill.solid("00FF00"),
            border=Border.all_sides("thick"),
            alignment=Alignment(horizontal="center"),
        )
        index = styles.ensure_cell_format(wanted)
        assert styles.cell_format(index) == wanted

    def test_asking_twice_adds_once(self) -> None:
        document = XmlDocument.parse(b"<styleSheet></styleSheet>")
        styles = Styles(document)
        wanted = CellFormat(font=Font(italic=True))
        first = styles.ensure_cell_format(wanted)
        count = styles.cell_format_count
        assert styles.ensure_cell_format(wanted) == first
        assert styles.cell_format_count == count

    def test_a_workbook_with_no_styles_part_refuses_a_format(
        self, book: Workbook, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Rather than silently dropping it: a format with nowhere to live
        would leave the cell looking unchanged with no error to explain it."""
        monkeypatch.setattr(Workbook, "styles", property(lambda _self: None))
        with pytest.raises(ValueError, match="no styles part"):
            book["Data"]["A1"].fill = "FF0000"


class TestStyleElementShape:
    def test_an_xf_writes_alignment_before_protection(self) -> None:
        """CT_Xf is a sequence."""
        document = XmlDocument.parse(b"<styleSheet></styleSheet>")
        styles = Styles(document)
        styles.ensure_cell_format(
            CellFormat(alignment=Alignment(wrap_text=True), protection=Protection(locked=False))
        )
        raw = document.to_bytes().decode()
        assert raw.index("<alignment") < raw.index("<protection")

    def test_an_empty_alignment_is_not_written(self) -> None:
        document = XmlDocument.parse(b"<styleSheet></styleSheet>")
        styles = Styles(document)
        styles.ensure_cell_format(CellFormat(font=Font(bold=True)))
        assert b"<alignment" not in document.to_bytes()
        assert b"<protection" not in document.to_bytes()

    def test_the_cell_element_gains_an_s_attribute(self, book: Workbook) -> None:
        sheet = book["Data"]
        sheet["J1"].value = 1
        sheet["J1"].fill = "FF0000"
        raw = book.package.read(sheet.part_name).decode()
        assert re.search(r'<c r="J1" s="\d+"', raw), raw[raw.index('r="J1"') - 20 :][:80]

    def test_formatting_an_absent_cell_creates_it(self, book: Workbook) -> None:
        sheet = book["Data"]
        assert sheet["Z80"].exists is False
        sheet["Z80"].fill = "FF0000"
        assert sheet["Z80"].exists is True
        reopened = Workbook.from_bytes(book.to_bytes())
        assert reopened["Data"]["Z80"].fill.foreground == Color.from_rgb("FF0000")
        assert reopened["Data"]["Z80"].value is None, "formatted but empty"
