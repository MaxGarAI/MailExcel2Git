from zipfile import ZipFile

import pytest

from kuchenland_importer.infrastructure.excel.calculation_settings import restore_calculation_mode


@pytest.mark.parametrize("mode", ["auto", "manual", "autoNoTable"])
def test_calculation_restore_preserves_rich_data_and_xml_namespaces(tmp_path, mode):
    target = tmp_path / "saved.xlsx"
    xml = (
        b'<workbook xmlns="urn:test" xmlns:xr="urn:revision" '
        b'xmlns:mc="urn:compatibility" mc:Ignorable="xr"><calcPr calcMode="manual"/>'
        b"</workbook>"
    )
    entries = {
        "xl/workbook.xml": xml,
        "xl/richData/rdrichvalue.xml": b"native metadata",
        "xl/media/image1.png": b"unchanged picture",
    }
    with ZipFile(target, "x") as book:
        book.comment = b"preserved archive comment"
        for name, content in entries.items():
            book.writestr(name, content)
    restore_calculation_mode(target, mode)
    with ZipFile(target) as book:
        assert book.comment == b"preserved archive comment"
        assert set(book.namelist()) == set(entries)
        for name, content in entries.items():
            if name != "xl/workbook.xml":
                assert book.read(name) == content
        saved = book.read("xl/workbook.xml")
        assert f'calcMode="{mode}"'.encode() in saved
        assert saved[: saved.index(b"<calcPr")] == xml[: xml.index(b"<calcPr")]


def test_missing_calculation_element_does_not_damage_working_copy(tmp_path):
    target = tmp_path / "saved.xlsx"
    with ZipFile(target, "x") as book:
        book.writestr("xl/workbook.xml", b"<workbook/>")
    original = target.read_bytes()
    with pytest.raises(ValueError, match="calcPr"):
        restore_calculation_mode(target, "auto")
    assert target.read_bytes() == original
    assert not list(tmp_path.glob(".kl-calc-*"))
