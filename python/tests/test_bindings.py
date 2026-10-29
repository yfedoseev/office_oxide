# SPDX-License-Identifier: MIT OR Apache-2.0
"""Behavioural tests for the Python binding's writer and editing surface."""

import office_oxide
import pytest
from office_oxide import Document, EditableDocument, XlsxWriter


def _cells(data: bytes) -> str:
    """The IR JSON of an XLSX byte string, for asserting cell kinds."""
    return Document.from_bytes(data, "xlsx").to_ir_json()


def test_writer_bool_is_written_as_a_boolean_not_a_number():
    # bool is a subclass of int, and the float arm accepted it first, so
    # True became the number 1.
    w = XlsxWriter()
    w.add_sheet("S")
    w.set_cell(0, 0, 0, True)
    w.set_cell_styled(0, 1, 0, False, True)
    text = Document.from_bytes(w.to_bytes(), "xlsx").plain_text()
    assert "TRUE" in text and "FALSE" in text, text


def test_writer_formula_api():
    w = XlsxWriter()
    w.add_sheet("S")
    w.set_cell(0, 0, 0, 2)
    w.set_cell(0, 1, 0, 3)
    w.set_formula(0, 2, 0, "=SUM(A1:A2)")
    w.set_cell(0, 3, 0, "=not a formula")
    ir = _cells(w.to_bytes())
    assert "SUM(A1:A2)" in ir
    assert "=not a formula" in ir


def test_int_beyond_double_precision_is_rejected():
    w = XlsxWriter()
    w.add_sheet("S")
    w.set_cell(0, 0, 0, 2**53)
    with pytest.raises(ValueError):
        w.set_cell(0, 1, 0, 2**53 + 1)
    with pytest.raises(ValueError):
        w.set_cell(0, 1, 0, 2**70)
    with pytest.raises(ValueError):
        w.set_cell_styled(0, 1, 0, -(2**60) - 1, False)


def _xlsx_bytes() -> bytes:
    w = XlsxWriter()
    w.add_sheet("S")
    w.set_cell(0, 0, 0, "old")
    return w.to_bytes()


def test_editable_document_from_bytes_and_to_bytes():
    ed = EditableDocument.from_bytes(_xlsx_bytes(), "xlsx")
    ed.set_cell(0, "A1", True)
    ed.set_cell(0, "B1", 7)
    with pytest.raises(ValueError):
        ed.set_cell(0, "C1", 2**53 + 1)
    out = ed.to_bytes()
    text = Document.from_bytes(out, "xlsx").plain_text()
    assert "TRUE" in text and "7" in text, text


def test_empty_find_is_rejected():
    md = "Hello world\n"
    import os
    import tempfile

    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "a.docx")
        office_oxide.create_from_markdown(md, "docx", path)
        with EditableDocument.open(path) as ed, pytest.raises(ValueError):
            ed.replace_text("", "X")


def test_markdown_with_images_is_typed_and_callable():
    w = XlsxWriter()
    w.add_sheet("S")
    w.set_cell(0, 0, 0, "x")
    doc = Document.from_bytes(w.to_bytes(), "xlsx")
    assert "x" in doc.to_markdown_with_images()


# --- per-document options -----------------------------------------------------


def _cut_off_docx() -> bytes:
    """A .docx whose body is cut off part-way (built with the writer API)."""
    import io
    import zipfile

    wml = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = f'<w:document xmlns:w="{wml}"><w:body><w:p><w:r><w:t>kept</w:t></w:r></w:p><w:p><w:r><w:t>cut'
    ct = (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    )
    rels = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    )
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", body)
    return out.getvalue()


def test_damage_lists_what_is_wrong_and_a_sound_document_has_none():
    doc = Document.from_bytes(_cut_off_docx(), "docx")
    kinds = {d["kind"] for d in doc.damage()}
    assert "unreadable_part" in kinds, doc.damage()
    assert any("word/document.xml" in (d["part"] or "") for d in doc.damage())

    w = XlsxWriter()
    w.add_sheet("S")
    w.set_cell(0, 0, 0, "x")
    sound = Document.from_bytes(w.to_bytes(), "xlsx")
    assert [d for d in sound.damage() if d["kind"] != "recovered"] == []


def test_strict_options_refuse_an_incomplete_document_with_a_typed_error():
    data = _cut_off_docx()
    with pytest.raises(office_oxide.IncompleteDocumentError):
        Document.from_bytes(
            data, "docx", options=office_oxide.ParseOptions(strict=True)
        )
    with pytest.raises(office_oxide.OfficeOxideError):  # it is also the base error
        Document.from_bytes(
            data, "docx", options=office_oxide.ParseOptions(on_incomplete="error")
        )
    # The default reports instead of raising.
    assert "kept" in Document.from_bytes(data, "docx").plain_text()


def test_options_are_set_by_name_and_bad_names_and_values_are_refused():
    opts = office_oxide.ParseOptions(text_chars=1000, include_hidden_text=True)
    opts.set("part_bytes", 1 << 20)
    assert "text_chars" in office_oxide.ParseOptions.names()
    with pytest.raises(ValueError):
        office_oxide.ParseOptions(no_such_limit=1)
    with pytest.raises(ValueError):
        office_oxide.ParseOptions(on_incomplete="sometimes")
    with pytest.raises(ValueError):
        opts.set("text_chars", "lots")


def test_a_limit_in_the_options_applies_to_that_document_only():
    w = XlsxWriter()
    w.add_sheet("S")
    for i in range(200):
        w.set_cell(0, i, 0, "x" * 200)
    data = w.to_bytes()
    small = Document.from_bytes(
        data, "xlsx", options=office_oxide.ParseOptions(text_chars=5000)
    )
    plain = Document.from_bytes(data, "xlsx")
    assert len(small.plain_text()) < len(plain.plain_text())
