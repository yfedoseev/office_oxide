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
