# SPDX-License-Identifier: MIT OR Apache-2.0
"""Behavioural tests for the Python binding's writer and editing surface."""

import functools

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
    # The process-wide limit, which new options start from, is untouched.
    assert office_oxide.ParseOptions().limits()["text_chars"] == 256 * 2**20


# Every option, as every binding names it.
OPTION_NAMES = [
    "include_hidden_text", "include_deleted_text", "on_incomplete", "lazy_pictures",
    "package_entries", "package_bytes", "part_bytes", "inflate_ratio",
    "embedded_depth", "embedded_packages", "text_chars", "memory_bytes",
]


def test_the_options_name_every_option():
    assert sorted(office_oxide.ParseOptions.names()) == sorted(OPTION_NAMES)
    for name in OPTION_NAMES:
        office_oxide.ParseOptions().set(name, "report" if name == "on_incomplete" else 1)


@functools.lru_cache(maxsize=1)
def _big_workbook() -> bytes:
    w = XlsxWriter()
    w.add_sheet("S")
    for r in range(60000):
        for c in range(8):
            w.set_cell(0, r, c, f"cell {r} {c}")
    return w.to_bytes()


def _python_ticks_inside(work) -> tuple[int, int]:
    """How many times a pure-Python thread ran while `work()` was inside the
    library, over repeated calls for about 0.3 s, and how many calls that was.
    A call that keeps the GIL starves every other Python thread for its whole
    length. The interpreter can still switch threads in the few bytecodes
    between taking a timestamp and entering the library, so only ticks well
    inside a call count, with a short switch interval to keep that edge small.
    """
    import sys
    import threading
    import time

    margin = 0.002
    old_interval = sys.getswitchinterval()
    sys.setswitchinterval(0.0001)

    ticks: list[float] = []
    stop = threading.Event()

    def spin():
        while not stop.is_set():
            ticks.append(time.perf_counter())
            time.sleep(0.0005)

    t = threading.Thread(target=spin)
    t.start()
    time.sleep(0.02)
    calls: list[tuple[float, float]] = []
    deadline = time.perf_counter() + 0.3
    while time.perf_counter() < deadline:
        start = time.perf_counter()
        work()
        calls.append((start, time.perf_counter()))
    stop.set()
    t.join()
    sys.setswitchinterval(old_interval)
    inside = sum(1 for x in ticks for a, b in calls if a + margin < x < b - margin)
    return inside, len(calls)


@pytest.mark.parametrize(
    "call",
    ["open", "plain_text", "to_markdown", "to_html", "to_ir_json", "damage"],
)
def test_library_calls_release_the_gil(call):
    # A parse that keeps the GIL starves every other Python thread, and a host
    # whose logging handler needs the GIL deadlocks against the parse thread.
    data = _big_workbook()
    doc = Document.from_bytes(data, "xlsx")
    work = {
        "open": lambda: Document.from_bytes(data, "xlsx"),
        "plain_text": doc.plain_text,
        "to_markdown": doc.to_markdown,
        "to_html": doc.to_html,
        "to_ir_json": doc.to_ir_json,
        "damage": doc.damage,
    }[call]
    inside, calls = _python_ticks_inside(work)
    assert inside >= 3, f"{call}: another Python thread ran {inside} times inside {calls} calls"


def test_a_logging_handler_that_needs_the_gil_does_not_hang_a_parse():
    # The parse runs on a helper thread; if the calling thread keeps the GIL
    # while it waits, a log record the helper emits through a handler that
    # takes the GIL never completes. Run the parse from several threads with
    # logging at DEBUG and a timeout, so a hang fails instead of sitting.
    import logging
    from concurrent.futures import ThreadPoolExecutor

    logging.basicConfig(level=logging.DEBUG)
    data = _big_workbook()
    with ThreadPoolExecutor(max_workers=4) as pool:
        futures = [
            pool.submit(lambda: len(Document.from_bytes(data, "xlsx").plain_text()))
            for _ in range(8)
        ]
        for f in futures:
            assert f.result(timeout=120) > 0


def test_options_report_their_limits_and_the_bounds_derived_from_them():
    opts = office_oxide.ParseOptions(memory_bytes=1000, part_bytes=2**64 - 1)
    limits = opts.limits()
    settings = {"include_hidden_text", "include_deleted_text", "on_incomplete", "lazy_pictures"}
    assert set(limits) == set(office_oxide.ParseOptions.names()) - settings
    assert limits["memory_bytes"] == 1000
    assert limits["part_bytes"] is None
    derived = opts.derived()
    assert derived["write_bytes"] == 500 and derived["image_bytes"] == 500
    assert office_oxide.ParseOptions(strict=True).limits()["embedded_depth"] == 1


def _astral_workbook(rows: int = 1) -> bytes:
    w = XlsxWriter()
    w.add_sheet("S")
    for r in range(rows):
        w.set_cell(0, r, 0, "omega \U0001d74e and \xe9 " + "x" * 200)
    return w.to_bytes()


@pytest.mark.parametrize(
    "as_str, as_bytes",
    [
        ("plain_text", "plain_text_bytes"),
        ("to_markdown", "to_markdown_bytes"),
        ("to_markdown_with_images", "to_markdown_with_images_bytes"),
        ("to_html", "to_html_bytes"),
        ("to_ir_json", "to_ir_json_bytes"),
    ],
)
def test_bytes_forms_equal_the_string_forms_encoded(as_str, as_bytes):
    doc = Document.from_bytes(_astral_workbook(), "xlsx")
    text = getattr(doc, as_str)()
    assert "\U0001d74e" in text
    data = getattr(doc, as_bytes)()
    assert isinstance(data, bytes)
    assert data == text.encode("utf-8")


@pytest.mark.parametrize(
    "write, as_str",
    [
        ("write_text", "plain_text"),
        ("write_markdown", "to_markdown"),
        ("write_html", "to_html"),
        ("write_ir_json", "to_ir_json"),
    ],
)
def test_writing_to_a_path_or_a_file_gives_the_string_form_encoded(tmp_path, write, as_str):
    import io

    # Over a megabyte of output, so a file object receives several chunks
    # (and under the 10,000 rows a worksheet's IR holds).
    doc = Document.from_bytes(_astral_workbook(rows=8_000), "xlsx")
    expected = getattr(doc, as_str)().encode("utf-8")
    assert len(expected) > 1 << 20
    path = tmp_path / "out"
    getattr(doc, write)(path)
    assert path.read_bytes() == expected
    getattr(doc, write)(str(path))
    assert path.read_bytes() == expected
    buf = io.BytesIO()
    getattr(doc, write)(buf)
    assert buf.getvalue() == expected


def test_markdown_with_images_can_be_written_and_bad_targets_are_refused(tmp_path):
    doc = Document.from_bytes(_astral_workbook(), "xlsx")
    path = tmp_path / "md"
    doc.write_markdown(path, images=True)
    assert path.read_bytes() == doc.to_markdown_with_images().encode("utf-8")
    with pytest.raises(TypeError):
        doc.write_text(42)
    with pytest.raises(OSError):
        doc.write_text(tmp_path / "missing-dir" / "out")


def test_lazy_pictures_change_no_output():
    from office_oxide import PptxWriter

    w = PptxWriter()
    w.add_slide()
    w.set_slide_title(0, "Pictures")
    png = (
        b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89"
        + b"\x00" * 200_000
    )
    w.add_slide_image(0, png, "png", 0, 0, 100, 100)
    data = w.to_bytes()
    assert "lazy_pictures" in office_oxide.ParseOptions.names()
    eager = Document.from_bytes(data, "pptx")
    lazy = Document.from_bytes(data, "pptx", options=office_oxide.ParseOptions(lazy_pictures=True))
    assert lazy.to_markdown_with_images() == eager.to_markdown_with_images()
    assert lazy.to_ir_json() == eager.to_ir_json()
    assert lazy.plain_text() == eager.plain_text()


def test_closing_a_document_while_another_thread_reads_it_waits_for_nothing_and_raises_nothing():
    """With the GIL released during a call, closing the document from
    another thread raised "Already borrowed"; the call now finishes on its
    own reference and later calls see the document closed."""
    import threading

    doc = Document.from_bytes(_big_workbook(), "xlsx")
    started = threading.Event()
    result = {}

    def read():
        started.set()
        try:
            result["text"] = doc.plain_text()
        except Exception as e:  # noqa: BLE001 - the test reports any
            result["error"] = e

    t = threading.Thread(target=read)
    t.start()
    started.wait()
    doc.close()
    t.join(timeout=120)
    if "error" in result:
        assert "closed" in str(result["error"]), result["error"]
    else:
        assert "cell 0 0" in result["text"]
    with pytest.raises(RuntimeError, match="closed"):
        doc.plain_text()


def test_an_editable_document_and_a_writer_take_calls_from_many_threads():
    """Edits and writer calls from several threads wait for each other, as
    the GIL made them, instead of raising "Already borrowed"."""
    from concurrent.futures import ThreadPoolExecutor

    w = XlsxWriter()
    w.add_sheet("S")
    for r in range(2000):
        w.set_cell(0, r, 0, f"word{r % 7}")
    data = w.to_bytes()
    ed = EditableDocument.from_bytes(data, "xlsx")
    with ThreadPoolExecutor(8) as pool:
        saved = list(pool.map(lambda _: len(ed.to_bytes()), range(16)))
        written = list(pool.map(lambda i: w.set_cell(0, 3000 + i, 1, "x") or len(w.to_bytes()), range(16)))
    assert all(n > 0 for n in saved) and all(n > 0 for n in written)


def test_a_file_object_that_raises_raises_its_own_exception():
    class Refused(Exception):
        pass

    class Sink:
        def write(self, b):
            raise Refused("full")

    doc = Document.from_bytes(_big_workbook(), "xlsx")
    with pytest.raises(Refused, match="full"):
        doc.write_text(Sink())


def test_a_file_object_that_takes_part_of_each_write_gets_all_of_it():
    class Trickle:
        def __init__(self):
            self.parts = []

        def write(self, b):
            self.parts.append(bytes(b[:777]))
            return min(len(b), 777)

    doc = Document.from_bytes(_big_workbook(), "xlsx")
    sink = Trickle()
    doc.write_text(sink)
    assert b"".join(sink.parts) == doc.plain_text().encode("utf-8")
