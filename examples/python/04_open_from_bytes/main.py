"""04_open_from_bytes: build a DOCX, an XLSX and a PPTX in memory and open each from bytes.

Run: python examples/python/04_open_from_bytes/main.py
"""
import base64
import os
import sys
import tempfile

from office_oxide import Document, EditableDocument, PptxWriter, XlsxWriter, create_from_markdown


MARKDOWN = """\
# Quarterly Report

Revenue grew in every region.

| Region | Q1  | Q2  |
|--------|-----|-----|
| North  | 120 | 135 |
| South  | 95  | 110 |
"""


def build_docx_bytes(title: str = "Quarterly Report") -> bytes:
    # The only DOCX writer is create_from_markdown, which writes a path; the
    # file lives in the temp dir just long enough to read its bytes back.
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        create_from_markdown(MARKDOWN.replace("Quarterly Report", title), "docx", path)
        with open(path, "rb") as f:
            return f.read()
    finally:
        os.unlink(path)

def build_xlsx_bytes() -> bytes:
    w = XlsxWriter()
    s = w.add_sheet("Sales")
    for col, head in enumerate(["Region", "Q1", "Q2", "Growth"]):
        w.set_cell_styled(s, 0, col, head, bold=True)  # bold header row
    for row, (region, q1, q2) in enumerate([("North", 120, 135), ("South", 95, 110)], 1):
        w.set_cell(s, row, 0, region)
        w.set_cell(s, row, 1, q1)
        w.set_cell(s, row, 2, q2)
    w.set_cell(s, 1, 3, True)  # D2
    w.set_cell(s, 3, 0, "Total")
    w.set_formula(s, 3, 1, "SUM(B2:B3)")
    w.set_formula(s, 3, 2, "SUM(C2:C3)")
    w.set_cell(s, 5, 0, "Notes")
    w.merge_cells(s, 5, 0, 1, 3)  # A6:C6
    w.set_column_width(s, 0, 18.0)
    return w.to_bytes()

PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="
)


def build_pptx_bytes() -> bytes:
    w = PptxWriter()
    w.set_presentation_size(12192000, 6858000)  # 16:9, in EMU
    one = w.add_slide()
    w.set_slide_title(one, "Welcome")
    w.add_slide_text(one, "Quarterly Report")
    two = w.add_slide()
    w.set_slide_title(two, "Results")
    w.add_slide_text(two, "Revenue grew in every region.")
    w.add_slide_image(two, PNG, "png", 914400, 2743200, 1828800, 1828800)
    return w.to_bytes()

def check(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def main() -> None:
    print("step 1: build a DOCX, an XLSX and a PPTX in memory")
    built = {"docx": build_docx_bytes(), "xlsx": build_xlsx_bytes(), "pptx": build_pptx_bytes()}

    print("step 2: open each from bytes")
    texts = {}
    for fmt, data in built.items():
        with Document.from_bytes(data, fmt) as doc:
            check(doc.format == fmt, "the opened format differs from the requested one")
            texts[fmt] = doc.plain_text()
        print(f"  format: {fmt}")
        print(f"  chars: {len(texts[fmt])}")

    print("step 3: check the text")
    check("Revenue grew in every region." in texts["docx"], "DOCX text is missing")
    check("North" in texts["xlsx"] and "Total" in texts["xlsx"], "XLSX text is missing")
    check("Welcome" in texts["pptx"] and "Results" in texts["pptx"], "PPTX text is missing")

    print("step 4: edit the DOCX in memory and reopen it")
    with EditableDocument.from_bytes(built["docx"], "docx") as ed:
        check(ed.replace_text("every region", "all regions") == 1, "replace_text did not report one replacement")
        edited = ed.to_bytes()
    with Document.from_bytes(edited, "docx") as doc:
        check("Revenue grew in all regions." in doc.plain_text(), "the edit is missing after reopening")

    print("OK 04_open_from_bytes")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"error: {exc!r}", file=sys.stderr)
        sys.exit(1)
