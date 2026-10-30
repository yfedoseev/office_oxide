"""04_open_from_bytes: build a DOCX, an XLSX and a PPTX in memory and open each from its bytes.

Run: python examples/python/04_open_from_bytes/main.py
"""
import base64
import os
import sys
import tempfile

from office_oxide import Document, OfficeOxideError, PptxWriter, XlsxWriter, create_from_markdown


MARKDOWN = """\
# Quarterly Report

Revenue grew in every region.

| Region | Q1  | Q2  |
|--------|-----|-----|
| North  | 120 | 135 |
| South  | 95  | 110 |
"""


PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="
)


def markdown_to_bytes(markdown: str, fmt: str) -> bytes:
    # create_from_markdown writes to a path, so the file lives in the temp dir
    # just long enough to read its bytes back.
    fd, path = tempfile.mkstemp(suffix=f".{fmt}")
    os.close(fd)
    try:
        create_from_markdown(markdown, fmt, path)
        with open(path, "rb") as f:
            return f.read()
    finally:
        os.unlink(path)


def build_docx_bytes(title: str = "Quarterly Report") -> bytes:
    return markdown_to_bytes(MARKDOWN.replace("Quarterly Report", title), "docx")


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


def build_pptx_bytes() -> bytes:
    w = PptxWriter()
    w.set_presentation_size(12192000, 6858000)  # 16:9, in EMU (914400 per inch)
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
        print(f"FAIL: {message}", file=sys.stderr)
        sys.exit(1)


def has(text: str, expected: str) -> None:
    check(expected in text, f'"{expected}" is missing from the text')


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
    has(texts["docx"], "Revenue grew in every region.")
    has(texts["xlsx"], "North")
    has(texts["xlsx"], "Total")
    has(texts["pptx"], "Welcome")
    has(texts["pptx"], "Results")

    print("step 4: check that a wrong format is an error")
    try:
        Document.from_bytes(built["xlsx"], "pptx")
    except OfficeOxideError as exc:
        print(f"  error: {exc}")
    else:
        check(False, "an XLSX opened as a PPTX was accepted")

    print("OK 04_open_from_bytes")


if __name__ == "__main__":
    main()
