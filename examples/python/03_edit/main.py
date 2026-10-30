"""03_edit: replace text in a DOCX and set cells in an XLSX, in memory.

Run: python examples/python/03_edit/main.py
"""
import os
import sys
import tempfile

from office_oxide import Document, EditableDocument, XlsxWriter, create_from_markdown


MARKDOWN = """\
# Quarterly Report

Revenue grew in every region.

| Region | Q1  | Q2  |
|--------|-----|-----|
| North  | 120 | 135 |
| South  | 95  | 110 |
"""


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


def check(condition: bool, message: str) -> None:
    if not condition:
        print(f"FAIL: {message}", file=sys.stderr)
        sys.exit(1)


def has(text: str, expected: str) -> None:
    check(expected in text, f'"{expected}" is missing from the text')


def main() -> None:
    print("step 1: build a DOCX in memory")
    docx = build_docx_bytes()

    print("step 2: replace text")
    with EditableDocument.from_bytes(docx, "docx") as ed:
        replaced = ed.replace_text("Quarterly Report", "Annual Report")
        print(f"  replacements: {replaced}")
        check(replaced == 1, "replace_text did not make exactly one replacement")

        print("step 3: save to bytes and check the text")
        edited = ed.to_bytes()
    with Document.from_bytes(edited, "docx") as doc:
        text = doc.plain_text()
    has(text, "Annual Report")
    check("Quarterly Report" not in text, "the old title is still present")

    print("step 4: build an XLSX in memory")
    xlsx = build_xlsx_bytes()

    print("step 5: set cells")
    with EditableDocument.from_bytes(xlsx, "xlsx") as ed:
        ed.set_cell(0, "B2", 999)
        ed.set_cell(0, "A2", "Central")

        print("step 6: save to bytes and check the cells")
        edited = ed.to_bytes()
    with Document.from_bytes(edited, "xlsx") as doc:
        text = doc.plain_text()
    has(text, "Central")
    has(text, "999")
    check("North" not in text, "the old A2 value is still present")

    print("OK 03_edit")


if __name__ == "__main__":
    main()
