"""05_xlsx_writer: write a workbook with typed cells, a formula, a style, a merge and a column width, then read it back.

Run: python examples/python/05_xlsx_writer/main.py
"""
import sys

from office_oxide import Document, XlsxWriter


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
    print("step 1: build an XLSX in memory")
    data = build_xlsx_bytes()

    print("step 2: open it from bytes")
    with Document.from_bytes(data, "xlsx") as doc:
        text = doc.plain_text()
        sheets = len(doc.to_ir()["sections"])
        print(f"  format: {doc.format}")
        print(f"  chars: {len(text)}")
        print(f"  sheets: {sheets}")

    print("step 3: check the cells")
    check(sheets == 1, "expected one sheet")
    for expected in ("Sales", "Region", "Q1", "Q2", "Growth", "North", "120", "135", "South", "95", "110", "Total", "Notes"):
        has(text, expected)
    check("TRUE" in text, "the boolean cell is missing")
    check("SUM(B2:B3)" in text and "SUM(C2:C3)" in text, "a formula is missing")

    print("OK 05_xlsx_writer")


if __name__ == "__main__":
    main()
