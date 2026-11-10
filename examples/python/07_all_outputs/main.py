"""07_all_outputs: render one document as plain text, Markdown, HTML and IR JSON.

Run: python examples/python/07_all_outputs/main.py
"""
import json
import os
import sys
import tempfile

from office_oxide import Document, create_from_markdown


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


def check(condition: bool, message: str) -> None:
    if not condition:
        print(f"FAIL: {message}", file=sys.stderr)
        sys.exit(1)


def has(text: str, expected: str) -> None:
    check(expected in text, f'"{expected}" is missing from the text')


def main() -> None:
    print("step 1: build a DOCX in memory")
    data = build_docx_bytes()

    print("step 2: open it from bytes")
    with Document.from_bytes(data, "docx") as doc:
        print(f"  format: {doc.format}")

        print("step 3: render every output")
        outputs = {
            "plain": doc.plain_text(),      # words only: no markup, no structure
            "markdown": doc.to_markdown(),  # headings and pipe tables survive as syntax
            "html": doc.to_html(),          # a fragment (no <html> wrapper), tables as <table>
            "ir_json": doc.to_ir_json(),    # the format-neutral tree as JSON, for programs
        }
    for name, value in outputs.items():
        print(f"  chars: {name} {len(value)}")

    print("step 4: check each output")
    for name, value in outputs.items():
        check(value.strip() != "", f"the {name} output is empty")
    has(outputs["plain"], "Quarterly Report")
    check("# Quarterly Report" in outputs["markdown"], "the Markdown is missing the heading")
    check("<table" in outputs["html"], "the HTML has no table")
    check(isinstance(json.loads(outputs["ir_json"]), dict), "the IR JSON does not parse to an object")

    print("OK 07_all_outputs")


if __name__ == "__main__":
    main()
