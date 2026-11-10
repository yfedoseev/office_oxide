"""01_extract: read a document as plain text, Markdown and IR JSON.

Run: python examples/python/01_extract/main.py
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
        check(doc.format == "docx", "the opened format differs from the requested one")

        print("step 3: extract plain text, Markdown and IR JSON")
        text = doc.plain_text()
        markdown = doc.to_markdown()
        ir_json = doc.to_ir_json()
        print(f"  chars: {len(text)}")

    print("step 4: check the output")
    has(text, "Revenue grew in every region.")
    has(markdown, "Quarterly Report")
    check(isinstance(json.loads(ir_json), dict), "the IR JSON is not an object")

    print("OK 01_extract")


if __name__ == "__main__":
    main()
