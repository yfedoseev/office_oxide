"""02_create_from_markdown: turn one Markdown source into a DOCX, an XLSX and a PPTX.

Run: python examples/python/02_create_from_markdown/main.py
"""
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


def check(condition: bool, message: str) -> None:
    if not condition:
        print(f"FAIL: {message}", file=sys.stderr)
        sys.exit(1)


def has(text: str, expected: str) -> None:
    check(expected in text, f'"{expected}" is missing from the text')


def main() -> None:
    print("step 1: write the Markdown source")
    markdown = MARKDOWN

    print("step 2: create a DOCX, an XLSX and a PPTX from it")
    created = {fmt: markdown_to_bytes(markdown, fmt) for fmt in ("docx", "xlsx", "pptx")}

    print("step 3: open each and check the text")
    for fmt, data in created.items():
        with Document.from_bytes(data, fmt) as doc:
            text = doc.plain_text()
        print(f"  format: {fmt}")
        print(f"  chars: {len(text)}")
        has(text, "North")

    print("OK 02_create_from_markdown")


if __name__ == "__main__":
    main()
