"""01_extract — Self-contained extract demo using the Python API.

Creates a DOCX from Markdown (written to a temp file), opens it with
Document.open, and prints the format, a text snippet, and Markdown output.
Exit 0 on success.
"""
import sys
import tempfile
import os

from office_oxide import Document, create_from_markdown


MARKDOWN = """\
# Office Oxide Extract Demo

This document was created from Markdown and parsed back via Python bindings.

## Features

- Plain text extraction
- Markdown conversion
- IR (intermediate representation) access
"""


def main() -> None:
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        tmp = f.name
    try:
        create_from_markdown(MARKDOWN, "docx", tmp)

        with Document.open(tmp) as doc:
            fmt = doc.format
            text = doc.plain_text()
            md = doc.to_markdown()
            ir_json = doc.to_ir_json()

        print(f"format: {fmt}")
        print("--- plain text (first 200 chars) ---")
        print(text[:200])
        print("--- markdown (first 200 chars) ---")
        print(md[:200])
        print(f"--- IR JSON length: {len(ir_json)} chars ---")

        assert "Office Oxide Extract Demo" in text, "heading missing from plain text"
        assert "Plain text extraction" in text, "bullet missing from plain text"

        print("\nAll checks passed.")
    finally:
        os.unlink(tmp)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        sys.exit(1)
