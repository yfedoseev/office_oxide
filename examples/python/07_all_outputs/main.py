"""07_all_outputs: one document, every output: plain text, Markdown, HTML and IR JSON.

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

def check(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def main() -> None:
    print("step 1: build a DOCX in memory")
    data = build_docx_bytes()

    print("step 2: open it from bytes")
    with Document.from_bytes(data, "docx") as doc:
        print(f"  format: {doc.format}")
        outputs = {
            "plain": doc.plain_text(),        # text only: no markup, no structure
            "markdown": doc.to_markdown(),    # headings, emphasis, pipe tables
            "html": doc.to_html(),            # a fragment (no <html>), tables as <table>
            "ir_json": doc.to_ir_json(),      # the full tree as JSON, for programs
        }

    print("step 3: check every output")
    for name, value in outputs.items():
        print(f"  chars: {name} {len(value)}")
        check(value.strip() != "", f"the {name} output is empty")
    check("Quarterly Report" in outputs["plain"], "the plain text is missing the title")
    check("<table" in outputs["html"], "the HTML has no table")
    check(isinstance(json.loads(outputs["ir_json"]), dict), "the IR JSON does not parse to an object")

    print("OK 07_all_outputs")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"error: {exc!r}", file=sys.stderr)
        sys.exit(1)
