"""02_create_from_markdown — Convert Markdown to DOCX, XLSX, and PPTX.

For each format: calls create_from_markdown, opens with Document.open,
prints a text snippet. Exit 0 on success.
"""
import sys
import tempfile
import os

from office_oxide import Document, create_from_markdown


MARKDOWN = """\
# Quarterly Report

Generated automatically from Markdown using Office Oxide.

## Highlights

- Revenue grew by **32%** year-over-year
- Customer satisfaction: 4.8 / 5.0
- New products launched: Widget Pro, Widget Lite

## Financial Summary

| Category   | Q3 2025 | Q4 2025 |
|------------|---------|---------|
| Revenue    | $1.2M   | $1.6M   |
| Expenses   | $0.8M   | $0.9M   |
| Net Profit | $0.4M   | $0.7M   |
"""

FORMATS = ["docx", "xlsx", "pptx"]


def main() -> None:
    tmps = {}
    try:
        for fmt in FORMATS:
            with tempfile.NamedTemporaryFile(suffix=f".{fmt}", delete=False) as f:
                tmps[fmt] = f.name

        for fmt in FORMATS:
            path = tmps[fmt]
            create_from_markdown(MARKDOWN, fmt, path)

            with Document.open(path) as doc:
                text = doc.plain_text()
                md_out = doc.to_markdown()

            assert text, f"{fmt}: plain text is empty"

            print(f"\n=== {fmt.upper()} ===")
            print(f"plain text length: {len(text)} chars")
            print(f"markdown length:   {len(md_out)} chars")
            print(f"first 100 chars: {text[:100]!r}")

        print("\nAll formats created and verified.")
    finally:
        for p in tmps.values():
            if os.path.exists(p):
                os.unlink(p)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        sys.exit(1)
