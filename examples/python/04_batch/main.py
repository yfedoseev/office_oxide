"""04_batch — Batch processing of multiple Office formats.

Creates DOCX, XLSX, and PPTX from Markdown, then opens each in a loop
and prints extracted text. Demonstrates the batch-processing pattern.
Exit 0 on success.
"""
import sys
import tempfile
import os

from office_oxide import Document, create_from_markdown


REPORTS = {
    "docx": """\
# Sales Report — Q4 2025

## Executive Summary

Strong performance across all product lines.

- Total Revenue: $4.8M (up 28%)
- New Customers: 340
- Churn Rate: 2.1%
""",
    "xlsx": """\
# Sales Data

| Product   | Revenue | Units |
|-----------|---------|-------|
| Widget A  | $1.2M   | 1200  |
| Widget B  | $2.1M   | 840   |
| Widget C  | $1.5M   | 3000  |
""",
    "pptx": """\
# Q4 2025 Investor Deck

## Key Metrics

- Revenue: $4.8M
- Growth: 28% YoY
- Customers: 1,240 active

## Next Steps

Expand into European markets in Q1 2026.
""",
}


def process_format(fmt: str, markdown: str) -> dict:
    with tempfile.NamedTemporaryFile(suffix=f".{fmt}", delete=False) as f:
        tmp = f.name
    try:
        create_from_markdown(markdown, fmt, tmp)
        with Document.open(tmp) as doc:
            return {
                "format": doc.format,
                "text_len": len(doc.plain_text()),
                "md_len": len(doc.to_markdown()),
                "text": doc.plain_text(),
            }
    finally:
        os.unlink(tmp)


def main() -> None:
    results = {}
    for fmt, md in REPORTS.items():
        print(f"Processing {fmt}...", end=" ", flush=True)
        info = process_format(fmt, md)
        results[fmt] = info
        print(f"OK — format={info['format']}, text={info['text_len']} chars, md={info['md_len']} chars")

    print("\n=== Batch Summary ===")
    for fmt, info in results.items():
        print(f"  {fmt}: {info['text_len']} text chars, {info['md_len']} md chars")
        assert info["text_len"] > 0, f"{fmt}: text is empty"

    print("\nBatch processing complete.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        sys.exit(1)
