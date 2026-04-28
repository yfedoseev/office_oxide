//! 06_from_markdown — Markdown → DOCX / XLSX / PPTX conversion.
//!
//! Converts a Markdown string (with heading, bullets, and table) to all three
//! OOXML formats via `create_from_markdown`, then reads each back and prints a
//! summary. Also demonstrates `DocumentIR::from_markdown` directly.
//!
//! Run: `cargo run --example 06_from_markdown`

use office_oxide::Document;
use office_oxide::create::create_from_markdown;
use office_oxide::format::DocumentFormat;
use office_oxide::ir::DocumentIR;

const MARKDOWN: &str = "# Quarterly Report

This document was generated automatically from Markdown.

## Highlights

- Revenue grew by **32%** year-over-year
- New products launched: *Widget Pro* and *Widget Lite*
- Customer satisfaction score: 4.8 / 5.0

## Financial Summary

| Category   | Q3 2025  | Q4 2025  |
|------------|----------|----------|
| Revenue    | $1.2M    | $1.6M    |
| Expenses   | $0.8M    | $0.9M    |
| Net Profit | $0.4M    | $0.7M    |

## Conclusion

Office Oxide makes it easy to turn structured data into Office documents.
";

fn main() {
    // ── DocumentIR::from_markdown demo ──────────────────────────────────────
    let ir = DocumentIR::from_markdown(MARKDOWN, DocumentFormat::Docx);
    println!(
        "IR from markdown: {} sections, {} elements in first section",
        ir.sections.len(),
        ir.sections.first().map(|s| s.elements.len()).unwrap_or(0)
    );

    // ── Create all three formats ─────────────────────────────────────────────
    let formats = [
        (DocumentFormat::Docx, "oo_example_06.docx"),
        (DocumentFormat::Xlsx, "oo_example_06.xlsx"),
        (DocumentFormat::Pptx, "oo_example_06.pptx"),
    ];

    for (fmt, filename) in &formats {
        let path = std::env::temp_dir().join(filename);
        create_from_markdown(MARKDOWN, *fmt, &path)
            .unwrap_or_else(|e| panic!("create {filename}: {e}"));

        let doc = Document::open(&path).unwrap_or_else(|e| panic!("open {filename}: {e}"));

        let text = doc.plain_text();
        let md_out = doc.to_markdown();

        assert!(!text.is_empty(), "{filename}: plain text is empty");

        println!("\n=== {filename} ({fmt:?}) ===");
        println!("format: {:?}", doc.format());
        println!("plain text length: {} chars", text.len());
        println!("markdown length:   {} chars", md_out.len());
        println!("first 120 chars of plain text:");
        let preview: String = text.chars().take(120).collect();
        println!("  {preview}");
    }

    println!("\nAll formats created and verified.");
}
