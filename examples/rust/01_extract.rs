//! 01_extract — Self-contained extract demo.
//!
//! Creates a DOCX in memory, reads it back, and prints plain text,
//! Markdown, and an IR summary (number of sections).
//!
//! Run: `cargo run --example 01_extract`

use std::io::Cursor;

use office_oxide::docx::write::DocxWriter;
use office_oxide::{Document, DocumentFormat};

fn main() {
    // ── Build a DOCX in memory ──────────────────────────────────────────────
    let mut writer = DocxWriter::new();
    writer.add_heading("Office Oxide Extract Demo", 1);
    writer.add_paragraph("This document was created in memory and read back via Document::from_reader.");
    writer.add_paragraph("It demonstrates plain-text extraction, Markdown conversion, and IR access.");

    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).expect("write DOCX to buffer");
    buf.set_position(0);

    // ── Read it back ────────────────────────────────────────────────────────
    let doc = Document::from_reader(buf, DocumentFormat::Docx).expect("parse DOCX");

    println!("format: {:?}", doc.format());

    let text = doc.plain_text();
    println!("--- plain text ---");
    println!("{text}");

    let md = doc.to_markdown();
    println!("--- markdown ---");
    println!("{md}");

    let ir = doc.to_ir();
    println!("--- IR ---");
    println!("sections: {}", ir.sections.len());
    for (i, s) in ir.sections.iter().enumerate() {
        println!("  section {i}: title={:?}, elements={}", s.title, s.elements.len());
    }

    // ── Verify ─────────────────────────────────────────────────────────────
    assert!(text.contains("Office Oxide Extract Demo"), "heading missing from plain text");
    assert!(text.contains("created in memory"), "paragraph missing from plain text");
    assert!(!ir.sections.is_empty(), "IR must have at least one section");

    println!("\nAll checks passed.");
}
