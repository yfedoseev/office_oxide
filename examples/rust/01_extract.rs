//! 01_extract: extract plain text, Markdown and the IR from a document.
//!
//! Run: `cargo run --example 01_extract`

use std::io::Cursor;

use office_oxide::create::create_from_markdown_to_writer;
use office_oxide::{Document, DocumentFormat};

const MARKDOWN: &str = "# Quarterly Report\n\nRevenue grew in every region.\n\n\
| Region | Q1  | Q2  |\n|--------|-----|-----|\n| North  | 120 | 135 |\n| South  | 95  | 110 |\n";

fn build_docx_bytes(markdown: &str) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    create_from_markdown_to_writer(markdown, DocumentFormat::Docx, &mut buf).expect("create DOCX");
    buf.into_inner()
}

/// Assert with a message; a failure exits non-zero.
fn check(condition: bool, message: &str) {
    if !condition {
        eprintln!("FAIL: {message}");
        std::process::exit(1);
    }
}

fn main() {
    println!("step 1: build a DOCX in memory");
    let bytes = build_docx_bytes(MARKDOWN);

    println!("step 2: open it from bytes");
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx).expect("open DOCX");
    println!("  format: {}", doc.format().extension());

    println!("step 3: extract text, Markdown and the IR");
    let text = doc.plain_text();
    let markdown = doc.to_markdown();
    let ir = doc.to_ir();
    println!("  chars: {}", text.len());

    println!("step 4: check the text");
    check(text.contains("Quarterly Report"), "docx text must contain the title");
    check(
        text.contains("Revenue grew in every region."),
        "docx text must contain the paragraph",
    );
    check(markdown.contains("North"), "markdown must contain the table");
    check(!ir.sections.is_empty(), "the IR must have at least one section");
    println!("OK 01_extract");
}
