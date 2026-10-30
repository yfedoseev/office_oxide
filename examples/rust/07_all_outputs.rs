//! 07_all_outputs: one document, every output (text, Markdown, HTML, IR JSON).
//!
//! Plain text keeps the words only; Markdown keeps headings and tables as light
//! markup; HTML keeps the same structure as tags; IR JSON is the format-neutral
//! tree (sections, elements, runs) for programs that need structure.
//!
//! Run: `cargo run --example 07_all_outputs`

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

    println!("step 3: produce every output");
    let text = doc.plain_text();
    let markdown = doc.to_markdown();
    let html = doc.to_html();
    let ir_json = serde_json::to_string(&doc.to_ir()).expect("IR serialises");
    for (name, out) in [
        ("text", &text),
        ("markdown", &markdown),
        ("html", &html),
        ("ir", &ir_json),
    ] {
        println!("  {name}: {}", out.len());
        check(!out.trim().is_empty(), "every output must be non-empty");
        check(out.contains("Quarterly Report"), "every output must contain the title");
    }

    println!("step 4: check each output");
    check(markdown.contains('#'), "markdown must keep the heading marker");
    check(html.contains('<') && html.contains('>'), "html must contain tags");
    let ir: serde_json::Value = serde_json::from_str(&ir_json).expect("the IR must be valid JSON");
    check(
        ir["sections"].as_array().is_some_and(|s| !s.is_empty()),
        "the IR must have sections",
    );
    println!("OK 07_all_outputs");
}
