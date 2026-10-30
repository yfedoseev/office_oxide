//! 01_extract: read a document as plain text, Markdown and IR JSON.
//!
//! Run: `cargo run --example 01_extract`

use office_oxide::create::create_from_markdown_to_writer;
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;

const MARKDOWN: &str = "# Quarterly Report\n\nRevenue grew in every region.\n\n\
| Region | Q1  | Q2  |\n|--------|-----|-----|\n| North  | 120 | 135 |\n| South  | 95  | 110 |\n";

fn markdown_to_bytes(markdown: &str, format: DocumentFormat) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    create_from_markdown_to_writer(markdown, format, &mut buf).expect("create from Markdown");
    buf.into_inner()
}

fn build_docx_bytes(title: &str) -> Vec<u8> {
    markdown_to_bytes(&MARKDOWN.replace("Quarterly Report", title), DocumentFormat::Docx)
}

/// `Document::from_reader` needs a `Read + Seek + Send + 'static` source, so the
/// bytes go in an owned `Cursor`, not a borrowed slice.
fn open_bytes(bytes: Vec<u8>, format: DocumentFormat) -> Document {
    Document::from_reader(Cursor::new(bytes), format).expect("open from bytes")
}

fn check(condition: bool, message: &str) {
    if !condition {
        eprintln!("FAIL: {message}");
        std::process::exit(1);
    }
}

fn has(text: &str, expected: &str) {
    check(text.contains(expected), &format!("\"{expected}\" is missing from the text"));
}

fn main() {
    println!("step 1: build a DOCX in memory");
    let data = build_docx_bytes("Quarterly Report");

    println!("step 2: open it from bytes");
    let doc = open_bytes(data, DocumentFormat::Docx);
    println!("  format: {}", doc.format().extension());
    check(
        doc.format() == DocumentFormat::Docx,
        "the opened format differs from the requested one",
    );

    println!("step 3: extract plain text, Markdown and IR JSON");
    let text = doc.plain_text();
    let markdown = doc.to_markdown();
    let ir_json = serde_json::to_string(&doc.to_ir()).expect("the IR serialises");
    println!("  chars: {}", text.len());

    println!("step 4: check the output");
    has(&text, "Revenue grew in every region.");
    has(&markdown, "Quarterly Report");
    let ir: serde_json::Value = serde_json::from_str(&ir_json).expect("the IR JSON parses");
    check(ir.is_object(), "the IR JSON is not an object");

    println!("OK 01_extract");
}
