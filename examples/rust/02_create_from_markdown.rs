//! 02_create_from_markdown: turn one Markdown source into a DOCX, an XLSX and a PPTX.
//!
//! Run: `cargo run --example 02_create_from_markdown`

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
    println!("step 1: write the Markdown source");
    let markdown = MARKDOWN;

    println!("step 2: create a DOCX, an XLSX and a PPTX from it");
    let formats = [
        DocumentFormat::Docx,
        DocumentFormat::Xlsx,
        DocumentFormat::Pptx,
    ];
    let created: Vec<Vec<u8>> = formats
        .iter()
        .map(|&f| markdown_to_bytes(markdown, f))
        .collect();

    println!("step 3: open each and check the text");
    for (format, data) in formats.into_iter().zip(created) {
        let text = open_bytes(data, format).plain_text();
        println!("  format: {}", format.extension());
        println!("  chars: {}", text.len());
        has(&text, "North");
    }

    println!("OK 02_create_from_markdown");
}
