//! 02_create_from_markdown: create DOCX, XLSX and PPTX from one Markdown text.
//!
//! Run: `cargo run --example 06_from_markdown`

use std::io::Cursor;

use office_oxide::create::create_from_markdown_to_writer;
use office_oxide::{Document, DocumentFormat};

const MARKDOWN: &str = "# Quarterly Report\n\nRevenue grew in every region.\n\n\
| Region | Q1  | Q2  |\n|--------|-----|-----|\n| North  | 120 | 135 |\n| South  | 95  | 110 |\n";

/// Assert with a message; a failure exits non-zero.
fn check(condition: bool, message: &str) {
    if !condition {
        eprintln!("FAIL: {message}");
        std::process::exit(1);
    }
}

fn main() {
    for (n, format) in [
        DocumentFormat::Docx,
        DocumentFormat::Xlsx,
        DocumentFormat::Pptx,
    ]
    .into_iter()
    .enumerate()
    {
        println!(
            "step {}: create a {} from Markdown and open it",
            n + 1,
            format.extension().to_uppercase()
        );
        let mut buf = Cursor::new(Vec::new());
        create_from_markdown_to_writer(MARKDOWN, format, &mut buf).expect("create from Markdown");
        let doc = Document::from_reader(Cursor::new(buf.into_inner()), format)
            .expect("open created file");
        let text = doc.plain_text();
        println!("  format: {}", doc.format().extension());
        println!("  chars: {}", text.len());
        check(text.contains("Quarterly Report"), "created file text must contain the title");
        check(text.contains("North"), "created file text must contain the table");
    }
    println!("OK 02_create_from_markdown");
}
