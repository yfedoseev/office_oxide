//! 07_all_outputs: render one document as plain text, Markdown, HTML and IR JSON.
//!
//! Run: `cargo run --example 07_all_outputs`

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

    println!("step 3: render every output");
    let outputs = [
        ("plain", doc.plain_text()),     // words only: no markup, no structure
        ("markdown", doc.to_markdown()), // headings and pipe tables survive as syntax
        ("html", doc.to_html()),         // a fragment (no <html> wrapper), tables as <table>
        // the format-neutral tree as JSON, for programs
        ("ir_json", serde_json::to_string(&doc.to_ir()).expect("the IR serialises")),
    ];
    for (name, value) in &outputs {
        println!("  chars: {name} {}", value.len());
    }

    println!("step 4: check each output");
    for (name, value) in &outputs {
        check(!value.trim().is_empty(), &format!("the {name} output is empty"));
    }
    has(&outputs[0].1, "Quarterly Report");
    check(
        outputs[1].1.contains("# Quarterly Report"),
        "the Markdown is missing the heading",
    );
    check(outputs[2].1.contains("<table"), "the HTML has no table");
    let ir: serde_json::Value = serde_json::from_str(&outputs[3].1).expect("the IR JSON parses");
    check(ir.is_object(), "the IR JSON does not parse to an object");

    println!("OK 07_all_outputs");
}
