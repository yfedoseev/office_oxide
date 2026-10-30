//! 08_limits_strict_damage: strict versus lenient opening of a truncated file.
//!
//! `ParseOptions::strict()` bounds resources and turns "known to be incomplete"
//! into an error; the default options report damage instead. Neither may hang
//! or crash: strict must fail, lenient may open (and say what is missing) or fail.
//!
//! Run: `cargo run --example 08_limits_strict_damage`

use std::io::Cursor;

use office_oxide::create::create_from_markdown_to_writer;
use office_oxide::options::ParseOptions;
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
    println!("step 1: build a DOCX in memory and keep the first half of its bytes");
    let bytes = build_docx_bytes(MARKDOWN);
    let damaged = bytes[..bytes.len() / 2].to_vec();

    println!("step 2: open the damaged bytes in strict mode");
    let strict = Document::from_reader_with(
        Cursor::new(damaged.clone()),
        DocumentFormat::Docx,
        &ParseOptions::strict(),
    );
    match &strict {
        Err(e) => println!("  error: {e}"),
        Ok(_) => println!("  error: none"),
    }
    check(strict.is_err(), "strict mode must reject a truncated archive");

    println!("step 3: open the same bytes leniently");
    match Document::from_reader_with(
        Cursor::new(damaged),
        DocumentFormat::Docx,
        &ParseOptions::default(),
    ) {
        Ok(doc) => {
            let damage = doc.damage();
            println!("  chars: {}", doc.plain_text().len());
            println!("  damage: {}", damage.len());
        },
        Err(e) => println!("  error: {e}"),
    }
    println!("OK 08_limits_strict_damage");
}
