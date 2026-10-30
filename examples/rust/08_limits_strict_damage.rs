//! 08_limits_strict_damage: strict mode refuses a damaged archive; a lenient open reports the damage.
//!
//! Run: `cargo run --example 08_limits_strict_damage`

use office_oxide::create::create_from_markdown_to_writer;
use office_oxide::options::ParseOptions;
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

const PART: &str = "word/_rels/document.xml.rels";

/// Overwrites the start of one part's compressed data. A first deflate byte of
/// 0xFF declares a reserved block type, so every inflater rejects the part, and
/// the central directory stays intact. The part is found by name, not by
/// offset, so the recipe does not depend on how big the other parts are.
fn damage_part(data: &[u8]) -> Vec<u8> {
    let name = PART.as_bytes();
    // The first hit is the part's local header.
    let at = data
        .windows(name.len())
        .position(|w| w == name)
        .unwrap_or(0);
    check(at > 0, "the part to damage is missing from the archive");
    // The extra-field length sits in the two bytes just before the name.
    let start = at + name.len() + usize::from(u16::from_le_bytes([data[at - 2], data[at - 1]]));
    let mut damaged = data.to_vec();
    damaged[start..start + 16].fill(0xFF);
    damaged
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

/// Opens `bytes` as a DOCX under `options`: the document, or the error text.
fn open_with(bytes: &[u8], options: &ParseOptions) -> Result<Document, String> {
    Document::from_reader_with(Cursor::new(bytes.to_vec()), DocumentFormat::Docx, options)
        .map_err(|e| e.to_string())
}

fn main() {
    println!("step 1: build a DOCX in memory");
    let good = build_docx_bytes("Quarterly Report");

    println!("step 2: damage one part: overwrite 16 bytes of {PART}");
    let damaged = damage_part(&good);

    println!("step 3: open it in strict mode");
    let strict = ParseOptions::strict();
    match open_with(&damaged, &strict) {
        Ok(_) => check(false, "strict mode accepted a damaged part"),
        Err(error) => println!("  error: {error}"),
    }

    println!("step 4: open it leniently");
    let doc = open_with(&damaged, &ParseOptions::default())
        .expect("a lenient open reads past a damaged part");
    let kinds: Vec<String> = doc
        .damage()
        .iter()
        .map(|d| {
            serde_json::to_string(&d.kind)
                .expect("the kind serialises")
                .trim_matches('"')
                .to_string()
        })
        .collect();
    println!("  damage: {}", kinds.join(", "));
    check(
        kinds.iter().any(|k| k == "unreadable_part"),
        "the damage report does not name an unreadable part",
    );
    has(&doc.plain_text(), "Quarterly Report");

    println!("step 5: truncate the archive to its first half and open it both ways");
    let half = &good[..good.len() / 2];
    for (options, message) in [
        (&strict, "strict mode accepted a truncated archive"),
        (&ParseOptions::default(), "a lenient open accepted a truncated archive"),
    ] {
        match open_with(half, options) {
            Ok(_) => check(false, message),
            Err(error) => println!("  error: {error}"),
        }
    }

    println!("step 6: a limit applies to one document only");
    let mut limited = ParseOptions::default();
    limited
        .set_option("part_bytes", 100)
        .expect("part_bytes is a known option");
    match open_with(&good, &limited) {
        Ok(_) => check(false, "the part_bytes limit was not applied"),
        Err(error) => println!("  error: {error}"),
    }
    let doc = open_with(&good, &ParseOptions::default()).expect("the default limits accept it");
    has(&doc.plain_text(), "Quarterly Report");

    println!("OK 08_limits_strict_damage");
}
