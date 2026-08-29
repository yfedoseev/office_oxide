//! inspect_doc — dump plain text + IR JSON for a given Office file.
//!
//! Run: `cargo run --example inspect_doc -- <path>`

use office_oxide::Document;

fn main() {
    let path = std::env::args().nth(1).expect("usage: inspect_doc <path>");
    let doc = Document::open(&path).expect("open document");

    println!("format: {:?}", doc.format());
    println!("--- plain text ---");
    println!("{}", doc.plain_text());
    println!("--- IR JSON (pretty) ---");
    let json = serde_json::to_string_pretty(&doc.to_ir()).expect("serialize IR");
    println!("{json}");
}
