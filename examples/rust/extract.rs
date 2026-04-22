//! Extract plain text and Markdown using the native Rust API.
//!
//! Run: `cargo run --release --example extract -- report.docx`

fn main() {
    let path = match std::env::args().nth(1) {
        Some(p) => p,
        None => {
            eprintln!("usage: extract <file>");
            std::process::exit(1);
        },
    };
    match office_oxide::Document::open(&path) {
        Ok(doc) => {
            println!("format: {:?}", doc.format());
            println!("--- plain text ---\n{}", doc.plain_text());
            println!("--- markdown ---\n{}", doc.to_markdown());
        },
        Err(e) => {
            eprintln!("error: {e}");
            std::process::exit(1);
        },
    }
}
