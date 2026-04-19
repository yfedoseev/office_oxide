//! Create a tiny DOCX fixture for binding smoke tests.
//!
//! Run: `cargo run --release --example make_smoke -- /tmp/ffi_smoke.docx`

fn main() {
    let path = std::env::args().nth(1).unwrap_or_else(|| "/tmp/ffi_smoke.docx".into());
    let mut w = office_oxide::docx::write::DocxWriter::new();
    w.add_heading("Hello FFI", 1);
    w.add_paragraph("This is a smoke-test DOCX created for FFI verification.");
    if let Err(e) = w.save(&path) {
        eprintln!("error: {e}");
        std::process::exit(1);
    }
    println!("wrote {path}");
}
